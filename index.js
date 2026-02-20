require('dotenv').config();
const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');
const MailComposer = require('nodemailer/lib/mail-composer');

const app = express();
app.use(cors());
app.use(express.json());

function getEnvValue(name) {
    const raw = process.env[name];
    if (!raw) return '';

    let value = String(raw).trim();

    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
        value = value.slice(1, -1).trim();
    }

    const prefixedKey = `${name}=`;
    if (value.startsWith(prefixedKey)) {
        value = value.slice(prefixedKey.length).trim();
    }

    return value;
}

function findSecretFile(fileName) {
    const candidates = [
        path.join('/etc/secrets', fileName),
        path.join(process.cwd(), fileName),
        path.join(__dirname, fileName),
        path.join(process.cwd(), 'server', fileName),
    ];

    return candidates.find((filePath) => fs.existsSync(filePath)) || null;
}

function readAuthorizedUserToken(fileName) {
    const tokenPath = findSecretFile(fileName);
    if (!tokenPath) return { tokenData: null, tokenPath: null };

    try {
        const raw = fs.readFileSync(tokenPath, 'utf8');
        const parsed = JSON.parse(raw);
        return { tokenData: parsed, tokenPath };
    } catch (error) {
        console.warn(`[Auth] Gagal parse token file ${tokenPath}: ${error.message}`);
        return { tokenData: null, tokenPath };
    }
}

function buildOAuthClient(config) {
    const {
        label,
        tokenFileName,
        envClientIdKey,
        envClientSecretKey,
        envRefreshTokenKey,
    } = config;

    const { tokenData, tokenPath } = readAuthorizedUserToken(tokenFileName);

    const clientId = tokenData?.client_id || getEnvValue(envClientIdKey);
    const clientSecret = tokenData?.client_secret || getEnvValue(envClientSecretKey);
    const refreshToken = tokenData?.refresh_token || getEnvValue(envRefreshTokenKey);

    if (!clientId || !clientSecret || !refreshToken) {
        throw new Error(`[Auth] Kredensial ${label} tidak lengkap. Cek ${tokenFileName} atau env ${envClientIdKey}/${envClientSecretKey}/${envRefreshTokenKey}.`);
    }

    const oauthClient = new google.auth.OAuth2(
        clientId,
        clientSecret,
        'https://developers.google.com/oauthplayground'
    );
    oauthClient.setCredentials({ refresh_token: refreshToken });

    return {
        client: oauthClient,
        meta: {
            source: tokenData ? `file:${tokenPath}` : `env:${envRefreshTokenKey}`,
            tokenFileFound: Boolean(tokenPath),
            hasRefreshToken: Boolean(refreshToken),
        }
    };
}

// ==============================================================================
// 1. KREDENSIAL "DOC" -> KHUSUS MEMBACA SHEETS DAN DOWNLOAD DRIVE
// Karena Sparta simpan PDF-nya pakai kredensial DOC, downloadnya wajib pakai DOC
// ==============================================================================
const docAuth = buildOAuthClient({
    label: 'DOC',
    tokenFileName: 'token_doc.json',
    envClientIdKey: 'DOC_GOOGLE_CLIENT_ID',
    envClientSecretKey: 'DOC_GOOGLE_CLIENT_SECRET',
    envRefreshTokenKey: 'DOC_GOOGLE_REFRESH_TOKEN',
});
const docOAuth2Client = docAuth.client;

const sheets = google.sheets({ version: 'v4', auth: docOAuth2Client });
const drive = google.drive({ version: 'v3', auth: docOAuth2Client }); // <-- KEMBALI PAKAI DOC


// ==============================================================================
// 2. KREDENSIAL "UTAMA" -> KHUSUS UNTUK GMAIL API (KIRIM EMAIL)
// ==============================================================================
const spartaAuth = buildOAuthClient({
    label: 'SPARTA',
    tokenFileName: 'token.json',
    envClientIdKey: 'GOOGLE_CLIENT_ID',
    envClientSecretKey: 'GOOGLE_CLIENT_SECRET',
    envRefreshTokenKey: 'GOOGLE_REFRESH_TOKEN',
});
const spartaOAuth2Client = spartaAuth.client;

const gmail = google.gmail({ version: 'v1', auth: spartaOAuth2Client });
const spartaDrive = google.drive({ version: 'v3', auth: spartaOAuth2Client });

console.log(`[Auth] DOC source: ${docAuth.meta.source}`);
console.log(`[Auth] SPARTA source: ${spartaAuth.meta.source}`);

// --- Helper Functions ---
function extractFileId(url) {
    if (!url) return null;
    const text = String(url).trim();

    if (/^[\w-]{20,}$/.test(text)) {
        return text;
    }

    const match = text.match(/(?:id=|\/d\/|file\/d\/)([\w-]{20,})/);
    return match ? match[1] : null;
}

async function downloadDriveFileWithClient(driveClient, fileId, label) {
    try {
        await driveClient.files.get({
            fileId,
            fields: 'id,name,mimeType',
            supportsAllDrives: true,
        });

        const response = await driveClient.files.get(
            { fileId, alt: 'media', supportsAllDrives: true },
            { responseType: 'arraybuffer' }
        );
        console.log(`[Drive] Berhasil mengunduh ID: ${fileId} via ${label}`);
        return Buffer.from(response.data);
    } catch (error) {
        const status = error?.response?.status;
        console.warn(`[Drive] Gagal via ${label} untuk ID ${fileId}: ${status || ''} ${error.message}`.trim());
        return null;
    }
}

async function downloadDriveFile(fileId) {
    if (!fileId) return null;

    const fromDoc = await downloadDriveFileWithClient(drive, fileId, 'DOC');
    if (fromDoc) return fromDoc;

    const fromSparta = await downloadDriveFileWithClient(spartaDrive, fileId, 'SPARTA');
    if (fromSparta) return fromSparta;

    console.error(`[Drive] Gagal mengunduh file ID ${fileId} pada semua kredensial.`);
    return null;
}

function normalizeString(str) {
    if (!str) return "";
    return String(str).replace(/-/g, "").replace(/\s/g, "").trim().toUpperCase();
}

const SPARTA_BACKEND_BASE_URL = getEnvValue('SPARTA_BACKEND_BASE_URL') || 'https://sparta-backend-5hdj.onrender.com';

function escapeHtml(text) {
    return String(text ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function buildRabApprovalEmailHtml({
    level,
    proyek,
    nomorUlok,
    approvalUrl,
    rejectionUrl,
    additionalInfo,
}) {
    const infoBlock = additionalInfo
        ? `<p style="font-style: italic;">${escapeHtml(additionalInfo)}</p>`
        : '';

    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        .button { padding: 10px 20px; text-decoration: none; color: white !important; border-radius: 5px; }
        .approve { background-color: #28a745; }
        .reject { background-color: #dc3545; }
    </style>
</head>
<body>
    <p>Yth. Bapak/Ibu ${escapeHtml(level)},</p>
    <p>
        Dokumen RAB untuk proyek
        <strong>${escapeHtml(proyek)}</strong>
        dengan Nomor Ulok <strong>${escapeHtml(nomorUlok)}</strong>
        memerlukan tinjauan dan persetujuan Anda.
    </p>
    ${infoBlock}
    <p>Silakan periksa detailnya pada file PDF yang terlampir dan pilih tindakan di bawah ini:</p>
    <br>
    <a href="${approvalUrl}" class="button approve">SETUJUI</a>
    <a href="${rejectionUrl}" class="button reject">TOLAK</a>
    <br><br>
    <p>Terima kasih.</p>
    <p><em>--- Email ini dibuat secara otomatis.---</em></p>
</body>
</html>
`;
}

function buildRabFinalApprovedEmailHtml({
    namaToko,
    proyek,
    lingkup,
    pdfNonSboFilename,
    pdfRekapFilename,
    linkPdfNonSbo,
    linkPdfRekap,
}) {
    return `
<p>Pengajuan RAB Toko <b>${escapeHtml(namaToko)}</b> untuk proyek <b>${escapeHtml(proyek)} - ${escapeHtml(lingkup)}</b> telah disetujui sepenuhnya.</p>
<p>Tiga versi file PDF RAB telah dilampirkan:</p>
<ul>
<li><b>${escapeHtml(pdfNonSboFilename)}</b>: Hanya berisi item pekerjaan di luar SBO.</li>
<li><b>${escapeHtml(pdfRekapFilename)}</b>: Rekapitulasi Total Biaya.</li>
</ul>
<p>Link Google Drive:</p>
<ul>
<li><a href="${escapeHtml(linkPdfNonSbo || '')}">Link PDF Non-SBO</a></li>
<li><a href="${escapeHtml(linkPdfRekap || '')}">Link PDF Rekapitulasi</a></li>
</ul>
`;
}

function buildRabFinalApprovedKontraktorHtml(baseBody) {
    return `${baseBody}
<p>Silakan upload Rekapitulasi RAB Termaterai & SPH melalui link berikut:</p>
<p><a href="https://materai-rab-pi.vercel.app/login" target="_blank">UPLOAD REKAP RAB TERMATERAI & SPH</a></p>`;
}

function buildDocApprovalEmailHtml({
    docType,
    level,
    proyek,
    nomorUlok,
    approvalUrl,
    rejectionUrl,
    additionalInfo,
}) {
    const infoBlock = additionalInfo
        ? `<p style="font-style: italic;">${escapeHtml(additionalInfo)}</p>`
        : '';

    return `
<!DOCTYPE html>
<html>
<head>
    <style>
        .button { padding: 10px 20px; text-decoration: none; color: white !important; border-radius: 5px; }
        .approve { background-color: #28a745; }
        .reject { background-color: #dc3545; }
    </style>
</head>
<body>
    <p>Yth. Bapak/Ibu ${escapeHtml(level)},</p>
    <p>
        Dokumen ${escapeHtml(docType)} untuk proyek
        <strong>${escapeHtml(proyek)}</strong>
        dengan Nomor Ulok <strong>${escapeHtml(nomorUlok)}</strong>
        memerlukan tinjauan dan persetujuan Anda.
    </p>
    ${infoBlock}
    <p>Silakan periksa detailnya pada file PDF yang terlampir dan pilih tindakan di bawah ini:</p>
    <br>
    <a href="${approvalUrl}" class="button approve">SETUJUI</a>
    <a href="${rejectionUrl}" class="button reject">TOLAK</a>
    <br><br>
    <p>Terima kasih.</p>
    <p><em>--- Email ini dibuat secara otomatis.---</em></p>
</body>
</html>
`;
}

function getHeaderIndex(headers, candidateNames) {
    const normalized = headers.map(h => String(h || '').trim().toUpperCase());
    for (const name of candidateNames) {
        const idx = normalized.indexOf(String(name || '').trim().toUpperCase());
        if (idx >= 0) return idx;
    }
    return -1;
}

function getCellByHeaders(row, headers, candidateNames, fallback = '') {
    const idx = getHeaderIndex(headers, candidateNames);
    if (idx < 0) return fallback;
    return row[idx] ?? fallback;
}

async function sendMailViaGmail(mailOptions) {
    const mail = new MailComposer(mailOptions);
    const messageBuffer = await mail.compile().build();

    const encodedMessage = messageBuffer.toString('base64')
        .replace(/\+/g, '-')
        .replace(/\//g, '_')
        .replace(/=+$/, '');

    const result = await gmail.users.messages.send({
        userId: 'me',
        requestBody: { raw: encodedMessage }
    });

    console.log(`[Email] Sukses terkirim via Gmail API. ID: ${result.data.id}`);
    return result.data.id;
}

async function getClientScopeInfo(oauthClient) {
    try {
        const accessToken = await oauthClient.getAccessToken();
        const tokenValue = typeof accessToken === 'string' ? accessToken : accessToken?.token;
        if (!tokenValue) {
            return { ok: false, message: 'Tidak bisa mengambil access token.', scopes: [] };
        }

        const info = await oauthClient.getTokenInfo(tokenValue);
        const scopes = Array.isArray(info?.scopes)
            ? info.scopes
            : String(info?.scope || '')
                .split(' ')
                .map(s => s.trim())
                .filter(Boolean);

        return { ok: true, message: 'OK', scopes };
    } catch (error) {
        return {
            ok: false,
            message: error.message,
            scopes: []
        };
    }
}

app.get('/api/debug/oauth-clients', async (req, res) => {
    try {
        const [docScopeInfo, spartaScopeInfo] = await Promise.all([
            getClientScopeInfo(docOAuth2Client),
            getClientScopeInfo(spartaOAuth2Client),
        ]);

        return res.status(200).json({
            doc: {
                source: docAuth.meta.source,
                tokenFileFound: docAuth.meta.tokenFileFound,
                hasRefreshToken: docAuth.meta.hasRefreshToken,
                scopeInfo: docScopeInfo,
            },
            sparta: {
                source: spartaAuth.meta.source,
                tokenFileFound: spartaAuth.meta.tokenFileFound,
                hasRefreshToken: spartaAuth.meta.hasRefreshToken,
                scopeInfo: spartaScopeInfo,
            },
        });
    } catch (error) {
        return res.status(500).json({ error: 'Gagal membaca status OAuth client.', details: error.message });
    }
});

// === ENDPOINT API ===
app.post('/api/resend-email', async (req, res) => {
    const { ulok, lingkup } = req.body;

    if (!ulok || !lingkup) {
        return res.status(400).json({ error: 'Ulok dan Lingkup Pekerjaan harus diisi.' });
    }

    try {
        console.log(`[API] Memproses Ulok: ${ulok}, Lingkup: ${lingkup}...`);

        const sheetId = process.env.DOC_SHEET_ID;

        // --- STEP A: Cari Data di form2 ---
        const responseForm2 = await sheets.spreadsheets.values.get({
            spreadsheetId: sheetId,
            range: 'form2!A:AA',
        });

        const rowsForm2 = responseForm2.data.values;
        if (!rowsForm2 || rowsForm2.length === 0) return res.status(404).json({ error: 'Data form2 kosong.' });
        const headersForm2 = rowsForm2[0].map(h => String(h || '').trim().toUpperCase());

        const normalizedTargetUlok = normalizeString(ulok);
        const normalizedTargetLingkup = String(lingkup).trim().toLowerCase();

        const dataRowsForm2 = rowsForm2.slice(1);
        const targetRowRelativeIndex = dataRowsForm2.findIndex(row => {
            const rowUlok = normalizeString(row[9]);
            const rowLingkup = String(row[13] || "").trim().toLowerCase();
            return rowUlok === normalizedTargetUlok && rowLingkup === normalizedTargetLingkup;
        });

        if (targetRowRelativeIndex === -1) return res.status(404).json({ error: 'Data tidak ditemukan di form2.' });

        const targetRow = dataRowsForm2[targetRowRelativeIndex];
        const sheetRowNumber = targetRowRelativeIndex + 2;

        const [
            status, timestamp, linkPdf, linkPdfNonSbo,
            emailKoord_old, waktuKoord, emailManager_old, waktuManager,
            emailPembuat, rowUlok, proyek, alamat, cabang, rowLingkup
        ] = targetRow;
        const idxLinkPdfRekap = headersForm2.indexOf('LINK PDF REKAPITULASI');
        const linkPdfRekap = idxLinkPdfRekap >= 0 ? targetRow[idxLinkPdfRekap] : targetRow[25];
        const idxNamaToko = headersForm2.indexOf('NAMA_TOKO');
        const namaToko = idxNamaToko >= 0 ? targetRow[idxNamaToko] : proyek;

        // --- STEP B: Tentukan Role ---
        let role = '';
        let targetJabatan = '';
        let approvalLevel = '';
        let isFinalApproved = false;

        if (status === 'Menunggu Persetujuan Koordinator') {
            role = 'Koordinator';
            targetJabatan = 'BRANCH BUILDING COORDINATOR';
            approvalLevel = 'coordinator';
        } else if (status === 'Menunggu Persetujuan Manager') {
            role = 'Manager';
            targetJabatan = 'BRANCH BUILDING & MAINTENANCE MANAGER';
            approvalLevel = 'manager';
        } else if (status === 'Disetujui') {
            role = 'Final Approved';
            isFinalApproved = true;
        } else {
            return res.status(200).json({ message: `Email tidak dikirim. Status saat ini: "${status}"` });
        }

        if (!cabang) return res.status(400).json({ error: 'Kolom cabang kosong.' });

        // --- STEP C: Cari BANYAK Email di Sheet Cabang (Menggunakan .filter) ---
        const responseCabang = await sheets.spreadsheets.values.get({
            spreadsheetId: sheetId,
            range: 'Cabang!A:Z',
        });

        const rowsCabang = responseCabang.data.values;
        if (!rowsCabang) return res.status(404).json({ error: 'Sheet Cabang kosong.' });

        const headersCabang = rowsCabang[0].map(h => String(h).trim().toUpperCase());
        const idxCabang = headersCabang.indexOf('CABANG');
        const idxJabatan = headersCabang.indexOf('JABATAN');
        const idxEmail = headersCabang.indexOf('EMAIL_SAT');

        const targetCabangUpper = String(cabang).trim().toUpperCase();
        const targetJabatanUpper = targetJabatan.toUpperCase();
        let recipientEmailsArray = [];

        if (!isFinalApproved) {
            const matchRowsCabang = rowsCabang.slice(1).filter(row => {
                const valCabang = String(row[idxCabang] || "").trim().toUpperCase();
                const valJabatan = String(row[idxJabatan] || "").trim().toUpperCase();
                return valCabang === targetCabangUpper && valJabatan === targetJabatanUpper;
            });

            if (matchRowsCabang.length === 0) {
                return res.status(404).json({ error: `Jabatan ${targetJabatan} tidak ditemukan di cabang ${cabang}.` });
            }

            recipientEmailsArray = matchRowsCabang
                .map(row => String(row[idxEmail] || "").trim())
                .filter(email => email !== "");
        } else {
            const allowedJabatan = new Set([
                'BRANCH BUILDING COORDINATOR',
                'BRANCH BUILDING & MAINTENANCE MANAGER',
            ]);

            const cabangTeamEmails = rowsCabang.slice(1)
                .filter(row => {
                    const valCabang = String(row[idxCabang] || "").trim().toUpperCase();
                    const valJabatan = String(row[idxJabatan] || "").trim().toUpperCase();
                    return valCabang === targetCabangUpper && allowedJabatan.has(valJabatan);
                })
                .map(row => String(row[idxEmail] || "").trim())
                .filter(Boolean);

            recipientEmailsArray = [
                String(emailPembuat || '').trim(),
                String(emailKoord_old || '').trim(),
                String(emailManager_old || '').trim(),
                ...cabangTeamEmails,
            ].filter(Boolean);
        }

        recipientEmailsArray = [...new Set(recipientEmailsArray)];

        if (recipientEmailsArray.length === 0) {
            return res.status(404).json({ error: `Email tujuan ditemukan tapi datanya kosong di sheet Cabang.` });
        }

        // Format akhirnya jadi: "email1@gmail.com, email2@gmail.com"
        const recipientEmailsStr = recipientEmailsArray.join(', ');
        console.log(`[API] Email tujuan ditemukan (${recipientEmailsArray.length} orang): ${recipientEmailsStr} (${role})`);

        const approverForLink = recipientEmailsArray[0];
        const encodedApprover = encodeURIComponent(approverForLink || '');
        const approvalUrl = `${SPARTA_BACKEND_BASE_URL}/api/handle_rab_approval?action=approve&row=${sheetRowNumber}&level=${approvalLevel}&approver=${encodedApprover}`;
        const rejectionUrl = `${SPARTA_BACKEND_BASE_URL}/api/reject_form/rab?row=${sheetRowNumber}&level=${approvalLevel}&approver=${encodedApprover}`;

        const additionalInfo = approvalLevel === 'manager' && emailKoord_old
            ? `Telah disetujui oleh Koordinator: ${emailKoord_old}`
            : '';


        // --- STEP D: Download PDF (Sekarang pakai DOC Auth) ---
        const attachments = [];
        const pdfId = extractFileId(linkPdf);
        const pdfNonSboId = extractFileId(linkPdfNonSbo);
        const pdfRekapId = extractFileId(linkPdfRekap);

        if (pdfId) {
            const pdfBuffer = await downloadDriveFile(pdfId);
            if (pdfBuffer) attachments.push({ filename: 'RAB_SBO.pdf', content: pdfBuffer });
        }
        if (pdfNonSboId) {
            const pdfNonSboBuffer = await downloadDriveFile(pdfNonSboId);
            if (pdfNonSboBuffer) attachments.push({ filename: 'RAB_NON_SBO.pdf', content: pdfNonSboBuffer });
        }
        if (pdfRekapId) {
            const pdfRekapBuffer = await downloadDriveFile(pdfRekapId);
            if (pdfRekapBuffer) attachments.push({ filename: 'REKAP_RAB.pdf', content: pdfRekapBuffer });
        }

        // --- STEP E: Kirim Email via GMAIL API ---
        const fromAddress = `"Sparta System RE-EMAIL" <${getEnvValue('EMAIL_USER')}>`;

        async function sendMailViaGmail(mailOptions) {
            const mail = new MailComposer(mailOptions);
            const messageBuffer = await mail.compile().build();

            const encodedMessage = messageBuffer.toString('base64')
                .replace(/\+/g, '-')
                .replace(/\//g, '_')
                .replace(/=+$/, '');

            const result = await gmail.users.messages.send({
                userId: 'me',
                requestBody: { raw: encodedMessage }
            });

            console.log(`[Email] Sukses terkirim via Gmail API. ID: ${result.data.id}`);
            return result.data.id;
        }

        let sentMessageIds = [];

        if (isFinalApproved) {
            const subject = `[FINAL - DISETUJUI] Pengajuan RAB Proyek ${namaToko}: ${proyek} - ${rowLingkup}`;
            const baseBody = buildRabFinalApprovedEmailHtml({
                namaToko,
                proyek,
                lingkup: rowLingkup,
                pdfNonSboFilename: 'RAB_NON_SBO.pdf',
                pdfRekapFilename: 'REKAP_RAB.pdf',
                linkPdfNonSbo: linkPdfNonSbo,
                linkPdfRekap: linkPdfRekap,
            });

            const kontraktorEmail = String(emailPembuat || '').trim();
            const teamRecipients = recipientEmailsArray.filter(email => {
                if (!kontraktorEmail) return true;
                return email.toLowerCase() !== kontraktorEmail.toLowerCase();
            });

            if (kontraktorEmail) {
                const kontraktorMessageId = await sendMailViaGmail({
                    from: fromAddress,
                    to: kontraktorEmail,
                    subject,
                    html: buildRabFinalApprovedKontraktorHtml(baseBody),
                    attachments,
                });
                sentMessageIds.push(kontraktorMessageId);
            }

            if (teamRecipients.length > 0) {
                const teamMessageId = await sendMailViaGmail({
                    from: fromAddress,
                    to: teamRecipients.join(', '),
                    subject,
                    html: baseBody,
                    attachments,
                });
                sentMessageIds.push(teamMessageId);
            }

            if (sentMessageIds.length === 0) {
                return res.status(404).json({ error: 'Tidak ada penerima final yang valid untuk status Disetujui.' });
            }
        } else {
            const singleMessageId = await sendMailViaGmail({
                from: fromAddress,
                to: recipientEmailsStr,
                subject: approvalLevel === 'coordinator'
                    ? `[TAHAP 1: PERLU PERSETUJUAN] RAB Proyek ${proyek} - ${rowLingkup}`
                    : `[TAHAP 2: PERLU PERSETUJUAN] RAB Proyek ${proyek} - ${rowLingkup}`,
                html: buildRabApprovalEmailHtml({
                    level: role,
                    proyek,
                    nomorUlok: rowUlok,
                    approvalUrl,
                    rejectionUrl,
                    additionalInfo,
                }),
                attachments,
            });
            sentMessageIds.push(singleMessageId);
        }

        return res.status(200).json({
            message: 'Email berhasil dikirim.',
            recipient: recipientEmailsStr,
            role: role,
            messageId: sentMessageIds[0],
            messageIds: sentMessageIds
        });

    } catch (error) {
        console.error('Terjadi kesalahan:', error);
        return res.status(500).json({ error: 'Terjadi kesalahan internal server.', details: error.message });
    }
});

app.post('/api/resend-email-spk', async (req, res) => {
    const { ulok, lingkup } = req.body;

    if (!ulok || !lingkup) {
        return res.status(400).json({ error: 'Nomor Ulok dan Lingkup Pekerjaan harus diisi.' });
    }

    try {
        console.log(`[API-SPK] Memproses Ulok: ${ulok}, Lingkup: ${lingkup}...`);

        const sheetId = getEnvValue('DOC_SHEET_ID');
        const fromAddress = `"Sparta System RE-EMAIL" <${getEnvValue('EMAIL_USER')}>`;

        const spkResp = await sheets.spreadsheets.values.get({
            spreadsheetId: sheetId,
            range: 'SPK_Data!A:AZ',
        });

        const spkRows = spkResp.data.values;
        if (!spkRows || spkRows.length === 0) {
            return res.status(404).json({ error: 'Sheet SPK_Data kosong.' });
        }

        const spkHeaders = spkRows[0];
        const dataRows = spkRows.slice(1);
        const targetUlok = normalizeString(ulok);
        const targetLingkup = String(lingkup).trim().toLowerCase();

        const targetRowRelativeIndex = dataRows.findIndex(row => {
            const rowUlok = normalizeString(getCellByHeaders(row, spkHeaders, ['Nomor Ulok']));
            const rowLingkup = String(getCellByHeaders(row, spkHeaders, ['Lingkup Pekerjaan', 'Lingkup_Pekerjaan']) || '').trim().toLowerCase();
            return rowUlok === targetUlok && rowLingkup === targetLingkup;
        });

        if (targetRowRelativeIndex === -1) {
            return res.status(404).json({ error: 'Data SPK tidak ditemukan untuk Nomor Ulok + Lingkup tersebut.' });
        }

        const row = dataRows[targetRowRelativeIndex];
        const sheetRowNumber = targetRowRelativeIndex + 2;

        const status = String(getCellByHeaders(row, spkHeaders, ['Status'])).trim();
        const cabang = String(getCellByHeaders(row, spkHeaders, ['Cabang'])).trim();
        const nomorUlok = String(getCellByHeaders(row, spkHeaders, ['Nomor Ulok'])).trim();
        const namaToko = String(getCellByHeaders(row, spkHeaders, ['Nama_Toko', 'nama_toko'])).trim();
        const kodeToko = String(getCellByHeaders(row, spkHeaders, ['Kode Toko', 'kode_toko'])).trim();
        const jenisToko = String(getCellByHeaders(row, spkHeaders, ['Jenis_Toko', 'Proyek'])).trim();
        const lingkupPekerjaan = String(getCellByHeaders(row, spkHeaders, ['Lingkup Pekerjaan', 'Lingkup_Pekerjaan'])).trim();
        const linkPdf = String(getCellByHeaders(row, spkHeaders, ['Link PDF'])).trim();
        const initiatorEmail = String(getCellByHeaders(row, spkHeaders, ['Dibuat Oleh', 'Email_Pembuat', 'EMAIL_PEMBUAT'])).trim();
        const approverEmail = String(getCellByHeaders(row, spkHeaders, ['Disetujui Oleh'])).trim();
        const alasanPenolakan = String(getCellByHeaders(row, spkHeaders, ['Alasan Penolakan'])).trim();

        if (!cabang) {
            return res.status(400).json({ error: 'Kolom Cabang SPK kosong.' });
        }

        const cabangResp = await sheets.spreadsheets.values.get({
            spreadsheetId: sheetId,
            range: 'Cabang!A:Z',
        });
        const cabangRows = cabangResp.data.values;
        if (!cabangRows || cabangRows.length === 0) {
            return res.status(404).json({ error: 'Sheet Cabang kosong.' });
        }

        const cabangHeaders = cabangRows[0].map(h => String(h || '').trim().toUpperCase());
        const idxCabang = cabangHeaders.indexOf('CABANG');
        const idxJabatan = cabangHeaders.indexOf('JABATAN');
        const idxEmail = cabangHeaders.indexOf('EMAIL_SAT');
        const targetCabangUpper = cabang.toUpperCase();

        const getEmailsByJabatan = (jabatanName) => cabangRows.slice(1)
            .filter(cRow => {
                const valCabang = String(cRow[idxCabang] || '').trim().toUpperCase();
                const valJabatan = String(cRow[idxJabatan] || '').trim().toUpperCase();
                return valCabang === targetCabangUpper && valJabatan === jabatanName.toUpperCase();
            })
            .map(cRow => String(cRow[idxEmail] || '').trim())
            .filter(Boolean);

        const branchManagerEmails = getEmailsByJabatan('BRANCH MANAGER');
        const managerEmails = getEmailsByJabatan('BRANCH BUILDING & MAINTENANCE MANAGER');
        const coordinatorEmails = getEmailsByJabatan('BRANCH BUILDING COORDINATOR');

        const attachments = [];
        const spkPdfId = extractFileId(linkPdf);
        if (spkPdfId) {
            const spkPdfBuffer = await downloadDriveFile(spkPdfId);
            if (spkPdfBuffer) {
                attachments.push({
                    filename: status === 'SPK Disetujui'
                        ? `SPK_DISETUJUI_${jenisToko || 'PROYEK'}_${nomorUlok || 'ULOK'}.pdf`.replace(/\s+/g, '_')
                        : `SPK_${jenisToko || 'PROYEK'}_${nomorUlok || 'ULOK'}.pdf`.replace(/\s+/g, '_'),
                    content: spkPdfBuffer,
                });
            }
        }

        if (status === 'Menunggu Persetujuan Branch Manager') {
            const bmEmail = branchManagerEmails[0] || approverEmail;
            if (!bmEmail) {
                return res.status(404).json({ error: `Email Branch Manager untuk cabang ${cabang} tidak ditemukan.` });
            }

            const approvalUrl = `${SPARTA_BACKEND_BASE_URL}/api/handle_spk_approval?action=approve&row=${sheetRowNumber}&approver=${encodeURIComponent(bmEmail)}`;
            const rejectionUrl = `${SPARTA_BACKEND_BASE_URL}/api/reject_form/spk?row=${sheetRowNumber}&approver=${encodeURIComponent(bmEmail)}`;
            const subject = `[PERLU PERSETUJUAN BM] SPK Proyek ${namaToko} (${kodeToko}): ${jenisToko} - ${lingkupPekerjaan}`;

            const messageId = await sendMailViaGmail({
                from: fromAddress,
                to: bmEmail,
                subject,
                html: buildDocApprovalEmailHtml({
                    docType: 'SPK',
                    level: 'Branch Manager',
                    proyek: jenisToko || namaToko,
                    nomorUlok,
                    approvalUrl,
                    rejectionUrl,
                }),
                attachments,
            });

            return res.status(200).json({
                message: 'Email SPK berhasil dikirim.',
                recipient: bmEmail,
                role: 'Branch Manager',
                messageId,
            });
        }

        if (status === 'SPK Disetujui') {
            const form2Resp = await sheets.spreadsheets.values.get({
                spreadsheetId: sheetId,
                range: 'form2!A:AA',
            });

            const form2Rows = form2Resp.data.values || [];
            const form2Headers = form2Rows[0] || [];
            let pembuatRabEmail = '';

            if (form2Rows.length > 1) {
                const form2DataRows = form2Rows.slice(1);
                const idxMatch = form2DataRows.findIndex(fRow => {
                    const fUlok = normalizeString(getCellByHeaders(fRow, form2Headers, ['Nomor Ulok', 'Lokasi']));
                    const fLingkup = String(getCellByHeaders(fRow, form2Headers, ['Lingkup Pekerjaan', 'Lingkup_Pekerjaan']) || '').trim().toLowerCase();
                    return fUlok === normalizeString(nomorUlok) && fLingkup === String(lingkupPekerjaan).trim().toLowerCase();
                });

                if (idxMatch >= 0) {
                    const form2Row = form2DataRows[idxMatch];
                    pembuatRabEmail = String(getCellByHeaders(form2Row, form2Headers, ['Email_Pembuat', 'EMAIL_PEMBUAT'])).trim();
                }
            }

            const bmEmail = approverEmail || branchManagerEmails[0] || '';
            const bbmManagerEmail = managerEmails[0] || '';
            const kontraktorList = pembuatRabEmail ? [pembuatRabEmail] : [];

            const subject = `[DISETUJUI] SPK Proyek ${namaToko} (${kodeToko}): ${jenisToko} - ${lingkupPekerjaan}`;
            const messageIds = [];

            const otherRecipients = new Set();
            if (initiatorEmail) otherRecipients.add(initiatorEmail);
            if (pembuatRabEmail) otherRecipients.add(pembuatRabEmail);

            if (bmEmail) {
                const bodyBm = `<p>SPK yang Anda setujui untuk Toko <b>${escapeHtml(namaToko)}</b> pada proyek <b>${escapeHtml(jenisToko)} - ${escapeHtml(lingkupPekerjaan)}</b> (${escapeHtml(nomorUlok)}) telah disetujui sepenuhnya dan final.</p><p>File PDF final terlampir.</p>`;
                messageIds.push(await sendMailViaGmail({ from: fromAddress, to: bmEmail, subject, html: bodyBm, attachments }));
                otherRecipients.delete(bmEmail);
            }

            if (bbmManagerEmail) {
                const bodyBbm = `<p>SPK yang diajukan untuk Toko <b>${escapeHtml(namaToko)}</b> pada proyek <b>${escapeHtml(jenisToko)} - ${escapeHtml(lingkupPekerjaan)}</b> (${escapeHtml(nomorUlok)}) telah disetujui oleh Branch Manager.</p><p>Silakan melakukan input PIC pengawasan melalui link berikut: <a href='https://frontend-form-virid.vercel.app/login-input_pic.html' target='_blank' rel='noopener noreferrer'>Input PIC Pengawasan</a></p><p>File PDF final terlampir.</p>`;
                messageIds.push(await sendMailViaGmail({ from: fromAddress, to: bbmManagerEmail, subject, html: bodyBbm, attachments }));
                otherRecipients.delete(bbmManagerEmail);
            }

            if (coordinatorEmails.length > 0) {
                const bodyCoord = `<p>SPK untuk Toko <b>${escapeHtml(namaToko)}</b> pada proyek <b>${escapeHtml(jenisToko)} - ${escapeHtml(lingkupPekerjaan)}</b> (${escapeHtml(nomorUlok)}) telah disetujui oleh Branch Manager.</p><p>File PDF final terlampir.</p>`;
                messageIds.push(await sendMailViaGmail({ from: fromAddress, to: coordinatorEmails.join(', '), subject, html: bodyCoord, attachments }));
                coordinatorEmails.forEach(email => otherRecipients.delete(email));
            }

            if (kontraktorList.length > 0) {
                const bodyOpname = `<p>SPK untuk Toko <b>${escapeHtml(namaToko)}</b> pada proyek <b>${escapeHtml(jenisToko)} - ${escapeHtml(lingkupPekerjaan)}</b> (${escapeHtml(nomorUlok)}) telah disetujui.</p><p>Silakan melakukan Opname melalui link berikut: <a href='https://sparta-alfamart.vercel.app' target='_blank' rel='noopener noreferrer'>Pengisian Opname</a></p><p>File PDF final terlampir.</p>`;
                messageIds.push(await sendMailViaGmail({ from: fromAddress, to: kontraktorList.join(', '), subject, html: bodyOpname, attachments }));
                kontraktorList.forEach(email => otherRecipients.delete(email));
            }

            if (otherRecipients.size > 0) {
                const bodyDefault = `<p>SPK yang Anda ajukan untuk Toko <b>${escapeHtml(namaToko)}</b> pada proyek <b>${escapeHtml(jenisToko)} - ${escapeHtml(lingkupPekerjaan)}</b> (${escapeHtml(nomorUlok)}) telah disetujui oleh Branch Manager.</p><p>File PDF final terlampir.</p>`;
                messageIds.push(await sendMailViaGmail({ from: fromAddress, to: Array.from(otherRecipients).join(', '), subject, html: bodyDefault, attachments }));
            }

            if (messageIds.length === 0) {
                return res.status(404).json({ error: 'Tidak ada penerima email SPK final yang valid.' });
            }

            const recipientSummary = [
                bmEmail,
                bbmManagerEmail,
                ...coordinatorEmails,
                ...kontraktorList,
                ...Array.from(otherRecipients),
            ].filter(Boolean);

            return res.status(200).json({
                message: 'Email SPK final disetujui berhasil dikirim.',
                recipient: [...new Set(recipientSummary)].join(', '),
                role: 'SPK Disetujui',
                messageId: messageIds[0],
                messageIds,
            });
        }

        if (status === 'SPK Ditolak') {
            if (!initiatorEmail) {
                return res.status(404).json({ error: 'Email pembuat SPK (Dibuat Oleh) tidak ditemukan untuk status ditolak.' });
            }

            const subject = `[DITOLAK] SPK untuk Proyek ${namaToko} (${kodeToko}): ${jenisToko} - ${lingkupPekerjaan}`;
            const body = `<p>SPK yang Anda ajukan untuk Toko <b>${escapeHtml(namaToko)}</b> pada proyek <b>${escapeHtml(jenisToko)} - ${escapeHtml(lingkupPekerjaan)}</b> (${escapeHtml(nomorUlok)}) telah ditolak oleh Branch Manager.</p><p><b>Alasan Penolakan:</b></p><p><i>${escapeHtml(alasanPenolakan || 'Tidak ada alasan yang diberikan.')}</i></p><p>Silakan ajukan revisi SPK Anda melalui link berikut:</p><p><a href='https://sparta-alfamart.vercel.app' target='_blank' rel='noopener noreferrer'>Input Ulang SPK</a></p>`;

            const messageId = await sendMailViaGmail({
                from: fromAddress,
                to: initiatorEmail,
                subject,
                html: body,
            });

            return res.status(200).json({
                message: 'Email SPK status ditolak berhasil dikirim.',
                recipient: initiatorEmail,
                role: 'SPK Ditolak',
                messageId,
            });
        }

        return res.status(200).json({
            message: `Email SPK tidak dikirim. Status saat ini: "${status}"`,
        });
    } catch (error) {
        console.error('Terjadi kesalahan SPK:', error);
        return res.status(500).json({ error: 'Terjadi kesalahan internal server (SPK).', details: error.message });
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server email resender berjalan di port ${PORT}`);
});