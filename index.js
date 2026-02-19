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
        const fromAddress = `"Sparta System Resend Email" <${getEnvValue('EMAIL_USER')}>`;

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
            const subject = `[RE-EMAIL][FINAL - DISETUJUI] Pengajuan RAB Proyek ${namaToko}: ${proyek} - ${rowLingkup}`;
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
                    ? `[RE-EMAIL][TAHAP 1: PERLU PERSETUJUAN] RAB Proyek ${proyek} - ${rowLingkup}`
                    : `[RE-EMAIL][TAHAP 2: PERLU PERSETUJUAN] RAB Proyek ${proyek} - ${rowLingkup}`,
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

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server email resender berjalan di port ${PORT}`);
});