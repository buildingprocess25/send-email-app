require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const MailComposer = require('nodemailer/lib/mail-composer');

const app = express();
app.use(cors());
app.use(express.json());

// ==============================================================================
// KREDENSIAL UTAMA (SPARTA) -> Akses Sheets, Drive (PDF_STORAGE), & Gmail
// ==============================================================================
const spartaOAuth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    "https://developers.google.com/oauthplayground"
);
// PENTING: GOOGLE_REFRESH_TOKEN di Render harus sama dengan yang di token.json
spartaOAuth2Client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });

// Semua layanan Google sekarang hanya pakai 1 kredensial utama
const sheets = google.sheets({ version: 'v4', auth: spartaOAuth2Client });
const drive = google.drive({ version: 'v3', auth: spartaOAuth2Client });
const gmail = google.gmail({ version: 'v1', auth: spartaOAuth2Client });

// --- Helper Functions ---
function extractFileId(url) {
    if (!url) return null;
    const match = url.match(/(?:id=|d\/|file\/d\/)([\w-]{25,})/);
    return match ? match[1] : null;
}

async function downloadDriveFile(fileId) {
    if (!fileId) return null;
    try {
        const response = await drive.files.get(
            { fileId: fileId, alt: 'media' },
            { responseType: 'arraybuffer' }
        );
        console.log(`[Drive] Berhasil mengunduh ID: ${fileId}`);
        return Buffer.from(response.data);
    } catch (error) {
        console.error(`[Drive] Gagal mengunduh file ID ${fileId}:`, error.message);
        return null;
    }
}

function normalizeString(str) {
    if (!str) return "";
    return String(str).replace(/-/g, "").replace(/\s/g, "").trim().toUpperCase();
}

// === ENDPOINT API ===
app.post('/api/resend-email', async (req, res) => {
    const { ulok, lingkup } = req.body;

    if (!ulok || !lingkup) {
        return res.status(400).json({ error: 'Ulok dan Lingkup Pekerjaan harus diisi.' });
    }

    try {
        console.log(`[API] Memproses Ulok: ${ulok}, Lingkup: ${lingkup}...`);

        // SEKARANG MENGGUNAKAN SPREADSHEET_ID UTAMA
        const sheetId = process.env.SPREADSHEET_ID;

        if (!sheetId) {
            return res.status(500).json({ error: 'SPREADSHEET_ID tidak ditemukan di environment variables.' });
        }

        // --- STEP A: Cari Data di form2 ---
        const responseForm2 = await sheets.spreadsheets.values.get({
            spreadsheetId: sheetId,
            range: 'form2!A:O',
        });

        const rowsForm2 = responseForm2.data.values;
        if (!rowsForm2 || rowsForm2.length === 0) return res.status(404).json({ error: 'Data form2 kosong.' });

        const normalizedTargetUlok = normalizeString(ulok);
        const normalizedTargetLingkup = String(lingkup).trim().toLowerCase();

        const targetRow = rowsForm2.slice(1).find(row => {
            const rowUlok = normalizeString(row[9]);
            const rowLingkup = String(row[13] || "").trim().toLowerCase();
            return rowUlok === normalizedTargetUlok && rowLingkup === normalizedTargetLingkup;
        });

        if (!targetRow) return res.status(404).json({ error: 'Data tidak ditemukan di form2.' });

        const [
            status, timestamp, linkPdf, linkPdfNonSbo,
            emailKoord_old, waktuKoord, emailManager_old, waktuManager,
            emailPembuat, rowUlok, proyek, alamat, cabang, rowLingkup
        ] = targetRow;

        // --- STEP B: Tentukan Role ---
        let role = '';
        let targetJabatan = '';

        if (status === 'Menunggu Persetujuan Koordinator') {
            role = 'Koordinator';
            targetJabatan = 'BRANCH BUILDING COORDINATOR';
        } else if (status === 'Menunggu Persetujuan Manager') {
            role = 'Manager';
            targetJabatan = 'BRANCH BUILDING & MAINTENANCE MANAGER';
        } else {
            return res.status(200).json({ message: `Email tidak dikirim. Status saat ini: "${status}"` });
        }

        if (!cabang) return res.status(400).json({ error: 'Kolom cabang kosong.' });

        // --- STEP C: Cari BANYAK Email di Sheet Cabang ---
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

        const matchRowsCabang = rowsCabang.slice(1).filter(row => {
            const valCabang = String(row[idxCabang] || "").trim().toUpperCase();
            const valJabatan = String(row[idxJabatan] || "").trim().toUpperCase();
            return valCabang === targetCabangUpper && valJabatan === targetJabatanUpper;
        });

        if (matchRowsCabang.length === 0) {
            return res.status(404).json({ error: `Jabatan ${targetJabatan} tidak ditemukan di cabang ${cabang}.` });
        }

        const recipientEmailsArray = matchRowsCabang
            .map(row => String(row[idxEmail] || "").trim())
            .filter(email => email !== "");

        if (recipientEmailsArray.length === 0) {
            return res.status(404).json({ error: `Email tujuan ditemukan tapi datanya kosong di sheet Cabang.` });
        }

        const recipientEmailsStr = recipientEmailsArray.join(', ');
        console.log(`[API] Email tujuan ditemukan (${recipientEmailsArray.length} orang): ${recipientEmailsStr} (${role})`);


        // --- STEP D: Download PDF (Sekarang 100% pakai Auth Utama) ---
        const attachments = [];
        const pdfId = extractFileId(linkPdf);
        const pdfNonSboId = extractFileId(linkPdfNonSbo);

        if (pdfId) {
            const pdfBuffer = await downloadDriveFile(pdfId);
            if (pdfBuffer) attachments.push({ filename: 'RAB_SBO.pdf', content: pdfBuffer });
        }
        if (pdfNonSboId) {
            const pdfNonSboBuffer = await downloadDriveFile(pdfNonSboId);
            if (pdfNonSboBuffer) attachments.push({ filename: 'RAB_NON_SBO.pdf', content: pdfNonSboBuffer });
        }

        // --- STEP E: Kirim Email via GMAIL API ---
        const mailOptions = {
            from: `"Sparta System" <${process.env.EMAIL_USER}>`,
            to: recipientEmailsStr,
            subject: `[REMINDER] Persetujuan RAB - ${proyek} - Cabang ${cabang}`,
            html: `
                <h3>Halo,</h3>
                <p>Ini adalah pengingat bahwa terdapat dokumen RAB yang membutuhkan persetujuan Anda sebagai <strong>${role}</strong>.</p>
                <ul>
                    <li><strong>Proyek:</strong> ${proyek}</li>
                    <li><strong>Cabang:</strong> ${cabang}</li>
                    <li><strong>Ulok:</strong> ${rowUlok}</li>
                    <li><strong>Lingkup Pekerjaan:</strong> ${rowLingkup}</li>
                    <li><strong>Pembuat:</strong> ${emailPembuat}</li>
                </ul>
                <p>Dokumen PDF terlampir pada email ini. Mohon segera diproses ke dalam sistem Sparta.</p>
                <br/>
                <p>Terima kasih,</p>
                <p><strong>Building & Maintenance Dept.</strong></p>
            `,
            attachments: attachments
        };

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

        return res.status(200).json({
            message: 'Email berhasil dikirim.',
            recipient: recipientEmailsStr,
            role: role,
            messageId: result.data.id
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