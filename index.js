require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const nodemailer = require('nodemailer');

const app = express();
app.use(cors());
app.use(express.json());

// ==============================================================================
// 1. Setup OAuth2 Client untuk SHEETS (Tetap pakai DOC_ jika sheetnya ada di sana)
// ==============================================================================
const docOAuth2Client = new google.auth.OAuth2(
    process.env.DOC_GOOGLE_CLIENT_ID,
    process.env.DOC_GOOGLE_CLIENT_SECRET,
    "https://developers.google.com/oauthplayground"
);
docOAuth2Client.setCredentials({ refresh_token: process.env.DOC_GOOGLE_REFRESH_TOKEN });

const sheets = google.sheets({ version: 'v4', auth: docOAuth2Client });

// ==============================================================================
// 2. Setup OAuth2 Client UTAMA (Untuk Email & DRIVE)
//    Kita pakai ini untuk Drive juga agar tidak error 404 saat download file
// ==============================================================================
const mailOAuth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    "https://developers.google.com/oauthplayground"
);
mailOAuth2Client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });

// Pakai auth utama untuk Drive
const drive = google.drive({ version: 'v3', auth: mailOAuth2Client });


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
        return Buffer.from(response.data);
    } catch (error) {
        // Log error tapi jangan biarkan aplikasi crash
        console.error(`Gagal mengunduh file ID ${fileId}:`, error.message);
        return null;
    }
}

// Fungsi untuk menyamakan format Ulok (Hapus strip dan spasi)
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

        const sheetId = process.env.DOC_SHEET_ID;

        // -------------------------------------------------------------
        // STEP A: AMBIL DATA DARI form2 UNTUK CEK STATUS & CABANG
        // -------------------------------------------------------------
        const responseForm2 = await sheets.spreadsheets.values.get({
            spreadsheetId: sheetId,
            range: 'form2!A:O',
        });

        const rowsForm2 = responseForm2.data.values;
        if (!rowsForm2 || rowsForm2.length === 0) {
            return res.status(404).json({ error: 'Data tidak ditemukan di Google Sheet (form2).' });
        }

        const normalizedTargetUlok = normalizeString(ulok);
        const normalizedTargetLingkup = String(lingkup).trim().toLowerCase();

        // Cari baris menggunakan INDEX: Ulok (Index 9) & Lingkup (Index 13)
        const targetRow = rowsForm2.slice(1).find(row => {
            const rowUlok = normalizeString(row[9]);
            const rowLingkup = String(row[13] || "").trim().toLowerCase();
            return rowUlok === normalizedTargetUlok && rowLingkup === normalizedTargetLingkup;
        });

        if (!targetRow) {
            return res.status(404).json({ error: 'Data dengan Ulok dan Lingkup Pekerjaan tersebut tidak ditemukan di form2.' });
        }

        const [
            status, timestamp, linkPdf, linkPdfNonSbo,
            emailKoord_old, waktuKoord, emailManager_old, waktuManager,
            emailPembuat, rowUlok, proyek, alamat, cabang, rowLingkup
        ] = targetRow;

        // -------------------------------------------------------------
        // STEP B: TENTUKAN TARGET JABATAN BERDASARKAN STATUS
        // -------------------------------------------------------------
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

        if (!cabang) {
            return res.status(400).json({ error: 'Kolom cabang di form2 kosong, tidak bisa mencari email tujuan.' });
        }

        // -------------------------------------------------------------
        // STEP C: CARI EMAIL DI SHEET "Cabang"
        // -------------------------------------------------------------
        const responseCabang = await sheets.spreadsheets.values.get({
            spreadsheetId: sheetId,
            range: 'Cabang!A:Z',
        });

        const rowsCabang = responseCabang.data.values;
        if (!rowsCabang || rowsCabang.length === 0) {
            return res.status(404).json({ error: 'Sheet Cabang kosong atau tidak ditemukan.' });
        }

        const headersCabang = rowsCabang[0].map(h => String(h).trim().toUpperCase());
        const idxCabang = headersCabang.indexOf('CABANG');
        const idxJabatan = headersCabang.indexOf('JABATAN');
        const idxEmail = headersCabang.indexOf('EMAIL_SAT');

        if (idxCabang === -1 || idxJabatan === -1 || idxEmail === -1) {
            return res.status(500).json({ error: 'Format header di sheet Cabang salah. Pastikan ada kolom CABANG, JABATAN, dan EMAIL_SAT.' });
        }

        const targetCabangUpper = String(cabang).trim().toUpperCase();
        const targetJabatanUpper = targetJabatan.toUpperCase();

        const matchRowCabang = rowsCabang.slice(1).find(row => {
            const valCabang = String(row[idxCabang] || "").trim().toUpperCase();
            const valJabatan = String(row[idxJabatan] || "").trim().toUpperCase();
            return valCabang === targetCabangUpper && valJabatan === targetJabatanUpper;
        });

        if (!matchRowCabang || !matchRowCabang[idxEmail]) {
            return res.status(404).json({ error: `Gagal mengirim. Email untuk jabatan ${targetJabatan} di cabang ${cabang} tidak ditemukan di sheet Cabang.` });
        }

        const recipientEmail = String(matchRowCabang[idxEmail]).trim();
        console.log(`[API] Email tujuan ditemukan: ${recipientEmail} (${role} - ${cabang})`);

        // -------------------------------------------------------------
        // STEP D: DOWNLOAD PDF & KIRIM EMAIL
        // -------------------------------------------------------------
        const attachments = [];
        const pdfId = extractFileId(linkPdf);
        const pdfNonSboId = extractFileId(linkPdfNonSbo);

        // Download file (sekarang pakai auth Utama, semoga 404 hilang)
        if (pdfId) {
            const pdfBuffer = await downloadDriveFile(pdfId);
            if (pdfBuffer) attachments.push({ filename: 'RAB_SBO.pdf', content: pdfBuffer });
        }

        if (pdfNonSboId) {
            const pdfNonSboBuffer = await downloadDriveFile(pdfNonSboId);
            if (pdfNonSboBuffer) attachments.push({ filename: 'RAB_NON_SBO.pdf', content: pdfNonSboBuffer });
        }

        const accessToken = await mailOAuth2Client.getAccessToken();

        // KONFIGURASI FIX IPV6: Tambahkan family: 4
        const transporter = nodemailer.createTransport({
            service: 'gmail',
            family: 4, // <--- INI PENTING! Memaksa pakai IPv4
            auth: {
                type: 'OAuth2',
                user: process.env.EMAIL_USER,
                clientId: process.env.GOOGLE_CLIENT_ID,
                clientSecret: process.env.GOOGLE_CLIENT_SECRET,
                refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
                accessToken: accessToken.token,
            },
        });

        const mailOptions = {
            from: `"Sparta System" <${process.env.EMAIL_USER}>`,
            to: recipientEmail,
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

        const result = await transporter.sendMail(mailOptions);

        return res.status(200).json({
            message: 'Email berhasil dikirim.',
            recipient: recipientEmail,
            role: role,
            messageId: result.messageId
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