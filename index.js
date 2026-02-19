require('dotenv').config();
const express = require('express');
const cors = require('cors');
const { google } = require('googleapis');
const nodemailer = require('nodemailer');

const app = express();
app.use(cors());
app.use(express.json()); // Agar bisa menerima payload JSON

// Setup Google OAuth2 Client
const oAuth2Client = new google.auth.OAuth2(
    process.env.GOOGLE_CLIENT_ID,
    process.env.GOOGLE_CLIENT_SECRET,
    "https://developers.google.com/oauthplayground"
);
oAuth2Client.setCredentials({ refresh_token: process.env.GOOGLE_REFRESH_TOKEN });

const sheets = google.sheets({ version: 'v4', auth: oAuth2Client });
const drive = google.drive({ version: 'v3', auth: oAuth2Client });

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
        console.error(`Gagal mengunduh file ID ${fileId}:`, error.message);
        return null;
    }
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
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: sheetId,
            range: 'form2!A:O',
        });

        const rows = response.data.values;
        if (!rows || rows.length === 0) {
            return res.status(404).json({ error: 'Data tidak ditemukan di Google Sheet.' });
        }

        const targetRow = rows.slice(1).find(row => row[10] === ulok && row[14] === lingkup);

        if (!targetRow) {
            return res.status(404).json({ error: 'Data dengan Ulok dan Lingkup Pekerjaan tersebut tidak ditemukan.' });
        }

        const [
            status, timestamp, linkPdf, linkPdfNonSbo,
            emailKoord, waktuKoord, emailManager, waktuManager,
            emailPembuat, nomor, rowUlok, proyek, alamat, cabang, rowLingkup
        ] = targetRow;

        let recipientEmail = '';
        let role = '';

        if (status === 'Menunggu Persetujuan Koordinator') {
            recipientEmail = emailKoord;
            role = 'Koordinator';
        } else if (status === 'Menunggu Persetujuan Manager') {
            recipientEmail = emailManager;
            role = 'Manager';
        } else {
            return res.status(200).json({ message: `Email tidak dikirim. Status saat ini: "${status}"` });
        }

        if (!recipientEmail) {
            return res.status(400).json({ error: `Email ${role} kosong di Sheet.` });
        }

        // Download PDF
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

        // Setup Nodemailer
        const accessToken = await oAuth2Client.getAccessToken();
        const transporter = nodemailer.createTransport({
            service: 'gmail',
            auth: {
                type: 'OAuth2',
                user: process.env.EMAIL_USER,
                clientId: process.env.GOOGLE_CLIENT_ID,
                clientSecret: process.env.GOOGLE_CLIENT_SECRET,
                refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
                accessToken: accessToken.token,
            },
        });

        // Kirim Email
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

// Port dari environment (penting untuk Render)
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server email resender berjalan di port ${PORT}`);
});