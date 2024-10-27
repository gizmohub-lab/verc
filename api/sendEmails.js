const express = require('express');
const nodemailer = require('nodemailer');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const router = express.Router();
const upload = multer({ dest: 'uploads/' });

router.post('/send-emails', upload.fields([{ name: 'file' }, { name: 'attachment' }]), async (req, res) => {
    const { file } = req.files;

    if (!file || file.length === 0) {
        return res.status(400).send('Excel file is required.');
    }

    const workbook = xlsx.readFile(file[0].path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Convert the sheet to JSON
    const data = xlsx.utils.sheet_to_json(worksheet);
    const attachmentPath = req.files.attachment ? path.join(__dirname, '../uploads/', req.files.attachment[0].path) : null;

    // Set up the nodemailer transporter
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: process.env.EMAIL,
            pass: process.env.PASSWORD,
        },
    });

    // Extract the message from the request body
    const messageText = req.body.message; // Assuming the message is sent as a form field

    try {
        // Iterate through the email addresses and send emails
        for (const row of data) {
            const email = row.Email; // Assuming the Excel file has a column named "Email"

            await transporter.sendMail({
                from: process.env.EMAIL,
                to: email,
                subject: 'Your Subject Here',
                text: messageText || 'Default message text here.', // Use the message from the body
                attachments: attachmentPath ? [{ path: attachmentPath }] : [], // Attach if there is a file
            });
        }
        res.status(200).send('Emails sent successfully!');
    } catch (error) {
        console.error(error);
        res.status(500).send('Error sending emails');
    } finally {
        // Clean up the uploaded files
        fs.unlinkSync(file[0].path);
        if (attachmentPath) fs.unlinkSync(attachmentPath);
    }
});

module.exports = router;
