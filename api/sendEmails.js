const nodemailer = require('nodemailer');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const { IncomingForm } = require('formidable');

// Set up the email transporter
const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
        user: process.env.EMAIL,
        pass: process.env.PASSWORD,
    },
});

module.exports = async (req, res) => {
    const form = new IncomingForm();

    form.parse(req, async (err, fields, files) => {
        if (err) return res.status(500).json({ message: 'Form parse error' });

        const { message } = fields;
        const excelFile = files.file;
        const attachment = files.attachment;

        if (!excelFile) {
            return res.status(400).json({ message: 'Excel file is required.' });
        }

        // Read Excel file data
        const workbook = xlsx.readFile(excelFile.filepath);
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(worksheet);

        // Prepare attachment
        const attachmentPath = attachment ? attachment.filepath : null;

        try {
            for (const row of data) {
                const email = row.Email;

                await transporter.sendMail({
                    from: process.env.EMAIL,
                    to: email,
                    subject: 'Your Subject Here',
                    text: message || 'Default message text here.',
                    attachments: attachmentPath ? [{ path: attachmentPath }] : [],
                });
            }

            res.status(200).json({ message: 'Emails sent successfully!' });
        } catch (error) {
            console.error(error);
            res.status(500).json({ message: 'Error sending emails' });
        } finally {
            // Clean up uploaded files
            fs.unlinkSync(excelFile.filepath);
            if (attachmentPath) fs.unlinkSync(attachmentPath);
        }
    });
};
