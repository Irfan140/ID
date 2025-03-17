const XLSX = require('xlsx');
const path = require('path');
const nodemailer = require('nodemailer');

const saveToExcel = (data, callback) => {
    const filePath = path.join(__dirname, '../data/users.xlsx');
    let workbook;

    try {
        workbook = XLSX.readFile(filePath);
    } catch (err) {
        workbook = XLSX.utils.book_new();
    }

    const worksheetName = 'Users';
    const worksheet = XLSX.utils.json_to_sheet([data]);

    if (workbook.SheetNames.includes(worksheetName)) {
        const existingSheet = workbook.Sheets[worksheetName];
        const existingData = XLSX.utils.sheet_to_json(existingSheet);
        existingData.push(data);

        const updatedSheet = XLSX.utils.json_to_sheet(existingData);
        workbook.Sheets[worksheetName] = updatedSheet;
    } else {
        XLSX.utils.book_append_sheet(workbook, worksheet, worksheetName);
    }

    XLSX.writeFile(workbook, filePath);

    // Send email after saving
    sendEmailWithAttachment(filePath);

    callback(null, 'Data saved successfully and email sent!');
};

const sendEmailWithAttachment = (filePath) => {
    const transporter = nodemailer.createTransport({
        service: 'Gmail',
        auth: {
            user: process.env.EMAIL,         // Add your email in .env file
            pass: process.env.EMAIL_PASS    // Add your email app password in .env file
        }
    });

    const mailOptions = {
        from: process.env.EMAIL,
        to: process.env.RECIPIENT_EMAIL,  // Your email or the desired recipient
        subject: 'New ID Card Data - Excel File',
        text: 'Please find the attached Excel file with the latest data.',
        attachments: [
            {
                filename: 'users.xlsx',
                path: filePath
            }
        ]
    };

    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            console.error('Error sending email:', error);
        } else {
            console.log('Email sent successfully:', info.response);
        }
    });
};

module.exports = saveToExcel;
