const XLSX = require('xlsx');
const path = require('path');
const nodemailer = require('nodemailer');

const saveToExcel = (data, callback) => {
    console.log('[INFO] Received data to save:', data);

    const filePath = path.join(__dirname, '../data/users.xlsx');
    let workbook;

    try {
        workbook = XLSX.readFile(filePath);
        console.log('[INFO] Existing workbook found and loaded.');
    } catch (err) {
        console.log('[WARNING] No existing workbook found. Creating a new one.');
        workbook = XLSX.utils.book_new();
    }

    const worksheetName = 'Users';
    const worksheet = XLSX.utils.json_to_sheet([data]);

    if (workbook.SheetNames.includes(worksheetName)) {
        const existingSheet = workbook.Sheets[worksheetName];
        const existingData = XLSX.utils.sheet_to_json(existingSheet);
        console.log(`[INFO] Existing sheet found. Current data count: ${existingData.length}`);

        existingData.push(data);

        const updatedSheet = XLSX.utils.json_to_sheet(existingData);
        workbook.Sheets[worksheetName] = updatedSheet;
    } else {
        console.log('[INFO] Creating new worksheet for data.');
        XLSX.utils.book_append_sheet(workbook, worksheet, worksheetName);
    }

    XLSX.writeFile(workbook, filePath);
    console.log('[SUCCESS] Data successfully saved to Excel:', filePath);

    // Send email after saving
    console.log('[INFO] Attempting to send email with attachment...');
    sendEmailWithAttachment(filePath, callback);
};

const sendEmailWithAttachment = (filePath, callback) => {
    console.log('[INFO] Configuring email transporter...');

    const transporter = nodemailer.createTransport({
        service: 'Gmail',
        auth: {
            user: process.env.EMAIL,
            pass: process.env.EMAIL_PASS
        }
    });

    const mailOptions = {
        from: process.env.EMAIL,
        to: process.env.RECIPIENT_EMAIL,
        subject: 'New ID Card Data - Excel File',
        text: 'Please find the attached Excel file with the latest data.',
        attachments: [
            {
                filename: 'users.xlsx',
                path: filePath
            }
        ]
    };

    console.log('[INFO] Sending email...');
    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            console.error('[ERROR] Error sending email:', error);
            callback(error, null); // Added proper callback error handling
        } else {
            console.log('[SUCCESS] Email sent successfully:', info.response);
            callback(null, 'Data saved successfully and email sent!');
        }
    });
};

module.exports = {saveToExcel};
