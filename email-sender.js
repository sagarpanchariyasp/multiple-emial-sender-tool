import nodemailer from 'nodemailer';
import XLSX from 'xlsx';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

// Get the directory name of the current module
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Configure Gmail SMTP transporter
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: 'techandfacts2000@gmail.com',
    pass: 'zdox sdit bcko emqt'
  },
  port: 587,
  secure: false,
  tls: {
    rejectUnauthorized: false
  }
});

// Path to the Excel file
const excelFilePath = join(__dirname, 'emails.xlsx');

// Read the Excel file
const workbook = XLSX.readFile(excelFilePath);
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

// Convert sheet to JSON format
const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

// Extract emails from the first column and convert to lowercase
const emailList = data.flat().filter(email => email.includes('@')).map(email => email.toLowerCase());

// Email template
const getEmailTemplate = () => `
  <h1>Hello!</h1>
  <p>This is a professionally designed email using HTML and CSS.</p>
  <p>Sent automatically by our Node.js tool.</p>
  <p>Best Regards,<br>Your Company</p>
`;

// Function to send email
const sendEmail = (recipient) => {
  const mailOptions = {
    from: 'techandfacts2000@gmail.com',
    to: recipient,
    subject: 'Professional Email from Node.js Tool',
    html: getEmailTemplate()
  };

  transporter.sendMail(mailOptions, (error, info) => {
    if (error) {
      console.log('Error occurred:', error);
    } else {
      console.log(`Email sent to: ${recipient}`);
    }
  });
};

// Send emails one by one with a delay of 1 minute (60000 milliseconds)
let emailIndex = 0;

const sendEmails = () => {
  if (emailIndex < emailList.length) {
    const recipient = emailList[emailIndex];
    sendEmail(recipient);
    emailIndex++;
    setTimeout(sendEmails, 60000); // Wait 1 minute before sending the next email
  } else {
    console.log('All emails have been sent!');
  }
};

// Start sending emails
sendEmails();
