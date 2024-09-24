const express = require('express');
const nodemailer = require('nodemailer');
const multer = require('multer');
const csv = require('csv-parser');
const fs = require('fs');
const path = require('path');
const axios = require('axios');

const app = express();
const port = 3004;

// Setup Multer for file uploads
const upload = multer({ dest: 'uploads/' });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Route to serve the HTML form
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Function to replace placeholders in the email template
function fillTemplate(template, data, defaultFields) {
  const defaults = defaultFields || {};
  return template
    .replace(/{{email}}/g, data.email || defaults.email || 'N/A')
    .replace(/{{name}}/g, data.name || defaults.name || 'N/A')
    .replace(/{{middle_name}}/g, data.middle_name || defaults.middle_name || 'N/A')
    .replace(/{{sur_name}}/g, data.sur_name || defaults.sur_name || 'N/A')
    .replace(/{{phone_no}}/g, data.phone_no || defaults.phone_no || 'N/A')
    .replace(/{{salutation}}/g, data.salutation || defaults.salutation || 'Dear')
    .replace(/{{position}}/g, data.position || defaults.position || 'Position Not Specified');
}

// Parse custom fields into an object
function parseCustomFields(customFieldsStr = '') {
  if (!customFieldsStr) return {};
  const fields = customFieldsStr.split(';');
  return fields.reduce((acc, field) => {
    const [key, value] = field.split('=');
    if (key && value) {
      acc[key.trim()] = value.trim();
    }
    return acc;
  }, {});
}

// Function to download a file from a remote URL and return a stream
async function downloadFile(url) {
  try {
    const response = await axios({
      url,
      method: 'GET',
      responseType: 'stream',
    });
    return response.data;
  } catch (error) {
    console.error(`Failed to download file from: ${url} - ${error.message}`);
    throw new Error(`Failed to download file from URL: ${url}`);
  }
}

// Function to read a local file and return a stream
function readLocalFile(filePath) {
  return new Promise((resolve, reject) => {
    const fileStream = fs.createReadStream(filePath);
    fileStream.on('error', (error) => {
      console.error(`Failed to read local file: ${filePath} - ${error.message}`);
      reject(new Error(`Failed to read local file: ${filePath}`));
    });
    fileStream.on('open', () => resolve(fileStream));
  });
}

// Clean and normalize field values, handling extra quotes and spaces
function cleanFieldValue(value, fieldName = '') {
  if (typeof value === 'string') {
    const cleanedValue = value.trim().replace(/^"+|"+$/g, ''); // Remove any surrounding quotes
    console.log(`Field "${fieldName}" cleaned:`, cleanedValue); // Log how the field is being cleaned
    return cleanedValue;
  }
  console.log(`Field "${fieldName}" is not a string, original value returned:`, value);
  return value;
}

// Normalize headers by trimming and converting to lowercase
function normalizeHeader(header) {
  return header.trim().toLowerCase();
}

// Handle form submission
app.post('/sendEmails', upload.single('csvFile'), async (req, res) => {
  const {
    smtpHost,
    smtpUser,
    smtpPass,
    subject,
    ccEmails = '',
    emailBody,
    customFields = '',
  } = req.body;

  const csvFilePath = req.file.path;
  const defaultFields = parseCustomFields(customFields);
  const ccEmailsList = ccEmails.split(',').map(email => email.trim()).filter(email => email);

  const portNumber = () => {
    return (smtpHost === "smtp.gmail.com" || smtpHost === "smtp.zoho.com") ? 465 : 587;
  };

  // Create transporter
  const transporter = nodemailer.createTransport({
    host: smtpHost,
    port: portNumber(),
    secure: smtpHost === "smtp.gmail.com" || smtpHost === "smtp.zoho.com",
    auth: {
      user: smtpUser,
      pass: smtpPass,
    },
  });

  let report = { success: [], failure: [] };

  // Function to send email with optional attachment
  async function sendEmail(to, template, attachmentStream, filename, data) {
    const filledBody = fillTemplate(template, data, defaultFields);

    const mailOptions = {
      from: smtpUser,
      to: to,
      cc: ccEmailsList,
      subject: subject,
      html: filledBody,
    };

    // Attach the file only if attachmentStream and filename are available
    if (attachmentStream && filename) {
      mailOptions.attachments = [
        {
          filename: filename,
          content: attachmentStream,
        },
      ];
    }

    try {
      await transporter.sendMail(mailOptions);
      console.log(`Email sent to: ${to}`);
      report.success.push(to);
    } catch (error) {
      console.error(`Failed to send email to: ${to} - ${error.message}`);
      report.failure.push({ email: to, error: error.message });
    }
  }

  // Read CSV file and process each row
  const emailPromises = [];

  // Start of CSV read process
  fs.createReadStream(csvFilePath, { encoding: 'utf-8' }) // Ensure UTF-8 encoding
    .pipe(csv({
      mapHeaders: ({ header }) => normalizeHeader(header), // Normalize headers to avoid case or space issues
    }))
    .on('headers', (headers) => {
      console.log('CSV Headers:', headers); // Log headers to ensure correct fields
    })
    .on('data', async (row) => {
      // Clean each field value before using it
      const email = cleanFieldValue(row['email'], 'email');
      const send_email = cleanFieldValue(row['send_email'], 'send_email').toLowerCase();

      // Debugging: Log the entire row
      console.log('Raw Row Data:', row);

      console.log(`Cleaned email: ${email}, send_email: ${send_email}`);

      // Check for missing email
      if (!email || email === 'undefined') {
        console.error('Email field is missing or undefined in the CSV row:', row);
        return;
      }

      // Proceed only if send_email is 'yes'
      if (send_email === 'yes') {
        const emailPromise = (async () => {
          const emailData = {
            email: email,
            name: cleanFieldValue(row['name'], 'name'),
            middle_name: cleanFieldValue(row['middle_name'], 'middle_name'),
            sur_name: cleanFieldValue(row['sur_name'], 'sur_name'),
            salutation: cleanFieldValue(row['salutation'], 'salutation'),
            position: cleanFieldValue(row['position'], 'position'),
          };

          // Debugging: Ensure that email and data fields are populated correctly
          console.log(`Preparing to send email to: ${email}`);
          
          // Call your email sending function
          await sendEmail(email, emailBody, null, null, emailData);
        })();

        emailPromises.push(emailPromise);
      }
    })
    .on('end', async () => {
      await Promise.all(emailPromises);
      fs.unlinkSync(csvFilePath); // Clean up uploaded file
      console.log('CSV processing complete. Report:', report);
      res.json({ report });
    })
    .on('error', (error) => {
      console.error('Error reading CSV file:', error);
      res.status(500).json({ error: 'Error processing CSV file' });
    });
});

// Start the server
app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
