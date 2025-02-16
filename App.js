const express = require('express');
const nodemailer = require('nodemailer');
const path = require('path');

const app = express();
const port = 3000;

// To parse JSON bodies in POST requests
app.use(express.json());

// Serve static files (including sample.xlsx and index.html)
app.use(express.static('views'));

// Endpoint to send an email for a single row of data
app.post('/sendEmail', async (req, res) => {
  try {
    // Expecting sender credentials and email data from the frontend
    const { senderEmail, senderPass, emailData } = req.body;
    if (!senderEmail || !senderPass) {
      return res.status(400).json({ error: 'Sender email and password are required.' });
    }
    if (!emailData || !emailData.email) {
      return res.status(400).json({ error: 'Recipient email is required in the Excel data.' });
    }
    
    // Create a transporter using the sender's credentials
    const transporter = nodemailer.createTransport({
      service: 'gmail',
      auth: {
        user: senderEmail,
        pass: senderPass
      }
    });

    // Extract details from emailData (adjust keys as per your Excel file)
    const { name, email, message } = emailData;

    const mailOptions = {
      from: `"Sender" <${senderEmail}>`,
      to: email,
      subject: 'Automated Email',
      text: `Hello ${name || ''},\n\n${message || 'No message provided.'}\n\nRegards,\nYour Team`
    };

    await transporter.sendMail(mailOptions);
    console.log(`Email sent to ${email}`);
    res.status(200).json({ success: true });
  } catch (error) {
    console.error('Error sending email:', error);
    res.status(500).json({ error: 'Error sending email' });
  }
});

// Start the server
app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
