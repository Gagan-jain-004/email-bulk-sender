import xlsx from "xlsx";
import nodemailer from "nodemailer";
import dotenv from "dotenv";
import fs from "fs/promises";
import readline from "readline-sync";

dotenv.config();

// === Step 1: Ask for Excel filename ===
const excelFile = readline.question("Enter Excel filename (with extension, e.g., emails.xlsx): ");

// Check if file exists
try {
  await fs.access(excelFile);
} catch {
  console.error("❌ File not found!");
  process.exit(1);
}

// === Step 2: Read Excel file ===
const workbook = xlsx.readFile(excelFile);
let emails = [];
const emailRegex = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i;

// Loop through all sheets and columns
for (let sheetName of workbook.SheetNames) {
  const sheet = workbook.Sheets[sheetName];
  const rows = xlsx.utils.sheet_to_json(sheet, { header: 1 }); // 2D array

  rows.forEach((row) => {
    row.forEach((cell) => {
      if (typeof cell === "string" && emailRegex.test(cell)) {
        emails.push(cell.match(emailRegex)[0]);
      }
    });
  });
}

// Remove duplicates
emails = [...new Set(emails)];

if (emails.length === 0) {
  console.log("No emails found in the file!");
  process.exit(0);
}

console.log("\n✅ Emails found:");
emails.forEach((email) => console.log(email));

// === Step 3: Ask for confirmation before sending ===
const confirm = readline.question("\nDo you want to send emails to these addresses? (yes/no): ");
if (confirm.toLowerCase() !== "yes") {
  console.log("Operation cancelled.");
  process.exit(0);
}

// === Step 4: Read HTML template ===
let emailTemplate;
try {
  emailTemplate = await fs.readFile("template.html", "utf8");
} catch (err) {
  console.error("Failed to read template.html", err);
  process.exit(1);
}

// === Step 5: Setup Nodemailer ===
const transporter = nodemailer.createTransport({
  service: "gmail",
  auth: {
    user: process.env.EMAIL_USER,
    pass: process.env.EMAIL_PASS,
  },
});

const mailOptions = {
  from: process.env.EMAIL_USER,
  subject: "Your Custom Subject",
  html: emailTemplate,
};

// === Step 6: Send emails with delay ===
async function sendEmails() {
  for (let email of emails) {
    try {
      await transporter.sendMail({ ...mailOptions, to: email });
      console.log(`✅ Email sent to: ${email}`);
      await new Promise((r) => setTimeout(r, 1500)); // 1.5 sec delay
    } catch (error) {
      console.log(`❌ Failed to send email to: ${email}`, error.message);
    }
  }
}

sendEmails();
