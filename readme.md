
```markdown
# Bulk Email Sender (Excel → Nodemailer)

A Node.js script to read emails from an Excel file and send a **custom HTML email template** to all recipients using **Nodemailer**. The script asks for the Excel filename in the terminal, shows all emails found, and sends the email after confirmation.

---

## **Features**

- Reads **all sheets and columns** of the Excel file to extract emails.
- Removes duplicate emails automatically.
- Sends **custom HTML emails**.
- Asks for **confirmation before sending**.
- Supports attaching files (PDF, images, etc.) if needed.

---

## **Project Structure**


email-bulk-sender/
│
├── index.js           # Main Node.js script
├── template.html      # Your email HTML template
├── .env               # Environment variables (email + password)
├── emails.xlsx        # Example Excel file (placed at root)
└── README.md

````

---

## **Setup Instructions**

### **1. Install Node.js**

Make sure Node.js is installed. Check by:

```bash
node -v
npm -v
````

---

### **2. Clone the project / create folder**

```bash
git clone <repo-url>
cd email-bulk-sender
```

or just create a folder and place the files there.

---

### **3. Install dependencies**

```bash
npm install xlsx nodemailer dotenv readline-sync
```

---

### **4. Setup `.env` file**

Create a `.env` file in the project root:

```env
EMAIL_USER=youremail@gmail.com
EMAIL_PASS=your-app-password
```

#### **How to get Gmail App Password**

1. Go to your **Google Account → Security → App Passwords**
2. Select **Mail** and **Other (Custom name)**, then generate.
3. Copy the password and paste as `EMAIL_PASS`.
4. Make sure **2FA is enabled** for your Google account, otherwise App Password won't work.

> ⚠️ **Do NOT use your normal Gmail password**. App Password is required.

---

### **5. Prepare the Excel file**

* Place the Excel file in the **root of the project** (same folder as `index.js`).
* Example: `emails.xlsx`
* It can have **any number of sheets**.
* Emails can be **in any column**.
* Example:

| Name | Email                                       | Notes     |
| ---- | ------------------------------------------- | --------- |
| John | [john@example.com](mailto:john@example.com) | Important |
| Jane | [jane@example.com](mailto:jane@example.com) | Follow up |

---

### **6. Prepare the HTML template**

* Place your **custom email template** in `template.html`
* Example: `template.html` contains your full recruiter invitation or any email content.
* Nodemailer will read this file and send it as HTML body.

---

## **Usage**

1. Run the script:

```bash
node index.js
```

2. The terminal will ask for **Excel filename**:

```
Enter Excel filename (with extension, e.g., emails.xlsx):
```

* Type your file name, e.g., `emails.xlsx`

3. The script will **show all emails found**:

```
✅ Emails found:
john@example.com
jane@example.com
...
```

4. You will be asked for **confirmation**:

```
Do you want to send emails to these addresses? (yes/no):
```

* Type `yes` → emails are sent
* Type `no` → script exits

5. Emails are sent with a **1.5-second delay** between each to avoid spam limits.

---

## **Supported File Types**

* **Email extraction:** `.xlsx`, `.xls`, `.csv`
* **Attachments (optional):** `.pdf`, `.docx`, `.xlsx`, `.png`, `.jpg`, `.zip`

---

## **Tips**

* Make sure the **Excel file has valid email addresses**.
* Avoid sending too many emails at once on Gmail to prevent **rate limits**.
* You can attach files in `index.js` using Nodemailer’s `attachments` array.
* You can personalize emails if your Excel has a `Name` column (optional enhancement).

---

## **Example Flow**

1. Place `emails.xlsx` and `template.html` in project root.
2. Fill `.env` with Gmail email and App Password.
3. Run `node index.js`.
4. Enter Excel filename when prompted.
5. Confirm the emails displayed in terminal.
6. Emails are sent to all extracted addresses.
