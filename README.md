# 🚀 AutoForge Email Sender Pro

Send personalized bulk emails in seconds with a clean, powerful desktop tool.

---

## ⚡ Features
- Bulk email sending from CSV
- Personalized messages using `{name}`, `{email}`, `{company}`
- File attachments support
- Exportable delivery logs
- Clean, professional UI

---

## 🛠 How It Works
1. Load your `contacts.csv`
2. Enter your Gmail + App Password
3. Write your message
4. Click **Send Emails**

---

## 📁 Project Structure
AutoForge-EmailSender/
│
├── src/
│ ├── main.py
│ ├── gui.py
│ ├── email_sender.py
│ └── utils.py
│
├── assets/
│ ├── email_sender_main.png
│ ├── email_sender_settings.png
│ ├── email_sender_previewcontacts.png
│ └── email_sender_deliverylogs.png
│
├── data/
│ └── contacts.csv
│
├── output/
├── config.json
├── requirements.txt
└── README.md

---

## 🖼 Screenshots

### Main Interface
![Main UI](assets/email%20sender_main.png)

### Settings
![Settings](assets/email%20sender_settings.png)

### Contact Preview
![Contacts](assets/email%20sender_previewcontacts.png)

### Delivery Logs
![Logs](assets/email%20sender_deliverylog.png)
---

## 📦 Requirements
pip install -r requirements.txt

---

## 🚀 Run the App
python src/main.py

---

## 📌 Notes
- Use a Gmail App Password (not your normal password)
- Ensure your CSV follows this format:

email,name,company  
example@gmail.com,John Doe,Company Inc

---

## 💼 Built With
- Python
- Tkinter
- Pandas

---

## 🔥 Author
AutoForge — Automated Tools Development