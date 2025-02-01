# automated-email-sender
This Python project automates the process of sending personalized emails to research participants using Gmail and an Excel file as both data source and database. I originally developed it to invite participants to surveys or interviews, as we depended on the CRM team for email invitations (which used Mailchimp or other third-party platforms). By implementing this solution, the **FTI Group Research Team** that I led gained full control over scheduling, tracking sent messages, and logging responses—resulting in significant time savings and independence from paid messaging services. The script reads recipient details, personalizes email content, and ensures no duplicate sends.

---

![creamy portfolio emailsender github](https://github.com/user-attachments/assets/8cc8d821-40d8-4a2f-8cae-50d257aa3fe8)


---

Once messages are sent, the tool automatically updates Excel and subsequently transfers the data to SQL, expanding the research database with a consolidated record of participant interactions. By providing comprehensive statistics on when recipients were contacted and if/how they responded, the team can measure the success of recruitment efforts and refine future outreach strategies.

---

![5485be215909111 6776c3b5e241b](https://github.com/user-attachments/assets/93f2139f-12e8-43f1-8aaf-74623408d08a)


---

## 🔸 Features
✅ Reads **recipient email, name, and other details** from an Excel file  
✅ Sends **customized emails** based on recipient data  
✅ Updates **"Customer Data"** sheet by marking emails as "Sent"  
✅ Logs all sent emails in **"Sent Emails Data"**, preserving formatting  
✅ **Skips empty rows** and **doesn't resend emails**  
✅ Supports **batch sending (max 4 emails per run for testing)**  


---

## 🔸 Installation

1️⃣ **Clone the Repository**
```bash
git clone https://github.com/ninryt/automated-email-sender
.git
cd automated-email-sender

```
2️⃣ **Create a Virtual Environment** (Optional but Recommended)

```bash
python3 -m venv venv
source venv/bin/activate  # macOS/Linux
venv\Scripts\activate     # Windows
```
3️⃣ **Required Packages**
```bash
pip install openpyxl
```
---
## 🔸 Setup Gmail SMTP Authenthication

Create an App Password for security:
[click here](https://accounts.google.com/v3/signin/challenge/pwd?TL=AO-GBTcH6IznI2l9eCaASOBU-tZ-Jp_8lwfpiPzNJUREDQueVW6ULyuU1xBq-Qz5&cid=2&continue=https%3A%2F%2Fmyaccount.google.com%2Fapppasswords&flowName=GlifWebSignIn&followup=https%3A%2F%2Fmyaccount.google.com%2Fapppasswords&ifkv=AVdkyDmsdu2m1LxZCZHnP6N4o43AxziCFbhEeO6uGRUqj5zgUonO5AcOWApoz7-tF-6DXi03Anjy&osid=1&rart=ANgoxcfMq2NguWng04csK2FyiHftiIG7mkAQBV1k-Mue2caDW3BRezRLpw-pyghIjaVtmeWKQLWmSdu_Z8SuqwwWpJN51NFVcnYw-Zk9sxtkTrNeovVEU2U&rpbg=1&service=accountsettings) 

```
EMAIL=your_email@gmail.com
PASSWORD=your_app_password
```
---
## 🔸 Download the excel file I prepared for you: 
[Excel_DataSource](./email_sender_datasource.xlsx)

- Open and review the Excel file (which has two sheets: "Customer Data" and "Sent Emails Data").
- Go to the "Customer Data" sheet.
- Fill in the columns: Email, Last Name, First Name, Gender.
- The script will automatically update the "Status" column to "Sent" after sending emails, don't fill in the column "Sent".
- Don't fill in the sheet "Sent Emails Data", the sheet will automatically update once your emails will be sent.

---

![excel portfolio emailsender github](https://github.com/user-attachments/assets/041564af-6015-4869-ba21-3635feb72420)

---
## 🔸 Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss your ideas.

## 🔸 LICENSE
This project is licensed under the MIT License. See the LICENSE file for more details.

---
## 👤 Author
👤 N.B. Ryttel
📧 [Email me](zerobughero@gmail.com)
🔗 [Github](https://github.com/ninryt)



  


