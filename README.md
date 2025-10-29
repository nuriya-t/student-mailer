# Student Email Distribution Script using Microsoft Graph API

This Python script automates personalized email distribution to students using **Microsoft Graph API**.  
It reads student data from an Excel file and sends customized messages via Outlook —  
securely and without using SMTP or Outlook COM objects.



Features

Sends personalized HTML emails to students  
- Filters recipients by:
- Email address
- Discipline
- Academic level (Bachelor / Master)
- Works with institutional or Microsoft 365 accounts  
- Uses **OAuth 2.0 Device Code Flow** for secure authentication  
- No need to store or use passwords  

---

## Requirements

- Python 3.9 or higher  
- A registered **App** in [Azure Entra ID (Microsoft 365)](https://entra.microsoft.com)  
- Delegated permission: `Mail.Send`
- **Allow public client flows**: Enabled



## 🗂 Excel File Format

The Excel file must contain the following columns:

| Column Name | Description |
|--------------|-------------|
| Email        | Student's email address |
| ФИО          | Full name of the student |
| Дисциплина   | Course / Discipline |
| Факультет    | Faculty |
| Уровень      | Academic level (Bachelor / Master) |

> Example file: `student_debt.xlsx`

---

Authentication

The script uses **Microsoft Graph API** with Device Code Flow.  
When launched, you will see a message like:

```
Go to https://microsoft.com/devicelogin
and enter the code: ABCD-EFGH
```

Open that link, sign in with your **account**,  
and grant permission for the app to send emails.

---

## How It Works

1. Reads the Excel file into a Pandas DataFrame  
2. Displays the full student table  
3. Asks which recipients to target (all, one, by course, or level)  
4. Authenticates via Microsoft Graph (OAuth 2.0 Device Code Flow)  
5. Builds a personalized HTML message for each student  
6. Sends messages via Graph API endpoint:  
   `https://graph.microsoft.com/v1.0/me/sendMail`
7. Displays a summary of successful and failed sends

---

## Example Usage

```bash
python script.py
```

You will be asked:

```
Who do you want to send emails to?
1 - All students
2 - Specific student by email
3 - Students by discipline
4 - Students by academic level
```

Then confirm sending:  
```
Continue sending? (y/n)
```

---

## 📄 Example Output

```
Go to https://microsoft.com/devicelogin and enter code: J7N8W5
Sent: 222035@astanait.edu.kz
Sent: 221076@astanait.edu.kz
Error: invalid email format
===== SUMMARY =====
Success: 12
Failed: 1
```

---

## ⚙️ Configuration

In `script.py`, you can edit:

```python
CLIENT_ID = "YOUR_CLIENT_ID"
TENANT_ID = "organizations"
EXCEL_PATH = r"C:\Users\YourName\Downloads\student_debt.xlsx"
```

Optional parameters:
```python
SEND_DELAY_SEC = 0.4  # delay between messages
DRY_RUN = False       # True = test mode (no sending)
```

---

## Dependencies

- `pandas`
- `msal`
- `requests`
- `openpyxl`

You can install them all with:
```bash
pip install -r requirements.txt
```

---

## Security Notes

- The script does **not** store any passwords.  
- Authentication happens via official Microsoft login.  
- Only the `Mail.Send` permission is required.  
- Works in compliance with Microsoft 365 security policies.

---



