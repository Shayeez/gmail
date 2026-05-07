# Gmail Extraction Tool

This tool automatically extracts emails from multiple Gmail accounts, categorizes them, and exports them into highly readable Excel files with priority highlighting. It supports both local execution and automated remote server execution via Google Drive syncing.

---

## 1. Project Requirements & Features

**Core Features:**
- **Multi-Account Support**: Authenticate and extract emails from multiple accounts sequentially.
- **Data Extraction**: Grabs To, CC, BCC, Subject, Date, Link, and a 5-line Summary of the body.
- **Categorization & Highlighting**: Groups emails into "Important" vs "Other" based on Inbox rules. Highlights critical rows automatically using rules defined in `priorities.json`.
- **Remote Execution**: Includes a `server_runner.py` script that downloads code from Google Drive, executes the extraction locally on the server, and uploads the results back to Drive automatically without token conflicts.
- **Secure Token Management**: Safely stores OAuth tokens in a dedicated `tokens/` directory.

**Technical Details:**
- Language: Python 3.x
- APIs: Gmail API, Google Drive API (`google-api-python-client`)
- Output format: `.xlsx` files organized by Account, Category, and Date.

---

## 2. Installation & Setup

The easiest way to get started is by using the automated setup script. This creates an isolated Python environment and installs all dependencies listed in `requirements.txt`.

### Windows Users
1. Double-click the `setup.bat` file, or run it from your terminal.
2. Once complete, activate your environment:
   ```cmd
   .venv\Scripts\activate
   ```

### Mac/Linux Users
1. Open your terminal in this folder and run:
   ```bash
   chmod +x setup.sh
   ./setup.sh
   source .venv/bin/activate
   ```

---

## 3. Prerequisites: `credentials.json`

Before running the script, you must have a `credentials.json` file. This acts as the "ID Card" for your application to talk to Google's servers. 

> [!IMPORTANT]
> You only need **one** `credentials.json` file for the entire application, no matter how many email accounts you extract from!

**How to get `credentials.json`:**
1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new Project.
3. Navigate to **APIs & Services > Library**, search for **Gmail API** and **Google Drive API**, and click **Enable** for both.
4. Navigate to **APIs & Services > OAuth consent screen**. Add `https://www.googleapis.com/auth/gmail.readonly` and `https://www.googleapis.com/auth/drive` to the scopes.
5. Navigate to **APIs & Services > Credentials** > **Create Credentials** > **OAuth client ID**.
6. Choose **Desktop app** and click Create.
7. Click **Download JSON**, rename it to `credentials.json`, and place it in the same folder as `mail.py`.

---

## 4. Local Run Methodology (`mail.py`)

When running on your local computer, use `mail.py` directly.

### Managing Accounts
To authorize a new email account, run:
```bash
python mail.py --add-account your_email@gmail.com
```
This opens your browser to log in. The token is saved in `tokens/your_email@gmail.com.json`.
To remove an account:
```bash
python mail.py --remove-account your_email@gmail.com
```

### Running Extraction Locally
You can control the script with the following commands:

| Command | Description |
|---|---|
| `python mail.py --run-all` | Runs email extraction for all authorized accounts |
| `python mail.py --run-all --today` | Extracts only today's emails |
| `python mail.py --run-all --after 2024/01/01` | Extracts emails received after a specific date |
| `python mail.py --run-all --before 2024/02/01` | Extracts emails received before a specific date |
| `python mail.py --run-all --max 50` | Limits extraction to 50 emails per account |
| `python mail.py --run-all --token-dir ./groupA` | Uses a custom token directory |

---

## 5. Server Run Methodology (`server_runner.py`)

When deploying to a remote server, we use `server_runner.py`. The server does NOT need Google Drive mounted locally.

### How it works:
1. `server_runner.py` uses the Google Drive API to download the latest `mail.py`, `priorities.json`, `credentials.json`, and your `tokens/` from the cloud.
2. It executes `python mail.py --run-all --today` locally on the server.
3. It uploads the newly generated Excel outputs and logs directly back to Google Drive.
4. It purposely *does not* upload tokens to prevent conflict with your local machine.

### Server Setup:
1. **On your local machine**, run:
   ```cmd
   python server_runner.py
   ```
   This will open your browser and generate a `drive_token.json` file.
2. Copy `server_runner.py`, `credentials.json`, and `drive_token.json` to your remote server.
3. Setup a Cron Job on the server to run the script automatically (see below).

#### Setting up a Cron Job (Linux/Mac)
To run the extraction twice a day (e.g., at 8:00 AM and 8:00 PM), open your server terminal and type `crontab -e`. Then, add this line at the bottom:
```bash
0 8,20 * * * cd /path/to/your/GmailFolder && source .venv/bin/activate && python server_runner.py >> cron_audit.log 2>&1
```
*(Replace `/path/to/your/GmailFolder` with the actual path on your server).*

---

## 6. Priorities & Output Format

The emails are saved into `output/` with filenames like `spzohous@gmail.com_Important_20260506.xlsx`. 
The newest emails are always at the top.

### `priorities.json`
If you want to automatically highlight specific emails (bold font, yellow background), add rules to `priorities.json`:
```json
[
  {"From": "boss@company.com"},
  {"Subject": "URGENT"}
]
```
