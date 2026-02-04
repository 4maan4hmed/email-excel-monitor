# Email Excel Monitor with Telegram Notifications

A cloud-hosted, automated system that monitors your email inbox for Excel attachments, scans them for specific names or keywords, and sends instant Telegram notifications when matches are found. Runs completely free on GitHub Actions with no server required.

## Features

- **Automated Monitoring**: Checks your inbox every 5 minutes automatically
- **Excel Scanning**: Scans all sheets in `.xlsx`, `.xls`, and `.xlsm` files
- **Smart Search**: Finds your name/keyword in any cell across all worksheets
- **Telegram Alerts**: Instant notifications with email details and match locations
- **Secure**: All credentials stored as encrypted GitHub Secrets
- **Cloud-Hosted**: Runs on GitHub Actions with no server needed
- **Free**: Uses GitHub's free tier (2,000 minutes/month)

## Quick Start Guide

### Step 1: Create Telegram Bot

1. Open Telegram and search for `@BotFather`
2. Send `/newbot` and follow the prompts
3. Copy the **Bot Token** (looks like `110201543:AAHdqTcvCH1vGWJxfSeofSAs0K5PALDsaw`)
4. Start a chat with your new bot by sending any message
5. Get your Chat ID:
   - Visit: `https://api.telegram.org/bot<YOUR_BOT_TOKEN>/getUpdates`
   - Find your `chat_id` in the JSON response (a number like `123456789`)

### Step 2: Set Up Email App Password

**For Gmail:**
1. Enable 2-Factor Authentication on your Google Account
2. Go to [Google App Passwords](https://myaccount.google.com/apppasswords)
3. Create a new app password for "Mail"
4. Copy the 16-character password

**For Outlook/Hotmail:**
1. Go to [Microsoft Account Security](https://account.microsoft.com/security)
2. Enable 2-step verification
3. Create an app password
4. IMAP Server: `imap-mail.outlook.com`

**For Yahoo:**
1. Go to Yahoo Account Security
2. Generate app password
3. IMAP Server: `imap.mail.yahoo.com`

### Step 3: Create GitHub Repository

1. Create a new **private repository** on GitHub
2. Click "Add file" → "Upload files"
3. Upload these three files from this project:
   - `email_monitor.py`
   - `requirements.txt`
   - `.github/workflows/monitor.yml`

Alternatively, use Git:
```bash
git clone https://github.com/4maan4hmed/email-excel-monitor
cd email-excel-monitor
# Copy the three files here
git add .
git commit -m "Initial setup"
git push
```

### Step 4: Configure GitHub Secrets

1. Go to your repository on GitHub
2. Click **Settings** → **Secrets and variables** → **Actions**
3. Click **New repository secret** and add each of these:

| Secret Name | Value | Example |
|-------------|-------|---------|
| `EMAIL_ADDRESS` | Your email address | `your.email@gmail.com` |
| `EMAIL_APP_PASSWORD` | App password from Step 2 | `abcd efgh ijkl mnop` |
| `SEARCH_NAME` | Name/keyword to search for | `John Smith (My neo ID in my case)` |
| `TELEGRAM_BOT_TOKEN` | Bot token from Step 1 | `110201543:AAHdqTcvCH...` |
| `TELEGRAM_CHAT_ID` | Your chat ID from Step 1 | `123456789` |
| `IMAP_SERVER` | IMAP server | `imap.gmail.com` |

**Important Notes:**
- For Gmail, use `imap.gmail.com`
- For Outlook, use `imap-mail.outlook.com`
- For Yahoo, use `imap.mail.yahoo.com`
- The search is case-insensitive

### Step 5: Enable GitHub Actions

1. Go to the **Actions** tab in your repository
2. Click "I understand my workflows, go ahead and enable them"
3. You should see the "Email Excel Monitor" workflow

### Step 6: Test the Setup

1. In the **Actions** tab, click on "Email Excel Monitor"
2. Click **Run workflow** → **Run workflow** (green button)
3. Wait for completion (typically 20-30 seconds)
4. Click on the workflow run to view logs
5. Verify successful connection

**Test with a real email:**
1. Send yourself an email with an Excel attachment containing your search name
2. Wait up to 5 minutes for the next scheduled run
3. You should receive a Telegram notification

## How It Works

```
Every 5 minutes:
  ├─ GitHub Actions triggers the workflow
  ├─ Python environment is set up
  ├─ Dependencies are installed
  ├─ Script connects to your email via IMAP
  ├─ Checks for new emails within the time window
  ├─ Downloads Excel attachments (.xlsx, .xls, .xlsm)
  ├─ Scans every sheet and cell for your name
  ├─ If match found:
  │  ├─ Sends Telegram notification with:
  │  │  ├─ Email subject and sender
  │  │  ├─ Attachment filename
  │  │  ├─ Sheet name and cell location
  │  │  └─ Preview of matched content
  │  └─ Marks email as processed
  └─ Workflow completes and waits for next run
```

## Schedule Configuration

The workflow runs on two schedules:

- **Every 5 minutes** from 7:00 AM to 11:00 PM
  - Checks emails from the last 90 minutes
- **Once at 6:00 AM**
  - Checks emails from the last 12 hours (catches overnight emails)

This ensures no emails are missed while minimizing GitHub Actions usage.

## Customization

### Change Scan Frequency

Edit `.github/workflows/monitor.yml`:

```yaml
# Every 10 minutes
- cron: '*/10 * * * *'

# Every hour
- cron: '0 * * * *'

# Every 30 minutes
- cron: '*/30 * * * *'

# Business hours only (9 AM - 5 PM, Mon-Fri)
- cron: '*/5 9-17 * * 1-5'
```

### Search Multiple Keywords

Update the Python script to search for multiple terms:

```python
search_names = os.getenv('SEARCH_NAME').split(',')
for name in search_names:
    matches = scan_excel_for_name(file_data, name.strip(), filename)
```

Then set `SEARCH_NAME` secret as: `John Smith,Jane Doe,Company Inc`

### Add Email Filters

Only scan emails from specific senders:

```python
# In process_emails function, add after getting from_email:
allowed_senders = ['boss@company.com', 'hr@company.com']
if not any(sender in from_email for sender in allowed_senders):
    continue
```

## Monitoring and Logs

### View Execution History
1. Go to **Actions** tab
2. Click on "Email Excel Monitor"
3. View all runs with status and timestamps

### View Detailed Logs
1. Click on any workflow run
2. Click on "monitor-emails" job
3. Expand steps to see detailed output

### Debug Issues
- Verify all secrets are set correctly
- Confirm app password is valid
- Ensure IMAP is enabled for your email provider
- Check Telegram bot token and chat ID
- Review workflow logs for error messages

## Security Best Practices

**Recommended:**
- Use a private repository
- Use app-specific passwords (never your main password)
- Regularly rotate your app passwords
- Review GitHub Actions logs periodically
- Consider using a dedicated email account

**Avoid:**
- Hardcoding credentials in the code
- Sharing your repository publicly with secrets configured
- Using your main email password
- Committing sensitive data to version control

## Limitations and Considerations

- **Polling Delay**: 5-minute intervals mean notifications are not instant (trade-off for free hosting)
- **GitHub Actions Limits**: 
  - Free tier: 2,000 minutes/month
  - This setup uses approximately 8-10 minutes/day (about 300 minutes/month)
- **Email Providers**: Some providers may rate-limit IMAP connections
- **Large Files**: Very large Excel files (over 10MB) may slow processing
- **Processing Behavior**: Only processes emails within the configured time window to avoid duplicates

## Troubleshooting

### "Login Failed" Error
- Verify app password is correct
- Ensure 2FA is enabled on your email account
- Check that IMAP is enabled in email settings
- Try regenerating the app password

### "No module named 'openpyxl'" Error
- Check that `requirements.txt` is in the repository root
- Verify the workflow file includes the correct pip install command

### No Telegram Notifications
- Verify bot token and chat ID are correct
- Ensure you have sent at least one message to the bot
- Check that the Telegram bot is not blocked

### Workflow Not Running
- Verify Actions are enabled in repository settings
- Check cron syntax in the workflow file
- Note that GitHub may delay scheduled runs by a few minutes

## Usage Statistics

Each workflow run:
- **Duration**: Approximately 20-30 seconds
- **Compute**: About 0.5 GitHub Actions minutes
- **Frequency**: Up to 288 runs/day (every 5 minutes during active hours)
- **Monthly**: Approximately 144 minutes (about 7% of free tier)

## Use Cases

- **Job Applications**: Receive notifications when your name appears in applicant spreadsheets
- **Business**: Monitor when you are mentioned in partner or vendor lists
- **Education**: Track when grades or assignments are posted
- **Sales**: Get alerts when your name appears in commission sheets
- **Contests**: Know immediately when results are published

## Contributing

Contributions are welcome. Feel free to fork this repository and customize it for your needs.

## License

MIT License - Free to use and modify as needed.

---

**Last Updated: February 2026**
