# 📧 Email Excel Monitor with Telegram Notifications

A cloud-hosted, automated system that monitors your email inbox for Excel attachments, scans them for specific names or keywords, and sends instant Telegram notifications when matches are found. Runs completely free on GitHub Actions with no server required!

## 🌟 Features

- ✅ **Automated Monitoring**: Checks your inbox every 5 minutes automatically
- 📊 **Excel Scanning**: Scans all sheets in `.xlsx`, `.xls`, and `.xlsm` files
- 🔍 **Smart Search**: Finds your name/keyword in any cell across all worksheets
- 📱 **Telegram Alerts**: Instant notifications with email details and match locations
- 🔒 **Secure**: All credentials stored as encrypted GitHub Secrets
- ☁️ **Cloud-Hosted**: Runs on GitHub Actions - no server needed
- 🆓 **Completely Free**: Uses GitHub's free tier (2,000 minutes/month)

## 🚀 Quick Start Guide

### Step 1: Create Telegram Bot

1. Open Telegram and search for `@BotFather`
2. Send `/newbot` and follow the prompts
3. Copy the **Bot Token** (looks like `110201543:AAHdqTcvCH1vGWJxfSeofSAs0K5PALDsaw`)
4. Start a chat with your new bot (send any message)
5. Get your Chat ID:
   - Visit: `https://api.telegram.org/bot<YOUR_BOT_TOKEN>/getUpdates`
   - Find your `chat_id` in the JSON response (it's a number like `123456789`)

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

1. **Create a new private repository** on GitHub
2. Click "Add file" → "Upload files"
3. Upload these three files from this project:
   - `email_monitor.py`
   - `requirements.txt`
   - `.github/workflows/monitor.yml`

Or use Git:
```bash
git clone <your-repo-url>
cd <your-repo-name>
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
| `SEARCH_NAME` | Name/keyword to search for | `John Smith` |
| `TELEGRAM_BOT_TOKEN` | Bot token from Step 1 | `110201543:AAHdqTcvCH...` |
| `TELEGRAM_CHAT_ID` | Your chat ID from Step 1 | `123456789` |
| `IMAP_SERVER` | (Optional) IMAP server | `imap.gmail.com` |

**Important Notes:**
- For Gmail, use `imap.gmail.com` (default)
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
3. Wait for it to complete (should take 20-30 seconds)
4. Click on the workflow run to see logs
5. Verify it connected successfully

**Test with a real email:**
1. Send yourself an email with an Excel attachment containing your search name
2. Wait up to 5 minutes for the next scheduled run
3. You should receive a Telegram notification!

## 📋 How It Works

```
Every 5 minutes:
  ├─ GitHub Actions triggers the workflow
  ├─ Python environment is set up
  ├─ Dependencies are installed
  ├─ Script connects to your email via IMAP
  ├─ Checks for UNREAD emails only
  ├─ Downloads Excel attachments (.xlsx, .xls, .xlsm)
  ├─ Scans every sheet and cell for your name
  ├─ If match found:
  │  ├─ Sends Telegram notification with:
  │  │  ├─ Email subject and sender
  │  │  ├─ Attachment filename
  │  │  ├─ Sheet name and cell location
  │  │  └─ Preview of matched content
  │  └─ Marks email as read
  └─ Workflow completes and waits for next run
```

## 🔧 Customization

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

## 📊 Monitoring & Logs

### View Execution History
1. Go to **Actions** tab
2. Click on "Email Excel Monitor"
3. See all runs with status and timestamps

### View Detailed Logs
1. Click on any workflow run
2. Click on "monitor-emails" job
3. Expand steps to see detailed output

### Debug Issues
- Check if secrets are set correctly
- Verify app password is valid
- Ensure IMAP is enabled for your email
- Check Telegram bot token and chat ID
- Review workflow logs for error messages

## 🔒 Security Best Practices

✅ **DO:**
- Use a private repository
- Use app-specific passwords (never your main password)
- Regularly rotate your app passwords
- Review GitHub Actions logs periodically
- Use a dedicated email if possible

❌ **DON'T:**
- Hardcode any credentials in the code
- Share your repository publicly with secrets
- Use your main email password
- Commit sensitive data to Git

## ⚠️ Limitations & Considerations

- **Polling Delay**: 5-minute intervals mean notifications aren't instant (trade-off for free hosting)
- **GitHub Actions Limits**: 
  - Free tier: 2,000 minutes/month
  - This setup uses ~8-10 minutes/day = ~300 minutes/month
- **Email Providers**: Some providers may rate-limit IMAP connections
- **Large Files**: Very large Excel files (>10MB) may slow processing
- **Unread Only**: Only processes unread emails (won't re-scan read emails)

## 🆘 Troubleshooting

### "Login Failed" Error
- Verify app password is correct
- Ensure 2FA is enabled on your email
- Check IMAP is enabled in email settings
- Try regenerating the app password

### "No module named 'openpyxl'" Error
- Check requirements.txt is in repository root
- Verify workflow file has correct pip install command

### No Telegram Notifications
- Verify bot token and chat ID are correct
- Make sure you've sent at least one message to the bot
- Check Telegram bot is not blocked

### Workflow Not Running
- Verify Actions are enabled in repository settings
- Check cron syntax in workflow file
- Remember: GitHub may delay scheduled runs by a few minutes

## 📈 Usage Statistics

Each workflow run:
- **Duration**: ~20-30 seconds
- **Compute**: ~0.5 GitHub Actions minutes
- **Frequency**: 288 runs/day (every 5 minutes)
- **Monthly**: ~144 minutes (~7% of free tier)

## 🤝 Contributing

Feel free to fork this repository and customize it for your needs!

## 📄 License

MIT License - feel free to use and modify as needed.

## 🎯 Use Cases

- 📊 **Job Applications**: Get notified when your name appears in applicant spreadsheets
- 💼 **Business**: Monitor when you're mentioned in partner/vendor lists
- 🏫 **Education**: Track when grades or assignments are posted
- 📈 **Sales**: Get alerts when your name appears in commission sheets
- 🏆 **Contests**: Know immediately when results are published

---

**Made with ❤️ for automated email monitoring**

*Last Updated: February 2026*
