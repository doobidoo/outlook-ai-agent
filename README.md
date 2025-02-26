# Outlook AI Agent

A sophisticated AI-powered assistant that automatically processes, analyzes, and organizes your Outlook emails using artificial intelligence.

![Outlook AI Agent Banner](https://cdn-icons-png.flaticon.com/512/732/732223.png)

## 🌟 Features

- **Intelligent Email Analysis**: Automatically summarizes, classifies, and prioritizes your emails
- **Sentiment Analysis**: Detects emotional tone in messages to prioritize urgent or negative communications
- **Context-Aware Processing**: Recognizes team and customer emails for appropriate handling
- **Folder Organization**: Automatically sorts emails into appropriate folders based on content
- **Calendar Integration**: Checks for conflicts with meeting invitations
- **Customizable Priority System**: Applies configurable rules to determine email importance
- **Auto-Response Drafting**: Creates context-appropriate draft replies for common email types
- **Statistical Insights**: Tracks email patterns and provides usage analytics
- **Real-Time Processing**: Can process emails as they arrive using Outlook events
- **Complete Folder Scanning**: Recursively scans all Outlook folders, not just the inbox

## 📋 Requirements

- Windows operating system (tested on Windows 10/11)
- Microsoft Outlook (Desktop version) installed and configured
- Python 3.8+ (3.10+ recommended)
- Administrator rights (for installing services)

## 🔧 Installation

### 1. Clone the Repository

```bash
git clone https://github.com/yourusername/outlook-ai-agent.git
cd outlook-ai-agent
```

Or simply download and extract to a location of your choice.

### 2. Create a Virtual Environment (Recommended)

```bash
python -m venv venv
venv\Scripts\activate
```

### 3. Install Required Dependencies

```bash
pip install -r requirements.txt
```

If you don't have a requirements.txt file, install the following packages:

```bash
pip install pywin32 transformers torch pandas tqdm email_reply_parser win10toast
```

Note: If you want to use GPU acceleration for AI models, install the CUDA-compatible version of PyTorch.

## ⚙️ Configuration

The application uses two main configuration files:

### 1. `config.ini`

The main configuration file is created automatically on first run with default settings. You can modify it to customize behavior:

```ini
[AI]
model = facebook/bart-large-cnn
use_llama = False
llama_path = 
max_summary_length = 100
min_summary_length = 30
use_gpu = True
lazy_loading = True

[OUTLOOK]
process_unread_only = True
max_emails_per_run = 50
body_max_length = 2000
folders = Urgent,Processed,Meetings,Newsletters,Questions,Reports
archive_processed = True
scan_all_folders = True
enable_event_handler = True
skip_folders = Deleted Items,Junk Email,Outbox,Sent Items
customer_priority_boost = 2
team_priority_boost = 3

[ANALYTICS]
track_statistics = True
stats_file = email_stats.csv
max_history_days = 30
```

### 2. `user_preferences.json`

This file contains your personal preferences and is also created automatically with default values:

```json
{
  "meeting_response_template": "Thanks for the invitation. I'll check my calendar and get back to you soon.",
  "priority_contacts": ["boss@company.com", "client@bigclient.com"],
  "auto_reply_enabled": true,
  "signature": "\n\nBest regards,\nYour Name",
  "working_hours": {"start": "09:00", "end": "17:00"},
  "vacation_mode": false,
  "time_zone": "Europe/Berlin"
}
```

## 🚀 Usage

### Basic Usage

To run the agent interactively:

```bash
python outlook_ai.py
```

This will:
1. Process all unread emails across your Outlook folders
2. Apply AI analysis to categorize and prioritize messages
3. Move emails to appropriate folders based on content
4. Create draft replies for certain email types
5. Provide a summary of actions taken

### Running as a Windows Service

For continuous background operation, you can install the script as a Windows service using NSSM (Non-Sucking Service Manager):

1. [Download NSSM](https://nssm.cc/download)
2. Install the service (run as Administrator):

```bash
nssm install "OutlookAIAgent" "C:\Path\To\Python.exe" "C:\REPOSITORIES\outlook-ai-agent\outlook_ai.py"
nssm set "OutlookAIAgent" AppDirectory "C:\REPOSITORIES\outlook-ai-agent\"
nssm set "OutlookAIAgent" DisplayName "Outlook AI Agent"
nssm set "OutlookAIAgent" Description "AI-powered Outlook email processor"
nssm set "OutlookAIAgent" Start SERVICE_AUTO_START
nssm start "OutlookAIAgent"
```

### Using Scheduled Tasks

Alternatively, you can set up a Windows Scheduled Task:

1. Open Task Scheduler
2. Create a new task with the following settings:
   - Run with highest privileges
   - Configure for: Windows 10/11
   - Trigger: At startup + Daily
   - Action: Start a program
     - Program/script: `C:\Path\To\Python.exe`
     - Arguments: `C:\REPOSITORIES\outlook-ai-agent\outlook_ai.py`
     - Start in: `C:\REPOSITORIES\outlook-ai-agent\`
   - Conditions: None (or as appropriate for your system)
   - Settings: Allow task to run on demand

## 📂 Folder Structure Recognition

The agent automatically detects special folder structures in your Outlook:

- **CUSTOMER folders**: These are treated with increased priority for customer-specific emails
- **TEAM folders**: These are given special handling for internal team communications

The agent can recognize these patterns in your folder hierarchy and apply appropriate processing rules.

## 🔄 Email Processing Rules

The agent categorizes emails into the following types:

1. **Urgent**: High-priority emails requiring immediate attention
2. **Meeting**: Meeting invitations and scheduling-related emails
3. **Question**: Emails containing questions that need responses
4. **Report**: Status updates and reporting information
5. **Newsletter**: Non-urgent informational emails
6. **Financial**: Invoices, payments, and financial matters
7. **General**: Default category for other emails

Each category has specific handling rules and folder destinations.

## 📊 Analytics

The agent collects statistics on processed emails, which are saved to `email_stats.csv` by default. This data can be analyzed to understand your email patterns and the agent's effectiveness.

You can view insights by running:

```bash
python -c "from outlook_ai import OutlookAIAgent; agent = OutlookAIAgent(); print(agent.generate_email_insights())"
```

## 🔧 Advanced Configuration

### Using External AI APIs

To use external AI APIs instead of local models, modify the `config.ini` file:

```ini
[AI]
use_external_api = True
api_provider = openai  # Options: openai, azure, anthropic
api_key = your_api_key_here
summarization_model = gpt-3.5-turbo
# Other settings...
```

This requires installing the appropriate API client:

```bash
pip install openai  # For OpenAI
# or
pip install anthropic  # For Anthropic Claude
```

### Customizing Email Templates

Edit the `user_preferences.json` file to customize auto-response templates.

## 🛠️ Troubleshooting

### Common Issues

1. **"Failed to connect to Outlook"**
   - Ensure Outlook is running and accessible
   - Verify you have the correct version of pywin32 installed
   - Try running as Administrator

2. **"Error loading AI models"**
   - Check internet connection (needed for first-time model download)
   - Ensure you have sufficient disk space
   - Try setting `lazy_loading = True` in config.ini

3. **"Event handler failed"**
   - This often happens due to COM issues with Outlook
   - Try setting `enable_event_handler = False` and use scheduled runs instead

4. **Performance Issues**
   - If the agent is slow, consider using external APIs
   - Set `max_emails_per_run` to a lower value
   - Enable GPU acceleration if available

### Log File

Check `outlook_ai_agent.log` for detailed error messages and operation history.

## 📚 Architecture

The Outlook AI Agent is built with a modular architecture:

- **Outlook Connectivity**: Uses Windows COM automation via pywin32
- **AI Processing**: Leverages transformer models for NLP tasks
- **Event Handling**: Real-time email processing using Outlook events
- **Analytics Engine**: Tracks and analyzes email statistics
- **Folder Management**: Intelligent organization of emails

## 🛡️ Security Considerations

- The agent runs locally on your machine, and no email data is sent to external servers (unless using external AI APIs)
- API keys are stored in plain text in config.ini - secure this file appropriately
- The agent requires access to your Outlook data; review the code if security is a concern

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📜 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 📬 Contact

For questions or support, please open an issue on this repository.

---

*"Bringing AI intelligence to your inbox, one email at a time."*
