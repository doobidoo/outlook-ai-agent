import win32com.client
import pythoncom
import os
import re
import datetime
import logging
import configparser
import json
from pathlib import Path
import time
import threading

# Optional dependencies with fallbacks
try:
    from tqdm import tqdm
    has_tqdm = True
except ImportError:
    has_tqdm = False
    print("Install tqdm for progress bars: pip install tqdm")

try:
    import pandas as pd
    has_pandas = True
except ImportError:
    has_pandas = False
    print("Install pandas for analytics: pip install pandas")

try:
    from win10toast import ToastNotifier
    has_notifications = True
except ImportError:
    has_notifications = False
    print("Install win10toast for desktop notifications (Windows only): pip install win10toast")

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='outlook_ai_agent.log',
    filemode='a'
)
logger = logging.getLogger('outlook_ai_agent')

# Add console handler for immediate feedback
console = logging.StreamHandler()
console.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
console.setFormatter(formatter)
logger.addHandler(console)

class OutlookMailEventHandler:
    def __init__(self, ai_agent):
        self.ai_agent = ai_agent
        logger.info("Event handler initialized")
    
    def OnNewMailEx(self, item_id):
        """Triggered when new mail arrives"""
        try:
            # Get the new email using the EntryID
            message = self.ai_agent.namespace.GetItemFromID(item_id)
            
            # Process the email immediately, before rules move it
            logger.info(f"New email received: {message.Subject}")
            self.ai_agent._process_single_email(message)
            
        except Exception as e:
            logger.error(f"Error processing new email: {e}")

class OutlookAIAgent:
    def __init__(self, config_path="config.ini"):
        # Load configuration
        self.config = self._load_config(config_path)
        
        # Connect to Outlook
        self._connect_outlook()
        
        # Initialize AI models (lazy-loaded when needed)
        self.models = {}
        
        # Load user preferences
        self.user_preferences = self._load_user_preferences()
        
        # Create statistics tracker
        self.stats = self._init_stats()
        
        # Dictionary to track folder types (team, customer, etc.)
        self.folder_types = {}
        self._analyze_folder_structure()
    
    def _init_stats(self):
        """Initialize or reset statistics"""
        return {
            "emails_processed": 0,
            "urgent_emails": 0,
            "meeting_emails": 0,
            "auto_replies_drafted": 0,
            "start_time": datetime.datetime.now(),
            "by_folder": {}
        }
    
    def _connect_outlook(self):
        """Connect or reconnect to Outlook"""
        try:
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook.GetNamespace("MAPI")
            self.inbox = self.namespace.GetDefaultFolder(6)  # 6 = Inbox
            logger.info("Connected to Outlook")
        except Exception as e:
            logger.error(f"Failed to connect to Outlook: {e}")
            logger.error("Please ensure that Outlook is running and accessible.")
            logger.error("Check if the win32com.client package is installed correctly.")
            raise e
    
    def _reconnect_outlook_if_needed(self):
        """Attempt to reconnect to Outlook if the connection is lost"""
        try:
            # Test the connection
            _ = self.inbox.Items.Count
            return True
        except Exception:
            logger.warning("Outlook connection lost, attempting to reconnect...")
            try:
                self._connect_outlook()
                return True
            except Exception as e:
                logger.error(f"Failed to reconnect to Outlook: {e}")
                return False
    
    def _load_config(self, config_path):
        """Load configuration from ini file or create default"""
        config = configparser.ConfigParser()
        
        if not os.path.exists(config_path):
            # Create default config
            config['AI'] = {
                'model': 'facebook/bart-large-cnn',
                'use_llama': 'False',
                'llama_path': '',
                'max_summary_length': '100',
                'min_summary_length': '30',
                'use_gpu': 'True',
                'lazy_loading': 'True'
            }
            config['OUTLOOK'] = {
                'process_unread_only': 'True',
                'max_emails_per_run': '50',
                'body_max_length': '2000',
                'folders': 'Urgent,Processed,Meetings,Newsletters,Questions,Reports',
                'archive_processed': 'True',
                'scan_all_folders': 'True',
                'enable_event_handler': 'True',
                'skip_folders': 'Deleted Items,Junk Email,Outbox,Sent Items,Calendar,Tasks,Aufgaben,Kalender,Contacts,Deleted Items,Junk Email,Outbox,Sent Items,Gelöschte Elemente,Postausgang,Gesendete Elemente,Kontakte,Recipient Cache,{06967759-274D-40B2-A3EB-D7F9E73727D7},{A9E2BC46-B3A0-4243-B315-60D991004455},Firmen,Organizational Contacts,GAL Contacts,PeopleCentricConversation Buddies,ExternalContacts,Journal,PersonMetadata,Notizen,Synchronisierungsprobleme,Konflikte,Lokale Fehler',
                'customer_priority_boost': '2',
                'team_priority_boost': '3'
            }
            config['ANALYTICS'] = {
                'track_statistics': 'True',
                'stats_file': 'email_stats.csv',
                'max_history_days': '30'
            }
            
            with open(config_path, 'w') as f:
                config.write(f)
            logger.info(f"Created default configuration at {config_path}")
        else:
            config.read(config_path)
        
        return config
    
    def _analyze_folder_structure(self):
        """Analyze folder structure to identify customer and team folders"""
        try:
            # Get root account folder
            root_email = self.inbox.Parent.Name
            logger.info(f"Analyzing folder structure for {root_email}")
            
            # Find Customer and TEAM folders
            for folder in self.inbox.Parent.Folders:
                if folder.Name.upper() == "CUSTOMER":
                    self._categorize_folder_tree(folder, "customer")
                elif folder.Name.upper() == "TEAM":
                    self._categorize_folder_tree(folder, "team")
            
            logger.info(f"Folder analysis complete. Found {len(self.folder_types)} categorized folders")
        except Exception as e:
            logger.error(f"Error analyzing folder structure: {e}")
    
    def _categorize_folder_tree(self, parent_folder, folder_type, depth=0, max_depth=5):
        """Recursively categorize folders by type"""
        if depth > max_depth:
            return
            
        # Store folder type
        entry_id = parent_folder.EntryID
        self.folder_types[entry_id] = folder_type
        
        # Process subfolders
        for folder in parent_folder.Folders:
            self._categorize_folder_tree(folder, folder_type, depth+1, max_depth)
    
    def _get_folder_type(self, folder):
        """Get the type of a folder (customer, team, etc.)"""
        try:
            return self.folder_types.get(folder.EntryID, "other")
        except:
            return "other"
    
    def _get_model(self, model_type):
        """Lazy-load AI models when needed"""
        if model_type in self.models:
            return self.models[model_type]
        
        # Import here to avoid loading if not used
        try:
            from transformers import pipeline
            import torch
            
            ai_config = self.config['AI']
            device = 0 if torch.cuda.is_available() and ai_config.getboolean('use_gpu', True) else -1
            
            logger.info(f"Loading {model_type} model on device {device}")
            
            if model_type == 'summarizer':
                if ai_config.getboolean('use_llama', False) and ai_config.get('llama_path'):
                    try:
                        model = pipeline("summarization", 
                                         model=ai_config.get('llama_path'),
                                         tokenizer=ai_config.get('llama_path'),
                                         device=device)
                        logger.info("Using LLaMA model for summarization")
                    except Exception as e:
                        logger.error(f"Failed to load LLaMA model: {e}")
                        model = pipeline("summarization", 
                                        model=ai_config.get('model', 'facebook/bart-large-cnn'),
                                        device=device)
                        logger.info(f"Fallback to {ai_config.get('model')} for summarization")
                else:
                    model = pipeline("summarization", 
                                    model=ai_config.get('model', 'facebook/bart-large-cnn'),
                                    device=device)
                    logger.info(f"Using {ai_config.get('model')} for summarization")
            
            elif model_type == 'sentiment':
                model = pipeline("sentiment-analysis", device=device)
                logger.info("Sentiment analyzer loaded")
            
            elif model_type == 'classifier':
                model = pipeline("text-classification", 
                                model="distilbert-base-uncased-finetuned-sst-2-english",
                                return_all_scores=True,
                                device=device)
                logger.info("Email classifier loaded")
            
            else:
                raise ValueError(f"Unknown model type: {model_type}")
            
            self.models[model_type] = model
            return model
            
        except ImportError as e:
            logger.error(f"Failed to import required libraries for AI models: {e}")
            logger.info("Make sure transformers and torch are installed: pip install transformers torch")
            raise
    

    def _load_user_preferences(self):
        """Load user preferences from file or use defaults"""
        prefs_file = "user_preferences.json"
        default_prefs = {
            "meeting_response_template": "Thanks for the invitation. I'll check my calendar and get back to you soon.",
            "priority_contacts": ["boss@company.com", "client@bigclient.com"],
            "auto_reply_enabled": True,
            "signature": "\n\nBest regards,\nHeinrich Krupp",
            "working_hours": {"start": "09:00", "end": "17:00"},
            "vacation_mode": False,
            "time_zone": "Europe/Berlin"
        }
        
        if os.path.exists(prefs_file):
            try:
                with open(prefs_file, 'r') as f:
                    prefs = json.load(f)
                logger.info(f"Loaded user preferences from {prefs_file}")
                return prefs
            except Exception as e:
                logger.error(f"Failed to load user preferences: {e}")
        
        # Save default preferences if file doesn't exist
        try:
            with open(prefs_file, 'w') as f:
                json.dump(default_prefs, f, indent=2)
            logger.info(f"Created default user preferences at {prefs_file}")
        except Exception as e:
            logger.error(f"Failed to save default user preferences: {e}")
        
        return default_prefs
    
    def process_emails(self):
        """Process all unread emails across folders"""
        # Reset statistics for this run
        self.stats = self._init_stats()
        
        # Ensure Outlook connection is active
        if not self._reconnect_outlook_if_needed():
            logger.error("Failed to connect to Outlook, aborting")
            return self.stats
        
        outlook_config = self.config['OUTLOOK']
        
        # Determine if we should scan all folders
        if outlook_config.getboolean('scan_all_folders', True):
            processed_count = self.scan_all_folders()
        else:
            # Just scan inbox
            processed_count = self._process_folder(self.inbox)
        
        # Save statistics
        if self.config['ANALYTICS'].getboolean('track_statistics', True) and has_pandas:
            self._save_statistics()
        
        # Provide a summary of what was done
        self.stats["end_time"] = datetime.datetime.now()
        self.stats["duration"] = (self.stats["end_time"] - self.stats["start_time"]).total_seconds()
        
        logger.info(f"Completed processing. Processed {processed_count} emails in {self.stats['duration']:.1f} seconds")
        return self.stats
    
    def scan_all_folders(self):
        """Scan all folders for unread emails"""
        # Start with the root folder
        root_folder = self.inbox.Parent
        
        # Get list of folders to skip
        skip_folders = self.config['OUTLOOK'].get('skip_folders', 'Deleted Items,Junk Email,Outbox,Sent Items')
        skip_folders = [folder.strip() for folder in skip_folders.split(',')]
        
        processed_count = 0
        
        # Process all folders recursively
        processed_count += self._scan_folder_recursive(root_folder, skip_folders)
        
        logger.info(f"Total emails processed across all folders: {processed_count}")
        return processed_count
    
    def _scan_folder_recursive(self, folder, skip_folders, depth=0, max_depth=10):
        """Recursively scan a folder and all its subfolders for unread emails"""
        if depth > max_depth:
            return 0
            
        processed_count = 0
        
        # Get the destination folders from config
        dest_folders = self.config['OUTLOOK'].get('folders', '').split(',')
        dest_folders = [f.strip() for f in dest_folders if f.strip()]
        
        # Skip specified folders, non-mail folders, or agent-managed folders
        if (folder.Name in skip_folders or 
            not self._is_mail_folder(folder) or 
            folder.Name in dest_folders):
            
            skip_reason = "in skip list"
            if not self._is_mail_folder(folder):
                skip_reason = "not a mail folder"
            elif folder.Name in dest_folders:
                skip_reason = "agent-managed folder"
                
            logger.info(f"Skipping folder: {folder.Name} ({skip_reason})")
            return 0
            
        # Process current folder
        try:
            processed_count += self._process_folder(folder)
        except Exception as e:
            logger.error(f"Error processing folder '{folder.Name}': {e}")
        
        # Recursively process subfolders
        try:
            for subfolder in folder.Folders:
                processed_count += self._scan_folder_recursive(subfolder, skip_folders, depth + 1, max_depth)
        except Exception as e:
            logger.error(f"Error accessing subfolders of '{folder.Name}': {e}")
        
        return processed_count
    
    def _process_folder(self, folder):
        """Process all unread emails in a specific folder"""
        outlook_config = self.config['OUTLOOK']
        processed_count = 0
        
        try:
            # Get folder name for logging
            folder_name = folder.Name
            folder_type = self._get_folder_type(folder)
            
            # Get unread emails in this folder
            emails = folder.Items.Restrict("[Unread] = true")
            
            # Sort by received time
            emails.Sort("[ReceivedTime]", True)  # True = descending
            
            # Count emails
            count = 0
            for _ in emails:
                count += 1
            
            if count == 0:
                return 0
                
            # Limit the number of emails processed per run
            max_emails = outlook_config.getint('max_emails_per_run', 50)
            limited_count = min(count, max_emails)
            
            logger.info(f"Processing {limited_count} unread emails in folder '{folder_name}' (type: {folder_type})")
            
            # Process emails with or without progress bar
            if has_tqdm:
                iterator = tqdm(enumerate(emails), total=limited_count, desc=f"Processing {folder_name}")
            else:
                iterator = enumerate(emails)
            
            for i, message in iterator:
                if i >= max_emails:
                    break
                    
                try:
                    # Process with folder context
                    self._process_single_email(message, folder=folder, folder_type=folder_type)
                    processed_count += 1
                    
                    # Track by folder
                    if folder_name not in self.stats["by_folder"]:
                        self.stats["by_folder"][folder_name] = 0
                    self.stats["by_folder"][folder_name] += 1
                    
                except Exception as e:
                    logger.error(f"Error processing email '{message.Subject}' in folder '{folder_name}': {e}")
                    
                # Reconnect if needed after each batch
                if (i + 1) % 10 == 0:
                    self._reconnect_outlook_if_needed()
            
        except Exception as e:
            logger.error(f"Error accessing folder '{folder.Name}': {e}")
        
        return processed_count
    
    def _process_single_email(self, message, folder=None, folder_type=None):
        """Process a single email message"""
        outlook_config = self.config['OUTLOOK']
        
        # Extract email information
        subject = message.Subject
        sender = message.SenderEmailAddress
        received_time = message.ReceivedTime
        body = message.Body[:outlook_config.getint('body_max_length', 2000)]
        
        folder_name = folder.Name if folder else "Unknown"
        folder_type = folder_type or (self._get_folder_type(folder) if folder else "other")
        
        logger.info(f"Processing email: '{subject}' from {sender} in folder '{folder_name}' ({folder_type})")
        
        # Clean text before processing
        cleaned_body = self._clean_email_text(body)
        
        # Check if we need AI analysis
        needs_ai = True
        
        # Skip AI for newsletters, automated notifications, etc.
        if any(kw in subject.lower() for kw in ['newsletter', 'noreply', 'automated']):
            email_type = "newsletter"
            priority = 1
            summary = "Automated message - no summary needed"
            sentiment = {'label': 'NEUTRAL', 'score': 1.0}
            needs_ai = False
        
        # For other emails, use AI to analyze
        if needs_ai:
            # Generate summary
            summary = self._generate_summary(cleaned_body)
            
            # Analyze sentiment (only first 512 chars for efficiency)
            sentiment = self._get_model('sentiment')(cleaned_body[:512])[0]
            
            # Classify email type/category
            email_type = self._classify_email_type(subject, cleaned_body)
            
            # Determine email priority with folder context
            priority = self._calculate_priority(sender, subject, sentiment, email_type, folder_type)
        
        # Handle email based on type and priority
        self._route_email(message, summary, sentiment, email_type, priority, folder_type)
        
        # Log the processing results
        logger.info(f"Email processed: Type={email_type}, Priority={priority}, Sentiment={sentiment['label']}")
    
    def _generate_summary(self, text):
        """Generate a summary of the email body"""
        ai_config = self.config['AI']
        max_length = ai_config.getint('max_summary_length', 100)
        min_length = ai_config.getint('min_summary_length', 30)
        
        # Skip summarization if text is already short
        if len(text.split()) < min_length:
            return text
        
        try:
            summarizer = self._get_model('summarizer')
            summary = summarizer(text, 
                               max_length=max_length, 
                               min_length=min_length, 
                               do_sample=False)[0]["summary_text"]
            return summary
        except Exception as e:
            logger.error(f"Summarization failed: {e}")
            # Return truncated text as fallback
            return text[:200] + "..."
    
    def _clean_email_text(self, text):
        """Clean email text by removing signatures, reply chains, etc."""
        try:
            # Try using email_reply_parser if available
            from email_reply_parser import EmailReplyParser
            return EmailReplyParser.parse_reply(text)
        except ImportError:
            # Fallback to manual cleaning
            # Remove email signatures
            text = re.sub(r'--+[\s\S]*$', '', text)
            
            # Remove reply headers
            text = re.sub(r'From:.*?Sent:.*?To:.*?Subject:.*?\n', '', text)
            
            # Remove excessive newlines
            text = re.sub(r'\n{3,}', '\n\n', text)
            
            return text.strip()
    
    def _classify_email_type(self, subject, body):
        """Classify the type of email using the subject and body"""
        # Combine subject and first part of body for classification
        text = f"{subject} {body[:200]}"
        
        # Check for common email types
        if re.search(r'\b(meeting|invite|calendar|appointment)\b', text, re.I):
            return "meeting"
        elif re.search(r'\b(urgent|asap|emergency|immediately)\b', text, re.I):
            return "urgent"
        elif re.search(r'\b(question|help|assistance|support)\b', text, re.I):
            return "question"
        elif re.search(r'\b(report|update|status)\b', text, re.I):
            return "report"
        elif re.search(r'\b(newsletter|subscription|weekly|monthly)\b', text, re.I):
            return "newsletter"
        elif re.search(r'\b(invoice|payment|receipt|order)\b', text, re.I):
            return "financial"
        else:
            # Use AI classifier for more nuanced classification
            try:
                result = self._get_model('classifier')(text)
                if result[0][0]['score'] > 0.7:  # High confidence
                    return result[0][0]['label']
            except Exception as e:
                logger.warning(f"Classification failed, using default: {e}")
            
            return "general"
    
    def _calculate_priority(self, sender, subject, sentiment, email_type, folder_type="other"):
        """Calculate priority score for the email with folder context"""
        priority = 0
        
        # Priority based on sender
        if sender in self.user_preferences["priority_contacts"]:
            priority += 5
        
        # Priority based on email type
        type_priority = {
            "urgent": 5,
            "meeting": 4,
            "question": 3,
            "report": 2,
            "financial": 3,
            "general": 1,
            "newsletter": 0
        }
        priority += type_priority.get(email_type, 1)
        
        # Priority based on sentiment
        if sentiment['label'] == 'NEGATIVE' and sentiment['score'] > 0.8:
            priority += 2
        elif sentiment['label'] == 'POSITIVE' and sentiment['score'] > 0.9:
            # High confidence positive emails might deserve attention too
            priority += 1
        
        # Priority based on subject keywords
        if any(word in subject.lower() for word in ["urgent", "important", "asap", "deadline"]):
            priority += 3
        
        # Priority based on folder type
        if folder_type == "customer":
            # Boost priority for customer folders
            customer_boost = self.config['OUTLOOK'].getint('customer_priority_boost', 2)
            priority += customer_boost
            logger.info(f"Applied customer priority boost: +{customer_boost}")
        elif folder_type == "team":
            # Boost priority for team folders
            team_boost = self.config['OUTLOOK'].getint('team_priority_boost', 3)
            priority += team_boost
            logger.info(f"Applied team priority boost: +{team_boost}")
        
        return min(priority, 10)  # Cap at 10
    
    def _route_email(self, message, summary, sentiment, email_type, priority, folder_type="other"):
        """Route email based on its type and priority with folder context"""
        # Handle based on email type
        if email_type == "urgent" or priority >= 8:
            self._handle_urgent_email(message, summary)
            
        elif email_type == "meeting":
            self._handle_meeting_email(message, summary)
            
        elif email_type == "question" and priority >= 4:
            self._handle_question_email(message, summary)
            
        elif email_type == "report":
            self._handle_report_email(message)
            
        # Special handling for team emails
        elif folder_type == "team" and priority >= 6:
            self._handle_team_email(message, summary, priority)
            
        # Special handling for customer emails
        elif folder_type == "customer" and priority >= 5:
            self._handle_customer_email(message, summary, priority)
            
        # Handle based on sentiment
        elif sentiment['label'] == 'NEGATIVE' and sentiment['score'] > 0.7:
            self._handle_negative_email(message, summary, sentiment)
            
        # Handle newsletters and low priority emails
        elif email_type == "newsletter" or priority <= 2:
            self._handle_low_priority_email(message)
        
        # For everything else, handle based on priority
        elif priority >= 5:
            self._handle_high_priority_email(message, summary, sentiment)
    
    def _handle_team_email(self, message, summary, priority):
        """Special handling for team member emails"""
        # Flag team emails that need attention
        message.FlagStatus = 2  # 2 = Flagged
        message.Categories = "Team Priority"
        
        # Set a reminder for team emails
        if priority >= 7:
            # High priority team email - reminder for today
            reminder_time = datetime.datetime.now() + datetime.timedelta(hours=2)
        else:
            # Standard team email - reminder for tomorrow
            tomorrow = datetime.datetime.now() + datetime.timedelta(days=1)
            
            # Set to morning working hours
            work_hours = self.user_preferences["working_hours"]
            hour, minute = map(int, work_hours["start"].split(":"))
            reminder_time = tomorrow.replace(hour=hour, minute=minute)
        
        message.ReminderSet = True
        message.ReminderTime = reminder_time
        
        # Add a note about the summary
        message.Body = f"[AI SUMMARY: {summary}]\n\n" + message.Body
        message.Save()
        
        logger.info(f"Processed team email with priority {priority}")
        # Don't mark as read - team emails may need attention
    
    def _handle_customer_email(self, message, summary, priority):
        """Special handling for customer emails"""
        # Categorize by customer importance
        if priority >= 7:
            message.Categories = "Key Customer"
        else:
            message.Categories = "Customer Inquiry"
        
        # Flag for follow-up
        message.FlagStatus = 2  # 2 = Flagged
        
        # Add reminder for high priority customer emails
        if priority >= 6:
            reminder_time = datetime.datetime.now() + datetime.timedelta(hours=4)
            message.ReminderSet = True
            message.ReminderTime = reminder_time
        
        # Add a note about the summary
        message.Body = f"[AI SUMMARY: {summary}]\n\n" + message.Body
        message.Save()
        
        logger.info(f"Processed customer email with priority {priority}")
        # Don't mark as read - customer emails may need attention
    
    def _handle_urgent_email(self, message, summary):
        """Handle emails marked as urgent"""
        # Move to urgent folder
        urgent_folder = self._get_or_create_folder("Urgent")
        message.Move(urgent_folder)
        
        # Create notification or alert
        if has_notifications:
            toaster = ToastNotifier()
            toaster.show_toast(
                "Urgent Email", 
                f"From: {message.SenderName}\nSubject: {message.Subject}\nSummary: {summary}",
                duration=10
            )
        
        # Draft a quick acknowledgment reply if enabled
        if self.user_preferences["auto_reply_enabled"]:
            reply = message.Reply()
            reply.Body = f"I've received your urgent email and will prioritize it.\n\n{self.user_preferences['signature']}"
            reply.Save()
            self.stats["auto_replies_drafted"] += 1
        
        self.stats["urgent_emails"] += 1
        message.UnRead = False
    
    def _handle_meeting_email(self, message, summary):
        """Handle meeting invitations and related emails"""
        # Move to meetings folder
        meetings_folder = self._get_or_create_folder("Meetings")
        
        # Check if this is a calendar invite
        try:
            if message.MessageClass == "IPM.Schedule.Meeting.Request":
                # Try to get meeting details
                try:
                    meeting_item = message.GetAssociatedAppointment(False)
                    meeting_start = meeting_item.Start
                    meeting_end = meeting_item.End
                    organizer = meeting_item.Organizer
                    location = meeting_item.Location
                    
                    # Check calendar conflicts
                    conflicts = self._check_calendar_conflicts(meeting_start, meeting_end)
                    
                    # Draft appropriate response
                    reply = message.Reply()
                    if conflicts:
                        reply.Body = f"Thank you for the meeting invitation. I notice I have a scheduling conflict:\n\n"
                        for conflict in conflicts:
                            reply.Body += f"- {conflict['subject']} from {conflict['start']} to {conflict['end']}\n"
                        reply.Body += f"\nCould we find an alternative time?{self.user_preferences['signature']}"
                    else:
                        reply.Body = f"{self.user_preferences['meeting_response_template']}{self.user_preferences['signature']}"
                    
                    reply.Save()
                    self.stats["auto_replies_drafted"] += 1
                except Exception as e:
                    logger.warning(f"Failed to process meeting invite details: {e}")
                    # Still move the message
                    message.Move(meetings_folder)
            else:
                # Handle other meeting-related emails
                message.Move(meetings_folder)
                if self.user_preferences["auto_reply_enabled"]:
                    reply = message.Reply()
                    reply.Body = f"Thank you for the information about the meeting. I have noted the details.\n\nSummary: {summary}{self.user_preferences['signature']}"
                    reply.Save()
                    self.stats["auto_replies_drafted"] += 1
        except Exception as e:
            logger.error(f"Error processing meeting email: {e}")
            # Still try to move the message
            try:
                message.Move(meetings_folder)
            except:
                pass
        
        self.stats["meeting_emails"] += 1
        message.UnRead = False
    
    def _handle_question_email(self, message, summary):
        """Handle emails containing questions"""
        # Move to questions folder
        questions_folder = self._get_or_create_folder("Questions")
        message.Move(questions_folder)
        
        # Flag for follow-up
        message.FlagStatus = 2  # 2 = Flagged
        
        # Add follow-up reminder for tomorrow during working hours
        tomorrow = datetime.datetime.now() + datetime.timedelta(days=1)
        
        # Parse working hours
        try:
            work_start = self.user_preferences["working_hours"]["start"].split(":")
            hours, minutes = int(work_start[0]), int(work_start[1])
            
            # Set reminder time to tomorrow at start of working hours
            tomorrow = tomorrow.replace(hour=hours, minute=minutes, second=0, microsecond=0)
            
            message.ReminderSet = True
            message.ReminderTime = tomorrow
            message.Save()
        except Exception as e:
            logger.warning(f"Failed to set reminder: {e}")
        
        message.UnRead = False

    def _handle_report_email(self, message):
        """Handle report emails"""
        # Move to reports folder
        reports_folder = self._get_or_create_folder("Reports")
        message.Move(reports_folder)
        message.UnRead = False

    def _handle_negative_email(self, message, summary, sentiment):
        """Handle emails with negative sentiment"""
        # These might need special attention
        message.FlagStatus = 2  # 2 = Flagged
        message.Categories = "Needs Attention"
        
        # Draft a empathetic response
        if self.user_preferences["auto_reply_enabled"]:
            reply = message.Reply()
            reply.Body = f"Thank you for your email. I understand your concerns and will address them promptly.\n\nBrief summary of your message: {summary}{self.user_preferences['signature']}"
            reply.Save()
            self.stats["auto_replies_drafted"] += 1
        
        message.Save()

    def _handle_high_priority_email(self, message, summary, sentiment):
        """Handle high priority emails"""
        # Flag for follow-up
        message.FlagStatus = 2  # 2 = Flagged
        
        # Add follow-up reminder
        tomorrow = datetime.datetime.now() + datetime.timedelta(days=1)
        message.ReminderSet = True
        message.ReminderTime = tomorrow
        
        message.Save()

    def _handle_low_priority_email(self, message):
        """Handle low priority emails like newsletters"""
        # Move to newsletters folder if it seems like a newsletter
        if "newsletter" in message.Subject.lower() or "weekly update" in message.Subject.lower():
            newsletters_folder = self._get_or_create_folder("Newsletters")
            message.Move(newsletters_folder)
        
        # Simply mark as read
        message.UnRead = False
        message.Save()

    def _get_or_create_folder(self, folder_path):
        """Get an Outlook folder or create it if it doesn't exist, supports nested folders"""
        folders = self.inbox.Folders
        parent = self.inbox
        
        # Handle nested folders
        for part in folder_path.split('/'):
            try:
                folder = parent.Folders[part]
                parent = folder
            except Exception:
                logger.info(f"Creating folder: {part} in {parent.Name}")
                folder = parent.Folders.Add(part)
                parent = folder
        
        return parent

    def _check_calendar_conflicts(self, start_time, end_time):
        """Check for calendar conflicts during the specified time period"""
        calendar = self.namespace.GetDefaultFolder(9)  # 9 = Calendar
        appointments = calendar.Items
        appointments.Sort("[Start]")
        appointments.IncludeRecurrences = True
        
        # Set filter for the date range
        filter_string = f"[Start] <= '{end_time.strftime('%m/%d/%Y %H:%M')}' AND [End] >= '{start_time.strftime('%m/%d/%Y %H:%M')}'"
        appointments = appointments.Restrict(filter_string)
        
        conflicts = []
        for appt in appointments:
            conflicts.append({
                'subject': appt.Subject,
                'start': appt.Start.strftime('%m/%d/%Y %H:%M'),
                'end': appt.End.strftime('%m/%d/%Y %H:%M'),
                'organizer': appt.Organizer
            })
        
        return conflicts

    def _save_statistics(self):
        """Save email processing statistics to CSV file"""
        if not has_pandas:
            logger.warning("Pandas not available, skipping statistics")
            return
            
        stats_file = self.config['ANALYTICS'].get('stats_file', 'email_stats.csv')
        
        # Add timestamp
        self.stats["timestamp"] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Create DataFrame and save
        stats_dict = {
            "timestamp": self.stats["timestamp"],
            "emails_processed": self.stats["emails_processed"],
            "urgent_emails": self.stats["urgent_emails"],
            "meeting_emails": self.stats["meeting_emails"],
            "auto_replies_drafted": self.stats["auto_replies_drafted"],
            "duration": self.stats.get("duration", 0),
        }
        
        # Add folder statistics
        for folder, count in self.stats["by_folder"].items():
            folder_key = f"folder_{folder.replace(' ', '_')}"
            stats_dict[folder_key] = count
        
        stats_df = pd.DataFrame([stats_dict])
        
        # Append to existing file or create new
        if os.path.exists(stats_file):
            stats_df.to_csv(stats_file, mode='a', header=False, index=False)
        else:
            stats_df.to_csv(stats_file, index=False)
        
        logger.info(f"Statistics saved to {stats_file}")
        
        # Clean up old statistics if configured
        try:
            max_days = self.config['ANALYTICS'].getint('max_history_days', 30)
            self._prune_old_stats(stats_file, max_days)
        except Exception as e:
            logger.warning(f"Failed to prune old statistics: {e}")

    def _prune_old_stats(self, stats_file, max_days):
        """Remove statistics older than max_days"""
        if not os.path.exists(stats_file) or not has_pandas:
            return
            
        df = pd.read_csv(stats_file)
        if 'timestamp' not in df.columns:
            return
            
        # Convert timestamps to datetime
        df['timestamp'] = pd.to_datetime(df['timestamp'])
        
        # Calculate cutoff date
        cutoff = datetime.datetime.now() - datetime.timedelta(days=max_days)
        
        # Filter and save
        df_recent = df[df['timestamp'] >= cutoff]
        if len(df_recent) < len(df):
            df_recent.to_csv(stats_file, index=False)
            logger.info(f"Pruned {len(df) - len(df_recent)} old statistics entries")

    def generate_email_insights(self):
        """Generate insights from email patterns and statistics"""
        if not has_pandas:
            return "Pandas not available for generating insights. Install with: pip install pandas"
            
        stats_file = self.config['ANALYTICS'].get('stats_file', 'email_stats.csv')
        
        if not os.path.exists(stats_file):
            return "No email statistics available yet."
        
        try:
            stats_df = pd.read_csv(stats_file)
            stats_df['timestamp'] = pd.to_datetime(stats_df['timestamp'])
            
            # Calculate some basic insights
            total_emails = stats_df['emails_processed'].sum()
            avg_urgent = stats_df['urgent_emails'].sum() / len(stats_df) if len(stats_df) > 0 else 0
            avg_meeting = stats_df['meeting_emails'].sum() / len(stats_df) if len(stats_df) > 0 else 0
            
            # Calculate emails per day
            stats_df['date'] = stats_df['timestamp'].dt.date
            emails_by_date = stats_df.groupby('date')['emails_processed'].sum()
            
            if len(emails_by_date) > 1:
                avg_per_day = emails_by_date.mean()
                max_day = emails_by_date.idxmax()
                max_day_count = emails_by_date.max()
            else:
                avg_per_day = total_emails
                max_day = "N/A"
                max_day_count = total_emails
            
            # Find most active folders
            folder_columns = [col for col in stats_df.columns if col.startswith('folder_')]
            folder_totals = {}
            
            for col in folder_columns:
                folder_name = col.replace('folder_', '').replace('_', ' ')
                folder_totals[folder_name] = stats_df[col].sum()
            
            top_folders = sorted(folder_totals.items(), key=lambda x: x[1], reverse=True)[:5]
            
            insights = f"""
            Email Processing Insights:
            - Total emails processed: {total_emails}
            - Average emails per day: {avg_per_day:.1f}
            - Busiest day: {max_day} with {max_day_count} emails
            - Average urgent emails per run: {avg_urgent:.2f}
            - Average meeting emails per run: {avg_meeting:.2f}
            - Auto-replies drafted: {stats_df['auto_replies_drafted'].sum()}
            
            Top Active Folders:
            """
            
            for folder, count in top_folders:
                insights += f"- {folder}: {count} emails\n"
            
            return insights
        except Exception as e:
            logger.error(f"Error generating insights: {e}")
            return f"Error generating insights: {e}"

    def start_event_handler(self):
        """Start the event handler for new incoming emails"""
        if not self.config['OUTLOOK'].getboolean('enable_event_handler', True):
            logger.info("Event handler disabled in config")
            return
            
        # Create a thread for the event handler
        event_thread = threading.Thread(target=self._run_event_handler)
        event_thread.daemon = True  # Allow the program to exit even if thread is running
        event_thread.start()
        logger.info("Started event handler thread")

    def _run_event_handler(self):
        """Run the event handler loop with improved COM handling"""
        try:
            # Initialize COM for this thread
            pythoncom.CoInitializeEx(pythoncom.COINIT_MULTITHREADED)
            
            # Define event sink class inline to avoid metaclass conflicts
            class OutlookEventSink:
                def __init__(self, agent):
                    self._agent = agent
                
                def OnNewMail(self):
                    try:
                        logger.info("New mail received event triggered")
                        # Process the most recent unread email in the inbox
                        inbox = self._agent.inbox
                        emails = inbox.Items.Restrict("[Unread] = true")
                        if emails.Count > 0:
                            # Sort newest first
                            emails.Sort("[ReceivedTime]", True)
                            newest = emails.GetFirst()
                            self._agent._process_single_email(newest)
                    except Exception as e:
                        logger.error(f"Error in OnNewMail handler: {e}")
            
            # Create and register the event sink
            self.outlook_sink = win32com.client.WithEvents(self.outlook, OutlookEventSink)
            self.outlook_sink._agent = self
            
            logger.info("Event handler running, waiting for new emails...")
            
            # Message pump
            while True:
                pythoncom.PumpWaitingMessages()
                time.sleep(0.5)
        except Exception as e:
            logger.error(f"Event handler failed: {e}")
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass
            
    def stop(self):
        """Clean up resources before exiting"""
        logger.info("Shutting down Outlook AI Agent")
        # Release COM objects
        try:
            del self.inbox
            del self.namespace
            del self.outlook
        except:
            pass

    def _is_mail_folder(self, folder):
        """Check if a folder contains mail items that can be processed"""
        try:
            # Try to access a mail-specific property or method
            items = folder.Items
            if items.Count > 0:
                # Try to get the first item and check if it has mail properties
                try:
                    item = items.GetFirst()
                    # Try to access a mail-specific property
                    _ = getattr(item, "ReceivedTime", None)
                    return True
                except:
                    return False
            return True  # Empty folders are assumed to be mail folders
        except:
            return False
            
# Main execution
def run():
    agent = None
    try:
        # Initialize and run the agent
        agent = OutlookAIAgent()
        
        # Start event handler for real-time processing
        agent.start_event_handler()
        
        # Initial scan of existing unread emails
        logger.info("Starting initial scan of all folders")
        stats = agent.process_emails()
        
        # Print summary
        print(f"\nProcessed {stats['emails_processed']} emails in {stats.get('duration', 0):.1f} seconds")
        print(f"Found {stats['urgent_emails']} urgent emails")
        print(f"Found {stats['meeting_emails']} meeting emails")
        print(f"Drafted {stats['auto_replies_drafted']} auto-replies")
        
        # Print folder breakdown
        print("\nEmails processed by folder:")
        for folder, count in stats["by_folder"].items():
            print(f"- {folder}: {count}")
        
        # Generate insights
        if has_pandas:
            insights = agent.generate_email_insights()
            print("\nInsights:")
            print(insights)
        
        # Keep running until user interrupts
        print("\nOutlook AI Agent is running in the background.")
        print("Press Ctrl+C to exit...")
        
        # Get scan interval from config
        scan_interval = agent.config['OUTLOOK'].getint('scan_interval', 300)  # Default to 5 minutes (300 seconds)

        while True:
            time.sleep(scan_interval)  # Use the configured interval
            
            # Periodic reprocessing
            current_time = datetime.datetime.now()
            print(f"\n[{current_time.strftime('%Y-%m-%d %H:%M:%S')}] Running periodic scan...")
            stats = agent.process_emails()
            print(f"Processed {stats['emails_processed']} new emails")
            
    except KeyboardInterrupt:
        print("\nExiting Outlook AI Agent...")
    except Exception as e:
        logging.error(f"Error running Outlook AI Agent: {e}", exc_info=True)
        print(f"Error: {e}")
        
        # Offer suggestions for common errors
        if "win32com" in str(e).lower():
            print("\nTroubleshooting:")
            print("- Ensure Outlook is installed and running")
            print("- Install pywin32: pip install pywin32")
        elif "transformers" in str(e).lower():
            print("\nTroubleshooting:")
            print("- Install transformers: pip install transformers")
            print("- For GPU support: pip install torch")
    finally:
        if agent:
            agent.stop()

if __name__ == "__main__":
    run()