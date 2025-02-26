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
                'skip_folders': 'Deleted Items,Junk Email,Outbox,Sent Items,Calendar,Tasks',
                'customer_priority_boost': '2',
                'team_priority_boost': '3',
                'scan_interval': '300'
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
