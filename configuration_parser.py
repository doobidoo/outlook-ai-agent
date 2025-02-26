import configparser

# Load the current config
config = configparser.ConfigParser()
config.read('config.ini')

# Folders to skip (combining the error list with standard system folders)
folders_to_skip = [
    # Folders from error logs
    "Kontakte", "Recipient Cache", "{06967759-274D-40B2-A3EB-D7F9E73727D7}",
    "{A9E2BC46-B3A0-4243-B315-60D991004455}", "Firmen", "Organizational Contacts",
    "GAL Contacts", "PeopleCentricConversation Buddies", "ExternalContacts",
    "Journal", "PersonMetadata", "Notizen",
    
    # Standard system folders and their German equivalents
    "Calendar", "Tasks", "Contacts", "Deleted Items", "Junk Email", 
    "Outbox", "Sent Items", "Aufgaben", "Kalender", "Gelöschte Elemente",
    "Postausgang", "Gesendete Elemente", "Entwürfe", "SharePoint Online", "Junk-E-Mail",
    
    # Problem folders from previous logs
    "Synchronisierungsprobleme", "Konflikte", "Lokale Fehler"
]

# Update the skip_folders setting
if 'OUTLOOK' not in config:
    config['OUTLOOK'] = {}
    
config['OUTLOOK']['skip_folders'] = ','.join(folders_to_skip)

# Save the config
with open('config.ini', 'w') as f:
    config.write(f)

print("Configuration updated successfully with the following folders to skip:")
for folder in folders_to_skip:
    print(f"- {folder}")