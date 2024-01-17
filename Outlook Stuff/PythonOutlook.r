import win32com.client
import os
import pandas as pd
from datetime import datetime

# Create an Outlook application object
Outlook = win32com.client.Dispatch("Outlook.Application")
namespace = Outlook.GetNamespace("MAPI")

# Access the Inbox and then the specific subfolder
inbox = namespace.GetDefaultFolder(6)  # 6 refers to the inbox
bhCatBondFolder = inbox.Folders["BH Cat Bond"]

# List to store email details
emails = []

# Loop to process unread messages
for message in bhCatBondFolder.Items:
    if message.UnRead:
        # Extract email details
        subject = message.Subject
        body = message.Body
        received_time = message.ReceivedTime

        # Format the date
        formatted_time = received_time.strftime('%Y-%m-%d %H:%M:%S')

        # Append to the list
        emails.append({
            "Timestamp": formatted_time,
            "Subject": subject,
            "Content": body
        })

        # Mark the message as read (optional)
        # message.UnRead = False
        # message.Save()

# Create a DataFrame
emails_df = pd.DataFrame(emails)

# Define the output CSV file path
output_csv = "S:/Touchstone/Catrader/Boston/Database/UnreadDatabaseEntryEmails.csv"

# Save the DataFrame to CSV
emails_df.to_csv(output_csv, index=False)
