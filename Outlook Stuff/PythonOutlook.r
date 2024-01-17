library(reticulate)

# Use reticulate to run Python code
python_code <- "
import win32com.client
import os

# Create an Outlook application object
Outlook = win32com.client.Dispatch('Outlook.Application')
namespace = Outlook.GetNamespace('MAPI')

# Access the Inbox and then the specific subfolder
inbox = namespace.GetDefaultFolder(6)  # 6 refers to the inbox
bhCatBondFolder = inbox.Folders['BH Cat Bond']

# Create a directory to store the text files
outputDir = 'S:/Touchstone/Catrader/Boston/Database/UnreadEmails'
os.makedirs(outputDir, exist_ok=True)

# Loop to process unread messages
for message in bhCatBondFolder.Items:
    if message.UnRead:
        # Prepare the filename, removing problematic characters
        subject = message.Subject.replace(':', '').replace('/', '').replace('\\\\', '')
        filename = f'{outputDir}/{subject}.txt'
        # Save the email as a text file
        message.SaveAs(filename, 0)  # 0 is the olTXT format
        # Mark the message as read (optional)
        # message.UnRead = False
        # message.Save()
"

# Run the Python code
py_run_string(python_code)

# Process the saved text files in R
emailFiles <- list.files(path = "S:/Touchstone/Catrader/Boston/Database/UnreadEmails", pattern = "*.txt", full.names = TRUE)

emails_df <- data.frame(Timestamp = character(),
                        Subject = character(),
                        Content = character(),
                        stringsAsFactors = FALSE)

for (emailFile in emailFiles) {
    emailContent <- readLines(emailFile, warn = FALSE)
    receivedTime <- file.info(emailFile)$mtime
    formattedTime <- format(as.POSIXct(receivedTime), "%Y-%m-%d %H:%M:%S")

    emails_df <- rbind(emails_df, data.frame(Timestamp = formattedTime,
                                             Subject = "Email Subject", # Modify as needed
                                             Content = paste(emailContent, collapse = "\n")))
    unlink(emailFile)
}

write.csv(emails_df, file = "S:/Touchstone/Catrader/Boston/Database/UnreadDatabaseEntryEmails.csv", row.names = FALSE)


ed string constant in:
"        # Save the email as a text file
        filename = f"{outputDir}/{message.Subject.replace(':', '').replace('"
>         message.SaveAs(filename, 0)  # 0 is the olTXT format
Error in message.SaveAs(filename, 0) : 
  could not find function "message.SaveAs"
