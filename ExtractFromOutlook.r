library(RDCOMClient)
library(rvest)

# Create an Outlook application object
Outlook <- COMCreate("Outlook.Application")
myNameSpace <- Outlook$GetNameSpace("MAPI")

# Access the Inbox and then the specific subfolder
inboxFolderIndex <- 1 # Adjust based on your Outlook setup
inbox <- myNameSpace$Folders(inboxFolderIndex)$Folders("Inbox")
bhCatBondFolder <- inbox$Folders("BH Cat Bond")

# Get all messages in the "BH Cat Bond" folder
messages <- bhCatBondFolder$Items()

# Initialize a data frame to store email details
emails_df <- data.frame(Timestamp = character(),
                        Subject = character(),
                        Content = character(),
                        stringsAsFactors = FALSE)

# Get the number of messages in the folder
num_messages <- messages$Count()

# Loop to read messages
for (i in 1:num_messages) {
    message <- messages$Item(i)
    
    # Process only if the message is unread
    if (message$UnRead() == TRUE) {
        # Retrieve HTML email content
        htmlContent <- message$HTMLBody()

        # Parse and extract content within the <html> tags
        if (!is.null(htmlContent) && htmlContent != "") {
            parsedHtml <- read_html(htmlContent)
            extractedContent <- html_text(parsedHtml)
        }

        # Retrieve and format the ReceivedTime
        receivedTime <- message$ReceivedTime()
        formattedTime <- format(as.POSIXct(receivedTime, origin = "1970-01-01"), "%Y-%m-%d %H:%M:%S")

        # Add email details to the data frame
        emails_df <- rbind(emails_df, data.frame(Timestamp = formattedTime,
                                                 Subject = message$Subject(),
                                                 Content = extractedContent))

        # Mark the message as read (optional)
        # message$UnRead(FALSE)
        # message$Save()
    }
}


# Write the data frame to a CSV file
write.csv(emails_df, file = "S:/Touchstone/Catrader/Boston/Database/UnreadDatabaseEntryEmails.csv", row.names = FALSE)
cat("Email processing completed. Data written to UnreadDatabaseEntryEmails.csv\n")
