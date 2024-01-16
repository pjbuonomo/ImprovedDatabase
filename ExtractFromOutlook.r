library(RDCOMClient)

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
        # Retrieve email content
        emailContent <- ifelse(is.null(message$HTMLBody()), message$Body(), message$HTMLBody())

        # Add email details to the data frame
        emails_df <- rbind(emails_df, data.frame(Timestamp = format(message$ReceivedTime(), "%Y-%m-%d %H:%M:%S"),
                                                 Subject = message$Subject(),
                                                 Content = emailContent))

        # Mark the message as read (optional)
        # message$UnRead(FALSE)
        # message$Save()
    }
}

# Write the data frame to a CSV file
write.csv(emails_df, file = "S:/Touchstone/Catrader/Boston/Database/UnreadDatabaseEntryEmails.csv", row.names = FALSE)


















 library(RDCOMClient)
> 
> # Create an Outlook application object
> Outlook <- COMCreate("Outlook.Application")
> myNameSpace <- Outlook$GetNameSpace("MAPI")
> 
> # Access the Inbox and then the specific subfolder
> inboxFolderIndex <- 1 # Adjust based on your Outlook setup
> inbox <- myNameSpace$Folders(inboxFolderIndex)$Folders("Inbox")
> bhCatBondFolder <- inbox$Folders("BH Cat Bond")
> 
> # Get all messages in the "BH Cat Bond" folder
> messages <- bhCatBondFolder$Items()
> 
> # Initialize a data frame to store email details
> emails_df <- data.frame(Timestamp = character(),
+                         Subject = character(),
+                         Content = character(),
+                         stringsAsFactors = FALSE)
> 
> # Get the number of messages in the folder
> num_messages <- messages$Count()
> 
> # Loop to read messages
> for (i in 1:num_messages) {
+     message <- messages$Item(i)
+     
+     # Process only if the message is unread
+     if (message$UnRead() == TRUE) {
+         # Retrieve email content
+         emailContent <- ifelse(is.null(message$HTMLBody()), message$Body(), message$HTMLBody())
+ 
+         # Add email details to the data frame
+         emails_df <- rbind(emails_df, data.frame(Timestamp = format(message$ReceivedTime(), "%Y-%m-%d %H:%M:%S"),
+                                                  Subject = message$Subject(),
+                                                  Content = emailContent))
+ 
+         # Mark the message as read (optional)
+         # message$UnRead(FALSE)
+         # message$Save()
+     }
+ }
Error in prettyNum(.Internal(format(x, trim, digits, nsmall, width, 3L,  : 
  invalid 'trim' argument
> 
> # Write the data frame to a CSV file
> write.csv(emails_df, file = "S:/Touchstone/Catrader/Boston/Database/UnreadDatabaseEntryEmails.csv", row.names = FALSE)
> 
