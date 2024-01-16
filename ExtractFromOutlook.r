library(RDCOMClient)

# Create an Outlook application object
Outlook <- COMCreate("Outlook.Application")
myNameSpace <- Outlook$GetNameSpace("MAPI")

# Access the Inbox
folder <- myNameSpace$Folders$Item(1)$Folders$Item("Inbox")
messages <- folder$Items()

# Filter for unread messages from a specific sender
target_sender <- "csbond@bloomberg.net"
filtered_messages <- messages[messages$UnRead() == TRUE & messages$SenderEmailAddress() == target_sender]

# Get the number of filtered messages
num_messages <- filtered_messages$Count()

# Loop to read and process messages
for (i in num_messages:1) {
    message <- filtered_messages$Item(i)
    
    # Check if the message is from the specified sender and is unread
    if (message$SenderEmailAddress() == target_sender && message$UnRead() == TRUE) {
        # Print message details
        cat("Subject:", message$Subject(), "\n")
        cat("Body:", message$Body(), "\n\n")
        
        # Process the message as needed
        # ... (your processing code here)

        # Mark the message as read
        message$UnRead(FALSE)
        message$Save()
    }
}
