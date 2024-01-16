library(RDCOMClient)

# Create an Outlook application object
Outlook <- COMCreate("Outlook.Application")
myNameSpace <- Outlook$GetNameSpace("MAPI")

# Access the Inbox - Ensure you're accessing the correct folder
inboxFolderIndex <- 1 # This is usually 1, but might vary depending on your Outlook setup
inbox <- myNameSpace$Folders(inboxFolderIndex)$Folders("Inbox")

# Get all messages in the Inbox
messages <- inbox$Items()

# Now, proceed to filter messages and process them

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
> library(RDCOMClient)
> 
> # Create an Outlook application object
> Outlook <- COMCreate("Outlook.Application")
> myNameSpace <- Outlook$GetNameSpace("MAPI")
> 
> # Access the Inbox - Ensure you're accessing the correct folder
> inboxFolderIndex <- 1 # This is usually 1, but might vary depending on your Outlook setup
> inbox <- myNameSpace$Folders(inboxFolderIndex)$Folders("Inbox")
> 
> # Get all messages in the Inbox
> messages <- inbox$Items()
> 
> # Now, proceed to filter messages and process them
> 
> # Filter for unread messages from a specific sender
> target_sender <- "csbond@bloomberg.net"
> filtered_messages <- messages[messages$UnRead() == TRUE & messages$SenderEmailAddress() == target_sender]
Error in h(simpleError(msg, call)) : 
  error in evaluating the argument 'i' in selecting a method for function '[': Cannot locate 0 name(s) UnRead in COM object (status = -2147352570)
> 
> # Get the number of filtered messages
> num_messages <- filtered_messages$Count()
Error: object 'filtered_messages' not found
> 
> # Loop to read and process messages
> for (i in num_messages:1) {
+     message <- filtered_messages$Item(i)
+     
+     # Check if the message is from the specified sender and is unread
+     if (message$SenderEmailAddress() == target_sender && message$UnRead() == TRUE) {
+         # Print message details
+         cat("Subject:", message$Subject(), "\n")
+         cat("Body:", message$Body(), "\n\n")
+         
+         # Process the message as needed
+         # ... (your processing code here)
+ 
+         # Mark the message as read
+         message$UnRead(FALSE)
+         message$Save()
+     }
+ }
