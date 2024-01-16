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
> library(RDCOMClient)
> 
> # Create an Outlook application object
> Outlook <- COMCreate("Outlook.Application")
> myNameSpace <- Outlook$GetNameSpace("MAPI")
> 
> # Access the Inbox
> folder <- myNameSpace$Folders$Item(1)$Folders$Item("Inbox")
Error in myNameSpace$Folders$Item : 
  object of type 'closure' is not subsettable
> messages <- folder$Items()
Error: object 'folder' not found
> 
> # Filter for unread messages from a specific sender
> target_sender <- "csbond@bloomberg.net"
> filtered_messages <- messages[messages$UnRead() == TRUE & messages$SenderEmailAddress() == target_sender]
Error: object 'messages' not found
> 
> # Get the number of filtered messages
> num_messages <- filtered_messages$Count()
Error: object 'filtered_messages' not found
> 
> # Define the number of messages to process (in this case, 3)
> num_to_process <- 3
> 
> # Loop to read and process messages
> for (i in num_messages:1) {
+   message <- filtered_messages$Item(i)
+   
+   # Check if the message is from the specified sender and is unread
+   if (message$SenderEmailAddress() == target_sender && message$UnRead() == TRUE) {
+     # Print message details
+     cat("Subject:", message$Subject(), "\n")
+     cat("Body:", message$Body(), "\n\n")
+     
+     # Process the message as needed
+     # ... (your processing code here)
+     
+     # Mark the message as read
+     message$UnRead(FALSE)
+     message$Save()
+     
+     # Decrease the number of messages to process
+     num_to_process <- num_to_process - 1
+     
+     # Exit the loop if we have processed the desired number of messages
+     if (num_to_process == 0) {
+       break
+     }
+   }
+ }
