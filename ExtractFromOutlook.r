library(RDCOMClient)

# Create an Outlook application object
Outlook <- COMCreate("Outlook.Application")
myNameSpace <- Outlook$GetNameSpace("MAPI")

# Access the Inbox
inboxFolderIndex <- 1 # This is usually 1, but might vary depending on your Outlook setup
inbox <- myNameSpace$Folders(inboxFolderIndex)$Folders("Inbox")

# Access the specific subfolder "BH Cat Bond" within the Inbox
bhCatBondFolder <- inbox$Folders("BH Cat Bond")

# Get all messages in the "BH Cat Bond" folder
messages <- bhCatBondFolder$Items()

# Get the number of messages in the folder
num_messages <- messages$Count()

# Loop to read messages
for (i in 1:num_messages) {
    message <- messages$Item(i)
    
    # Print message details
    cat("Subject:", message$Subject(), "\n")
    cat("Body:", message$Body(), "\n\n")
    
    # Here you can add code to process these details
    # For example, storing them in an SQL database
}
