if (!requireNamespace("remotes", quietly = TRUE))
    install.packages("remotes")
remotes::install_github("omegahat/RDCOMClient")


library(RDCOMClient)

# Create an Outlook application object
Outlook <- COMCreate("Outlook.Application")
myNameSpace <- Outlook$GetNameSpace("MAPI")

# Access the Inbox
folder <- myNameSpace$Folders$Item(1)$Folders$Item("Inbox")
messages <- folder$Items()

# Define the number of messages to read
num_messages <- messages$Count()

# Loop to read messages
for (i in 1:num_messages) {
    message <- messages$Item(i)
    cat("Subject:", message$Subject(), "\n")
    cat("Body:", message$Body(), "\n\n")
    # Here you can add code to process or store these details
    # For example, parsing the body of the email and storing it in an SQL database
}
