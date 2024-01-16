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

        # Retrieve and format the ReceivedTime
        receivedTime <- message$ReceivedTime()
        formattedTime <- format(as.POSIXct(receivedTime, origin = "1970-01-01"), "%Y-%m-%d %H:%M:%S")

        # Add email details to the data frame
        emails_df <- rbind(emails_df, data.frame(Timestamp = formattedTime,
                                                 Subject = message$Subject(),
                                                 Content = emailContent))

        # Mark the message as read (optional)
        # message$UnRead(FALSE)
        # message$Save()
    }
}

# Write the data frame to a CSV file
write.csv(emails_df, file = "S:/Touchstone/Catrader/Boston/Database/UnreadDatabaseEntryEmails.csv", row.names = FALSE)
cat("Email processing completed. Data written to UnreadDatabaseEntryEmails.csv\n")
<html><head><title></title></head><body><!-- rte-version 0.2 9947551637294008b77bce25eb683dac --><div class="rte-style-maintainer rte-pre-wrap" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small;" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small; color: rgb(0, 0, 0);">Trading Res Re at 100.45 and axed to buy more there. Also can improve Montoya 24 to 101 bid. Once again many thanks for the focus..<div><br><div class="rte-style-maintainer" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-size: small; font-family: &quot;Courier New&quot;, Courier, &quot;BB.FixedWidth&quot;;" style="font-size: small; font-family: &quot;Courier New&quot;, Courier, &quot;BB.FixedWidth&quot;; color: rgb(0, 0, 0);"><div class="rte-style-maintainer rte-pre-wrap" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small;" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small; color: rgb(0, 0, 0);"><div class="rte-style-maintainer" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-size: small; font-family: &quot;Courier New&quot;, Courier, &quot;BB.FixedWidth&quot;;" style="font-size: small; font-family: &quot;Courier New&quot;, Courier, &quot;BB.FixedWidth&quot;; color: rgb(0, 0, 0);"><div class="rte-style-maintainer rte-pre-wrap" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small;" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small; color: rgb(0, 0, 0);"><span class="bbScopedStyle8017738174112958">Please show in all offerings. Many thanks for the focus.</span></div><div class="rte-style-maintainer rte-pre-wrap" data-color="global-default" bbg-color="default" data-bb-font-size="medium" bbg-font-size="medium" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small;" style="font-family: Arial, &quot;BB.Proportional&quot;; white-space: pre-wrap; font-size: small; color: rgb(0, 0, 0);"><div><span class="bbScopedStyle8017738174112958">Alamo 2023-1 A (011395AJ9) bid at 102.50</span></div><div><span class="bbScopedStyle8017738174112958">Blue Sky 2023-1 (XS2728630596) bid at 100.15</span></div><div><span class="bbScopedStyle8017738174112958">Bonanza 2022-1 A (09785EAJ0) bid at 90.00</span></div><div><span class="bbScopedStyle8017738174112958">Bonanza 2023-1 A (09785EAK7) bid at 99.90</span></div><div><span class="bbScopedStyle8017738174112958">Citrus 2023-1 B (177510AM6) bid at 102.40</span></div><div><span class="bbScopedStyle8017738174112958">Easton 2024-1 A (27777AAA9) bid at 100.25<br>First Coast 2021-1 (31971CAA1) bid at 96.15</span></div><div><span class="bbScopedStyle8017738174112958">First Coast 2023-1 (31969UAA5) bid at 101.10</span></div><div><span class="bbScopedStyle8017738174112958">Galileo 2023-1 B (36354TAP7) bid at 100.25</span></div><div><span class="bbScopedStyle8017738174112958">Galileo 2023-1 A (36354TAN2) bid at 100.25</span></div><div><span class="bbScopedStyle8017738174112958">Hypatia 2023-1 A (44914CAC0) bid at 104.35</span></div><div><span class="bbScopedStyle8017738174112958">Hexagon 2023-1 A (428270AA0) bid at 100.50</span></div><div><span class="bbScopedStyle8017738174112958">Lightning 2023-1 A (532242AA2) bid at 106.30</span></div><div><span class="bbScopedStyle8017738174112958">Matterhorn 2022-I B (577092AQ2) bid at 98.50</span></div><div><span class="bbScopedStyle8017738174112958">Merna 2022-2A (59013MAF9) bid at 98.65 </span></div><div><span class="bbScopedStyle8017738174112958">Merna 2023-2 A (59013MAJ1) bid at 104.35</span></div><div><span class="bbScopedStyle8017738174112958">Mona Lisa 2023-1 B (608800AG3) bid at 107.75</span></div><div><span class="bbScopedStyle8017738174112958">Montoya 2022-2 (613752AB0) bid at 108.60</span></div><div><span class="bbScopedStyle8017738174112958">Montoya 2024-1 A (613752AC8) bid at 100.25</span></div><div><span class="bbScopedStyle8017738174112958">Ocelot 2023-1 A (675951AA5) bid at 100.30</span></div><div><span class="bbScopedStyle8017738174112958">Residential Re 2023-2 5 (76090WAC4) bid at 100.35</span></div><div><span class="bbScopedStyle8017738174112958">Tailwind 2022-1 B (87403TAE6) bid at 96.90</span></div><div><span class="bbScopedStyle8017738174112958">Tailwind 2022-1 C (87403TAE) bid at 98.15</span></div><div><span class="bbScopedStyle8017738174112958">Titania 2021-1 A (888329AA7) bid at 100.40</span></div><div><span class="bbScopedStyle8017738174112958">Titania 2021-2 A (888329AB5) bid at 97.50</span></div><div><span class="bbScopedStyle8017738174112958">Titania 2023-1 A (888329AC3) bid at 107.75</span></div><div><span class="bbScopedStyle8017738174112958">Ursa 2023-1 C (90323WAM2) bid at 100.45</span></div><div><span class="bbScopedStyle8017738174112958">Ursa 2023-3 D (90323WAQ3) bid at 100.35</span></div></div></div></div></div></div></div></body></html>
