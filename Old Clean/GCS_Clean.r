library(dplyr)
library(DBI)
library(RODBC)
library(readr)
library(readxl)

db_conn <- odbcConnect("ILS")

# Read data from Excel file
table <- read_excel("S:/Touchstone/Catrader/Boston/Database/CatBond/Shiny/GCS_Weekly/GCS20231229.xlsx", sheet = 2, col_names = FALSE, skip = 6) %>%
  data.frame()

# Add date to the front of the table (Needs to be manually edited)
Date <- rep(as.Date("2023-12-29"), times = nrow(table))
table <- cbind(Date, table)

# Get column names from the database and set them as column names for the table
ColNames <- sqlColumns(db_conn, "GCS") %>%
  select('COLUMN_NAME')
colnames(table) <- ColNames$COLUMN_NAME

# Limit BidSpread and OfferSpread to 2 decimal places and convert any N/A values to 0
table$BidSpread[is.na(table$BidSpread)] <- 0
table$OfferSpread[is.na(table$OfferSpread)] <- 0
table$BidSpread <- formatC(c(table$BidSpread), digits = 2, format = 'f')
table$OfferSpread <- formatC(c(table$OfferSpread), digits = 2, format = 'f')

# Save the table to the database
sqlSave(db_conn, table, tablename = "GCS", rownames = FALSE, append = TRUE, fast = FALSE)

# Close the database connection
odbcClose(db_conn)
