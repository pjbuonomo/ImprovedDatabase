library(dplyr)
library(DBI)
library(RODBC)
library(readr)
library(readxl)

# Establish a connection to the database
db_conn <- odbcConnect("ILS")

# Read the Excel file into a data frame, remove NA values, and create a Date column
table <- read_excel("S:/Touchstone/Catrader/Boston/Database/CatBond/Shiny/RBC_Weekly/RBC20231215.xlsx", col_names = TRUE) %>%
  na.omit() %>%
  data.frame()

# Add a date column with "2023-12-15" to the front of the table (Needs to be manually edited)
Date <- rep(as.Date("2023-12-15"), times = nrow(table))
table <- cbind(Date, table)

# Retrieve column names from the database table and set them as column names for the table
ColNames <- sqlColumns(db_conn, "RBC") %>%
  select('COLUMN_NAME')
colnames(table) <- ColNames$COLUMN_NAME

# Save the table to the database as "RBC"
sqlSave(db_conn, table, tablename = "RBC", rownames = FALSE, append = TRUE, fast = FALSE, verbose = TRUE)

# Close the database connection
odbcClose(db_conn)
