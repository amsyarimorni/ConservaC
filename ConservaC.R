# Load the required library
if (!requireNamespace("writexl", quietly = TRUE)) {
  install.packages("writexl")
}
library(writexl)

if (!requireNamespace("readxl", quietly = TRUE)) {
  install.packages("readxl")
}
library(readxl)

if (!requireNamespace("openxlsx", quietly = TRUE)) {
  install.packages("openxlsx")
}
library(openxlsx)

if (!requireNamespace("readxlsb", quietly = TRUE)) {
  install.packages("readxlsb")
}
library(readxlsb)

if (!requireNamespace("xlsx", quietly = TRUE)) {
  install.packages("xlsx")
}
library(xlsx)


# Specify the file paths for Excel_1 and Excel_2
excel_1_path <- "Species query.xlsx"
excel_2_path <- "MasterData_IUCN_ENDEMICITY_9.2023.xlsx"

# Specify the password for the Excel file
password <- "1beconly"  # Replace with the actual password

# Read data from the Excel file
excel_2 <- xlsx::read.xlsx(excel_2_path, sheetName = "sheet1", password = password)# Specify the password for the Excel file

# Load Excel_1
excel_1 <- readxl::read_xlsx(excel_1_path)

# Perform VLOOKUP using merge
Result <- merge(excel_1, excel_2, by="Species", all.x=TRUE)

# Replace NAs with "NOT AVAILABLE"
Result[is.na(Result)] <- "NO MATCH"

# Specify the file path for Excel_3
Result_path <- "Result.xlsx"

# Save the result to Excel_3
writexl::write_xlsx(Result, Result_path)

# Export Excel_3 to a CSV file
write.csv(Result, "Result.csv", row.names = FALSE)

# Open the CSV file in Excel
shell.exec("Result.csv")

#Custom message
cat("\033[32mYour query is completed, Database update: Sept 2023.\033[32m\n")
cat("\033[32mAll the best for your research -Morni M.A.(IBEC, UNIMAS)\033[32m\n")
