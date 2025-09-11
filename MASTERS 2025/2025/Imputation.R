# Install and load necessary packages
# install.packages("missForest")
# install.packages("readxl")
# install.packages("openxlsx")
library(missForest)
library(readxl)
library(openxlsx)

setwd("C:\\Users\\AdonsMB\\OneDrive - University of the Free State\\Desktop\\Masters\\MASTERS 2025\\2025\\Data")
# Create a list of your 7 file names. Replace "file1.xlsx", etc., with your actual file names.
file_list <- c("Bethlehem Data.xlsx", "Bloem Stad Data.xlsx", "Bloem Wo Data.xlsx", "Fauresmith Data.xlsx", "Gariep Dam Data.xlsx", "Vrede Data.xlsx", "Welkom Data.xlsx")

# Loop through each file
for (file_name in file_list) {
  tryCatch({
    # Read the data from the "Maximum" sheet
    sheet_name_max <- "Maximum"
    data_max <- read_excel(file_name, sheet = sheet_name_max)
    
    # Read the data from the "Precipitation" sheet
    sheet_name_precip <- "Precipitation"
    data_precip <- read_excel(file_name, sheet = sheet_name_precip)
    
    # Convert data frames to numeric for missForest imputation
    data_max_numeric <- data.frame(sapply(data_max, as.numeric))
    data_precip_numeric <- data.frame(sapply(data_precip, as.numeric))
    
    # Perform missForest imputation on "Maximum" sheet
    imputed_max <- missForest(data_max_numeric)
    imputed_max_df <- imputed_max$ximp
    
    # Perform missForest imputation on "Precipitation" sheet
    imputed_precip <- missForest(data_precip_numeric)
    imputed_precip_df <- imputed_precip$ximp
    
    # Create a new workbook and add the imputed sheets
    wb <- createWorkbook()
    addWorksheet(wb, sheet_name_max)
    writeData(wb, sheet = sheet_name_max, x = imputed_max_df)
    
    addWorksheet(wb, sheet_name_precip)
    writeData(wb, sheet = sheet_name_precip, x = imputed_precip_df)
    
    # Construct the new file name with "_new"
    new_file_name <- gsub("\\.xlsx$", "_new.xlsx", file_name)
    
    # Save the new workbook
    saveWorkbook(wb, new_file_name, overwrite = TRUE)
    
    cat(paste("Successfully imputed and saved:", new_file_name, "\n"))
    
  }, error = function(e) {
    cat(paste("An error occurred for file:", file_name, "\n"))
    cat(paste("Error message:", e$message, "\n"))
  })
}