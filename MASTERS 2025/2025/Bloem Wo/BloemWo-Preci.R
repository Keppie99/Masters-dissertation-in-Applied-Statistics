library(readxl)
library(dplyr)
library(tidyverse)
library(missForest)
library(trend)
library(precintcon)
library(writexl)
library(fmsb)
library(strucchange)
library(robustrank)
library(openxlsx)
library(Kendall)
library(EnvStats)
library(modifiedmk)


#####PRECIPITATION#####
##GariepDam##
GariepDam = read_excel("C:\\Users\\AdonsMB\\OneDrive - University of the Free State\\Desktop\\Masters\\MASTERS 2025\\2025\\Bloem Wo\\Bloem Wo Data.xlsx")
GariepDam = data.frame(sapply(GariepDam, as.numeric))
#GariepDam$Year = as.factor(GariepDam$Year)
GariepDam = missForest(GariepDam) #MISS_FOREST IMPUTATION
GariepDam = GariepDam$ximp
GariepDam = cbind(GariepDam,"Winter"=rowMeans(GariepDam[c(6:8)]))
GariepDam = cbind(GariepDam,"Spring"=rowMeans(GariepDam[c(9:11)]))
GariepDam = cbind(GariepDam,"Summer"=rowMeans(GariepDam[c(12,13,2)]))
GariepDam = cbind(GariepDam,"Autumn"=rowMeans(GariepDam[c(3:5)]))
GariepDam = cbind(GariepDam,"Annual"=rowMeans(GariepDam[c(2:13)]))
write_xlsx(GariepDam,"C:\\Users\\AdonsMB\\OneDrive - University of the Free State\\Desktop\\Masters\\MASTERS 2025\\2025\\Bloem Wo\\BloemWo-Preci.xlsx")
GariepDam1 = GariepDam
GariepDam1$Year = as.factor(GariepDam1$Year)
# Initialize an empty data frame to store the results
sens_slope_results <- data.frame(
  Variable = character(),
  Sen_Slope = numeric(),
  P_Value = numeric(),
  stringsAsFactors = FALSE
)
mk.test(GariepDam$Summer)
# Get the column names to iterate over (exclude 'Year')
columns_to_analyze <- names(GariepDam)[!names(GariepDam) %in% "Year"]

# Loop through each column and perform Sens Slope analysis
for (col_name in columns_to_analyze) {
  # Ensure the data for sens.slope is numeric
  data_vector <- as.numeric(GariepDam[[col_name]])
  
  # Perform Sens slope test
  sens_test <- sens.slope(data_vector)
  
  # Extract relevant information
  slope_value <- sens_test$estimates[1]
  p_value <- sens_test$p.value
  
  # Add to results data frame
  sens_slope_results <- rbind(sens_slope_results,
                              data.frame(Variable = col_name,
                                         Sen_Slope = slope_value,
                                         P_Value = p_value))
}

# Print the results to see them in the console
print(sens_slope_results)

# Write the results to an Excel file
write_xlsx(sens_slope_results,"C:\\Users\\AdonsMB\\OneDrive - University of the Free State\\Desktop\\Masters\\MASTERS 2025\\2025\\Bloem Wo\\Sens_Slope_Results.xlsx")

cat("\nSens Slope results saved to:", "C:\\Users\\AdonsMB\\OneDrive - University of the Free State\\Desktop\\Masters\\MASTERS 2025\\2025\\Bloem Wo\\Sens_Slope_Results.xlsx", "\n")

#######################################################################################
#######################################################################################
columns_to_test <- names(GariepDam)[!names(GariepDam) %in% c("Year")]

# Create an empty list to store the results
mk_results <- list()

# Loop through each column and perform the Mann-Kendall test
for (col_name in columns_to_test) {
  # Get the data vector for the current column
  data_vector <- as.numeric(GariepDam[[col_name]])
  
  # Remove any NA values from the vector before passing to mk.test
  # mk.test can handle NAs, but it's good practice to be aware if many exist
  # If you have very few non-NA values, the test might not be meaningful.
  data_vector_clean <- na.omit(data_vector)
  
  # Check if there's enough data to perform the test (mk.test requires at least 4 observations)
  if (length(data_vector_clean) >= 4 && sd(data_vector_clean) > 0) { # sd > 0 ensures variability
    # Perform Mann-Kendall test
    test_result <- tryCatch({
      mk.test(data_vector_clean)
    }, error = function(e) {
      message(paste("Error running mk.test for column:", col_name, "-", e$message))
      return(NULL) # Return NULL if an error occurs
    })
    
    if (!is.null(test_result)) {
      # Extract relevant information, providing defaults if an estimate is missing
      tau_val <- if (!is.null(test_result$estimates) && "tau" %in% names(test_result$estimates)) test_result$estimates["tau"] else NA
      s_val <- if (!is.null(test_result$estimates) && "S" %in% names(test_result$estimates)) test_result$estimates["S"] else NA
      vars_val <- if (!is.null(test_result$estimates) && "varS" %in% names(test_result$estimates)) test_result$estimates["varS"] else NA
      p_value_val <- if (!is.null(test_result$p.value)) test_result$p.value else NA
      z_val <- if (!is.null(test_result$statistic)) test_result$statistic else NA
      
      mk_results[[col_name]] <- data.frame(
        Variable = col_name,
        Tau = tau_val,
        S = s_val,
        VarS = vars_val,
        p_value = p_value_val,
        Z = z_val,
        stringsAsFactors = FALSE
      )
    } else {
      message(paste("Skipping column", col_name, "due to test failure or insufficient data."))
      mk_results[[col_name]] <- data.frame(
        Variable = col_name,
        Tau = NA,
        S = NA,
        VarS = NA,
        p_value = NA,
        Z = NA,
        stringsAsFactors = FALSE
      )
    }
  } else {
    message(paste("Skipping column", col_name, "due to insufficient data points or no variability (all values are the same)."))
    mk_results[[col_name]] <- data.frame(
      Variable = col_name,
      Tau = NA,
      S = NA,
      VarS = NA,
      p_value = NA,
      Z = NA,
      stringsAsFactors = FALSE
    )
  }
}

# Check if any results were collected before attempting to bind
if (length(mk_results) > 0) {
  # Combine all results into a single data frame
  all_mk_results_df <- do.call(rbind, mk_results)
  
  # Define the output file path
  
  
  # Write the data frame to an Excel file
  write_xlsx(all_mk_results_df, "C:\\Users\\AdonsMB\\OneDrive - University of the Free State\\Desktop\\Masters\\MASTERS 2025\\2025\\Bloem Wo\\MannKendall_Results.xlsx")
  
  cat(paste0("Mann-Kendall test results saved to: ", "C:\\Users\\AdonsMB\\OneDrive - University of the Free State\\Desktop\\Masters\\MASTERS 2025\\2025\\Bloem Wo\\MannKendall_Results.xlsx", "\n"))
} else {
  message("No Mann-Kendall results to save. All columns might have been skipped.")
}
pettitt.test(GariepDam$JAN)
########################################################################################################
########################################################################################################
pettitt_results <- list()

# Loop through all columns except 'Year' and 'Year_numeric' to run the Pettitt test
# We'll skip the first column ('Year') and the 'Year_numeric' column.
# If you have other non-numeric columns, adjust the loop accordingly.
columns_to_test <- names(GariepDam)[!(names(GariepDam) %in% c("Year", "Year_numeric"))]

for (col_name in columns_to_test) {
  # Extract the data for the current column
  data_series <- GariepDam[[col_name]]
  
  # Run the Pettitt test
  pettitt_test_result <- pettitt.test(data_series)
  
  # Extract information
  p_value <- pettitt_test_result$p.value
  # The Pettitt test 'k' value indicates the *index* of the probable change point.
  # We need to map this index back to the 'Year' column.
 
  
  # Store the results in the list
  pettitt_results[[col_name]] <- data.frame(
    Column = col_name,
    Probable_Change_Year = NA,
    P_Value = p_value,
    Average_Before_Change = NA,
    Average_After_Change = NA
  )
}

# Combine all results into a single data frame
final_results_df <- do.call(rbind, pettitt_results)

write_xlsx(final_results_df, "C:\\Users\\AdonsMB\\OneDrive - University of the Free State\\Desktop\\Masters\\MASTERS 2025\\2025\\Bloem Wo\\Pettitt_Test_Results.xlsx")

cat("Pettitt test results saved to:", "C:\\Users\\AdonsMB\\OneDrive - University of the Free State\\Desktop\\Masters\\MASTERS 2025\\2025\\Bloem Wo\\Pettitt_Test_Results.xlsx", "\n")

################################################################################################
################################################################################################
################################################################################################
################################################################################################
################################################################################################
################################################################################################
################################################################################################
################################################################################################
################################################################################################
######BUISHAND TEST#################
lrt_results <- list()

# Get the names of the columns to test (excluding 'Year')
# Assuming 'Year' is named 'Year'
columns_to_test <- names(GariepDam)[!names(GariepDam) %in% c("Year")]

# Loop through each column and perform the structural change test
for (col_name in columns_to_test) {
  # Construct the formula for the test: dependent variable ~ 1 (intercept only model)
  
  formula_str <- as.formula(paste(col_name, "~ 1"))
  
  tryCatch({
    
    
    # Fit a linear model (intercept only) for each variable
    model <- lm(formula_str, data = GariepDam)
    
    
    test_result <- sctest(model)
    
    lrt_results[[col_name]] <- data.frame(
      Column = col_name,
      Statistic = test_result$statistic,
      P_Value = test_result$p.value,
      Method = test_result$method
    )
  }, error = function(e) {
    lrt_results[[col_name]] <- data.frame(
      Column = col_name,
      Statistic = NA,
      P_Value = NA,
      Method = paste("Error:", e$message)
    )
    warning(paste("Could not run sctest for column", col_name, ":", e$message))
  })
}

# Combine all results into a single data frame
results_df <- do.call(rbind, lrt_results)

# Save the results to an Excel file
output_file <- "C:\\Users\\AdonsMB\\OneDrive - University of the Free State\\Desktop\\Masters\\MASTERS 2025\\2025\\Bloem Wo\\Likelihood_Ratio_Test_Results.xlsx"
write_xlsx(results_df, output_file)

print(paste("Likelihood Ratio Test results (structural change tests) saved to:", output_file))

#############################################################################################################
#############################################################################################################

regression_results <- list()

# Loop through each column (excluding 'Year') and perform linear regression
for (col_name in names(GariepDam)) {
  if (col_name != "Year") {
    # Create the formula for the linear regression
    formula_str <- paste0("`", col_name, "` ~ Year") # Use backticks for column names with special characters or spaces if any
    
    # Run the linear regression model
    model <- lm(as.formula(formula_str), data = GariepDam)
    
    # Get the summary of the model
    model_summary <- summary(model)
    
    # Extract the slope (estimate for 'Year') and its p-value
    # The coefficients table has rows for (Intercept) and 'Year'
    # The columns are Estimate, Std. Error, t value, Pr(>|t|)
    
    # Check if 'Year' coefficient exists (it should, but good practice)
    if ("Year" %in% rownames(model_summary$coefficients)) {
      slope_estimate <- model_summary$coefficients["Year", "Estimate"]
      p_value <- model_summary$coefficients["Year", "Pr(>|t|)"]
      
      # Store the results
      regression_results[[col_name]] <- data.frame(
        Column_Name = col_name,
        Estimate_Slope = slope_estimate,
        P_Value = p_value
      )
    } else {
      # Handle cases where 'Year' might not be a valid predictor (e.g., if it's constant)
      regression_results[[col_name]] <- data.frame(
        Column_Name = col_name,
        Estimate_Slope = NA,
        P_Value = NA
      )
      warning(paste("Could not find 'Year' coefficient for column:", col_name))
    }
  }
}

# Combine all results into a single data frame
final_results_df <- do.call(rbind, regression_results)

# Define the output file path
output_file <- "C:\\Users\\AdonsMB\\OneDrive - University of the Free State\\Desktop\\Masters\\MASTERS 2025\\2025\\Bloem Wo\\Regression_Results.xlsx"

# Write the data frame to an Excel file
write_xlsx(final_results_df, path = output_file)

cat("Linear regression results saved to:", output_file, "\n")

##############################################################################################
#############################################################################################
# Prepare a data frame to store the MMK results
mmk_results <- data.frame(
  Variable = character(),
  Corrected_Zc = numeric(),
  New_P_value = numeric(),
  N_N_star = numeric(),
  Original_Z = numeric(),
  Old_P_value = numeric(),
  Tau = numeric(),
  Sen_Slope = numeric(),
  Old_Variance = numeric(),
  New_Variance = numeric(),
  stringsAsFactors = FALSE
)

# Loop through all columns (except 'Year') and apply Modified Mann-Kendall test
for (col_name in names(GariepDam1)) {
  if (col_name != "Year") { # Exclude the 'Year' column
    data_vector <- GariepDam1[[col_name]]
    
    # Ensure the data_vector is numeric and has no NA values
    if (is.numeric(data_vector) && !any(is.na(data_vector))) {
      # Perform the Modified Mann-Kendall test
      mmk_test_result <- mmkh(data_vector)
      
      # --- FIX STARTS HERE ---
      # Access elements by index instead of $ operator
      # The order of elements in the atomic vector output from mmkh is assumed to be:
      # 1: Corrected.Zc
      # 2: new.P.value
      # 3: N.N.
      # 4: Original.Z
      # 5: Old.P.value
      # 6: Tau
      # 7: Sens.Slope
      # 8: old.variance
      # 9: new.variance
      # You can verify this by running mmkh(your_vector) and inspecting the output directly.
      
      mmk_results <- rbind(mmk_results, data.frame(
        Variable = col_name,
        Corrected_Zc = mmk_test_result[1],
        New_P_value = mmk_test_result[2],
        N_N_star = mmk_test_result[3],
        Original_Z = mmk_test_result[4],
        Old_P_value = mmk_test_result[5],
        Tau = mmk_test_result[6],
        Sen_Slope = mmk_test_result[7],
        Old_Variance = mmk_test_result[8],
        New_Variance = mmk_test_result[9]
      ))
      # --- FIX ENDS HERE ---
      
    } else {
      warning(paste("Skipping column '", col_name, "' due to non-numeric data or NA values.", sep=""))
    }
  }
}

# The rest of your code remains the same for saving to Excel
# Define the output file path
output_file <- "C:\\Users\\AdonsMB\\OneDrive - University of the Free State\\Desktop\\Masters\\MASTERS 2025\\2025\\Bloem Wo\\MMK_Results.xlsx"

# Save the results to an Excel file
write.xlsx(mmk_results, file = output_file, rowNames = FALSE)

cat("Modified Mann-Kendall test results saved to:", output_file, "\n")