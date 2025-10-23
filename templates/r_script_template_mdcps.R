# Title:          Generic starter template

# Purpose:        Joy

#####################################
# COMMITTMENT ----
#####################################



#####################################
# NOTES ----
#####################################



#####################################
# TO DO ----
#####################################



#####################################
# OTHER DEPENDENCIES (e.g., files) ----
#####################################

# 1. ...
# 2. ...

#####################################
# LIBRARIES & UTILITIES ----
#####################################

# Load packages and generic functions
source("D:/repo_childmetrix/r_utilities/loader.R")

# Load functions specific to this project
# source(file.path(util_root, "project_specific", "cfsr_profile.R"), chdir = FALSE)

# Add helper to insert month column first
add_month_col <- function(df) {
  df %>%
    mutate(report_month = as.Date(paste0(folder_date, "_01"), format = "%Y_%m_%d"),
           .before = 1)
  }

########################################
# FOLDER PATHS & DIRECTORY STRUCTURE ----
########################################

# Base data folder
base_folder <- "D:/repo_mdcps_suspension_period/r_2.9.a/data"

# File name elements (e.g., 2024_01 - [commitment] - [commitment_description] - 2024-02-15.csv")
# e.g., save_to_folder_run(claiming_df)
commitment <- "2.9.a"
commitment_description <- "MIC Manual Review Verification"

# Establish current period and set up folders and global variables
my_setup <- setup_folders("2025_09")

# # Select quarter(s) to run
# quarters <- sprintf("2024_Q%d", 1:4)
#   for (q in quarters) {
#   setup_folders(q)

# Select months to run
# months <- sprintf("2024_%02d", 1:12)
#   for (m in months) {
#   setup <- setup_folders(m)
    
########################################
# LOAD FILES ----
########################################

# Absolute path and file example
# file_2024_q2 <- "S:/Folders/2024 - 2nd Continuing Suspension Period - MDCPS/2. Child Safety & Maltreatment in Care/2.8.b/Q2 2024/2.8.b.1 and 2.8.b.2 2024 Q2 2024 Export Reporting.xlsx"

# Absolute path and file string example (or change csv to xlsx)
# dir_path <- "S:/Folders/2024 - 2nd Continuing Suspension Period - MDCPS/9. Monitoring/9.6/Q4 2024"
# csv_file <- list.files(
#   path        = dir_path,
#   pattern     = "_expanded_population_details.*\\.csv$",
#   ignore.case = TRUE,
#   full.names  = TRUE
# )
# custody_df <- read_csv(csv_file)

# --------------------------------------
# Files in project folder
# --------------------------------------

# data_df <- find_file(keyword = "my file", "raw", file_type = "excel")

# --------------------------------------
# Eric files
# --------------------------------------

# Detail
base_path <- "S:/Folders/2024 - 2nd Continuing Suspension Period - Kurt/2.7"
file_string <- "Review 20"
load_sharefile_file(base_path, file_string, "eric_detail_df")

# DQ
base_path <- "S:/Folders/2024 - 2nd Continuing Suspension Period - Kurt/2.7"
file_string <- "DQ Results"
load_sharefile_file(base_path, file_string, "eric_dq_df", "Possibly Subject to Review")

# Custom when no quarterly subfolder
# dir_path <- "S:/Folders/2024 - 2nd Continuing Suspension Period - Kurt/6.3.b.2 & 6.3.b.3/2024 CY/2025-03-28"
# xlsx_file <- list.files(dir_path, pattern = "Referrals 20.*\\.xlsx$", ignore.case = TRUE, full.names = TRUE)
# sheet_name <- "Monitor Detail Data"  # or sheet_name <- NULL
# eric_detail_df <- if (!is.null(sheet_name)) {
#   read_excel(xlsx_file, sheet = sheet_name)
# } else {
#   read_excel(xlsx_file)
# }

# --------------------------------------
# MDCPS main files
# --------------------------------------

# Custody file (cumulative)
base_path <- "S:/Folders/2024 - 2nd Continuing Suspension Period - MDCPS/9. Monitoring/9.6"
file_string <- "_expanded_population_details"
load_sharefile_file(base_path, file_string, "custody_df")

# Allegations file (cumulative)
base_path <- "S:/Folders/2024 - 2nd Continuing Suspension Period - MDCPS/2. Child Safety & Maltreatment in Care/2.8.a"
file_string <- "_ane_"
load_sharefile_file(base_path, file_string, "allegations_df")

# --------------------------------------
# MDCPS commitment files
# --------------------------------------

########################################
# SANITIZE & DATA CLEANING
########################################
  
# Clean names with janitor
# mdcps_df <- mdcps_df %>% clean_names()

# Clean custody file (overlapping dates, dedup, etc.)
res <- process_custody(custody_df, folder_date_quarter, return_details = TRUE)
list2env(res, .GlobalEnv)


########################################
# CREATE FIELDS, FLAGS, CALCULATIONS ----
########################################

# --------------------------------------
# Detail file
# --------------------------------------



# --------------------------------------
# Summary file
# --------------------------------------


########################################
# CREATE AND SAVE QUARTERLY WORKBOOK & CUMULATIVE VERSION
########################################

sheets_to_save <- list(
  "Population Details"        = detail_df,
  "Manual Review Children"    = mr_children_df,
  "Manual Review Allegations" = mr_allegations_df
)

out <- run_quarter_and_cumulative(
  sheets_to_save           = sheets_to_save,
  base_data_dir            = base_data_dir,
  folder_date              = folder_date,              # e.g., "2025_Q2"
  commitment               = commitment,               # e.g., "2.9.a"
  commitment_description   = commitment_description,   # e.g., "MIC Manual Review Verification"
  folder_cumulative        = folder_cumulative,
  create_latest_copies     = FALSE
)

# Access paths if needed:
# out$quarterly_xlsx
# out$cumulative_xlsx

########################################
# COMPARE TO ERIC ----
########################################

# Define metrics
metrics <- list(
  all      = expr(TRUE),
  timely   = expr(action_timely == TRUE),
  untimely = expr(action_timely == FALSE)
)

output_file <- file.path(
  run_folder,
  paste(
    folder_date,
    commitment,
    commitment_description,
    "Comparison",
    format(run_date, "%Y-%m-%d"),
    sep = " - "
  ) %>% paste0(".xlsx")
)

# out_file <- file.path(
#   folder_processed,
#   paste0(folder_date, "_", commitment, "_compare_", today_date, ".xlsx")
# )

# Prep the two data sets if needed
detail_base <- detail_df %>% 
  filter(flag_error == FALSE) # %>%
# distinct(resource_id, .keep_all = TRUE)

eric_base <- eric_detail_df %>% 
  filter(exclude == FALSE) # %>%
# distinct(resource_id, .keep_all = TRUE)

# Run the comparison + write workbook
write_comparisons(
  df1       = detail_base,
  df2       = eric_base,
  id_cols   = c("resource_id"),
  metrics   = metrics,
  out_path  = out_file,
  df1_label = "kurt",
  df2_label = "eric"
)

}