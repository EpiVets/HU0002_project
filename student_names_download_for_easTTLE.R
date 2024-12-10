#################################################################################################
#################################################################################################

# HU0002 LEARNING MATTERS DATA EXTRACTION PROJECT

# EXTRACTION OF ALPHABETIC NAMES FROM GOOGLE DRIVE FILES SO LEARNING MATTERS CAN TIDY UP NAME SPELLING ISSUES WITH E-ASTTLE DATASET

#################################################################################################
#################################################################################################

# SETUP

library(tidyverse)
library(stringr)
library(lubridate)
library(gtools)
library(arsenal)
library(survival)
library(dplyr)
library(plyr)
library(readxl)
library(consort)
library(stringi)
library(gt)
library(gtsummary)
library(webshot2)
library(janitor)
library(consort)
library(flextable)
library(lmtest)
library(car)
library(emmeans)
library(openxlsx)
library(janitor)
library(ggpubr)
library(DescTools)
library(naniar)
library(googledrive)
library(writexl)
library(googlesheets4)


mixedrank <- function(x) order(gtools::mixedorder(x))


########################################################################################################################################################
########################################################################################################################################################
# Authenticate with Google Drive (if not already authenticated)
drive_auth()

# drive_deauth()  # De-authenticate first
# drive_auth(scopes = "https://www.googleapis.com/auth/drive.readonly")  # Request read-only access to Drive
# 
# gs4_deauth()  # De-authenticate first
# gs4_auth(scopes = "https://www.googleapis.com/auth/spreadsheets.readonly")  # Request read-only access to Sheets
# 
# drive_auth(scopes = "https://www.googleapis.com/auth/drive.readonly")
# gs4_auth(scopes = "https://www.googleapis.com/auth/spreadsheets.readonly")


# List files in "Learning Matters data for extraction/Master Spreadsheet ALL" folder. 
# Setting recursive = TRUE means it looks in subfolders.
files <- drive_ls(path = "Learning Matters data for extraction/Master Spreadsheet ALL", recursive = TRUE)

# Filter the list to include only .xlsx files
xlsx_files <- files[grepl(".xlsx", files$name), ]

# Build a dictionary of parent-child relationships
path_dict <- list()
for (i in 1:nrow(files)) {
  file <- files[i, ]
  path_dict[[file$id]] <- file$name
}

# Function to get the full path of a file (to track where it comes from)
get_full_path <- function(file) {
  full_path <- file$name
  current_id <- file$id
  
  while (TRUE) {
    parent_info <- drive_get(as_id(current_id))
    if (length(parent_info$drive_resource[[1]]$parents) == 0) break
    parent_id <- parent_info$drive_resource[[1]]$parents[[1]]
    parent_name <- path_dict[[parent_id]]
    full_path <- paste(parent_name, full_path, sep = "/")
    current_id <- parent_id
  }
  
  return(full_path)
}


# Function to download and read the 'Tier 2 Student Details' sheet, skipping files that don't have it
read_tier2_student_details_xlsx <- function(file) {
  # Extract file information
  file_id <- file$id
  file_name <- file$name
  file_path <- get_full_path(file)  # Assuming you have the get_full_path function from earlier
  
  # Create a temporary file path to download the xlsx file
  temp_file <- tempfile(fileext = ".xlsx")
  
  # Download the file from Google Drive
  drive_download(as_id(file_id), path = temp_file, overwrite = TRUE)
  
  # Check if "Tier 2 Student Details" sheet exists
  sheet_names <- excel_sheets(temp_file)
  if (!"Tier 2 Student Details" %in% sheet_names) {
    message(paste("Skipping file:", file_name, "- 'Tier 2 Student Details' sheet not found"))
    return(list(data = NULL, path = file_path))  # Return NULL and file path if the sheet is not present
  }
  
  # Read the 'Tier 2 Student Details' sheet
  sheet_data <- read_excel(temp_file, sheet = "Tier 2 Student Details")
  
  # Return the dataframe along with the file path
  return(list(data = sheet_data, path = file_path))
}


# xlsx_files2 <- xlsx_files[1:3, ]


########################################################################################################################################################

# Apply the function to each .xlsx file and store the results in a list
tier2_data_list <- lapply(1:nrow(xlsx_files), function(i) read_tier2_student_details_xlsx(xlsx_files[i, ]))



# Separate the data and paths
data_list <- lapply(tier2_data_list, function(x) x$data)
path_list <- lapply(tier2_data_list, function(x) x$path)

# Remove NULL entries (skipped files) from the data
data_list <- data_list[!sapply(data_list, is.null)]
path_list <- path_list[!sapply(data_list, is.null)]  # Keep the corresponding paths


# For each item in data_list, add a column with the corresponding path
for (i in seq_along(data_list)) {
  data_list[[i]]$path <- path_list[[i]]
}


# Combine all dataframes into a single dataframe
final_df <- bind_rows(data_list) %>%
  select(c(1:6, 22)) %>%
  # Set first row as column names
  setNames(.[1, ]) %>%
  clean_names() %>%
  # Remove rows where national_student_number == "National Student number"
  filter(national_student_number != "National Student number") %>%
  # Mutate national_student_number to numeric
  mutate(national_student_number = as.numeric(national_student_number)) %>%
  # Split student_name into first and last name columns
  #separate(student_name, into = c("Firstname", "Surname"), sep = " ", remove = FALSE) %>%
  separate(student_name, into = c("Firstname", "Surname"), sep = " ", extra = "merge", remove = FALSE) %>%
  dplyr::rename(Source = 9)

length(unique(final_df$student_name))
# Write to .xlsx file
write_xlsx(final_df, "student_names.xlsx")

# Check if all names in paste0(eas$real_firstname, eas$real_surname) appear in final_df$student_name 
paste(eas$real_firstname, eas$real_surname, sep = " ") %in% final_df$student_name

########################################################################################################################################################
########################################################################################################################################################

