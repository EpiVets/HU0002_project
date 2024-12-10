#################################################################################################
#################################################################################################

# HU0002 LEARNING MATTERS DATA EXTRACTION PROJECT

# DATA IMPORTATION FROM GOOGLE DRIVE

#################################################################################################
#################################################################################################

# SETUP

library(plyr)
library(tidyverse)
library(stringr)
library(lubridate)
library(gtools)
library(arsenal)
library(survival)
library(dplyr)
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

mixedrank <- function(x) order(gtools::mixedorder(x))


#################################################################################################
#################################################################################################

# This works in two stages: 1) extract data from the LM Google Drive account and save as .xlsx files locally, 2) process the .xlsx files to make the report.


# EXTRACT DATA FROM THE LM GOOGLE DRIVE ACCOUNT

# Authenticate with Google
drive_auth()

# List files in "Learning Matters data for extraction/Master Spreadsheet ALL" folder. Setting recursive = T means it looks in subfolders.
files <- drive_ls(path = "Learning Matters data for extraction/Master Spreadsheet ALL", recursive = T)

# Filter the list to include only .xlsx files that have more than one sheet
xlsx_files <- files[grepl(".xlsx", files$name), ]


# Create a local directory to store downloaded files
local_dir <- "downloaded_xlsx_files"
dir.create(local_dir, showWarnings = FALSE)

# Build a dictionary of parent-child relationships
path_dict <- list()

for (i in 1:nrow(files)) {
  file <- files[i, ]
  path_dict[[file$id]] <- file$name
}

# Function to get the full path of a file
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

# Function to anonymise student names
# Function to replace letters with their numeral equivalents
anonymise_name <- function(name) {
  alphabet <- c(letters, LETTERS)
  numerals <- c(1:26, 1:26)
  
  anonymised_name <- sapply(strsplit(name, NULL)[[1]], function(char) {
    if (char %in% alphabet) {
      numerals[which(alphabet == char)]
    } else {
      char # Leave non-letter characters unchanged
    }
  })
  
  return(paste0(anonymised_name, collapse = ""))
}

# Function to anonymize specific rows and columns in the sheet
anonymise_sheet <- function(wb, sheet, cols, rows) {
  data <- read.xlsx(wb, sheet = sheet)
  
  # Ensure rows and cols are within the bounds of the data frame
  for (col in cols) {
    # Check if column exists in the data frame
    if (col <= ncol(data)) {
      for (row in rows) {
        # Check if the row exists in the data frame
        if (row <= nrow(data)) {
          # Check if the cell is not NA and is a character string
          if (!is.na(data[row, col])) {
            if (nzchar(as.character(data[row, col]))) {
              data[row, col] <- anonymise_name(data[row, col])
            }
          }
        }
      }
    }
  }
  
  # Write back the modified data while preserving formatting
  writeData(wb, sheet = sheet, x = data, colNames = TRUE, rowNames = FALSE)
}

# List of sheets to keep (I will drop irrelevant sheets in case the contain information that identifies students) (not used because removing sheets upsets formatting!)
sheets_to_keep <- c("Baseline Assessment ORF", "Tier 2 Student Details", "Common Assessment (AB)", "Student Session Attendance")

#############################################################

# Function to process each .xlsx file
process_xlsx_file <- function(file) {
  # Extract file information
  file_id <- file$id
  file_name <- file$name
  file_path <- get_full_path(file)
  
  # Print the file information
  message(paste("Starting to process file", file_name, "in folder", file_path))
  
  # Determine where this file is in the list (i.e., its row number) so I can give it a unique identifier (many have the same name)
  file_index <- which(xlsx_files$id == file_id)
  
  # Download the file
  local_path <- file.path(local_dir, file_name)
  drive_download(as_id(file_id), path = local_path, overwrite = TRUE)
  
  
  # Load the .xlsx file
  wb <- loadWorkbook(local_path)
  
  
    # Use presence of "Baseline Assessment ORF" sheet to determine relevance
  prefix <- ifelse("Baseline Assessment ORF" %in% excel_sheets(path = local_path), "Relevant", "Irrelevant")
  
  
  # Identify sheets to clear the contents of in case they contain names (those not in sheets_to_keep). I cannot remove those sheets as it upsets formatting.
  # Get all sheets in the workbook
  all_sheets      <- excel_sheets(path = local_path)
  sheets_to_clear <- setdiff(all_sheets, sheets_to_keep)
  
  # Remove sheets with "Salisbury" in their name (error caused by Suzanne Pentecost and Shona McInnes)
  salisbury_sheets_to_remove <- grep("Salisbury", all_sheets, value = TRUE)
  
  for (sheet in salisbury_sheets_to_remove) {
    message(paste("Removing sheet:", sheet))
    removeWorksheet(wb, sheet)
  }

  
  # Loop through the sheets to anonymise
  # for (sheet in sheets_to_anonymise) {
  #   if (sheet %in% names(wb)) {
  #     # Read data from the sheet
  #     data <- tryCatch({
  #       read.xlsx(wb, sheet = sheet)
  #     }, error = function(e) {
  #       message(paste("Error reading sheet:", sheet, "-", e$message))
  #       return(NULL)
  #     })
  #     
  #     # Check if data is NULL or empty
  #     if (!is.null(data) && nrow(data) > 0) {
  #       # Replace all cells with NA
  #       data[,] <- NA
  #       
  #       # Write back the modified data while preserving formatting
  #       writeData(wb, sheet = sheet, x = data, colNames = TRUE, rowNames = FALSE)
  #     } else {
  #       message(paste("No data found in sheet:", sheet))
  #     }
  #   }
  # }
  message(paste("Clearing data from irrelevant sheets"))
  for (sheet in sheets_to_clear) {
    if (sheet %in% names(wb)) {
      message(paste("Working on sheet:", sheet))
      # Read data from the sheet
      data <- tryCatch({
        read.xlsx(wb, sheet = sheet)
        }, error = function(e) {
        message(paste("Error reading sheet:", sheet, "-", e$message))
        return(NULL)
      })
      
      # Check if data is NULL or empty
      if (!is.null(data) && nrow(data) > 0) {
        message(paste("Data found in sheet:", sheet))
        # Get the dimensions of the data
        rows <- nrow(data)
        cols <- ncol(data)
        
        # Create a new empty matrix with the same dimensions
        empty_data <- matrix(NA, nrow = rows, ncol = cols)
        
        # Write the empty data back to the sheet
        message(paste("Overwriting sheet:", sheet))
        writeData(wb, sheet = sheet, x = empty_data, colNames = FALSE, rowNames = FALSE)
      } else {
        message(paste("No data found in sheet:", sheet))
      }
    }
  }

  # Repeat to tidy up
  message(paste("Repeat clearing process"))
  for (sheet in sheets_to_clear) {
    if (sheet %in% names(wb)) {
      message(paste("Working on sheet:", sheet))
      # Read data from the sheet
      data <- tryCatch({
        read.xlsx(wb, sheet = sheet)
        }, error = function(e) {
        message(paste("Error reading sheet:", sheet, "-", e$message))
        return(NULL)
      })
      
      # Check if data is NULL or empty
      if (!is.null(data) && nrow(data) > 0) {
        message(paste("Data found in sheet:", sheet))
        # Get the dimensions of the data
        rows <- nrow(data)
        cols <- ncol(data)
        
        # Create a new empty matrix with the same dimensions
        empty_data <- matrix(NA, nrow = rows, ncol = cols)
        
        # Write the empty data back to the sheet
        message(paste("Overwriting sheet:", sheet))
        writeData(wb, sheet = sheet, x = empty_data, colNames = FALSE, rowNames = FALSE)
      } else {
        message(paste("No data found in sheet:", sheet))
      }
    }
  }
  

  # Anonymise students by replacing upper or lower case letters in names with the equivalent number in the alphabet. Leave non-alphabetic characters and spaces unchanged.
  message(paste("Anonymising student names in relevant sheets"))
  
  # Define the sheets and columns where student names are located
  sheets_with_names <- c("Baseline Assessment ORF", "Tier 2 Student Details", "Common Assessment (AB)", "Student Session Attendance")

  # Define the mapping for sheets, columns, and rows
  sheets_info <- list(
    "Baseline Assessment ORF" = list(cols = 1, rows = 2:1000),  
    "Tier 2 Student Details" = list(cols = 2, rows = 2:1000),
    "Common Assessment (AB)" = list(cols = 2, rows = 2:1000),
    "Student Session Attendance" = list(cols =  5:10, rows = 2)
  )
  
  # Loop through the sheets_info to anonymise names
  for (sheet in names(sheets_info)) {
    message(paste("Working on sheet:", sheet))
    if (sheet %in% names(wb)) {
      cols <- sheets_info[[sheet]]$cols
      rows <- sheets_info[[sheet]]$rows
      anonymise_sheet(wb, sheet, cols, rows)
    }
  }
  
  


  # Add a new sheet with the folder path so it can be tracked
  message(paste("Adding sheet with file path"))
  addWorksheet(wb, "File Path")
  writeData(wb, sheet = "File Path", x = file_path)
  
  # Add a new sheet with the sheets removed for tracking
  message(paste("Adding sheet with sheets removed"))
  addWorksheet(wb, "Sheets Removed")
  writeData(wb, sheet = "Sheets Removed", x = salisbury_sheets_to_remove)
  
  
  # Save the workbook
  message(paste("Saving the workbook"))
  saveWorkbook(wb, local_path, overwrite = TRUE)
  
  # Return the path of the modified file
  #return(new_local_path)
  
  # Rename the file to include a unique identifier
  message(paste("Renaming the file to include a unique identifier"))
  unique_file_name  <- paste0(prefix, "_", file_index, "_", file_name)
  unique_local_path <- file.path(local_dir, unique_file_name)
  file.rename(local_path, unique_local_path)
  
  
  # Return the path of the modified file
  return(unique_local_path)  
}





# Use a reduced dataset for practice
# xlsx_files <- xlsx_files[1:5, ]
# xlsx_files2 <- xlsx_files %>% filter(name %in% c("Master SS -Amie Campbell  Hira School.xlsx", "Copy of Master SS - Salisbury - Suzanne Pentecost and Shona McInnes.xlsx"))

# Process each .xlsx file and store the paths of modified files
modified_files <- sapply(1:nrow(xlsx_files), function(i) process_xlsx_file(xlsx_files[i, ]))

###################################################################################################################################################
###################################################################################################################################################

# # COLLATE XLSX FILES INTO A SINGLE DATAFRAME CONTAINING RELEVANT DATA
# 
# #1. Read in files in the folder "downloaded_xlsx_files2" that contain "Relevant" in the name and extract the sheets called "Tier 2 Student Details" and "File Path".
# # Then add a column called "Region" to the sheet called "Tier 2 Student Details" that contains the text betwen the first "/" and the characters " ALL" in the "File Path" sheet.
# 
# # Define the path to the folder containing the .xlsx files
# folder_path <- "downloaded_xlsx_files"
# 
# # List all .xlsx files in the folder that contain "Relevant" in the name, BUT REMOVE FILES WITH "COPY" IN THE NAME
# files2 <- list.files(path = folder_path, pattern = "Relevant.*\\.xlsx$", full.names = TRUE) %>%
#   setdiff(list.files(path = folder_path, pattern = "Copy", full.names = TRUE))
# 
# # List of files not in files2
# excluded_files_copies <- list.files(path = folder_path, pattern = "Relevant.*\\.xlsx$", full.names = TRUE) %>%
#   setdiff(files2)
# # Extract file paths from the "File Path" sheet of files in excluded_files_copies and save as a data frame
# excluded_files_copies_paths <- lapply(excluded_files_copies, function(file) {
#   wb <- loadWorkbook(file)
#   file_path <- read.xlsx(wb, sheet = "File Path")
#   return(file_path)
# })
# 
# length(files2)
# 
# # Initialize an empty list to store data frames
# data_list <- list()
# 
# # Create a vector to hold name of spreadsheet being worked on in case of a bug
# spreadsheet_name <- c()
# 
# 
# 
# ###################################################################################################################################################
# # Loop through each file
# for (file in files2) {
#   
#   message(noquote(paste0("Spreadsheet = ", "'", file, "'")))
#   
#   spreadsheet_name <- file
#   
#   # Function for tidying colnames
#   replace_x_colnames <- function(colnames) {
#     for (i in seq_along(colnames)) {
#       if (startsWith(colnames[i], "x")) {
#         # Find the previous column name that does not start with "x"
#         for (j in (i-1):1) {
#           if (!startsWith(colnames[j], "x")) {
#             colnames[i] <- colnames[j]
#             break
#           }
#         }
#       }
#     }
#     return(colnames)
#   }
#   ##############################################################################################
#   # Read the "Tier 2 Student Details" sheet
#   print(noquote(paste0("Reading sheet 'Tier 2 Student Details'")))
#   tier2_details <- read_excel(file, sheet = "Tier 2 Student Details", col_types = c("text")) %>%
#     clean_names()
# 
#   
#   # Tidy up colnames
#   new_colnames <- replace_x_colnames(colnames(tier2_details))
#   colnames(tier2_details) <- new_colnames
#   
#   # Get first row values to add to colnames
#   first_row <- tier2_details[1, ]
#   
#   # Modify column names by appending the first row values
#   colnames(tier2_details) <- paste(colnames(tier2_details), first_row, sep = "_")
#   
#   # Remove the first row
#   tier2_details <- tier2_details[-1, ]
#   
#   # Replace spaces in colnames with underscores
#   colnames(tier2_details) <- str_replace_all(colnames(tier2_details), " ", "_")
#     
#   
#   # Read the "File Path" sheet and convert to a vector
#   file_path_vector <- read_excel(file, sheet = "File Path") %>%
#     colnames()
#   
#   print(noquote(paste0("File path = ", "'", file_path_vector, "'")))
#   
#   # Check for columns that have been duplicated and remove said columns
#   # tier2_details <- tier2_details %>%
#   #   select(-contains("x"))
#   
#   # Remove duplicate columns if table contains data, otherwise leave it as is
#   if (nrow(tier2_details) > 0) {
#     tier2_details <- tier2_details[, !duplicated(names(tier2_details))]
#   }
#   
#   
#   ##############################################################################################
#   # Extract the "Region" from the "File Path" sheet.
#   print(noquote(paste0("Exracting region from 'File Path' sheet")))
#   region <- case_when(grepl("Canterbury", file_path_vector) ~ "Canterbury",
#                       grepl("Hawkes", file_path_vector) ~ "Hawkes Bay",
#                       grepl("Manawatu", file_path_vector) ~ "Manawatu/Rangitikei/Whanganui",
#                       grepl("Nelson", file_path_vector) ~ "Nelson/Marlborough/West Coast")
#   
#   # Extract teacher from the "File Path" sheet
#   teacher <- sub("^(.*/)([^/]*/[^/]+)$", "\\2", file_path_vector)
#   
#   # Remove text including and after "/" in teacher
#   teacher <- sub("/.*", "", teacher)
#   
#   ##############################################################################################
#   # Extract "neurodiverse_diagnosis"
#   print(noquote(paste0("Extracting 'neurodiverse_diagnosis' and 'neurodiverse_suspicion'")))
#   neurodiverse_diagnosis <- tier2_details %>%
#     dplyr::select(student_details_Student_Name, starts_with("neuro_diverse_diagnosis"), -neuro_diverse_diagnosis_Ethnicity) %>%
#     # Rename columns with only the characters after the last "_"
#     rename_with(~str_extract(., "(?<=_)[^_]+$"), .cols = everything()) %>%
#     pivot_longer(cols = !Name, names_to = "Diagnosis", values_to = "value") %>%
#     filter(!is.na(value) & !tolower(value) %in% c("no", "none", "n/a", "n", "x")) %>%  # Exclude "No", "no", "None", "none", "N/A", "N", "x"
#     group_by(Name) %>%
#     mutate(neurodiverse_diagnosis = paste(Diagnosis, collapse = ", ")) %>%
#     dplyr::select(Name, neurodiverse_diagnosis) %>%
#     slice_head(n=1) %>%
#     ungroup()
#   
#   neurodiverse_suspicion <- tier2_details %>%
#     dplyr::select(student_details_Student_Name, starts_with("neuro_diverse_suspicion")) %>%
#     rename_with(~str_extract(., "(?<=_)[^_]+$"), .cols = everything()) %>%
#     pivot_longer(cols = -Name, names_to = "Suspicion", values_to = "value") %>%
#     filter(!is.na(value) & !tolower(value) %in% c("no", "none", "n/a", "n", "x")) %>%  # Exclude "No", "no", "None", "none", "N/A", "N", "x"
#     group_by(Name) %>%
#     mutate(neurodiverse_suspicion = paste(Suspicion, collapse = ", ")) %>%
#     dplyr::select(Name, neurodiverse_suspicion) %>%
#     slice_head(n=1) %>%
#     ungroup()
#   
#   
#   
#   # Add "Region" and "Teacher" columns to the "Tier 2 Student Details" data frame and add neurodiverse diagnosis and suspicion
#   tier2_details <- tier2_details %>% 
#     mutate(Region_tier2 = region, Teacher_tier2 = teacher) %>%
#     left_join(neurodiverse_diagnosis %>% dplyr::select(Name, neurodiverse_diagnosis), by = c("student_details_Student_Name" = "Name")) %>%
#     left_join(neurodiverse_suspicion %>% dplyr::select(Name, neurodiverse_suspicion), by = c("student_details_Student_Name" = "Name"))
#   
#   # Edit tier2_details variables 
#   tier2_details <- tier2_details %>%
#     mutate(student_details_National_Student_number = as.numeric(student_details_National_Student_number)) %>%
#     mutate(student_details_Year_level = as.integer(student_details_Year_level)) %>%
#     mutate(student_details_Gender = case_when(student_details_Gender=="m" ~ "M",
#                                               student_details_Gender=="Male" ~ "M",
#                                               student_details_Gender=="male" ~ "M",
#                                               student_details_Gender=="f" ~ "F",
#                                               grepl("fem|Fem", student_details_Gender) ~ "F",
#                                               .default = student_details_Gender)) %>%
#     dplyr::rename(student_details_National_Student_number_tier2 = student_details_National_Student_number)
#   
#   ##############################################################################################
#   
#   # Extract adaptive Bryant score from the "Common Assessment (AB)" sheet
#   print(noquote(paste0("Extracting adaptive Bryant score from 'Common Assessment (AB)' sheet")))
#   ab <- read_excel(file, sheet = "Common Assessment (AB)") %>%
#     clean_names()
#   
#   # Tidy up colnames
#   new_colnames <- replace_x_colnames(colnames(ab))
#   colnames(ab) <- new_colnames
#   
#   # Get first row values to add to colnames
#   first_row <- ab[1, ]
#   
#   # Modify column names by appending the first row values
#   colnames(ab) <- paste(colnames(ab), first_row, sep = "_")
#   
#   # Remove the first row
#   ab <- ab[-1, ]
#   
#   # Replace spaces in colnames with underscores
#   colnames(ab) <- str_replace_all(colnames(ab), " ", "_")
#   
#   # Tidy up NSN and add region, teacher and neurodiverse_diagnosis
#   ab <- ab %>%
#     mutate(student_details_National_Student_number = as.numeric(student_details_National_Student_number)) %>%
#     dplyr::rename(student_details_National_Student_number_ad_bry = student_details_National_Student_number) %>%
#     mutate(Region_ad_bry = region, Teacher_ad_bry = teacher) 
#   
#   ##############################################################################################
#   
#   # Extract student attendance data from the "Student Session Attendance" sheet
#   print(noquote(paste0("Extracting student attendance data from 'Student Session Attendance' sheet")))
#   ssa <- read_excel(file, sheet = "Student Session Attendance", skip = 1) %>%
#     clean_names() %>%
#     dplyr::select(-no_of_sessions_wk)
#   
#   # Check if there are any attendance data
#   attendance_data <- ssa[-1,] %>%
#     pivot_longer(cols = starts_with("student_name"), names_to = "Name", values_to = "n_sessions")
#   
#   attendance_data <- ifelse(sum(!is.na(attendance_data$n_sessions)) > 0, "Yes", "No")
#   
#   # If attendance_data == "No", create a null dataframe containing the columns "Name", "Total available", and "Total attended" containing one row of NA values
#   # If attendance_data == "Yes", run the code below.
#   
#   if (attendance_data == "No") {
#     
#     ssa <- data.frame(`Name` = NA, `Total available` = NA, `Total attended` = NA)
#     colnames(ssa) <- c("Name", "Total available", "Total attended")
#     
#   } else if (attendance_data == "Yes") {
#     
#     # Step 1: Extract the first row to use as new column names
#     new_colnames <- sapply(ssa[1, ], function(x) ifelse(is.na(x), NA, as.character(x)))
#     
#     # Step 2: Replace existing column names where the first row is not NA
#     colnames(ssa) <- ifelse(is.na(new_colnames), colnames(ssa), new_colnames)
#     
#     # Remove duplicate column names
#     ssa <- ssa[, !duplicated(names(ssa))]
#     
#     # Remove columns with names starting with "x" or "student_name"
#     ssa <- ssa %>%
#       select(-starts_with("x"), -starts_with("student_name"))
#     
#     # Remove the first row from the dataframe
#     ssa <- ssa[-1, ]
#     
#       # Tidy up sessions (some entered as numbers, others as days)
#     convert_sessions <- function(sessions) {
#       if (is.na(sessions)) {
#         return(NA_integer_)
#       } else if (str_detect(sessions, "(?i)^o\\s*-\\s*overseas$")) {
#         # If the string is "O - overseas" (case insensitive), return 0
#         return(0)
#       } else if (str_detect(sessions, "(?i)away")) {
#         # If the string contains "Away" or "away", return 0
#         return(0)
#       } else if (str_detect(sessions, "^✔+$")) {
#         # If it contains only check marks, count them
#         return(str_count(sessions, "✔"))
#       } else if (str_detect(sessions, "\\d+")) {
#         # If it contains a number, extract and convert to numeric
#         return(as.numeric(str_extract(sessions, "\\d+")))
#       } else {
#         # If it is text, count the number of items and convert to integer
#         days <- str_split(sessions, ",\\s*")[[1]]
#         return(length(days))
#       }
#     }
#     
#     
#     
#       ssa <- ssa %>%
#         mutate(sessions_taught_per_week = sapply(sessions_taught_per_week, convert_sessions))
# 
#       # Pivot longer the columns containing after sessions_taught_per_week
#       # Find the position of the 'sessions_taught_per_week' column
#       sessions_taught_per_week_position <- which(names(ssa) == "sessions_taught_per_week")
# 
#       ssa <- ssa %>%
#         pivot_longer(cols = (sessions_taught_per_week_position + 1):ncol(ssa), names_to = "Name", values_to = "n_sessions") %>%
#         mutate(n_sessions = sapply(n_sessions, convert_sessions)) %>%
#         #mutate(n_sessions = as.numeric(n_sessions)) %>%
#         select(Name, sessions_taught_per_week, n_sessions) %>%
#         group_by(Name) %>%
#         dplyr::summarise(`Total available` = sum(sessions_taught_per_week, na.rm = TRUE), `Total attended` = sum(n_sessions, na.rm = TRUE))
#      
#       colnames(ssa) <- c("Name", "Total available", "Total attended")
#       
#       
#       
#     }
#   
#   
#   ##############################################################################################
#   
#   # Join "final_adapted_bryant_assessment_Overall_Score_/50" to tier2_details using "student_details_Student_Name", keeping all
#   print(noquote(paste0("Joining 'final_adapted_bryant_assessment_Overall_Score_/50' to 'Tier 2 Student Details'")))
#   data <- full_join(tier2_details, ab, by = "student_details_Student_Name") %>%
#     mutate(National_Student_number = ifelse(is.na(student_details_National_Student_number_tier2), 
#                                             student_details_National_Student_number_ad_bry,
#                                             student_details_National_Student_number_tier2)) %>%
#     mutate(Region = ifelse(is.na(Region_tier2), Region_ad_bry, Region_tier2)) %>%
#     mutate(Teacher = ifelse(is.na(Teacher_tier2), Teacher_ad_bry, Teacher_tier2)) %>%
#     
#     # Split student name into first name and last name using the first space as the delimiter
#     # This is for joining e-asTTLE data, which is split into separate columns for firstname and surname.
#     separate(student_details_Student_Name, into = c("Firstname", "Surname"), sep = " ", extra = "merge", remove = FALSE) %>%
#     
#     # Handle missing Surnames by replacing NA with "Surname not provided"
#     mutate(Surname = replace_na(Surname, "Surname not provided")) %>%
#     
#     # Add a new School column with NA
#     mutate(School = NA) %>%
#     
#     # Relocate columns for a specific order
#     dplyr::relocate(Firstname, Surname, Region, School, Teacher, student_details_Year_level, 
#                     student_details_Gender, student_details_Ethnicity_1, student_details_Ethnicity_2)
#   
#   
# 
#   # Join the student session attendance data
#   data <- left_join(data, ssa, by = c("student_details_Student_Name" = "Name")) %>%
#     # Add column capturing the source of the data
#     mutate(Source = spreadsheet_name)
# 
#   # Append the modified data frame to the list
#   data_list <- append(data_list, list(data))
# }
# 
# 
# ###########################################################################################################################################
# ###########################################################################################################################################
# 
# # Combine all data frames together by row
# length(data_list)
# combined_data <- bind_rows(data_list)
# length(unique(combined_data$Teacher)) # was 86 teachers, now 80
# # combined_data$student_details_Student_Name
# 
# # Identify rows with the same value for "Firstname" but different values for "Surname" and one of the values for "Surname" is "Surname not provided"
# combined_data <- combined_data %>%
#   group_by(Firstname, Teacher) %>%
#   mutate(same_child = ifelse(n_distinct(Surname) > 1 & any(Surname == "Surname not provided"), 1, 0)) %>%
#   ungroup() %>%
#   relocate(same_child)
# 
# # If same_child==1, replace "Surname not provided" with the other value of "Surname"
# combined_data <- combined_data %>%
#   mutate(Surname = ifelse(Surname == "Surname not provided", NA, Surname)) %>%
#   group_by(Firstname, Teacher) %>%
#   tidyr::fill(Surname, .direction = "updown") %>%
#   ungroup()
# 
# View(combined_data %>% filter(same_child == 1) %>% arrange(Firstname, Surname))
# View(combined_data %>% filter(Firstname=="1051919931" & Surname=="1714523"))
# 
# combined_data <- combined_data %>%
#   group_by(Firstname, Surname, Teacher) %>%
#   fill(everything(), .direction = "downup") %>%
#   mutate(student_details_Student_Name = paste(Firstname, Surname, sep = " ")) %>%
#   distinct() %>%
#   ungroup()
# 
# 
# # Deal with Fiona Wilder
# anonymise_name("Fiona"); anonymise_name("Wilder")
# View(combined_data[combined_data$Firstname == "6915141" & combined_data$Surname=="239124518" & !is.na(combined_data$Firstname), ])
# # Was 51, now 45 Fiona Wilders
# 
# # Tidy up rogue variables and remove Fiona Wilder
# combined_data <- combined_data %>%
#   mutate(student_details_Gender = ifelse(grepl("Fem|fem", student_details_Gender), "F", student_details_Gender)) %>%
#   # Remove characters including and after the "/" in columns containing "Overall_Score_/50"
#   mutate_at(vars(contains("Overall_Score_/50")), ~str_remove(., "/.*")) %>%
#   # Convert same columns to numeric
#   mutate_at(vars(contains("Overall_Score_/50")), as.numeric) %>%
#   filter(!(coalesce(Firstname, "") == "6915141" & coalesce(Surname, "") == "239124518")) # Use coalesce to handle NA values!
# 
# 
# table(combined_data$student_details_Year_level)
# 
# ###########################################################################################################################################
# ###########################################################################################################################################
# 
# # IMPORT E-ASTTLE DATA AND MERGE WITH COMBINED DATA
# # NB student_name is the correct name
# eas <- read_excel("e-asTTLe Data Term 4 Beg_End.xlsx", sheet = 1) %>%
#   clean_names() %>%
#   dplyr::rename(east = "overall_scale_score") %>%
#   #dplyr::select(Firstname, Surname, east) %>%
#   # Identify duplicate rows
#   group_by(student_name) %>%
#   mutate(duplicate = n() > 1) %>%
#   ungroup() %>%
#   # Remove whitespace from east and convert to numeric (values of "absent" become NA)
#   mutate(east = as.numeric(str_remove_all(east, "\\s+")))
# 
# 
# # Anonymise names in eas
# eas <- eas %>%
#   # mutate(real_firstname = Firstname, real_surname = Surname) %>%
#   mutate(student_name = sapply(student_name, anonymise_name)) %>%
#   # Identify rows that appear in combined_data using student_name
#   mutate(in_google_drive_data = ifelse(student_name %in% combined_data$student_details_Student_Name, "Yes", "No")) %>%
#   arrange(in_google_drive_data)
# 
# # anonymise_name("Flynn Brydon")
# 
# 
# # table(eas$in_google_drive_data)
# # Quantify number without a value for east
# # eas %>%
# # select(east) %>%
# #   filter(is.na(east)) %>%
# #   nrow() 
# 
# # 1+96+16-2+11-8
# 
# 
# # Export .xlsx for LM to track down missing students
# # write_xlsx(eas, "eas.xlsx")
# 
# # Remove the duplicated row for 13251 4122919-201212381185 by taking only the first iteration (not applicable after LM tidied this up)
# # eas <- eas %>% 
# #   group_by(Firstname, Surname) %>%
# #   slice_head(n=1) %>%
# #   ungroup() %>%
# #   select(-duplicate)
# 
# 
# # Merge the e-asTTLe data with the combined data
# combined_data <- left_join(combined_data, eas %>% select(student_name, east), by = c("student_details_Student_Name" = "student_name"))
# 
# combined_data %>%
#   select(east) %>%
#   filter(!is.na(east)) %>%
#   nrow()
# 
# combined_data %>%
#   select(east) %>%
#   filter(east == 0) %>%
#   nrow()
# 
# 
# # 
# # combined_data %>%
# #   select(east) %>%
# #   filter(!is.na(east)) %>%
# #   nrow()
# 
# # Identify rows of eas that did not merge with combined_data
# eas_not_merged <- eas %>%
#   anti_join(combined_data, by = c("student_name" = "student_details_Student_Name"))
# # Chananthida Buntha not in main dataset.
# 
# ###########################################################################################################################################
# ###########################################################################################################################################
# 
# # Create the final dataframe needed by LM
# DATA <- combined_data %>%
#   dplyr::select(Firstname, Surname, student_details_Student_Name, Region, School, Teacher, student_details_Year_level, student_details_Gender, student_details_Ethnicity_1, student_details_Ethnicity_2,
#                 neurodiverse_diagnosis, neurodiverse_suspicion, `Total available`, `Total attended`, 
#                 `pre_adapted_bryant_assessment_Overall_Score_/50`, `mid_adapted_bryant_assessment_Overall_Score_/50`, 
#                 `final_adapted_bryant_assessment_Overall_Score_/50`, east) %>%
#   dplyr::rename(`Year level` = student_details_Year_level, `Ethnicity 1` = student_details_Ethnicity_1, `Ethnicity 2` = student_details_Ethnicity_2, 
#                 Gender = student_details_Gender,
#                 `Neurodiversity diagnosis` = neurodiverse_diagnosis, `Neurodiversity suspicion` = neurodiverse_suspicion,
#                 `Pre AB score` = `pre_adapted_bryant_assessment_Overall_Score_/50`, `Mid AB score` = `mid_adapted_bryant_assessment_Overall_Score_/50`, 
#                 `Final AB score` = `final_adapted_bryant_assessment_Overall_Score_/50`,
#                 `e-asTTle overall scaled score` = east) %>%
#   # If Gender == "M (F)", replace with "M" 
#   mutate(Gender = ifelse(Gender=="M (F)", "M", Gender))
# 
# # sum(!is.na(DATA$`e-asTTle overall scaled score`))
# # length(unique(paste0(DATA$Firstname, DATA$Surname)))
# 
# table(DATA$`Year level`, useNA = "always")
# table(DATA$Gender, useNA = "always")
# 
# ###########################################################################################################################################
# ###########################################################################################################################################
# 
# # ETHNICITY SHEET
# 
# # Ethnicity table - NB this only takes the first ethnicity value
# ethnicity <- DATA %>% 
#   dplyr::select(Gender, `Year level`, `Ethnicity 1`) %>%
#   dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
#   mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
#                                grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
#                                grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
#                                is.na(Ethnicity) ~ "Unknown",
#                                .default = "Other"))) %>%
#   mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
#   mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#   mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#   # If Gender == "M (F)", replace with "M" 
#   #mutate(Gender = ifelse(Gender=="M (F)", "M", Gender)) %>%
#   mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
#   mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown")))
#   
# 
# table(ethnicity$Gender, useNA = "always")
#   
# # tbl_strata(data = ethnicity,
# #            strata = Gender,
# #            .tbl_fun =
# #              ~ .x %>%
# #              tbl_summary(by = Ethnicity, 
# #                          missing = "no",
# #                          statistic = list(all_categorical() ~ "{n}/{N} ({p}%)"),
# #                          percent = "row") %>%
# #              bold_labels() %>%
# #              modify_header(label = ""),
# #            .header = "**{strata}**, N = {n}") %>%
# #   modify_footnote(everything() ~ NA) %>%
# #   as_hux_xlsx("ethnicity_summary.xlsx")
# 
# 
# # Use dplyr instead
# 
# eth_tbl_n <- ethnicity %>%
#   group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
#   dplyr::summarise(n = n(), .groups = "drop") %>%
#   ungroup() %>%
#   pivot_wider(names_from = c(Ethnicity, Gender), values_from = "n") %>%
#   mutate(Total = rowSums(.[-1])); eth_tbl_n
# 
# tot_row <- c(NA, colSums(eth_tbl_n[, -1], na.rm = T))
#   
# eth_tbl_n <- rbind(eth_tbl_n, tot_row) %>%
#   mutate(`Year level` = as.character(`Year level`)) %>%
#   mutate(`Year level` = factor(ifelse(is.na(`Year level`), "Total", `Year level`))) %>%
#   # Make first column row names
#   column_to_rownames(var = "Year level")
# 
# # Calculate row sums for all columns except the last one
# row_sums <- rowSums(eth_tbl_n[, -ncol(eth_tbl_n)])
# 
# # Calculate percentages for each cell by dividing by the row sum (excluding the last column)
# percentages <- round(sweep(eth_tbl_n[, -ncol(eth_tbl_n)], 1, row_sums, FUN = "/") * 100, 0)
# 
# # Calculate percentages for the last column as percentages of the total
# total_percentage <- round((eth_tbl_n[, ncol(eth_tbl_n)] / sum(eth_tbl_n[-length(eth_tbl_n$Total), ]$Total)) * 100, 0)
# 
# # Create a new dataframe to store combined counts and percentages
# df_combined <- eth_tbl_n
# 
# # Loop through columns (excluding the last one)
# for (col in names(eth_tbl_n)[-ncol(eth_tbl_n)]) {
#   # Create formatted string with count and percentage
#   combined_col <- paste0(eth_tbl_n[[col]], " (", round(percentages[[col]], 1), "%)")
#   # Assign the formatted string to the column in the combined dataframe
#   df_combined[[col]] <- combined_col
# }
# 
# # Create formatted string for the Total column
# combined_total <- sprintf("%d (%.1f%%)", eth_tbl_n$Total, total_percentage)
# df_combined$Total <- combined_total
# 
# colnames(df_combined) <- c("Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Total")
# 
# # Set rownames as first column
# ethn_all <- df_combined %>%
#   rownames_to_column(var = "Year level") %>%
#   mutate(Region = "All regions") %>%
#   dplyr::relocate(Region)
# 
# ###########################################################################################################################################
# # Now repeat for Canterbury, Hawkes Bay, Manawatu/Rangitikei/Whanganui, Nelson/Marlborough/West Coast
# 
# #1. Canterbury
# ethnicity <- DATA %>% 
#   dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`) %>%
#   filter(Region=="Canterbury") %>%
#   dplyr::select(-Region) %>%
#   dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
#   mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
#                                          grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
#                                          grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
#                                          is.na(Ethnicity) ~ "Unknown",
#                                          .default = "Other"))) %>%
#   mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
#   mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#   mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#   mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
#   mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown")))
# 
# eth_tbl_n <- ethnicity %>%
#   group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
#   dplyr::summarise(n = n(), .groups = "drop") %>%
#   ungroup() %>%
#   pivot_wider(names_from = c(Ethnicity, Gender), values_from = "n") %>%
#   mutate(Total = rowSums(.[-1])); eth_tbl_n
# 
# tot_row <- c(NA, colSums(eth_tbl_n[, -1], na.rm = T))
# 
# eth_tbl_n <- rbind(eth_tbl_n, tot_row) %>%
#   mutate(`Year level` = as.character(`Year level`)) %>%
#   mutate(`Year level` = factor(ifelse(is.na(`Year level`), "Total", `Year level`))) %>%
#   # Make first column row names
#   column_to_rownames(var = "Year level")
# 
# # Calculate row sums for all columns except the last one
# row_sums <- rowSums(eth_tbl_n[, -ncol(eth_tbl_n)])
# 
# # Calculate percentages for each cell by dividing by the row sum (excluding the last column)
# percentages <- round(sweep(eth_tbl_n[, -ncol(eth_tbl_n)], 1, row_sums, FUN = "/") * 100, 0)
# 
# # Calculate percentages for the last column as percentages of the total
# total_percentage <- round((eth_tbl_n[, ncol(eth_tbl_n)] / sum(eth_tbl_n[-length(eth_tbl_n$Total), ]$Total)) * 100, 0)
# 
# # Create a new dataframe to store combined counts and percentages
# df_combined <- eth_tbl_n
# 
# # Loop through columns (excluding the last one)
# for (col in names(eth_tbl_n)[-ncol(eth_tbl_n)]) {
#   # Create formatted string with count and percentage
#   combined_col <- paste0(eth_tbl_n[[col]], " (", round(percentages[[col]], 1), "%)")
#   # Assign the formatted string to the column in the combined dataframe
#   df_combined[[col]] <- combined_col
# }
# 
# # Create formatted string for the Total column
# combined_total <- sprintf("%d (%.1f%%)", eth_tbl_n$Total, total_percentage)
# df_combined$Total <- combined_total
# 
# colnames(df_combined) <- c("Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Total")
# 
# # Set rownames as first column
# ethn_cant <- df_combined %>%
#   rownames_to_column(var = "Year level") %>%
#   mutate(Region = "Canterbury") %>%
#   dplyr::relocate(Region)
# 
# #2. Hawkes Bay
# ethnicity <- DATA %>% 
#   dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`) %>%
#   filter(Region=="Hawkes Bay") %>%
#   dplyr::select(-Region) %>%
#   dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
#   mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
#                                          grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
#                                          grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
#                                          is.na(Ethnicity) ~ "Unknown",
#                                          .default = "Other"))) %>%
#   mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
#   mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#   mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#   mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
#   mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown")))
# 
# eth_tbl_n <- ethnicity %>%
#   group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
#   dplyr::summarise(n = n(), .groups = "drop") %>%
#   ungroup() %>%
#   pivot_wider(names_from = c(Ethnicity, Gender), values_from = "n") %>%
#   mutate(Total = rowSums(.[-1])); eth_tbl_n
# 
# tot_row <- c(NA, colSums(eth_tbl_n[, -1], na.rm = T))
# 
# eth_tbl_n <- rbind(eth_tbl_n, tot_row) %>%
#   mutate(`Year level` = as.character(`Year level`)) %>%
#   mutate(`Year level` = factor(ifelse(is.na(`Year level`), "Total", `Year level`))) %>%
#   # Make first column row names
#   column_to_rownames(var = "Year level")
# 
# # Calculate row sums for all columns except the last one
# row_sums <- rowSums(eth_tbl_n[, -ncol(eth_tbl_n)])
# 
# # Calculate percentages for each cell by dividing by the row sum (excluding the last column)
# percentages <- round(sweep(eth_tbl_n[, -ncol(eth_tbl_n)], 1, row_sums, FUN = "/") * 100, 0)
# 
# # Calculate percentages for the last column as percentages of the total
# total_percentage <- round((eth_tbl_n[, ncol(eth_tbl_n)] / sum(eth_tbl_n[-length(eth_tbl_n$Total), ]$Total)) * 100, 0)
# 
# # Create a new dataframe to store combined counts and percentages
# df_combined <- eth_tbl_n
# 
# # Loop through columns (excluding the last one)
# for (col in names(eth_tbl_n)[-ncol(eth_tbl_n)]) {
#   # Create formatted string with count and percentage
#   combined_col <- paste0(eth_tbl_n[[col]], " (", round(percentages[[col]], 1), "%)")
#   # Assign the formatted string to the column in the combined dataframe
#   df_combined[[col]] <- combined_col
# }
# 
# # Create formatted string for the Total column
# combined_total <- sprintf("%d (%.1f%%)", eth_tbl_n$Total, total_percentage)
# df_combined$Total <- combined_total
# 
# colnames(df_combined) <- c("Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Total")
# 
# # Set rownames as first column
# ethn_HB <- df_combined %>%
#   rownames_to_column(var = "Year level")  %>%
#   mutate(Region = "Hawkes Bay") %>%
#   dplyr::relocate(Region)
# 
# #3. Manawatu/Rangitikei/Whanganui
# ethnicity <- DATA %>% 
#   dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`) %>%
#   filter(Region=="Manawatu/Rangitikei/Whanganui") %>%
#   dplyr::select(-Region) %>%
#   dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
#   mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
#                                          grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
#                                          grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
#                                          is.na(Ethnicity) ~ "Unknown",
#                                          .default = "Other"))) %>%
#   mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
#   mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#   mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#   mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
#   mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown")))
# 
# eth_tbl_n <- ethnicity %>%
#   group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
#   dplyr::summarise(n = n(), .groups = "drop") %>%
#   ungroup() %>%
#   pivot_wider(names_from = c(Ethnicity, Gender), values_from = "n") %>%
#   mutate(Total = rowSums(.[-1])); eth_tbl_n
# 
# tot_row <- c(NA, colSums(eth_tbl_n[, -1], na.rm = T))
# 
# eth_tbl_n <- rbind(eth_tbl_n, tot_row) %>%
#   mutate(`Year level` = as.character(`Year level`)) %>%
#   mutate(`Year level` = factor(ifelse(is.na(`Year level`), "Total", `Year level`))) %>%
#   # Make first column row names
#   column_to_rownames(var = "Year level")
# 
# # Calculate row sums for all columns except the last one
# row_sums <- rowSums(eth_tbl_n[, -ncol(eth_tbl_n)])
# 
# # Calculate percentages for each cell by dividing by the row sum (excluding the last column)
# percentages <- round(sweep(eth_tbl_n[, -ncol(eth_tbl_n)], 1, row_sums, FUN = "/") * 100, 0)
# 
# # Calculate percentages for the last column as percentages of the total
# total_percentage <- round((eth_tbl_n[, ncol(eth_tbl_n)] / sum(eth_tbl_n[-length(eth_tbl_n$Total), ]$Total)) * 100, 0)
# 
# # Create a new dataframe to store combined counts and percentages
# df_combined <- eth_tbl_n
# 
# # Loop through columns (excluding the last one)
# for (col in names(eth_tbl_n)[-ncol(eth_tbl_n)]) {
#   # Create formatted string with count and percentage
#   combined_col <- paste0(eth_tbl_n[[col]], " (", round(percentages[[col]], 1), "%)")
#   # Assign the formatted string to the column in the combined dataframe
#   df_combined[[col]] <- combined_col
# }
# 
# # Create formatted string for the Total column
# combined_total <- sprintf("%d (%.1f%%)", eth_tbl_n$Total, total_percentage)
# df_combined$Total <- combined_total
# 
# colnames(df_combined) <- c("Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Total")
# 
# # Set rownames as first column
# ethn_man <- df_combined %>%
#   rownames_to_column(var = "Year level") %>% 
#   mutate(Region = "Manawatu/Rangitikei/Whanganui") %>%
#   dplyr::relocate(Region)
# 
# #4. Nelson
# ethnicity <- DATA %>% 
#   dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`) %>%
#   filter(Region=="Nelson/Marlborough/West Coast") %>%
#   dplyr::select(-Region) %>%
#   dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
#   mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
#                                          grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
#                                          grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
#                                          is.na(Ethnicity) ~ "Unknown",
#                                          .default = "Other"))) %>%
#   mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
#   mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#   mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#   mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
#   mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown")))
# 
# eth_tbl_n <- ethnicity %>%
#   group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
#   dplyr::summarise(n = n(), .groups = "drop") %>%
#   ungroup() %>%
#   pivot_wider(names_from = c(Ethnicity, Gender), values_from = "n") %>%
#   mutate(Total = rowSums(.[-1])); eth_tbl_n
# 
# tot_row <- c(NA, colSums(eth_tbl_n[, -1], na.rm = T))
# 
# eth_tbl_n <- rbind(eth_tbl_n, tot_row) %>%
#   mutate(`Year level` = as.character(`Year level`)) %>%
#   mutate(`Year level` = factor(ifelse(is.na(`Year level`), "Total", `Year level`))) %>%
#   # Make first column row names
#   column_to_rownames(var = "Year level")
# 
# # Calculate row sums for all columns except the last one
# row_sums <- rowSums(eth_tbl_n[, -ncol(eth_tbl_n)])
# 
# # Calculate percentages for each cell by dividing by the row sum (excluding the last column)
# percentages <- round(sweep(eth_tbl_n[, -ncol(eth_tbl_n)], 1, row_sums, FUN = "/") * 100, 0)
# 
# # Calculate percentages for the last column as percentages of the total
# total_percentage <- round((eth_tbl_n[, ncol(eth_tbl_n)] / sum(eth_tbl_n[-length(eth_tbl_n$Total), ]$Total)) * 100, 0)
# 
# # Create a new dataframe to store combined counts and percentages
# df_combined <- eth_tbl_n
# 
# # Loop through columns (excluding the last one)
# for (col in names(eth_tbl_n)[-ncol(eth_tbl_n)]) {
#   # Create formatted string with count and percentage
#   combined_col <- paste0(eth_tbl_n[[col]], " (", round(percentages[[col]], 1), "%)")
#   # Assign the formatted string to the column in the combined dataframe
#   df_combined[[col]] <- combined_col
# }
# 
# # Create formatted string for the Total column
# combined_total <- sprintf("%d (%.1f%%)", eth_tbl_n$Total, total_percentage)
# df_combined$Total <- combined_total
# 
# colnames(df_combined) <- c("Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Total")
# 
# # Set rownames as first column
# ethn_nel <- df_combined %>%
#   rownames_to_column(var = "Year level")  %>% 
#   mutate(Region = "Nelson/Marlborough/West Coasti") %>%
#   dplyr::relocate(Region)
# 
# 
# ################################################################################################################################################################
# ################################################################################################################################################################
# 
# # LEARNING DIFFERENCES SHEET
# 
# ld <- DATA %>%
#   dplyr::select(`Year level`, `Neurodiversity diagnosis`, `Neurodiversity suspicion`) %>%
#   mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#   mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#   group_by(`Year level`, .drop = F) %>%
#   dplyr::summarise(`Neurodiversity diagnosis` = sum(!is.na(`Neurodiversity diagnosis`), na.rm = TRUE),
#                    `Neurodiversity suspicion` = sum(!is.na(`Neurodiversity suspicion`), na.rm = TRUE),
#                    `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`) %>%
#   ungroup()
# 
# summary_row <- ld %>%
#   summarise(Region = "Total", `Year level` = "Total", `Neurodiversity diagnosis` = sum(`Neurodiversity diagnosis`, na.rm=T),
#             `Neurodiversity suspicion` = sum(`Neurodiversity suspicion`, na.rm=T),
#             `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`)
# 
# ld_all <- bind_rows(ld, summary_row) %>%
#   mutate(Region = "All regions") %>%
#   relocate(Region)
#   
# # Repeat for each region
# #1. Canterbury
# ld <- DATA %>%
#   dplyr::select(Region, `Year level`, `Neurodiversity diagnosis`, `Neurodiversity suspicion`) %>%
#   filter(Region=="Canterbury") %>%
#   dplyr::select(-Region) %>%
#   mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#   mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#   group_by(`Year level`, .drop = F) %>%
#   dplyr::summarise(`Neurodiversity diagnosis` = sum(!is.na(`Neurodiversity diagnosis`), na.rm = TRUE),
#                    `Neurodiversity suspicion` = sum(!is.na(`Neurodiversity suspicion`), na.rm = TRUE),
#                    `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`) %>%
#   ungroup()
# 
# summary_row <- ld %>%
#   summarise(`Year level` = "Total", `Neurodiversity diagnosis` = sum(`Neurodiversity diagnosis`, na.rm=T),
#             `Neurodiversity suspicion` = sum(`Neurodiversity suspicion`, na.rm=T),
#             `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`)
# 
# ld_can <- bind_rows(ld, summary_row) %>%
#   mutate(Region = "Canterbury") %>%
#   relocate(Region)
# 
# #2. Hawkes Bay
# ld <- DATA %>%
#   dplyr::select(Region, `Year level`, `Neurodiversity diagnosis`, `Neurodiversity suspicion`) %>%
#   filter(Region=="Hawkes Bay") %>%
#   dplyr::select(-Region) %>%
#   mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#   mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#   group_by(`Year level`, .drop = F) %>%
#   dplyr::summarise(`Neurodiversity diagnosis` = sum(!is.na(`Neurodiversity diagnosis`), na.rm = TRUE),
#                    `Neurodiversity suspicion` = sum(!is.na(`Neurodiversity suspicion`), na.rm = TRUE),
#                    `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`) %>%
#   ungroup()
# 
# summary_row <- ld %>%
#   summarise(`Year level` = "Total", `Neurodiversity diagnosis` = sum(`Neurodiversity diagnosis`, na.rm=T),
#             `Neurodiversity suspicion` = sum(`Neurodiversity suspicion`, na.rm=T),
#             `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`)
# 
# ld_HB <- bind_rows(ld, summary_row) %>%
#   mutate(Region = "Hawkes Bay") %>%
#   relocate(Region)
# 
# 
# #3. Manawatu/Rangitikei/Whanganui
# ld <- DATA %>%
#   dplyr::select(Region, `Year level`, `Neurodiversity diagnosis`, `Neurodiversity suspicion`) %>%
#   filter(Region=="Manawatu/Rangitikei/Whanganui") %>%
#   dplyr::select(-Region) %>%
#   mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#   mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#   group_by(`Year level`, .drop = F) %>%
#   dplyr::summarise(`Neurodiversity diagnosis` = sum(!is.na(`Neurodiversity diagnosis`), na.rm = TRUE),
#                    `Neurodiversity suspicion` = sum(!is.na(`Neurodiversity suspicion`), na.rm = TRUE),
#                    `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`) %>%
#   ungroup()
# 
# summary_row <- ld %>%
#   summarise(`Year level` = "Total", `Neurodiversity diagnosis` = sum(`Neurodiversity diagnosis`, na.rm=T),
#             `Neurodiversity suspicion` = sum(`Neurodiversity suspicion`, na.rm=T),
#             `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`)
# 
# ld_man <- bind_rows(ld, summary_row) %>%
#   mutate(Region = "Manawatu/Rangitikei/Whanganui") %>%
#   relocate(Region)
# 
# 
# #4. Nelson/Marlborough/West Coast
# ld <- DATA %>%
#   dplyr::select(Region, `Year level`, `Neurodiversity diagnosis`, `Neurodiversity suspicion`) %>%
#   filter(Region=="Nelson/Marlborough/West Coast") %>%
#   dplyr::select(-Region) %>%
#   mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#   mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#   group_by(`Year level`, .drop = F) %>%
#   dplyr::summarise(`Neurodiversity diagnosis` = sum(!is.na(`Neurodiversity diagnosis`), na.rm = TRUE),
#                    `Neurodiversity suspicion` = sum(!is.na(`Neurodiversity suspicion`), na.rm = TRUE),
#                    `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`) %>%
#   ungroup()
# 
# summary_row <- ld %>%
#   summarise(`Year level` = "Total", `Neurodiversity diagnosis` = sum(`Neurodiversity diagnosis`, na.rm=T),
#             `Neurodiversity suspicion` = sum(`Neurodiversity suspicion`, na.rm=T),
#             `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`)
# 
# ld_nel <- bind_rows(ld, summary_row) %>%
#   mutate(Region = "Nelson/Marlborough/West Coast") %>%
#   relocate(Region)
# 
# 
# ################################################################################################################################################################
# ################################################################################################################################################################
# 
# # ATTENDANCE SHEET
# 
# att <- DATA %>%
#   dplyr::select(`Year level`, `Total available`, `Total attended`) %>%
#   mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#   mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#   group_by(`Year level`, .drop = F) %>%
#   dplyr::summarise(`Intervention Session Attendance` = round(100*sum(`Total attended`, na.rm = TRUE)/sum(`Total available`, na.rm = T), 1)) %>%
#   ungroup() %>%
#   mutate(`Intervention Session Attendance` = ifelse(is.nan(`Intervention Session Attendance`), "No data", `Intervention Session Attendance`)) %>%
#   mutate(`Intervention Session Attendance` = ifelse(`Intervention Session Attendance` != "No data", paste0(`Intervention Session Attendance`, "%"),
#          `Intervention Session Attendance`)) %>%
#   mutate(Region = "All regions") %>%
#   relocate(Region)
# 
# # Repeat for each region
# att_can <- DATA %>%
#   dplyr::select(Region, `Year level`, `Total available`, `Total attended`) %>%
#   filter(Region=="Canterbury") %>%
#   dplyr::select(-Region) %>%
#   mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#   mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#   group_by(`Year level`, .drop = F) %>%
#   dplyr::summarise(`Intervention Session Attendance` = round(100*sum(`Total attended`, na.rm = TRUE)/sum(`Total available`, na.rm = T), 1)) %>%
#   ungroup() %>%
#   mutate(`Intervention Session Attendance` = ifelse(is.nan(`Intervention Session Attendance`), "No data", `Intervention Session Attendance`)) %>%
#   mutate(`Intervention Session Attendance` = ifelse(`Intervention Session Attendance` != "No data", paste0(`Intervention Session Attendance`, "%"),
#                                                     `Intervention Session Attendance`)) %>%
#   mutate(Region = "Canterbury") %>%
#   relocate(Region)
# 
# 
# att_HB <- DATA %>%
#   dplyr::select(Region, `Year level`, `Total available`, `Total attended`) %>%
#   filter(Region=="Hawkes Bay") %>%
#   dplyr::select(-Region) %>%
#   mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#   mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#   group_by(`Year level`, .drop = F) %>%
#   dplyr::summarise(`Intervention Session Attendance` = round(100*sum(`Total attended`, na.rm = TRUE)/sum(`Total available`, na.rm = T), 1)) %>%
#   ungroup() %>%
#   mutate(`Intervention Session Attendance` = ifelse(is.nan(`Intervention Session Attendance`), "No data", `Intervention Session Attendance`)) %>%
#   mutate(`Intervention Session Attendance` = ifelse(`Intervention Session Attendance` != "No data", paste0(`Intervention Session Attendance`, "%"),
#                                                     `Intervention Session Attendance`)) %>%
#   mutate(Region = "Hawkes Bay") %>%
#   relocate(Region)
# 
# att_man <- DATA %>%
#   dplyr::select(Region, `Year level`, `Total available`, `Total attended`) %>%
#   filter(Region=="Manawatu/Rangitikei/Whanganui") %>%
#   dplyr::select(-Region) %>%
#   mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#   mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#   group_by(`Year level`, .drop = F) %>%
#   dplyr::summarise(`Intervention Session Attendance` = round(100*sum(`Total attended`, na.rm = TRUE)/sum(`Total available`, na.rm = T), 1)) %>%
#   ungroup() %>%
#   mutate(`Intervention Session Attendance` = ifelse(is.nan(`Intervention Session Attendance`), "No data", `Intervention Session Attendance`)) %>%
#   mutate(`Intervention Session Attendance` = ifelse(`Intervention Session Attendance` != "No data", paste0(`Intervention Session Attendance`, "%"),
#                                                     `Intervention Session Attendance`)) %>%
#   mutate(Region = "Manawatu/Rangitikei/Whanganui") %>%
#   relocate(Region)
# 
# att_nel <- DATA %>%
#   dplyr::select(Region, `Year level`, `Total available`, `Total attended`) %>%
#   filter(Region=="Nelson/Marlborough/West Coast") %>%
#   dplyr::select(-Region) %>%
#   mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#   mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#   group_by(`Year level`, .drop = F) %>%
#   dplyr::summarise(`Intervention Session Attendance` = round(100*sum(`Total attended`, na.rm = TRUE)/sum(`Total available`, na.rm = T), 1)) %>%
#   ungroup() %>%
#   mutate(`Intervention Session Attendance` = ifelse(is.nan(`Intervention Session Attendance`), "No data", `Intervention Session Attendance`)) %>%
#   mutate(`Intervention Session Attendance` = ifelse(`Intervention Session Attendance` != "No data", paste0(`Intervention Session Attendance`, "%"),
#                                                     `Intervention Session Attendance`)) %>%
#   mutate(Region = "Nelson/Marlborough/West Coast") %>%
#   relocate(Region)
# 
# ################################################################################################################################################################
# ################################################################################################################################################################
# 
# # ADPATIVE BRYANT SHEET
# 
# # Create function to save repetitive code
# 
# create_region_assessment_df <- function(region, assessment) {
#   
#   # Dynamic column name for assessment score
#   score_column <- paste0(assessment, " AB score")
#   
#   # Desired column names
#   desired_columns <- c(
#     "Year level", 
#     "N_Māori_F", paste0(assessment, " AB score_Māori_F"),
#     "N_Māori_M", paste0(assessment, " AB score_Māori_M"),
#     "N_Māori_Unknown", paste0(assessment, " AB score_Māori_Unknown"),
#     "N_Pasifika_F", paste0(assessment, " AB score_Pasifika_F"),
#     "N_Pasifika_M", paste0(assessment, " AB score_Pasifika_M"),
#     "N_Pasifika_Unknown", paste0(assessment, " AB score_Pasifika_Unknown"),
#     "N_NZ European_F", paste0(assessment, " AB score_NZ European_F"),
#     "N_NZ European_M", paste0(assessment, " AB score_NZ European_M"),
#     "N_NZ European_Unknown", paste0(assessment, " AB score_NZ European_Unknown"),
#     "N_Other_F", paste0(assessment, " AB score_Other_F"),
#     "N_Other_M", paste0(assessment, " AB score_Other_M"),
#     "N_Other_Unknown", paste0(assessment, " AB score_Other_Unknown"),
#     "N_Unknown_F", paste0(assessment, " AB score_Unknown_F"),
#     "N_Unknown_M", paste0(assessment, " AB score_Unknown_M"),
#     "N_Unknown_Unknown", paste0(assessment, " AB score_Unknown_Unknown"))
#   
#   # Filter and preprocess data
#   ad_data <- DATA %>%
#     dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, !!score_column) %>%
#     { if (region != "National") filter(., Region == region) else . } %>% # Region filter if not National
#     dplyr::select(-Region) %>%
#     dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
#     mutate(Ethnicity = as.factor(case_when(
#       grepl("Maori|Māori", Ethnicity) ~ "Māori",
#       grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
#       grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
#       is.na(Ethnicity) ~ "Unknown",
#       .default = "Other"
#     ))) %>%
#     mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
#     mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#     mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#     mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
#     mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
#     mutate(across(contains("score"), as.numeric)) %>%
#     filter(!is.na(!!as.name(score_column)))
#   
#   # Overall column
#   col <- ad_data %>%
#     group_by(`Year level`, .drop = F) %>%
#     dplyr::summarise(N = n(), Overall = median(!!as.name(score_column), na.rm = TRUE))
#   
#   # Pivot the data wider by Ethnicity and Gender
#   ad_wide <- ad_data %>%
#     group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
#     dplyr::summarise(N = n(), !!score_column := median(!!as.name(score_column), na.rm = TRUE)) %>%
#     pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(N, !!as.name(score_column))) %>%
#     ungroup() %>%
#     select(all_of(desired_columns)) %>%
#     bind_cols(col[2:3])
#   
#   # Overall row
#   row <- ad_data %>%
#     group_by(Gender, Ethnicity, .drop = F) %>%
#     dplyr::summarise(N = n(), !!score_column := median(!!as.name(score_column), na.rm = TRUE)) %>%
#     pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(N, !!as.name(score_column))) %>%
#     ungroup() %>%
#     mutate(`Year level` = "Overall") %>%
#     relocate(`Year level`) %>%
#     select(all_of(desired_columns)) %>%
#     mutate(N = sum(!is.na(ad_data %>% pull(!!score_column))), Overall = median(ad_data %>% pull(!!score_column), na.rm = TRUE))
#   
#   
#   # Combine the wide data and overall row, and add region information
#   ad_wide <- rbind(ad_wide, row) %>%
#     mutate(Region = region) %>%
#     relocate(Region)
#   
#   # Rename columns
#   colnames(ad_wide) <- c("Region", "Year level", "N Māori F", "Median Māori F", "N Māori M", "Median Māori M", "N Māori U", "Median Māori U", 
#                          "N Pasifika F", "Median Pasifika F", "N Pasifika M", "Median Pasifika M", "N Pasifika U", "Median Pasifika U", 
#                          "N NZ European F", "Median NZ European F", "N NZ European M", "Median NZ European M", "N NZ European U", "Median NZ European U", 
#                          "N Other F", "Median Other F", "N Other M", "Median Other M", "N Other U", "Median Other U",
#                          "N Unknown F", "Median Unknown F", "N Unknown M", "Median Unknown M", "N Unknown U", "Median Unknown U",
#                          "N overall", "Median overall")
#   
# 
#   return(ad_wide)
# }
# 
# 
# #1. National
# ad_pre_w_all <- create_region_assessment_df("National", "Pre")
# ad_mid_w_all <- create_region_assessment_df("National", "Mid")
# ad_fin_w_all <- create_region_assessment_df("National", "Final")
# 
# #2. Canterbury
# ad_pre_w_can <- create_region_assessment_df("Canterbury", "Pre")
# ad_mid_w_can <- create_region_assessment_df("Canterbury", "Mid")
# ad_fin_w_can <- create_region_assessment_df("Canterbury", "Final")
# 
# #3. Hawkes Bay
# ad_pre_w_HB <- create_region_assessment_df("Hawkes Bay", "Pre")
# ad_mid_w_HB <- create_region_assessment_df("Hawkes Bay", "Mid")
# ad_fin_w_HB <- create_region_assessment_df("Hawkes Bay", "Final")
# 
# #4. Manawatu/Rangitikei/Whanganui
# ad_pre_w_man <- create_region_assessment_df("Manawatu/Rangitikei/Whanganui", "Pre")
# ad_mid_w_man <- create_region_assessment_df("Manawatu/Rangitikei/Whanganui", "Mid")
# ad_fin_w_man <- create_region_assessment_df("Manawatu/Rangitikei/Whanganui", "Final")
# 
# #5. Nelson/Marlborough/West Coast
# ad_pre_w_nel <- create_region_assessment_df("Nelson/Marlborough/West Coast", "Pre")
# ad_mid_w_nel <- create_region_assessment_df("Nelson/Marlborough/West Coast", "Mid")
# ad_fin_w_nel <- create_region_assessment_df("Nelson/Marlborough/West Coast", "Final")
# 
# 
# ##################################################################################################################
# ##################################################################################################################
# 
# # ADD E-ASTTLE DATA SHEET
# 
# # Function to save repetitive code
# easttle_df <- function(region) {
#   
#   # Desired column names
#   desired_columns <- c(
#     "Year level", 
#     "N_Māori_F", "e-asTTle overall scaled score_Māori_F",
#     "N_Māori_M", "e-asTTle overall scaled score_Māori_M",
#     "N_Māori_Unknown", "e-asTTle overall scaled score_Māori_Unknown",
#     "N_Pasifika_F", "e-asTTle overall scaled score_Pasifika_F",
#     "N_Pasifika_M", "e-asTTle overall scaled score_Pasifika_M",
#     "N_Pasifika_Unknown", "e-asTTle overall scaled score_Pasifika_Unknown",
#     "N_NZ European_F", "e-asTTle overall scaled score_NZ European_F",
#     "N_NZ European_M", "e-asTTle overall scaled score_NZ European_M",
#     "N_NZ European_Unknown", "e-asTTle overall scaled score_NZ European_Unknown",
#     "N_Other_F", "e-asTTle overall scaled score_Other_F",
#     "N_Other_M", "e-asTTle overall scaled score_Other_M",
#     "N_Other_Unknown", "e-asTTle overall scaled score_Other_Unknown",
#     "N_Unknown_F", "e-asTTle overall scaled score_Unknown_F",
#     "N_Unknown_M", "e-asTTle overall scaled score_Unknown_M",
#     "N_Unknown_Unknown", "e-asTTle overall scaled score_Unknown_Unknown")
#   
#   easttle <- DATA %>% 
#     filter(`e-asTTle overall scaled score` != 0) %>%
#     dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `e-asTTle overall scaled score`) %>%
#     { if (region != "National") filter(., Region == region) else . } %>% # Region filter if not National
#     dplyr::select(-Region) %>%
#     dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
#     mutate(Ethnicity = as.factor(case_when(
#       grepl("Maori|Māori", Ethnicity) ~ "Māori",
#       grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
#       grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
#       is.na(Ethnicity) ~ "Unknown",
#       .default = "Other"
#     ))) %>%
#     mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
#     mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#     mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#     mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
#     mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
#     filter(!is.na(`e-asTTle overall scaled score`))
#   
#   
#   # Repeat chunk above but modify so it counts the number of students with a score of 0 and can be added as the column "No. students with no score"
#   no_score <- DATA %>% 
#     filter(`e-asTTle overall scaled score` == 0) %>%
#     dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `e-asTTle overall scaled score`) %>%
#     { if (region != "National") filter(., Region == region) else . } %>% # Region filter if not National
#     dplyr::select(-Region) %>%
#     dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
#     mutate(Ethnicity = as.factor(case_when(
#       grepl("Maori|Māori", Ethnicity) ~ "Māori",
#       grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
#       grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
#       is.na(Ethnicity) ~ "Unknown",
#       .default = "Other"
#     ))) %>%
#     mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
#     mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
#     mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
#     mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
#     mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) 
#   
#   
#   # Column aggregation - scores
#   col <- easttle %>%
#     group_by(`Year level`, .drop = F) %>%
#     dplyr::summarise(N = n(), Overall = median(`e-asTTle overall scaled score`, na.rm = TRUE))
#   
#   
#   # Column aggregation - no scores
#   no_score_col <- no_score %>%
#     group_by(`Year level`, .drop = F) %>%
#     dplyr::summarise(`No. students with no score` = n())
#   
#   
#   # Pivoting and reshaping
#   easttle_w <- easttle %>%
#     group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
#     dplyr::summarise(N = n(), `e-asTTle overall scaled score` = median(`e-asTTle overall scaled score`, na.rm = TRUE)) %>%
#     pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(N, `e-asTTle overall scaled score`)) %>%
#     ungroup() %>%
#     select(all_of(desired_columns)) %>%
#     bind_cols(col[2:3]) %>%
#     bind_cols(no_score_col[2])
#   
#   
#   # Adding "Overall" row
#   row <- easttle %>%
#     group_by(Gender, Ethnicity, .drop = F) %>%
#     dplyr::summarise(N = n(), `e-asTTle overall scaled score` = median(`e-asTTle overall scaled score`, na.rm = TRUE)) %>%
#     pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(N, `e-asTTle overall scaled score`)) %>%
#     ungroup() %>%
#     mutate(`Year level` = "Overall") %>%
#     relocate(`Year level`) %>%
#     select(all_of(desired_columns)) %>%
#     mutate(N = sum(!is.na(easttle$`e-asTTle overall scaled score`)), 
#            Overall = median(as.numeric(easttle$`e-asTTle overall scaled score`), na.rm = TRUE),
#            `No. students with no score` = sum(no_score_col$`No. students with no score`))
#   
#   # Combine everything and rename columns
#   easttle_w <- rbind(easttle_w, row) %>% 
#     mutate(Region = region) %>%
#     relocate(Region)
#   
#   colnames(easttle_w) <- c(
#     "Region", "Year level", "N Māori F", "Median Māori F", "N Māori M", "Median Māori M", "N Māori U", "Median Māori U", 
#     "N Pasifika F", "Median Pasifika F", "N Pasifika M", "Median Pasifika M", "N Pasifika U", "Median Pasifika U", 
#     "N NZ European F", "Median NZ European F", "N NZ European M", "Median NZ European M", "N NZ European U", "Median NZ European U", 
#     "N Other F", "Median Other F", "N Other M", "Median Other M", "N Other U", "Median Other U", 
#     "N Unknown F", "Median Unknown F", "N Unknown M", "Median Unknown M", "N Unknown U", "Median Unknown U", "N overall", "Median overall",
#     "No. students with no score")
#   
#   return(easttle_w)
# }
# 
# # DATA$`e-asTTle overall scaled score`
# # as.numeric(DATA$`e-asTTle overall scaled score`)
# 
# # Make the dataframes
# easttle_all <- easttle_df("National")
# easttle_can <- easttle_df("Canterbury")
# easttle_HB  <- easttle_df("Hawkes Bay")
# easttle_man <- easttle_df("Manawatu/Rangitikei/Whanganui")
# easttle_nel <- easttle_df("Nelson/Marlborough/West Coast")
# 
# 
# ##################################################################################################################
# ##################################################################################################################
# 
# # IDENTIFY TEACHERS WHOSE NEURODIVERSITY FIELDS ARE ALL EMPTY
# 
# # Use combined_data, group by Teacher, and make a list of teachers who have all NA for columns starting with "Neurodiversity"
# empty_neurodiversity <- DATA %>%
#   group_by(Teacher) %>%
#   filter(all(is.na(across(starts_with("Neurodiversity"))))) %>%
#   arrange(Teacher) %>%
#   pull(Teacher) %>%
#   unique() 
# 
# # Print as a vector I can paste into Excel
# cat(paste(empty_neurodiversity, collapse = ", "))
# 
# ##################################################################################################################
# ##################################################################################################################
# 
# 
# # CREATE WORKBOOK
# 
# wb <- createWorkbook()
# 
# # Add ethnicity worksheet to the workbook
# addWorksheet(wb, sheetName = "Ethnicity")
# 
# # Write each dataframe to the worksheet with spacing
# writeData(wb, sheet = "Ethnicity", x = ethn_all, startRow = 1, startCol = 1)
# writeData(wb, sheet = "Ethnicity", x = ethn_cant, startRow = nrow(df_combined) + 4, startCol = 1)
# writeData(wb, sheet = "Ethnicity", x = ethn_HB, startRow = nrow(df_combined) + nrow(ethn_cant) + 6, startCol = 1)
# writeData(wb, sheet = "Ethnicity", x = ethn_man, startRow = nrow(df_combined) + nrow(ethn_cant) + nrow(ethn_HB) + 9, startCol = 1)
# writeData(wb, sheet = "Ethnicity", x = ethn_nel, startRow = nrow(df_combined) + nrow(ethn_cant) + nrow(ethn_HB) + nrow(ethn_man) + 12, startCol = 1)
# 
# 
# # Add learning differences worksheet to the workbook
# addWorksheet(wb, sheetName = "Learning differences")
# 
# # Write each dataframe to the worksheet with spacing
# writeData(wb, sheet = "Learning differences", x = ld_all, startRow = 1, startCol = 1)
# writeData(wb, sheet = "Learning differences", x = ld_can, startRow = nrow(ld_all) + 4, startCol = 1)
# writeData(wb, sheet = "Learning differences", x = ld_HB, startRow = nrow(ld_all) + nrow(ld_can) + 6, startCol = 1)
# writeData(wb, sheet = "Learning differences", x = ld_man, startRow = nrow(ld_all) + nrow(ld_can) + nrow(ld_HB) + 9, startCol = 1)
# writeData(wb, sheet = "Learning differences", x = ld_nel, startRow = nrow(ld_all) + nrow(ld_can) + nrow(ld_HB) + nrow(ld_man) + 12, startCol = 1)
# 
# 
# # Add attendance worksheet to the workbook
# addWorksheet(wb, sheetName = "Attendance")
# 
# # Write each dataframe to the worksheet with spacing
# writeData(wb, sheet = "Attendance", x = att, startRow = 1, startCol = 1)
# writeData(wb, sheet = "Attendance", x = att_can, startRow = nrow(att) + 4, startCol = 1)
# writeData(wb, sheet = "Attendance", x = att_HB, startRow = nrow(att) + nrow(att_can) + 6, startCol = 1)
# writeData(wb, sheet = "Attendance", x = att_man, startRow = nrow(att) + nrow(att_can) + nrow(att_HB) + 9, startCol = 1)
# writeData(wb, sheet = "Attendance", x = att_nel, startRow = nrow(att) + nrow(att_can) + nrow(att_HB) + nrow(att_man) + 12, startCol = 1)
# 
# 
# # Add adaptive Bryant worksheet to the workbook
# addWorksheet(wb, sheetName = "Adaptive Bryant")
# 
# # Create header dataframes, one for pre, mid, and final assessments
# header_pre <- data.frame("PRE ASSESSMENTS" = double(), "(Median score)" = double(), check.names = F)
# header_mid <- data.frame("MID ASSESSMENTS" = double(), "(Median score)" = double(), check.names = F)
# header_fin <- data.frame("FINAL ASSESSMENTS" = double(), "(Median score)" = double(), check.names = F)
# 
# # Write each dataframe to the worksheet with spacing
# writeData(wb, sheet = "Adaptive Bryant", x = header_pre, startRow = 1, startCol = 1)
# writeData(wb, sheet = "Adaptive Bryant", x = ad_pre_w_all, startRow = 3, startCol = 1)
# writeData(wb, sheet = "Adaptive Bryant", x = header_mid, startRow = 1, startCol = ncol(ad_pre_w_all) + 3)
# writeData(wb, sheet = "Adaptive Bryant", x = ad_mid_w_all, startRow = 3, startCol = ncol(ad_pre_w_all) + 3)
# writeData(wb, sheet = "Adaptive Bryant", x = header_fin, startRow = 1, startCol = ncol(ad_pre_w_all) + ncol(ad_mid_w_all) + 6)
# writeData(wb, sheet = "Adaptive Bryant", x = ad_fin_w_all, startRow = 3, startCol = ncol(ad_pre_w_all) + ncol(ad_mid_w_all) + 6)
# 
# writeData(wb, sheet = "Adaptive Bryant", x = ad_pre_w_can, startRow = nrow(ad_pre_w_all) + 5, startCol = 1)
# writeData(wb, sheet = "Adaptive Bryant", x = ad_mid_w_can, startRow = nrow(ad_mid_w_all) + 5, startCol = ncol(ad_pre_w_all) + 3)
# writeData(wb, sheet = "Adaptive Bryant", x = ad_fin_w_can, startRow = nrow(ad_fin_w_all) + 5, startCol = ncol(ad_pre_w_all) + ncol(ad_mid_w_all) + 6)
# 
# writeData(wb, sheet = "Adaptive Bryant", x = ad_pre_w_HB, startRow = nrow(ad_pre_w_all) + nrow(ad_pre_w_can) + 7, startCol = 1)
# writeData(wb, sheet = "Adaptive Bryant", x = ad_mid_w_HB, startRow = nrow(ad_mid_w_all) + nrow(ad_pre_w_can) + 7, startCol = ncol(ad_pre_w_all) + 3)
# writeData(wb, sheet = "Adaptive Bryant", x = ad_fin_w_HB, startRow = nrow(ad_fin_w_all) + nrow(ad_pre_w_can) + 7, startCol = ncol(ad_pre_w_all) + ncol(ad_mid_w_all) + 6)
# 
# writeData(wb, sheet = "Adaptive Bryant", x = ad_pre_w_man, startRow = nrow(ad_pre_w_all) + nrow(ad_pre_w_can) + nrow(ad_pre_w_HB) + 9, startCol = 1)
# writeData(wb, sheet = "Adaptive Bryant", x = ad_mid_w_man, startRow = nrow(ad_mid_w_all) + nrow(ad_pre_w_can) + nrow(ad_pre_w_HB) + 9, startCol = ncol(ad_pre_w_all) + 3)
# writeData(wb, sheet = "Adaptive Bryant", x = ad_fin_w_man, startRow = nrow(ad_fin_w_all) + nrow(ad_pre_w_can) + nrow(ad_pre_w_HB) + 9, startCol = ncol(ad_pre_w_all) + ncol(ad_mid_w_all) + 6)
# 
# writeData(wb, sheet = "Adaptive Bryant", x = ad_pre_w_nel, startRow = nrow(ad_pre_w_all) + nrow(ad_pre_w_can) + nrow(ad_pre_w_HB) + nrow(ad_pre_w_man) + 11, startCol = 1)
# writeData(wb, sheet = "Adaptive Bryant", x = ad_mid_w_nel, startRow = nrow(ad_mid_w_all) + nrow(ad_pre_w_can) + nrow(ad_pre_w_HB) + nrow(ad_pre_w_man) + 11, startCol = ncol(ad_pre_w_all) + 3)
# writeData(wb, sheet = "Adaptive Bryant", x = ad_fin_w_nel, startRow = nrow(ad_fin_w_all) + nrow(ad_pre_w_can) + nrow(ad_pre_w_HB) + nrow(ad_pre_w_man) + 11, startCol = ncol(ad_pre_w_all) + ncol(ad_mid_w_all) + 6)
# 
# 
# # Add e-asTTle worksheet to the workbook
# addWorksheet(wb, sheetName = "e-asTTle data")
# 
# # Write each dataframe to the worksheet with spacing
# writeData(wb, sheet = "e-asTTle data", x = easttle_all, startRow = 1, startCol = 1)
# writeData(wb, sheet = "e-asTTle data", x = easttle_can, startRow = nrow(easttle_all) + 4, startCol = 1)
# writeData(wb, sheet = "e-asTTle data", x = easttle_HB, startRow = nrow(easttle_all) + nrow(easttle_can) + 6, startCol = 1)
# writeData(wb, sheet = "e-asTTle data", x = easttle_man, startRow = nrow(easttle_all) + nrow(easttle_can) + nrow(easttle_HB) + 9, startCol = 1)
# writeData(wb, sheet = "e-asTTle data", x = easttle_nel, startRow = nrow(easttle_all) + nrow(easttle_can) + nrow(easttle_HB) + nrow(easttle_man) + 12, startCol = 1)
# 
# 
# 
# 
# # Add DATA and full sheets
# # # Anomymise DATA$Firsname and DATA$Surname by replacing each unique firstname with a unique random string and each unique surname with a unique random string
# # # Step 1: Create a mapping of unique first names to anonymized codes
# # first_name_mapping <- DATA %>%
# #   distinct(Firstname) %>%
# #   mutate(AnonFirstName = stri_rand_strings(n(), length = 8, pattern = "[A-Za-z0-9]"))
# # 
# # # Step 2: Create a mapping of unique surnames to anonymized codes
# # surname_mapping <- DATA %>%
# #   distinct(Surname) %>%
# #   mutate(AnonSurname = stri_rand_strings(n(), length = 8, pattern = "[A-Za-z0-9]"))
# # 
# # # Step 3: Replace the first names and surnames with their anonymized codes
# # DATA_anonymised <- DATA %>%
# #   left_join(first_name_mapping, by = "Firstname") %>%
# #   left_join(surname_mapping, by = "Surname") %>%
# #   mutate(Firstname = AnonFirstName, Surname = AnonSurname) %>%
# #   select(-AnonFirstName, -AnonSurname)  # Drop temporary columns
# 
# addWorksheet(wb, sheetName = "DATA")
# writeData(wb, sheet = "DATA", x = DATA, startRow = 1, startCol = 1)
# 
# 
# # # Anomymise combined_data$Firstname and combined_data$Surname by replacing each unique firstname with a unique random string and each unique surname with a unique random string
# # # Step 1: Create a mapping of unique first names to anonymized codes
# # first_name_mapping <- combined_data %>%
# #   distinct(Firstname) %>%
# #   mutate(AnonFirstName = stri_rand_strings(n(), length = 8, pattern = "[A-Za-z0-9]"))
# # 
# # # Step 2: Create a mapping of unique surnames to anonymized codes
# # surname_mapping <- combined_data %>%
# #   distinct(Surname) %>%
# #   mutate(AnonSurname = stri_rand_strings(n(), length = 8, pattern = "[A-Za-z0-9]"))
# # 
# # # Step 3: Replace the first names and surnames with their anonymized codes
# # combined_data_anonymised <- combined_data %>%
# #   left_join(first_name_mapping, by = "Firstname") %>%
# #   left_join(surname_mapping, by = "Surname") %>%
# #   mutate(Firstname = AnonFirstName, Surname = AnonSurname) %>%
# #   select(-AnonFirstName, -AnonSurname, -student_details_Student_Name, -student_details_NA)  
# 
# 
# addWorksheet(wb, sheetName = "Raw data")
# writeData(wb, sheet = "Raw data", x = combined_data, startRow = 1, startCol = 1)
# 
# # Add an excluded files sheet, using the excluded_files dataframe
# addWorksheet(wb, sheetName = "Excluded .xlsx files")
# writeData(wb, sheet = "Excluded .xlsx files", x = excluded_files_copies_paths, startRow = 1, startCol = 1)
# 
# # Add a notes sheet
# notes <- data.frame("NOTES" = c("Gender is expressed as F (female), M (male) and U (unknown - gender was not entered).",
#                                 "Year level contains all years entered by the teachers (i.e., some may be out of scope). Students with no year are set as 'unknown'.",
#                                 "A value of 'NaN' appears in some percentages because there were no students in that category (0 divided by 0 produces an error).",
#                                 "For the Ethnicity sheet, percentages are calculated row-wise, except for the total column. Only the first ethnicity value was used.",
#                                 "Where there were two students with the same first name and teacher, but one student was missing the surname, the student missing the surname was assumed to be the same as the student with the surname.",
#                                 "Median values are presented for summaries of adaptive Bryant and e-asTTle scores across groups.",
#                                 "Excluded files are listed in the excluded files sheet. They were excluded due to being copies."))
# 
# addWorksheet(wb, sheetName = "Notes")
# writeData(wb, sheet = "Notes", x = notes, startRow = 1, startCol = 1)
# 
# # Define style and apply to all sheets
# style <- createStyle(fontName = "Calibri", fontSize = 10)
# 
# sheets <- names(wb)
# 
# getRows <- function(wb, sheet) {
#   sheet_data <- read.xlsx(wb, sheet = sheet)
#   return(max(which(rowSums(!is.na(sheet_data)) > 0)))
# }
# 
# getCols <- function(wb, sheet) {
#   sheet_data <- read.xlsx(wb, sheet = sheet)
#   return(ncol(sheet_data))
# }
# 
# for (sheet in sheets) {
#   # Get the number of rows and columns used in the sheet
#   usedRows <- getRows(wb, sheet) + 10
#   usedCols <- getCols(wb, sheet)
#   
#   addStyle(wb, sheet = sheet, style = style, rows = 1:usedRows, cols = 1:usedCols, gridExpand = TRUE)
#   setColWidths(wb, sheet = sheet, cols = 1:usedCols, widths = "auto")
#   
# }
# 
# 
# # Save the workbook to a file
# saveWorkbook(wb, file = "report.xlsx", overwrite = TRUE)
# 
# length(unique(DATA$Teacher))
# 
# ##################################################################################################################
# ##################################################################################################################
# 
