#################################################################################################
#################################################################################################

# HU0002 LEARNING MATTERS DATA EXTRACTION PROJECT

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

mixedrank <- function(x) order(gtools::mixedorder(x))


theme <- theme_bw()+
  theme(legend.position = "bottom") +
  theme(axis.title.x = element_text(face="bold", colour="black", size=10),
        axis.text.y  = element_text(size=10),
        axis.text.x  = element_text(size=10),
        axis.title.y = element_text(face="bold", colour="black", size=10),
        panel.grid.minor=element_blank(),
        panel.grid.major=element_blank())

#################################################################################################
#################################################################################################

# EXTRACT DATA FROM THE LM GOOGLE DRIVE ACCOUNT


# Authenticate with Google
drive_auth()

# List files in "Learning Matters data for extraction/Master Spreadsheet ALL" folder. Setting recursive = T means it looks in subfolders.
files <- drive_ls(path = "Learning Matters data for extraction/Master Spreadsheet ALL", recursive = T)

# Filter the list to include only .xlsx files that have more than one sheet
xlsx_files <- files[grepl(".xlsx", files$name), ]


# Create a local directory to store downloaded files
local_dir <- "downloaded_xlsx_files2"
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

# Function to process each .xlsx file
process_xlsx_file <- function(file) {
  # Extract file information
  file_id <- file$id
  file_name <- file$name
  file_path <- get_full_path(file)
  
  # Determine where this file is in the list (i.e., its row number) so I can give it a unique identifier (many have the same name)
  file_index <- which(xlsx_files$id == file_id)
  
  # Download the file
  local_path <- file.path(local_dir, file_name)
  drive_download(as_id(file_id), path = local_path, overwrite = TRUE)
  
  # Load the .xlsx file
  wb <- loadWorkbook(local_path)
  
  # Check the number of sheets
  #sheet_count <- length(getSheetNames(wb))
  #prefix <- ifelse(length(getSheetNames(wb))<2, "Irrelevant", "Relevant")
  prefix <- ifelse("Baseline Assessment ORF" %in% excel_sheets(path = local_path), "Relevant", "Irrelevant")
  
  # Option 1: Rename the file to include the name of its folder
  #new_file_name <- paste0(gsub("/", "_", file_path), "_", file_name)
  #new_local_path <- file.path(local_dir, new_file_name)
  #saveWorkbook(wb = wb, file = new_file_name, overwrite = TRUE)
  
  # Option 2: Add a new sheet with the folder path
  addWorksheet(wb, "File Path")
  writeData(wb, sheet = "File Path", x = file_path)
  saveWorkbook(wb, local_path, overwrite = TRUE)
  
  # Return the path of the modified file
  #return(new_local_path)
  
  # Rename the file to include a unique identifier
  unique_file_name <- paste0(prefix, "_", file_index, "_", file_name)
  unique_local_path <- file.path(local_dir, unique_file_name)
  file.rename(local_path, unique_local_path)
  
  # Return the path of the modified file
  return(unique_local_path)  
}

# Use a reduced dataset for practice
#xlsx_files <- xlsx_files[1:12, ]

# Process each .xlsx file and store the paths of modified files
modified_files <- sapply(1:nrow(xlsx_files), function(i) process_xlsx_file(xlsx_files[i, ]))

###################################################################################################################################################
###################################################################################################################################################

# COLLATE XLSX FILES INTO A SINGLE DATAFRAME CONTAINING RELEVANT DATA

#1. Read in files in the folder "downloaded_xlsx_files2" that contain "Relevant" in the name and extract the sheets called "Tier 2 Student Details" and "File Path".
# Then add a column called "Region" to the sheet called "Tier 2 Student Details" that contains the text betwen the first "/" and the characters " ALL" in the "File Path" sheet.

# Define the path to the folder containing the .xlsx files
folder_path <- "downloaded_xlsx_files"

# List all .xlsx files in the folder that contain "Relevant" in the name
files2 <- list.files(path = folder_path, pattern = "Relevant.*\\.xlsx$", full.names = TRUE)

# Initialize an empty list to store data frames
data_list <- list()

# Create a vector to hold name of spreadsheet being worked on in case of a bug
spreadsheet_name <- c()

###################################################################################################################################################
# Loop through each file
for (file in files2) {
  
  print(noquote(paste0("Spreadsheet = ", "'", file, "'")))
  
  spreadsheet_name <- file
  
  # Function for tidying colnames
  replace_x_colnames <- function(colnames) {
    for (i in seq_along(colnames)) {
      if (startsWith(colnames[i], "x")) {
        # Find the previous column name that does not start with "x"
        for (j in (i-1):1) {
          if (!startsWith(colnames[j], "x")) {
            colnames[i] <- colnames[j]
            break
          }
        }
      }
    }
    return(colnames)
  }
  ##############################################################################################
  # Read the "Tier 2 Student Details" sheet
  tier2_details <- read_excel(file, sheet = "Tier 2 Student Details", col_types = c("text")) %>%
    clean_names()
  
  # Tidy up colnames
  new_colnames <- replace_x_colnames(colnames(tier2_details))
  colnames(tier2_details) <- new_colnames
  
  # Get first row values to add to colnames
  first_row <- tier2_details[1, ]
  
  # Modify column names by appending the first row values
  colnames(tier2_details) <- paste(colnames(tier2_details), first_row, sep = "_")
  
  # Remove the first row
  tier2_details <- tier2_details[-1, ]
  
  # Replace spaces in colnames with underscores
  colnames(tier2_details) <- str_replace_all(colnames(tier2_details), " ", "_")
    
  
  # Read the "File Path" sheet and convert to a vector
  file_path_vector <- read_excel(file, sheet = "File Path") %>%
    colnames()
  
  print(noquote(paste0("File path = ", "'", file_path_vector, "'")))
  
  ##############################################################################################
  # Extract the "Region" from the "File Path" sheet.
  region <- case_when(grepl("Canterbury", file_path_vector) ~ "Canterbury",
                      grepl("Hawkes", file_path_vector) ~ "Hawkes Bay",
                      grepl("Manawatu", file_path_vector) ~ "Manawatu/Rangitikei/Whanganui",
                      grepl("Nelson", file_path_vector) ~ "Nelson/Marlborough/West Coast")
  
  # Extract teacher from the "File Path" sheet
  teacher <- sub("^(.*/)([^/]*/[^/]+)$", "\\2", file_path_vector)
  
  # Remove text including and after "/" in teacher
  teacher <- sub("/.*", "", teacher)
  
  ##############################################################################################
  # Extract "neurodiverse_diagnosis"
  neurodiverse_diagnosis <- tier2_details %>%
    dplyr::select(student_details_Student_Name, starts_with("neuro_diverse_diagnosis"), - neuro_diverse_diagnosis_Ethnicity) %>%
    # Rename columns with only the characters after the last "_"
    rename_with(~str_extract(., "(?<=_)[^_]+$"), .cols = everything()) %>%
    pivot_longer(cols = !Name, names_to = "Diagnosis", values_to = "value") %>%
    filter(!is.na(value)) %>%
    group_by(Name) %>%
    mutate(neurodiverse_diagnosis = paste(Diagnosis, collapse = ", ")) %>%
    dplyr::select(Name, neurodiverse_diagnosis) %>%
    ungroup()
  
  # Extract "neurodiverse suspicion"
  neurodiverse_suspicion <- tier2_details %>%
    dplyr::select(student_details_Student_Name, starts_with("neuro_diverse_suspicion")) %>%
    # Rename columns with only the characters after the last "_"
    rename_with(~str_extract(., "(?<=_)[^_]+$"), .cols = everything()) %>%
    pivot_longer(cols = !Name, names_to = "Suspicion", values_to = "value") %>%
    filter(!is.na(value)) %>%
    group_by(Name) %>%
    mutate(neurodiverse_suspicion = paste(Suspicion, collapse = ", ")) %>%
    dplyr::select(Name, neurodiverse_suspicion) %>%
    ungroup()
  
  # Add "Region" and "Teacher" columns to the "Tier 2 Student Details" data frame and add neurodiverse diagnosis and suspicion
  tier2_details <- tier2_details %>% 
    mutate(Region_tier2 = region, Teacher_tier2 = teacher) %>%
    left_join(neurodiverse_diagnosis %>% dplyr::select(Name, neurodiverse_diagnosis), by = c("student_details_Student_Name" = "Name")) %>%
    left_join(neurodiverse_suspicion %>% dplyr::select(Name, neurodiverse_suspicion), by = c("student_details_Student_Name" = "Name"))
  
  # Edit tier2_details variables 
  tier2_details <- tier2_details %>%
    mutate(student_details_National_Student_number = as.numeric(student_details_National_Student_number)) %>%
    mutate(student_details_Year_level = as.integer(student_details_Year_level)) %>%
    mutate(student_details_Gender = case_when(student_details_Gender=="m" ~ "M",
                                              student_details_Gender=="Male" ~ "M",
                                              student_details_Gender=="male" ~ "M",
                                              student_details_Gender=="f" ~ "F",
                                              grepl("fem|Fem", student_details_Gender) ~ "F",
                                              .default = student_details_Gender)) %>%
    dplyr::rename(student_details_National_Student_number_tier2 = student_details_National_Student_number)
  
  ##############################################################################################
  
  # Extract adaptive Bryant score from the "Common Assessment (AB)" sheet
  ab <- read_excel(file, sheet = "Common Assessment (AB)") %>%
    clean_names()
  
  # Tidy up colnames
  new_colnames <- replace_x_colnames(colnames(ab))
  colnames(ab) <- new_colnames
  
  # Get first row values to add to colnames
  first_row <- ab[1, ]
  
  # Modify column names by appending the first row values
  colnames(ab) <- paste(colnames(ab), first_row, sep = "_")
  
  # Remove the first row
  ab <- ab[-1, ]
  
  # Replace spaces in colnames with underscores
  colnames(ab) <- str_replace_all(colnames(ab), " ", "_")
  
  # Tidy up NSN and add region, teacher and neurodiverse_diagnosis
  ab <- ab %>%
    mutate(student_details_National_Student_number = as.numeric(student_details_National_Student_number)) %>%
    dplyr::rename(student_details_National_Student_number_ad_bry = student_details_National_Student_number) %>%
    mutate(Region_ad_bry = region, Teacher_ad_bry = teacher) 
  
  ##############################################################################################
  
  # Extract student attendance data from the "Student Session Attendance" sheet
  ssa <- read_excel(file, sheet = "Student Session Attendance", skip = 1) %>%
    clean_names() %>%
    dplyr::select(-no_of_sessions_wk)
  
  # Check if there are any attendance data
  attendance_data <- ssa %>%
    pivot_longer(cols = starts_with("student_name"), names_to = "Name", values_to = "n_sessions")
  
  attendance_data <- ifelse(sum(!is.na(attendance_data$n_sessions)) > 0, "Yes", "No")
  
  # If attendance_data == "No", create a null dataframe containing the columns "Name", "Total available", and "Total attended" containing one row of NA values
  # If attendance_data == "Yes", run the code below.
  
  if (attendance_data == "No") {
    
    ssa <- data.frame(`Name` = NA, `Total available` = NA, `Total attended` = NA)
    colnames(ssa) <- c("Name", "Total available", "Total attended")
    
  } else if (attendance_data == "Yes") {
    
    # Step 1: Extract the first row to use as new column names
    new_colnames <- sapply(ssa[1, ], function(x) ifelse(is.na(x), NA, as.character(x)))
    
    # Step 2: Replace existing column names where the first row is not NA
    colnames(ssa) <- ifelse(is.na(new_colnames), colnames(ssa), new_colnames)
    
    # Remove duplicate column names
    ssa <- ssa[, !duplicated(names(ssa))]
    
    # Remove columns with names starting with "x" or "student_name"
    ssa <- ssa %>%
      select(-starts_with("x"), -starts_with("student_name"))
    
    # Remove the first row from the dataframe
    ssa <- ssa[-1, ]
    
      # Tidy up sessions (some entered as numbers, others as days)
      convert_sessions <- function(sessions) {
        if (is.na(sessions)) {
          return(NA_integer_)
        } else if (str_detect(sessions, "^\\d+$")) {
          # If it is a number, convert to integer
          return(as.integer(sessions))
        } else {
          # If it is text, count the number of items and convert to integer
          days <- str_split(sessions, ",\\s*")[[1]]
          return(length(days))
        }
      }

      ssa <- ssa %>%
        mutate(sessions_taught_per_week = sapply(sessions_taught_per_week, convert_sessions))

      # Pivot longer the columns containing after sessions_taught_per_week
      # Find the position of the 'sessions_taught_per_week' column
      sessions_taught_per_week_position <- which(names(ssa) == "sessions_taught_per_week")

      ssa <- ssa %>%
        pivot_longer(cols = (sessions_taught_per_week_position + 1):ncol(ssa), names_to = "Name", values_to = "n_sessions") %>%
        mutate(n_sessions = sapply(n_sessions, convert_sessions)) %>%
        select(Name, sessions_taught_per_week, n_sessions) %>%
        group_by(Name) %>%
        dplyr::summarise(`Total available` = sum(sessions_taught_per_week, na.rm = TRUE), `Total attended` = sum(n_sessions, na.rm = TRUE))
     
      colnames(ssa) <- c("Name", "Total available", "Total attended")
      
      
      
    }
  
  
  ##############################################################################################
  
  # Join "final_adapted_bryant_assessment_Overall_Score_/50" to tier2_details using "student_details_Student_Name", keeping all
  data <- full_join(tier2_details, ab, by = "student_details_Student_Name") %>%
    mutate(National_Student_number = ifelse(is.na(student_details_National_Student_number_tier2), student_details_National_Student_number_ad_bry,
           student_details_National_Student_number_tier2)) %>%
    mutate(Region = ifelse(is.na(Region_tier2), Region_ad_bry, Region_tier2)) %>%
    mutate(Teacher = ifelse(is.na(Teacher_tier2), Teacher_ad_bry, Teacher_tier2)) %>%
    # Split student name into a column for the first name and a column for the last name
    separate(student_details_Student_Name, into = c("Firstname", "Surname"), sep = " ", remove = FALSE) %>%
    mutate(Surname = replace_na(Surname, "Surname not provided")) %>%
    mutate(School = NA) %>%
    dplyr::relocate(Firstname, Surname, Region, School, Teacher, student_details_Year_level, student_details_Gender, student_details_Ethnicity_1,
                    student_details_Ethnicity_2)


  # Join the student session attendance data
  data <- left_join(data, ssa, by = c("student_details_Student_Name" = "Name"))

  # Append the modified data frame to the list
  data_list <- append(data_list, list(data))
}


###########################################################################################################################################
###########################################################################################################################################

# Combine all data frames together by row
length(data_list)
combined_data <- bind_rows(data_list)
length(unique(combined_data$Teacher)) #90 teachers.

# Identify rows with the same value for "Firstname" but different values for "Surname" and one of the values for "Surname" is "Surname not provided"
combined_data <- combined_data %>%
  group_by(Firstname, Teacher) %>%
  mutate(same_child = ifelse(n_distinct(Surname) > 1 & any(Surname == "Surname not provided"), 1, 0)) %>%
  ungroup() %>%
  relocate(same_child)

# If same_child==1, replace "Surname not provided" with the other value of "Surname"
combined_data <- combined_data %>%
  mutate(Surname = ifelse(Surname == "Surname not provided", NA, Surname)) %>%
  group_by(Firstname, Teacher) %>%
  tidyr::fill(Surname, .direction = "updown") %>%
  ungroup()

#View(combined_data %>% filter(student_details_National_Student_number_tier2==144075364))

combined_data <- combined_data %>%
  group_by(Firstname, Surname, Teacher) %>%
  fill(everything(), .direction = "downup") %>%
  mutate(student_details_Student_Name = paste(Firstname, Surname, sep = " ")) %>%
  distinct() %>%
  ungroup()

# Tidy up rogue variables
combined_data <- combined_data %>%
  mutate(student_details_Gender = ifelse(grepl("Fem|fem", student_details_Gender), "F", student_details_Gender)) %>%
  # Remove characters including and after the "/" in columns containing "Overall_Score_/50"
  mutate_at(vars(contains("Overall_Score_/50")), ~str_remove(., "/.*")) %>%
  # Convert same columns to numeric
  mutate_at(vars(contains("Overall_Score_/50")), as.numeric)

table(combined_data$student_details_Year_level)
###########################################################################################################################################
###########################################################################################################################################


# Create the final dataframe needed by LM
DATA <- combined_data %>%
  dplyr::select(Firstname, Surname, Region, School, Teacher, student_details_Year_level, student_details_Gender, student_details_Ethnicity_1, student_details_Ethnicity_2,
                neurodiverse_diagnosis, neurodiverse_suspicion, `Total available`, `Total attended`, 
                `pre_adapted_bryant_assessment_Overall_Score_/50`, `mid_adapted_bryant_assessment_Overall_Score_/50`, 
                `final_adapted_bryant_assessment_Overall_Score_/50`) %>%
  dplyr::rename(`Year level` = student_details_Year_level, `Ethnicity 1` = student_details_Ethnicity_1, `Ethnicity 2` = student_details_Ethnicity_2, 
                Gender = student_details_Gender,
                `Neurodiversity diagnosis` = neurodiverse_diagnosis, `Neurodiversity suspicion` = neurodiverse_suspicion,
                `Pre AB score` = `pre_adapted_bryant_assessment_Overall_Score_/50`, `Mid AB score` = `mid_adapted_bryant_assessment_Overall_Score_/50`, 
                `Final AB score` = `final_adapted_bryant_assessment_Overall_Score_/50`)

table(DATA$Gender)

# Ethnicity table - NB this only takes the first ethnicity value
ethnicity <- DATA %>% 
  dplyr::select(Gender, `Year level`, `Ethnicity 1`) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                               grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                               grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                               is.na(Ethnicity) ~ "Unknown",
                               .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown")))
  
# tbl_strata(data = ethnicity,
#            strata = Gender,
#            .tbl_fun =
#              ~ .x %>%
#              tbl_summary(by = Ethnicity, 
#                          missing = "no",
#                          statistic = list(all_categorical() ~ "{n}/{N} ({p}%)"),
#                          percent = "row") %>%
#              bold_labels() %>%
#              modify_header(label = ""),
#            .header = "**{strata}**, N = {n}") %>%
#   modify_footnote(everything() ~ NA) %>%
#   as_hux_xlsx("ethnicity_summary.xlsx")


# Use dplyr instead

eth_tbl_n <- ethnicity %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(n = n(), .groups = "drop") %>%
  ungroup() %>%
  pivot_wider(names_from = c(Ethnicity, Gender), values_from = "n") %>%
  mutate(Total = rowSums(.[-1])); eth_tbl_n

tot_row <- c(NA, colSums(eth_tbl_n[, -1], na.rm = T))
  
eth_tbl_n <- rbind(eth_tbl_n, tot_row) %>%
  mutate(`Year level` = as.character(`Year level`)) %>%
  mutate(`Year level` = factor(ifelse(is.na(`Year level`), "Total", `Year level`))) %>%
  # Make first column row names
  column_to_rownames(var = "Year level")

# Calculate row sums for all columns except the last one
row_sums <- rowSums(eth_tbl_n[, -ncol(eth_tbl_n)])

# Calculate percentages for each cell by dividing by the row sum (excluding the last column)
percentages <- round(sweep(eth_tbl_n[, -ncol(eth_tbl_n)], 1, row_sums, FUN = "/") * 100, 0)

# Calculate percentages for the last column as percentages of the total
total_percentage <- round((eth_tbl_n[, ncol(eth_tbl_n)] / sum(eth_tbl_n[-length(eth_tbl_n$Total), ]$Total)) * 100, 0)

# Create a new dataframe to store combined counts and percentages
df_combined <- eth_tbl_n

# Loop through columns (excluding the last one)
for (col in names(eth_tbl_n)[-ncol(eth_tbl_n)]) {
  # Create formatted string with count and percentage
  combined_col <- paste0(eth_tbl_n[[col]], " (", round(percentages[[col]], 1), "%)")
  # Assign the formatted string to the column in the combined dataframe
  df_combined[[col]] <- combined_col
}

# Create formatted string for the Total column
combined_total <- sprintf("%d (%.1f%%)", eth_tbl_n$Total, total_percentage)
df_combined$Total <- combined_total

colnames(df_combined) <- c("Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Total")

# Set rownames as first column
ethn_all <- df_combined %>%
  rownames_to_column(var = "Year level") %>%
  mutate(Region = "All regions") %>%
  dplyr::relocate(Region)

###########################################################################################################################################
# Now repeat for Canterbury, Hawkes Bay, Manawatu/Rangitikei/Whanganui, Nelson/Marlborough/West Coast

#1. Canterbury
ethnicity <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`) %>%
  filter(Region=="Canterbury") %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown")))

eth_tbl_n <- ethnicity %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(n = n(), .groups = "drop") %>%
  ungroup() %>%
  pivot_wider(names_from = c(Ethnicity, Gender), values_from = "n") %>%
  mutate(Total = rowSums(.[-1])); eth_tbl_n

tot_row <- c(NA, colSums(eth_tbl_n[, -1], na.rm = T))

eth_tbl_n <- rbind(eth_tbl_n, tot_row) %>%
  mutate(`Year level` = as.character(`Year level`)) %>%
  mutate(`Year level` = factor(ifelse(is.na(`Year level`), "Total", `Year level`))) %>%
  # Make first column row names
  column_to_rownames(var = "Year level")

# Calculate row sums for all columns except the last one
row_sums <- rowSums(eth_tbl_n[, -ncol(eth_tbl_n)])

# Calculate percentages for each cell by dividing by the row sum (excluding the last column)
percentages <- round(sweep(eth_tbl_n[, -ncol(eth_tbl_n)], 1, row_sums, FUN = "/") * 100, 0)

# Calculate percentages for the last column as percentages of the total
total_percentage <- round((eth_tbl_n[, ncol(eth_tbl_n)] / sum(eth_tbl_n[-length(eth_tbl_n$Total), ]$Total)) * 100, 0)

# Create a new dataframe to store combined counts and percentages
df_combined <- eth_tbl_n

# Loop through columns (excluding the last one)
for (col in names(eth_tbl_n)[-ncol(eth_tbl_n)]) {
  # Create formatted string with count and percentage
  combined_col <- paste0(eth_tbl_n[[col]], " (", round(percentages[[col]], 1), "%)")
  # Assign the formatted string to the column in the combined dataframe
  df_combined[[col]] <- combined_col
}

# Create formatted string for the Total column
combined_total <- sprintf("%d (%.1f%%)", eth_tbl_n$Total, total_percentage)
df_combined$Total <- combined_total

colnames(df_combined) <- c("Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Total")

# Set rownames as first column
ethn_cant <- df_combined %>%
  rownames_to_column(var = "Year level") %>%
  mutate(Region = "Canterbury") %>%
  dplyr::relocate(Region)

#2. Hawkes Bay
ethnicity <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`) %>%
  filter(Region=="Hawkes Bay") %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown")))

eth_tbl_n <- ethnicity %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(n = n(), .groups = "drop") %>%
  ungroup() %>%
  pivot_wider(names_from = c(Ethnicity, Gender), values_from = "n") %>%
  mutate(Total = rowSums(.[-1])); eth_tbl_n

tot_row <- c(NA, colSums(eth_tbl_n[, -1], na.rm = T))

eth_tbl_n <- rbind(eth_tbl_n, tot_row) %>%
  mutate(`Year level` = as.character(`Year level`)) %>%
  mutate(`Year level` = factor(ifelse(is.na(`Year level`), "Total", `Year level`))) %>%
  # Make first column row names
  column_to_rownames(var = "Year level")

# Calculate row sums for all columns except the last one
row_sums <- rowSums(eth_tbl_n[, -ncol(eth_tbl_n)])

# Calculate percentages for each cell by dividing by the row sum (excluding the last column)
percentages <- round(sweep(eth_tbl_n[, -ncol(eth_tbl_n)], 1, row_sums, FUN = "/") * 100, 0)

# Calculate percentages for the last column as percentages of the total
total_percentage <- round((eth_tbl_n[, ncol(eth_tbl_n)] / sum(eth_tbl_n[-length(eth_tbl_n$Total), ]$Total)) * 100, 0)

# Create a new dataframe to store combined counts and percentages
df_combined <- eth_tbl_n

# Loop through columns (excluding the last one)
for (col in names(eth_tbl_n)[-ncol(eth_tbl_n)]) {
  # Create formatted string with count and percentage
  combined_col <- paste0(eth_tbl_n[[col]], " (", round(percentages[[col]], 1), "%)")
  # Assign the formatted string to the column in the combined dataframe
  df_combined[[col]] <- combined_col
}

# Create formatted string for the Total column
combined_total <- sprintf("%d (%.1f%%)", eth_tbl_n$Total, total_percentage)
df_combined$Total <- combined_total

colnames(df_combined) <- c("Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Total")

# Set rownames as first column
ethn_HB <- df_combined %>%
  rownames_to_column(var = "Year level")  %>%
  mutate(Region = "Hawkes Bay") %>%
  dplyr::relocate(Region)

#3. Manawatu/Rangitikei/Whanganui
ethnicity <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`) %>%
  filter(Region=="Manawatu/Rangitikei/Whanganui") %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown")))

eth_tbl_n <- ethnicity %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(n = n(), .groups = "drop") %>%
  ungroup() %>%
  pivot_wider(names_from = c(Ethnicity, Gender), values_from = "n") %>%
  mutate(Total = rowSums(.[-1])); eth_tbl_n

tot_row <- c(NA, colSums(eth_tbl_n[, -1], na.rm = T))

eth_tbl_n <- rbind(eth_tbl_n, tot_row) %>%
  mutate(`Year level` = as.character(`Year level`)) %>%
  mutate(`Year level` = factor(ifelse(is.na(`Year level`), "Total", `Year level`))) %>%
  # Make first column row names
  column_to_rownames(var = "Year level")

# Calculate row sums for all columns except the last one
row_sums <- rowSums(eth_tbl_n[, -ncol(eth_tbl_n)])

# Calculate percentages for each cell by dividing by the row sum (excluding the last column)
percentages <- round(sweep(eth_tbl_n[, -ncol(eth_tbl_n)], 1, row_sums, FUN = "/") * 100, 0)

# Calculate percentages for the last column as percentages of the total
total_percentage <- round((eth_tbl_n[, ncol(eth_tbl_n)] / sum(eth_tbl_n[-length(eth_tbl_n$Total), ]$Total)) * 100, 0)

# Create a new dataframe to store combined counts and percentages
df_combined <- eth_tbl_n

# Loop through columns (excluding the last one)
for (col in names(eth_tbl_n)[-ncol(eth_tbl_n)]) {
  # Create formatted string with count and percentage
  combined_col <- paste0(eth_tbl_n[[col]], " (", round(percentages[[col]], 1), "%)")
  # Assign the formatted string to the column in the combined dataframe
  df_combined[[col]] <- combined_col
}

# Create formatted string for the Total column
combined_total <- sprintf("%d (%.1f%%)", eth_tbl_n$Total, total_percentage)
df_combined$Total <- combined_total

colnames(df_combined) <- c("Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Total")

# Set rownames as first column
ethn_man <- df_combined %>%
  rownames_to_column(var = "Year level") %>% 
  mutate(Region = "Manawatu/Rangitikei/Whanganui") %>%
  dplyr::relocate(Region)

#4. Nelson
ethnicity <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`) %>%
  filter(Region=="Nelson/Marlborough/West Coast") %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown")))

eth_tbl_n <- ethnicity %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(n = n(), .groups = "drop") %>%
  ungroup() %>%
  pivot_wider(names_from = c(Ethnicity, Gender), values_from = "n") %>%
  mutate(Total = rowSums(.[-1])); eth_tbl_n

tot_row <- c(NA, colSums(eth_tbl_n[, -1], na.rm = T))

eth_tbl_n <- rbind(eth_tbl_n, tot_row) %>%
  mutate(`Year level` = as.character(`Year level`)) %>%
  mutate(`Year level` = factor(ifelse(is.na(`Year level`), "Total", `Year level`))) %>%
  # Make first column row names
  column_to_rownames(var = "Year level")

# Calculate row sums for all columns except the last one
row_sums <- rowSums(eth_tbl_n[, -ncol(eth_tbl_n)])

# Calculate percentages for each cell by dividing by the row sum (excluding the last column)
percentages <- round(sweep(eth_tbl_n[, -ncol(eth_tbl_n)], 1, row_sums, FUN = "/") * 100, 0)

# Calculate percentages for the last column as percentages of the total
total_percentage <- round((eth_tbl_n[, ncol(eth_tbl_n)] / sum(eth_tbl_n[-length(eth_tbl_n$Total), ]$Total)) * 100, 0)

# Create a new dataframe to store combined counts and percentages
df_combined <- eth_tbl_n

# Loop through columns (excluding the last one)
for (col in names(eth_tbl_n)[-ncol(eth_tbl_n)]) {
  # Create formatted string with count and percentage
  combined_col <- paste0(eth_tbl_n[[col]], " (", round(percentages[[col]], 1), "%)")
  # Assign the formatted string to the column in the combined dataframe
  df_combined[[col]] <- combined_col
}

# Create formatted string for the Total column
combined_total <- sprintf("%d (%.1f%%)", eth_tbl_n$Total, total_percentage)
df_combined$Total <- combined_total

colnames(df_combined) <- c("Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Total")

# Set rownames as first column
ethn_nel <- df_combined %>%
  rownames_to_column(var = "Year level")  %>% 
  mutate(Region = "Nelson/Marlborough/West Coasti") %>%
  dplyr::relocate(Region)

################################################################################

# Write to a workbook
wb <- createWorkbook()

# Add a worksheet to the workbook
addWorksheet(wb, sheetName = "Ethnicity")

# Write each dataframe to the worksheet with spacing
writeData(wb, sheet = "Ethnicity", x = ethn_all, startRow = 1, startCol = 1)
writeData(wb, sheet = "Ethnicity", x = ethn_cant, startRow = nrow(df_combined) + 4, startCol = 1)
writeData(wb, sheet = "Ethnicity", x = ethn_HB, startRow = nrow(df_combined) + nrow(ethn_cant) + 6, startCol = 1)
writeData(wb, sheet = "Ethnicity", x = ethn_man, startRow = nrow(df_combined) + nrow(ethn_cant) + nrow(ethn_HB) + 9, startCol = 1)
writeData(wb, sheet = "Ethnicity", x = ethn_nel, startRow = nrow(df_combined) + nrow(ethn_cant) + nrow(ethn_HB) + nrow(ethn_man) + 12, startCol = 1)


################################################################################################################################################################
################################################################################################################################################################

ld <- DATA %>%
  dplyr::select(`Year level`, `Neurodiversity diagnosis`, `Neurodiversity suspicion`) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(`Neurodiversity diagnosis` = sum(!is.na(`Neurodiversity diagnosis`), na.rm = TRUE),
                   `Neurodiversity suspicion` = sum(!is.na(`Neurodiversity suspicion`), na.rm = TRUE),
                   `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`) %>%
  ungroup()

summary_row <- ld %>%
  summarise(Region = "Total", `Year level` = "Total", `Neurodiversity diagnosis` = sum(`Neurodiversity diagnosis`, na.rm=T),
            `Neurodiversity suspicion` = sum(`Neurodiversity suspicion`, na.rm=T),
            `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`)

ld_all <- bind_rows(ld, summary_row) %>%
  mutate(Region = "All regions") %>%
  relocate(Region)
  
# Repeat for each region
#1. Canterbury
ld <- DATA %>%
  dplyr::select(Region, `Year level`, `Neurodiversity diagnosis`, `Neurodiversity suspicion`) %>%
  filter(Region=="Canterbury") %>%
  dplyr::select(-Region) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(`Neurodiversity diagnosis` = sum(!is.na(`Neurodiversity diagnosis`), na.rm = TRUE),
                   `Neurodiversity suspicion` = sum(!is.na(`Neurodiversity suspicion`), na.rm = TRUE),
                   `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`) %>%
  ungroup()

summary_row <- ld %>%
  summarise(`Year level` = "Total", `Neurodiversity diagnosis` = sum(`Neurodiversity diagnosis`, na.rm=T),
            `Neurodiversity suspicion` = sum(`Neurodiversity suspicion`, na.rm=T),
            `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`)

ld_can <- bind_rows(ld, summary_row) %>%
  mutate(Region = "Canterbury") %>%
  relocate(Region)

#2. Hawkes Bay
ld <- DATA %>%
  dplyr::select(Region, `Year level`, `Neurodiversity diagnosis`, `Neurodiversity suspicion`) %>%
  filter(Region=="Hawkes Bay") %>%
  dplyr::select(-Region) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(`Neurodiversity diagnosis` = sum(!is.na(`Neurodiversity diagnosis`), na.rm = TRUE),
                   `Neurodiversity suspicion` = sum(!is.na(`Neurodiversity suspicion`), na.rm = TRUE),
                   `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`) %>%
  ungroup()

summary_row <- ld %>%
  summarise(`Year level` = "Total", `Neurodiversity diagnosis` = sum(`Neurodiversity diagnosis`, na.rm=T),
            `Neurodiversity suspicion` = sum(`Neurodiversity suspicion`, na.rm=T),
            `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`)

ld_HB <- bind_rows(ld, summary_row) %>%
  mutate(Region = "Hawkes Bay") %>%
  relocate(Region)


#3. Manawatu/Rangitikei/Whanganui
ld <- DATA %>%
  dplyr::select(Region, `Year level`, `Neurodiversity diagnosis`, `Neurodiversity suspicion`) %>%
  filter(Region=="Manawatu/Rangitikei/Whanganui") %>%
  dplyr::select(-Region) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(`Neurodiversity diagnosis` = sum(!is.na(`Neurodiversity diagnosis`), na.rm = TRUE),
                   `Neurodiversity suspicion` = sum(!is.na(`Neurodiversity suspicion`), na.rm = TRUE),
                   `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`) %>%
  ungroup()

summary_row <- ld %>%
  summarise(`Year level` = "Total", `Neurodiversity diagnosis` = sum(`Neurodiversity diagnosis`, na.rm=T),
            `Neurodiversity suspicion` = sum(`Neurodiversity suspicion`, na.rm=T),
            `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`)

ld_man <- bind_rows(ld, summary_row) %>%
  mutate(Region = "Manawatu/Rangitikei/Whanganui") %>%
  relocate(Region)


#4. Nelson/Marlborough/West Coast
ld <- DATA %>%
  dplyr::select(Region, `Year level`, `Neurodiversity diagnosis`, `Neurodiversity suspicion`) %>%
  filter(Region=="Nelson/Marlborough/West Coast") %>%
  dplyr::select(-Region) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(`Neurodiversity diagnosis` = sum(!is.na(`Neurodiversity diagnosis`), na.rm = TRUE),
                   `Neurodiversity suspicion` = sum(!is.na(`Neurodiversity suspicion`), na.rm = TRUE),
                   `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`) %>%
  ungroup()

summary_row <- ld %>%
  summarise(`Year level` = "Total", `Neurodiversity diagnosis` = sum(`Neurodiversity diagnosis`, na.rm=T),
            `Neurodiversity suspicion` = sum(`Neurodiversity suspicion`, na.rm=T),
            `No. learners presenting learning challenges` = `Neurodiversity diagnosis` + `Neurodiversity suspicion`)

ld_nel <- bind_rows(ld, summary_row) %>%
  mutate(Region = "Nelson/Marlborough/West Coast") %>%
  relocate(Region)

##################################

# Add a worksheet to the workbook
addWorksheet(wb, sheetName = "Learning differences")

# Write each dataframe to the worksheet with spacing
writeData(wb, sheet = "Learning differences", x = ld_all, startRow = 1, startCol = 1)
writeData(wb, sheet = "Learning differences", x = ld_can, startRow = nrow(ld_all) + 4, startCol = 1)
writeData(wb, sheet = "Learning differences", x = ld_HB, startRow = nrow(ld_all) + nrow(ld_can) + 6, startCol = 1)
writeData(wb, sheet = "Learning differences", x = ld_man, startRow = nrow(ld_all) + nrow(ld_can) + nrow(ld_HB) + 9, startCol = 1)
writeData(wb, sheet = "Learning differences", x = ld_nel, startRow = nrow(ld_all) + nrow(ld_can) + nrow(ld_HB) + nrow(ld_man) + 12, startCol = 1)

################################################################################################################################################################
################################################################################################################################################################

att <- DATA %>%
  dplyr::select(`Year level`, `Total available`, `Total attended`) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(`Intervention Session Attendance` = round(100*sum(`Total attended`, na.rm = TRUE)/sum(`Total available`, na.rm = T), 1)) %>%
  ungroup() %>%
  mutate(`Intervention Session Attendance` = ifelse(is.nan(`Intervention Session Attendance`), "No data", `Intervention Session Attendance`)) %>%
  mutate(`Intervention Session Attendance` = ifelse(`Intervention Session Attendance` != "No data", paste0(`Intervention Session Attendance`, "%"),
         `Intervention Session Attendance`)) %>%
  mutate(Region = "All regions") %>%
  relocate(Region)

# Repeat for each region
att_can <- DATA %>%
  dplyr::select(Region, `Year level`, `Total available`, `Total attended`) %>%
  filter(Region=="Canterbury") %>%
  dplyr::select(-Region) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(`Intervention Session Attendance` = round(100*sum(`Total attended`, na.rm = TRUE)/sum(`Total available`, na.rm = T), 1)) %>%
  ungroup() %>%
  mutate(`Intervention Session Attendance` = ifelse(is.nan(`Intervention Session Attendance`), "No data", `Intervention Session Attendance`)) %>%
  mutate(`Intervention Session Attendance` = ifelse(`Intervention Session Attendance` != "No data", paste0(`Intervention Session Attendance`, "%"),
                                                    `Intervention Session Attendance`)) %>%
  mutate(Region = "Canterbury") %>%
  relocate(Region)


att_HB <- DATA %>%
  dplyr::select(Region, `Year level`, `Total available`, `Total attended`) %>%
  filter(Region=="Hawkes Bay") %>%
  dplyr::select(-Region) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(`Intervention Session Attendance` = round(100*sum(`Total attended`, na.rm = TRUE)/sum(`Total available`, na.rm = T), 1)) %>%
  ungroup() %>%
  mutate(`Intervention Session Attendance` = ifelse(is.nan(`Intervention Session Attendance`), "No data", `Intervention Session Attendance`)) %>%
  mutate(`Intervention Session Attendance` = ifelse(`Intervention Session Attendance` != "No data", paste0(`Intervention Session Attendance`, "%"),
                                                    `Intervention Session Attendance`)) %>%
  mutate(Region = "Hawkes Bay") %>%
  relocate(Region)

att_man <- DATA %>%
  dplyr::select(Region, `Year level`, `Total available`, `Total attended`) %>%
  filter(Region=="Manawatu/Rangitikei/Whanganui") %>%
  dplyr::select(-Region) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(`Intervention Session Attendance` = round(100*sum(`Total attended`, na.rm = TRUE)/sum(`Total available`, na.rm = T), 1)) %>%
  ungroup() %>%
  mutate(`Intervention Session Attendance` = ifelse(is.nan(`Intervention Session Attendance`), "No data", `Intervention Session Attendance`)) %>%
  mutate(`Intervention Session Attendance` = ifelse(`Intervention Session Attendance` != "No data", paste0(`Intervention Session Attendance`, "%"),
                                                    `Intervention Session Attendance`)) %>%
  mutate(Region = "Manawatu/Rangitikei/Whanganui") %>%
  relocate(Region)

att_nel <- DATA %>%
  dplyr::select(Region, `Year level`, `Total available`, `Total attended`) %>%
  filter(Region=="Nelson/Marlborough/West Coast") %>%
  dplyr::select(-Region) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(`Intervention Session Attendance` = round(100*sum(`Total attended`, na.rm = TRUE)/sum(`Total available`, na.rm = T), 1)) %>%
  ungroup() %>%
  mutate(`Intervention Session Attendance` = ifelse(is.nan(`Intervention Session Attendance`), "No data", `Intervention Session Attendance`)) %>%
  mutate(`Intervention Session Attendance` = ifelse(`Intervention Session Attendance` != "No data", paste0(`Intervention Session Attendance`, "%"),
                                                    `Intervention Session Attendance`)) %>%
  mutate(Region = "Nelson/Marlborough/West Coast") %>%
  relocate(Region)

##################################

# Add a worksheet to the workbook
addWorksheet(wb, sheetName = "Attendance")

# Write each dataframe to the worksheet with spacing
writeData(wb, sheet = "Attendance", x = att, startRow = 1, startCol = 1)
writeData(wb, sheet = "Attendance", x = att_can, startRow = nrow(att) + 4, startCol = 1)
writeData(wb, sheet = "Attendance", x = att_HB, startRow = nrow(att) + nrow(att_can) + 6, startCol = 1)
writeData(wb, sheet = "Attendance", x = att_man, startRow = nrow(att) + nrow(att_can) + nrow(att_HB) + 9, startCol = 1)
writeData(wb, sheet = "Attendance", x = att_nel, startRow = nrow(att) + nrow(att_can) + nrow(att_HB) + nrow(att_man) + 12, startCol = 1)

################################################################################################################################################################
################################################################################################################################################################

ad_pre <- DATA %>% 
  dplyr::select(Gender, `Year level`, `Ethnicity 1`, `Pre AB score`) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
  mutate(across(contains("score"), as.numeric))

# Overall column and row
col <- ad_pre %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(Overall = median(`Pre AB score`, na.rm = TRUE))

ad_pre_w <- ad_pre %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(`Pre AB score` = median(`Pre AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Pre AB score`)) %>%
  ungroup() %>%
  bind_cols(col[2])

row <- ad_pre %>%
  group_by(Gender, Ethnicity, .drop = F) %>%
  dplyr::summarise(`Pre AB score` = median(`Pre AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Pre AB score`)) %>%
  ungroup() %>%
  mutate(`Year level` = "Overall") %>%
  relocate(`Year level`) %>%
  mutate(Overall = median(DATA[DATA$Region == region, ]$`Pre AB score`, na.rm = TRUE))

ad_pre_w <- rbind(ad_pre_w, row) %>% 
  mutate(
    #Assessment = "Pre assessment", 
    Region = "All regions") %>%
  relocate(Region)

colnames(ad_pre_w) <- c("Region", "Year level", "Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", 
                        "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Overall")

ad_pre_w_all <- ad_pre_w

ad_mid <- DATA %>% 
  dplyr::select(Gender, `Year level`, `Ethnicity 1`, `Mid AB score`) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
  mutate(across(contains("score"), as.numeric))

# Overall column and row
col <- ad_mid %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(Overall = median(`Mid AB score`, na.rm = TRUE))

ad_mid_w <- ad_mid %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(`Mid AB score` = median(`Mid AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Mid AB score`)) %>%
  ungroup() %>%
  bind_cols(col[2])

row <- ad_mid %>%
  group_by(Gender, Ethnicity, .drop = F) %>%
  dplyr::summarise(`Mid AB score` = median(`Mid AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Mid AB score`)) %>%
  ungroup() %>%
  mutate(`Year level` = "Overall") %>%
  relocate(`Year level`) %>%
  mutate(Overall = median(DATA[DATA$Region == region, ]$`Mid AB score`, na.rm = TRUE))

ad_mid_w <- rbind(ad_mid_w, row) %>% 
  mutate(
    #Assessment = "Mid assessment", 
    Region = "All regions") %>%
  relocate(Region)

colnames(ad_mid_w) <- c( "Region", "Year level", "Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", 
                        "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Overall")

ad_mid_w_all <- ad_mid_w

ad_fin <- DATA %>% 
  dplyr::select(Gender, `Year level`, `Ethnicity 1`, `Final AB score`) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
  mutate(across(contains("score"), as.numeric))

# Overall column and row
col <- ad_fin %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(Overall = median(`Final AB score`, na.rm = TRUE))

ad_fin_w <- ad_fin %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(`Final AB score` = median(`Final AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Final AB score`)) %>%
  ungroup() %>%
  bind_cols(col[2])

row <- ad_fin %>%
  group_by(Gender, Ethnicity, .drop = F) %>%
  dplyr::summarise(`Final AB score` = median(`Final AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Final AB score`)) %>%
  ungroup() %>%
  mutate(`Year level` = "Overall") %>%
  relocate(`Year level`) %>%
  mutate(Overall = median(DATA[DATA$Region == region, ]$`Final AB score`, na.rm = TRUE))

ad_fin_w <- rbind(ad_fin_w, row) %>% 
  mutate(
    #Assessment = "Final assessment", 
    Region = "All regions") %>%
  relocate(Region)

colnames(ad_fin_w) <- c("Region", "Year level", "Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", 
                        "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Overall")

ad_fin_w_all <- ad_fin_w

#########################################################
# Repeat for each region
#1. Canterbury
region <- "Canterbury"

ad_pre <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `Pre AB score`) %>%
  filter(Region==region) %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
  mutate(across(contains("score"), as.numeric))

# Overall column and row
col <- ad_pre %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(Overall = median(`Pre AB score`, na.rm = TRUE))

ad_pre_w <- ad_pre %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(`Pre AB score` = median(`Pre AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Pre AB score`)) %>%
  ungroup() %>%
  bind_cols(col[2])

row <- ad_pre %>%
  group_by(Gender, Ethnicity, .drop = F) %>%
  dplyr::summarise(`Pre AB score` = median(`Pre AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Pre AB score`)) %>%
  ungroup() %>%
  mutate(`Year level` = "Overall") %>%
  relocate(`Year level`) %>%
  mutate(Overall = median(DATA[DATA$Region == region, ]$`Pre AB score`, na.rm = TRUE))

ad_pre_w <- rbind(ad_pre_w, row)  %>% 
  mutate(
    #Assessment = "Pre assessment", 
    Region = region) %>%
  relocate(Region)

colnames(ad_pre_w) <- c("Region", "Year level", "Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", 
                        "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Overall")
ad_pre_w_can <- ad_pre_w

ad_mid <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `Mid AB score`) %>%
  filter(Region==region) %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
  mutate(across(contains("score"), as.numeric))

# Overall column and row
col <- ad_mid %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(Overall = median(`Mid AB score`, na.rm = TRUE))

ad_mid_w <- ad_mid %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(`Mid AB score` = median(`Mid AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Mid AB score`)) %>%
  ungroup() %>%
  bind_cols(col[2])

row <- ad_mid %>%
  group_by(Gender, Ethnicity, .drop = F) %>%
  dplyr::summarise(`Mid AB score` = median(`Mid AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Mid AB score`)) %>%
  ungroup() %>%
  mutate(`Year level` = "Overall") %>%
  relocate(`Year level`) %>%
  mutate(Overall = median(DATA[DATA$Region == region, ]$`Mid AB score`, na.rm = TRUE))

ad_mid_w <- rbind(ad_mid_w, row) %>% 
  mutate(
    #Assessment = "Mid assessment", 
    Region = region) %>%
  relocate(Region)

colnames(ad_mid_w) <- c("Region", "Year level", "Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", 
                        "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Overall")
ad_mid_w_can <- ad_mid_w

ad_fin <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `Final AB score`) %>%
  filter(Region==region) %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
  mutate(across(contains("score"), as.numeric))

# Overall column and row
col <- ad_fin %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(Overall = median(`Final AB score`, na.rm = TRUE))

ad_fin_w <- ad_fin %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(`Final AB score` = median(`Final AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Final AB score`)) %>%
  ungroup() %>%
  bind_cols(col[2])

row <- ad_fin %>%
  group_by(Gender, Ethnicity, .drop = F) %>%
  dplyr::summarise(`Final AB score` = median(`Final AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Final AB score`)) %>%
  ungroup() %>%
  mutate(`Year level` = "Overall") %>%
  relocate(`Year level`) %>%
  mutate(Overall = median(DATA[DATA$Region == region, ]$`Final AB score`, na.rm = TRUE))

ad_fin_w <- rbind(ad_fin_w, row) %>% 
  mutate(
    #Assessment = "Final assessment", 
    Region = region) %>%
  relocate(Region)

colnames(ad_fin_w) <- c("Region", "Year level", "Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", 
                        "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Overall")
ad_fin_w_can <- ad_fin_w

#########################################################
#2. Hawkes Bay
region <- "Hawkes Bay"

ad_pre <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `Pre AB score`) %>%
  filter(Region==region) %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
  mutate(across(contains("score"), as.numeric))

# Overall column and row
col <- ad_pre %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(Overall = median(`Pre AB score`, na.rm = TRUE))

ad_pre_w <- ad_pre %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(`Pre AB score` = median(`Pre AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Pre AB score`)) %>%
  ungroup() %>%
  bind_cols(col[2])

row <- ad_pre %>%
  group_by(Gender, Ethnicity, .drop = F) %>%
  dplyr::summarise(`Pre AB score` = median(`Pre AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Pre AB score`)) %>%
  ungroup() %>%
  mutate(`Year level` = "Overall") %>%
  relocate(`Year level`) %>%
  mutate(Overall = median(DATA[DATA$Region == region, ]$`Pre AB score`, na.rm = TRUE))

ad_pre_w <- rbind(ad_pre_w, row) %>% 
  mutate(
    #Assessment = "Pre assessment", 
    Region = region) %>%
  relocate(Region)

colnames(ad_pre_w) <- c("Region", "Year level", "Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", 
                        "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Overall")
ad_pre_w_HB <- ad_pre_w

ad_mid <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `Mid AB score`) %>%
  filter(Region==region) %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
  mutate(across(contains("score"), as.numeric))

# Overall column and row
col <- ad_mid %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(Overall = median(`Mid AB score`, na.rm = TRUE))

ad_mid_w <- ad_mid %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(`Mid AB score` = median(`Mid AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Mid AB score`)) %>%
  ungroup() %>%
  bind_cols(col[2])

row <- ad_mid %>%
  group_by(Gender, Ethnicity, .drop = F) %>%
  dplyr::summarise(`Mid AB score` = median(`Mid AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Mid AB score`)) %>%
  ungroup() %>%
  mutate(`Year level` = "Overall") %>%
  relocate(`Year level`) %>%
  mutate(Overall = median(DATA[DATA$Region == region, ]$`Mid AB score`, na.rm = TRUE))

ad_mid_w <- rbind(ad_mid_w, row) %>% 
  mutate(
    #Assessment = "Mid assessment", 
    Region = region) %>%
  relocate(Region)

colnames(ad_mid_w) <- c("Region", "Year level", "Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", 
                        "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Overall")
ad_mid_w_HB <- ad_mid_w

ad_fin <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `Final AB score`) %>%
  filter(Region==region) %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
  mutate(across(contains("score"), as.numeric))

# Overall column and row
col <- ad_fin %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(Overall = median(`Final AB score`, na.rm = TRUE))

ad_fin_w <- ad_fin %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(`Final AB score` = median(`Final AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Final AB score`)) %>%
  ungroup() %>%
  bind_cols(col[2])

row <- ad_fin %>%
  group_by(Gender, Ethnicity, .drop = F) %>%
  dplyr::summarise(`Final AB score` = median(`Final AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Final AB score`)) %>%
  ungroup() %>%
  mutate(`Year level` = "Overall") %>%
  relocate(`Year level`) %>%
  mutate(Overall = median(DATA[DATA$Region == region, ]$`Final AB score`, na.rm = TRUE))

ad_fin_w <- rbind(ad_fin_w, row) %>% 
  mutate(
    #Assessment = "Final assessment", 
    Region = region) %>%
  relocate(Region)

colnames(ad_fin_w) <- c("Region", "Year level", "Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", 
                        "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Overall")
ad_fin_w_HB <- ad_fin_w

#########################################################
#3. Manawatu/Rangitikei/Whanganui
region <- "Manawatu/Rangitikei/Whanganui"

ad_pre <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `Pre AB score`) %>%
  filter(Region==region) %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
  mutate(across(contains("score"), as.numeric))

# Overall column and row
col <- ad_pre %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(Overall = median(`Pre AB score`, na.rm = TRUE))

ad_pre_w <- ad_pre %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(`Pre AB score` = median(`Pre AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Pre AB score`)) %>%
  ungroup() %>%
  bind_cols(col[2])

row <- ad_pre %>%
  group_by(Gender, Ethnicity, .drop = F) %>%
  dplyr::summarise(`Pre AB score` = median(`Pre AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Pre AB score`)) %>%
  ungroup() %>%
  mutate(`Year level` = "Overall") %>%
  relocate(`Year level`) %>%
  mutate(Overall = median(DATA[DATA$Region == region, ]$`Pre AB score`, na.rm = TRUE))

ad_pre_w <- rbind(ad_pre_w, row) %>% 
  mutate(
    #Assessment = "Pre assessment", 
    Region = region) %>%
  relocate(Region)

colnames(ad_pre_w) <- c("Region", "Year level", "Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", 
                        "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Overall")
ad_pre_w_man <- ad_pre_w

ad_mid <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `Mid AB score`) %>%
  filter(Region==region) %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
  mutate(across(contains("score"), as.numeric))

# Overall column and row
col <- ad_mid %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(Overall = median(`Mid AB score`, na.rm = TRUE))

ad_mid_w <- ad_mid %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(`Mid AB score` = median(`Mid AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Mid AB score`)) %>%
  ungroup() %>%
  bind_cols(col[2])

row <- ad_mid %>%
  group_by(Gender, Ethnicity, .drop = F) %>%
  dplyr::summarise(`Mid AB score` = median(`Mid AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Mid AB score`)) %>%
  ungroup() %>%
  mutate(`Year level` = "Overall") %>%
  relocate(`Year level`) %>%
  mutate(Overall = median(DATA[DATA$Region == region, ]$`Mid AB score`, na.rm = TRUE))

ad_mid_w <- rbind(ad_mid_w, row) %>% 
  mutate(
    #Assessment = "Mid assessment", 
    Region = region) %>%
  relocate(Region)

colnames(ad_mid_w) <- c("Region", "Year level", "Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", 
                        "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Overall")
ad_mid_w_man <- ad_mid_w

ad_fin <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `Final AB score`) %>%
  filter(Region==region) %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
  mutate(across(contains("score"), as.numeric))

# Overall column and row
col <- ad_fin %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(Overall = median(`Final AB score`, na.rm = TRUE))

ad_fin_w <- ad_fin %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(`Final AB score` = median(`Final AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Final AB score`)) %>%
  ungroup() %>%
  bind_cols(col[2])

row <- ad_fin %>%
  group_by(Gender, Ethnicity, .drop = F) %>%
  dplyr::summarise(`Final AB score` = median(`Final AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Final AB score`)) %>%
  ungroup() %>%
  mutate(`Year level` = "Overall") %>%
  relocate(`Year level`) %>%
  mutate(Overall = median(DATA[DATA$Region == region, ]$`Final AB score`, na.rm = TRUE))

ad_fin_w <- rbind(ad_fin_w, row) %>% 
  mutate(
    #Assessment = "Mid assessment", 
    Region = region) %>%
  relocate(Region)

colnames(ad_fin_w) <- c("Region", "Year level", "Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", 
                        "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Overall")
ad_fin_w_man <- ad_fin_w

#########################################################
#4. Nelson/Marlborough/West Coast
region <- "Nelson/Marlborough/West Coast"

ad_pre <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `Pre AB score`) %>%
  filter(Region==region) %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
  mutate(across(contains("score"), as.numeric))

# Overall column and row
col <- ad_pre %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(Overall = median(`Pre AB score`, na.rm = TRUE))

ad_pre_w <- ad_pre %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(`Pre AB score` = median(`Pre AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Pre AB score`)) %>%
  ungroup() %>%
  bind_cols(col[2])

row <- ad_pre %>%
  group_by(Gender, Ethnicity, .drop = F) %>%
  dplyr::summarise(`Pre AB score` = median(`Pre AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Pre AB score`)) %>%
  ungroup() %>%
  mutate(`Year level` = "Overall") %>%
  relocate(`Year level`) %>%
  mutate(Overall = median(DATA[DATA$Region == region, ]$`Pre AB score`, na.rm = TRUE))

ad_pre_w <- rbind(ad_pre_w, row) %>% 
  mutate(
    #Assessment = "Pre assessment", 
    Region = region) %>%
  relocate(Region)

colnames(ad_pre_w) <- c("Region", "Year level", "Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", 
                        "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Overall")
ad_pre_w_nel <- ad_pre_w

ad_mid <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `Mid AB score`) %>%
  filter(Region==region) %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
  mutate(across(contains("score"), as.numeric))

# Overall column and row
col <- ad_mid %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(Overall = median(`Mid AB score`, na.rm = TRUE))

ad_mid_w <- ad_mid %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(`Mid AB score` = median(`Mid AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Mid AB score`)) %>%
  ungroup() %>%
  bind_cols(col[2])

row <- ad_mid %>%
  group_by(Gender, Ethnicity, .drop = F) %>%
  dplyr::summarise(`Mid AB score` = median(`Mid AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Mid AB score`)) %>%
  ungroup() %>%
  mutate(`Year level` = "Overall") %>%
  relocate(`Year level`) %>%
  mutate(Overall = median(DATA[DATA$Region == region, ]$`Mid AB score`, na.rm = TRUE))

ad_mid_w <- rbind(ad_mid_w, row) %>% 
  mutate(
    #Assessment = "Mid assessment", 
    Region = region) %>%
  relocate(Region)

colnames(ad_mid_w) <- c("Region", "Year level", "Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", 
                        "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Overall")
ad_mid_w_nel <- ad_mid_w

ad_fin <- DATA %>% 
  dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `Final AB score`) %>%
  filter(Region==region) %>%
  dplyr::select(-Region) %>%
  dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
  mutate(Ethnicity = as.factor(case_when(grepl("Maori|Māori", Ethnicity) ~ "Māori",
                                         grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
                                         grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
                                         is.na(Ethnicity) ~ "Unknown",
                                         .default = "Other"))) %>%
  mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other", "Unknown"))) %>%
  mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
  mutate(`Year level` = factor(`Year level`, levels = c(11:1, "Unknown"))) %>%
  mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
  mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
  mutate(across(contains("score"), as.numeric))

# Overall column and row
col <- ad_fin %>%
  group_by(`Year level`, .drop = F) %>%
  dplyr::summarise(Overall = median(`Final AB score`, na.rm = TRUE))

ad_fin_w <- ad_fin %>%
  group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
  dplyr::summarise(`Final AB score` = median(`Final AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Final AB score`)) %>%
  ungroup() %>%
  bind_cols(col[2])

row <- ad_fin %>%
  group_by(Gender, Ethnicity, .drop = F) %>%
  dplyr::summarise(`Final AB score` = median(`Final AB score`, na.rm = TRUE)) %>%
  pivot_wider(names_from = c(`Ethnicity`, Gender), values_from = c(`Final AB score`)) %>%
  ungroup() %>%
  mutate(`Year level` = "Overall") %>%
  relocate(`Year level`) %>%
  mutate(Overall = median(DATA[DATA$Region == region, ]$`Final AB score`, na.rm = TRUE))

ad_fin_w <- rbind(ad_fin_w, row) %>% 
  mutate(
    #Assessment = "Final assessment", 
    Region = region) %>%
  relocate(Region)

colnames(ad_fin_w) <- c("Region", "Year level", "Māori F", "Māori M", "Māori U", "Pasifika F", "Pasifika M", "Pasifika U", 
                        "NZ European F", "NZ European M", "NZ European U", "Other F", "Other M", "Other U", "Unknown F", "Unknown M", "Unknown U", "Overall")
ad_fin_w_nel <- ad_fin_w

##############################################

# Add a worksheet to the workbook
addWorksheet(wb, sheetName = "Adaptive Bryant")

# Create header dataframes, one for pre, mid, and final assessments
header_pre <- data.frame("PRE ASSESSMENTS" = double(), "(Median score)" = double(), check.names = F)
header_mid <- data.frame("MID ASSESSMENTS" = double(), "(Median score)" = double(), check.names = F)
header_fin <- data.frame("FINAL ASSESSMENTS" = double(), "(Median score)" = double(), check.names = F)

# Write each dataframe to the worksheet with spacing
writeData(wb, sheet = "Adaptive Bryant", x = header_pre, startRow = 1, startCol = 1)
writeData(wb, sheet = "Adaptive Bryant", x = ad_pre_w_all, startRow = 3, startCol = 1)
writeData(wb, sheet = "Adaptive Bryant", x = header_mid, startRow = 1, startCol = ncol(ad_pre_w) + 3)
writeData(wb, sheet = "Adaptive Bryant", x = ad_mid_w_all, startRow = 3, startCol = ncol(ad_pre_w) + 3)
writeData(wb, sheet = "Adaptive Bryant", x = header_fin, startRow = 1, startCol = ncol(ad_pre_w) + ncol(ad_mid_w) + 6)
writeData(wb, sheet = "Adaptive Bryant", x = ad_fin_w_all, startRow = 3, startCol = ncol(ad_pre_w) + ncol(ad_mid_w) + 6)

writeData(wb, sheet = "Adaptive Bryant", x = ad_pre_w_can, startRow = nrow(ad_pre_w) + 5, startCol = 1)
writeData(wb, sheet = "Adaptive Bryant", x = ad_mid_w_can, startRow = nrow(ad_mid_w) + 5, startCol = ncol(ad_pre_w) + 3)
writeData(wb, sheet = "Adaptive Bryant", x = ad_fin_w_can, startRow = nrow(ad_fin_w) + 5, startCol = ncol(ad_pre_w) + ncol(ad_mid_w) + 6)

writeData(wb, sheet = "Adaptive Bryant", x = ad_pre_w_HB, startRow = nrow(ad_pre_w) + nrow(ad_pre_w_can) + 7, startCol = 1)
writeData(wb, sheet = "Adaptive Bryant", x = ad_mid_w_HB, startRow = nrow(ad_mid_w) + nrow(ad_pre_w_can) + 7, startCol = ncol(ad_pre_w) + 3)
writeData(wb, sheet = "Adaptive Bryant", x = ad_fin_w_HB, startRow = nrow(ad_fin_w) + nrow(ad_pre_w_can) + 7, startCol = ncol(ad_pre_w) + ncol(ad_mid_w) + 6)

writeData(wb, sheet = "Adaptive Bryant", x = ad_pre_w_man, startRow = nrow(ad_pre_w) + nrow(ad_pre_w_can) + nrow(ad_pre_w_HB) + 9, startCol = 1)
writeData(wb, sheet = "Adaptive Bryant", x = ad_mid_w_man, startRow = nrow(ad_mid_w) + nrow(ad_pre_w_can) + nrow(ad_pre_w_HB) + 9, startCol = ncol(ad_pre_w) + 3)
writeData(wb, sheet = "Adaptive Bryant", x = ad_fin_w_man, startRow = nrow(ad_fin_w) + nrow(ad_pre_w_can) + nrow(ad_pre_w_HB) + 9, startCol = ncol(ad_pre_w) + ncol(ad_mid_w) + 6)

writeData(wb, sheet = "Adaptive Bryant", x = ad_pre_w_nel, startRow = nrow(ad_pre_w) + nrow(ad_pre_w_can) + nrow(ad_pre_w_HB) + nrow(ad_pre_w_man) + 11, startCol = 1)
writeData(wb, sheet = "Adaptive Bryant", x = ad_mid_w_nel, startRow = nrow(ad_mid_w) + nrow(ad_pre_w_can) + nrow(ad_pre_w_HB) + nrow(ad_pre_w_man) + 11, startCol = ncol(ad_pre_w) + 3)
writeData(wb, sheet = "Adaptive Bryant", x = ad_fin_w_nel, startRow = nrow(ad_fin_w) + nrow(ad_pre_w_can) + nrow(ad_pre_w_HB) + nrow(ad_pre_w_man) + 11, startCol = ncol(ad_pre_w) + ncol(ad_mid_w) + 6)

##################################################################################################################
##################################################################################################################


# CREATE WORKBOOK

# Add DATA and full sheets
addWorksheet(wb, sheetName = "DATA")
writeData(wb, sheet = "DATA", x = DATA, startRow = 1, startCol = 1)

addWorksheet(wb, sheetName = "Raw data")
writeData(wb, sheet = "Raw data", x = combined_data, startRow = 1, startCol = 1)

# Add a notes sheet
notes <- data.frame("NOTES" = c("Gender is expressed as F (female), M (male) and U (unknown - gender was not entered).",
                                "Year level contains all years entered by the teachers (i.e., some may be out of scope). Students with no year are set as 'unknown'.",
                                "A value of 'NaN' appears in some percentages because there were no students in that category (0 divided by 0 produces an error).",
                                "For the Ethnicity sheet, percentages are calculated row-wise, except for the total column. Only the first ethnicity value was used.",
                                "Where there were two students with the same first name and teacher, but one student was missing the surname, the student missing the surname was assumed to be the same as the student with the surname."))

addWorksheet(wb, sheetName = "Notes")
writeData(wb, sheet = "Notes", x = notes, startRow = 1, startCol = 1)

# Define style and apply to all sheets
style <- createStyle(fontName = "Calibri", fontSize = 10)

sheets <- names(wb)

getRows <- function(wb, sheet) {
  sheet_data <- read.xlsx(wb, sheet = sheet)
  return(max(which(rowSums(!is.na(sheet_data)) > 0)))
}

getCols <- function(wb, sheet) {
  sheet_data <- read.xlsx(wb, sheet = sheet)
  return(ncol(sheet_data))
}

for (sheet in sheets) {
  # Get the number of rows and columns used in the sheet
  usedRows <- getRows(wb, sheet) + 10
  usedCols <- getCols(wb, sheet)
  
  addStyle(wb, sheet = sheet, style = style, rows = 1:usedRows, cols = 1:usedCols, gridExpand = TRUE)
  setColWidths(wb, sheet = sheet, cols = 1:usedCols, widths = "auto")
  
}

# Save the workbook to a file
saveWorkbook(wb, file = "Learning_Matters_summary_table.xlsx", overwrite = TRUE)

##################################################################################################################
##################################################################################################################

