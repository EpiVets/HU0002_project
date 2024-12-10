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

# Function to anonymise student names
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


#################################################################################################
#################################################################################################

# This works in two stages: 1) extract data from the LM Google Drive account and save as .xlsx files locally, 2) process the .xlsx files to make the report.


# COLLATE XLSX FILES INTO A SINGLE DATAFRAME CONTAINING RELEVANT DATA

#1. Read in files in the folder "downloaded_xlsx_files2" that contain "Relevant" in the name and extract the sheets called "Tier 2 Student Details" and "File Path".
# Then add a column called "Region" to the sheet called "Tier 2 Student Details" that contains the text betwen the first "/" and the characters " ALL" in the "File Path" sheet.

# Define the path to the folder containing the .xlsx files
folder_path <- "downloaded_xlsx_files"

# List all .xlsx files in the folder that contain "Relevant" in the name, BUT REMOVE FILES WITH "COPY" IN THE NAME
files2 <- list.files(path = folder_path, pattern = "Relevant.*\\.xlsx$", full.names = TRUE) %>%
  setdiff(list.files(path = folder_path, pattern = "Copy", full.names = TRUE))

# List of files not in files2
excluded_files_copies <- list.files(path = folder_path, pattern = "Relevant.*\\.xlsx$", full.names = TRUE) %>%
  setdiff(files2)

# Extract file paths from the "File Path" sheet of files in excluded_files_copies and save as a data frame
# Extract file paths from the "File path" sheet and store in a dataframe
excluded_file_paths_df <- lapply(excluded_files_copies, function(file) {
  # Read the "File path" sheet and extract the value in cell A1
  file_path <- read_excel(file, sheet = "File Path", range = "A1", col_names = FALSE) %>%
    pull(1) # Extract the first column value
  
  # Return a data frame with the file and the extracted path
  data.frame(`Excluded file` = file_path, stringsAsFactors = FALSE)
}) %>%
  bind_rows() %>%
  dplyr::rename(`Excluded file path` = `Excluded.file`)


# Initialize an empty list to store data frames
data_list <- list()

# Create a vector to hold name of spreadsheet being worked on in case of a bug
spreadsheet_name <- c()

# Abbreviated files2 for testing:
# file <- files2[1]
  

###################################################################################################################################################
# Loop through each file
for (file in files2) {
  
  message(noquote(paste0("Spreadsheet = ", "'", file, "'")))
  
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
  print(noquote(paste0("Reading sheet 'Tier 2 Student Details'")))
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
  
  # Check for columns that have been duplicated and remove said columns
  # tier2_details <- tier2_details %>%
  #   select(-contains("x"))
  
  # Remove duplicate columns if table contains data, otherwise leave it as is
  if (nrow(tier2_details) > 0) {
    tier2_details <- tier2_details[, !duplicated(names(tier2_details))]
  }
  
  
  ##############################################################################################
  # Extract the "Region" from the "File Path" sheet.
  print(noquote(paste0("Exracting region from 'File Path' sheet")))
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
  print(noquote(paste0("Extracting 'neurodiverse_diagnosis' and 'neurodiverse_suspicion'")))
  neurodiverse_diagnosis <- tier2_details %>%
    dplyr::select(student_details_Student_Name, starts_with("neuro_diverse_diagnosis"), -neuro_diverse_diagnosis_Ethnicity) %>%
    # Rename columns with only the characters after the last "_"
    rename_with(~str_extract(., "(?<=_)[^_]+$"), .cols = everything()) %>%
    pivot_longer(cols = !Name, names_to = "Diagnosis", values_to = "value") %>%
    filter(!is.na(value) & !tolower(value) %in% c("no", "none", "n/a", "n", "x")) %>%  # Exclude "No", "no", "None", "none", "N/A", "N", "x"
    group_by(Name) %>%
    mutate(neurodiverse_diagnosis = paste(Diagnosis, collapse = ", ")) %>%
    dplyr::select(Name, neurodiverse_diagnosis) %>%
    slice_head(n=1) %>%
    ungroup()
  
  neurodiverse_suspicion <- tier2_details %>%
    dplyr::select(student_details_Student_Name, starts_with("neuro_diverse_suspicion")) %>%
    rename_with(~str_extract(., "(?<=_)[^_]+$"), .cols = everything()) %>%
    pivot_longer(cols = -Name, names_to = "Suspicion", values_to = "value") %>%
    filter(!is.na(value) & !tolower(value) %in% c("no", "none", "n/a", "n", "x")) %>%  # Exclude "No", "no", "None", "none", "N/A", "N", "x"
    group_by(Name) %>%
    mutate(neurodiverse_suspicion = paste(Suspicion, collapse = ", ")) %>%
    dplyr::select(Name, neurodiverse_suspicion) %>%
    slice_head(n=1) %>%
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
  
  # Extract Adaptive Bryant score from the "Common Assessment (AB)" sheet
  print(noquote(paste0("Extracting adaptive Bryant score from 'Common Assessment (AB)' sheet")))
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
  print(noquote(paste0("Extracting student attendance data from 'Student Session Attendance' sheet")))
  ssa <- read_excel(file, sheet = "Student Session Attendance", skip = 1) %>%
    clean_names() %>%
    dplyr::select(-no_of_sessions_wk)
  
  # Check if there are any attendance data
  attendance_data <- ssa[-1,] %>%
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
      } else if (str_detect(sessions, "(?i)^o\\s*-\\s*overseas$")) {
        # If the string is "O - overseas" (case insensitive), return 0
        return(0)
      } else if (str_detect(sessions, "(?i)away")) {
        # If the string contains "Away" or "away", return 0
        return(0)
      } else if (str_detect(sessions, "^✔+$")) {
        # If it contains only check marks, count them
        return(str_count(sessions, "✔"))
      } else if (str_detect(sessions, "\\d+")) {
        # If it contains a number, extract and convert to numeric
        return(as.numeric(str_extract(sessions, "\\d+")))
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
        #mutate(n_sessions = as.numeric(n_sessions)) %>%
        select(Name, sessions_taught_per_week, n_sessions) %>%
        group_by(Name) %>%
        dplyr::summarise(`Total available` = sum(sessions_taught_per_week, na.rm = TRUE), 
                         `Total attended`  = sum(n_sessions, na.rm = TRUE),
                         N_weeks           = sum(!is.na(n_sessions) & n_sessions > 0, na.rm = TRUE))
     
      colnames(ssa) <- c("Name", "Total available", "Total attended", "N weeks")
      
      # Anonymise names that weren't anonymised in the previous step
      ssa <- ssa %>%
        # If Name contains any letters, run anonymise_name function
        rowwise() %>%
        mutate(Name2 = ifelse(any(grepl("[a-zA-Z]", Name)), anonymise_name(Name), Name)) %>%
        ungroup()
      
    }
  
  
  ##############################################################################################
  
  # Join "final_adapted_bryant_assessment_Overall_Score_/50" to tier2_details using "student_details_Student_Name", keeping all
  print(noquote(paste0("Joining 'final_adapted_bryant_assessment_Overall_Score_/50' to 'Tier 2 Student Details'")))
  data <- full_join(tier2_details, ab, by = "student_details_Student_Name") %>%
    mutate(National_Student_number = ifelse(is.na(student_details_National_Student_number_tier2), 
                                            student_details_National_Student_number_ad_bry,
                                            student_details_National_Student_number_tier2)) %>%
    mutate(Region = ifelse(is.na(Region_tier2), Region_ad_bry, Region_tier2)) %>%
    mutate(Teacher = ifelse(is.na(Teacher_tier2), Teacher_ad_bry, Teacher_tier2)) %>%
    
    # Split student name into first name and last name using the first space as the delimiter
    # This is for joining e-asTTLE data, which is split into separate columns for firstname and surname.
    separate(student_details_Student_Name, into = c("Firstname", "Surname"), sep = " ", extra = "merge", remove = FALSE) %>%
    
    # Handle missing Surnames by replacing NA with "Surname not provided"
    mutate(Surname = replace_na(Surname, "Surname not provided")) %>%
    
    # Add a new School column with NA
    mutate(School = NA) %>%
    
    # Relocate columns for a specific order
    dplyr::relocate(Firstname, Surname, Region, School, Teacher, student_details_Year_level, 
                    student_details_Gender, student_details_Ethnicity_1, student_details_Ethnicity_2)
  
  

  # Join the student session Adaptive Bryant data
  data <- left_join(data, ssa, by = c("student_details_Student_Name" = "Name")) %>%
    # Add column capturing the source of the data
    mutate(Source = spreadsheet_name)

  # Append the modified data frame to the list
  data_list <- append(data_list, list(data))
}


###########################################################################################################################################
###########################################################################################################################################

# Combine all data frames together by row
length(data_list)
combined_data <- bind_rows(data_list)
length(unique(combined_data$Teacher)) # was 86 teachers, now 80
# combined_data$student_details_Student_Name

# Explore probably bogus entries with no name
combined_data %>%
  filter(is.na(student_details_Student_Name) | student_details_Student_Name == "Student Name") %>%
  View()

# Remove the row with NA for student_details_Student_Name
combined_data <- combined_data %>%
  filter(!is.na(student_details_Student_Name))

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

# View(combined_data %>% filter(same_child == 1) %>% arrange(Firstname, Surname))
# View(combined_data %>% filter(Firstname=="1051919931" & Surname=="1714523"))

combined_data <- combined_data %>%
  group_by(Firstname, Surname, Teacher) %>%
  fill(everything(), .direction = "downup") %>%
  mutate(student_details_Student_Name = paste(Firstname, Surname, sep = " ")) %>%
  distinct() %>%
  ungroup()


# Deal with Fiona Wilder
# anonymise_name("Fiona"); anonymise_name("Wilder")
# View(combined_data[combined_data$Firstname == "6915141" & combined_data$Surname=="239124518" & !is.na(combined_data$Firstname), ])
# Was 51, now 45 Fiona Wilders

# Tidy up rogue variables and remove Fiona Wilder
combined_data <- combined_data %>%
  mutate(student_details_Gender = ifelse(grepl("Fem|fem", student_details_Gender), "F", student_details_Gender)) %>%
  # Remove characters including and after the "/" in columns containing "Overall_Score_/50"
  mutate_at(vars(contains("Overall_Score_/50")), ~str_remove(., "/.*")) %>%
  # Convert same columns to numeric
  mutate_at(vars(contains("Overall_Score_/50")), as.numeric) %>%
  filter(!(coalesce(Firstname, "") == "6915141" & coalesce(Surname, "") == "239124518")) # Use coalesce to handle NA values!


table(combined_data$student_details_Year_level)


###########################################################################################################################################
###########################################################################################################################################

# IMPORT E-ASTTLE DATA AND MERGE WITH COMBINED DATA
# NB student_name is the correct name
eas_beg <- read_excel("e-asTTLe Data Term 4 Beg_End.xlsx", sheet = 1) %>%
  clean_names() %>%
  dplyr::rename(east = "overall_scale_score") %>%
  #dplyr::select(Firstname, Surname, east) %>%
  # Identify duplicate rows
  group_by(student_name) %>%
  mutate(duplicate = n() > 1) %>%
  ungroup() %>%
  # Remove whitespace from east and convert to numeric (values of "absent" become NA)
  mutate(east = as.numeric(str_remove_all(east, "\\s+"))) %>%
  # Anonymise names
  mutate(student_name = sapply(student_name, anonymise_name)) %>%
  # Identify rows that appear in combined_data using student_name
  mutate(in_google_drive_data = ifelse(student_name %in% combined_data$student_details_Student_Name, "Yes", "No")) %>%
  select(-firstname, -surname)

eas_end <- read_excel("e-asTTLe Data Term 4 Beg_End.xlsx", sheet = 2) %>%
  clean_names() %>%
  dplyr::rename(east = "overall_scale_score") %>%
  #dplyr::select(Firstname, Surname, east) %>%
  # Identify duplicate rows
  group_by(student_name) %>%
  mutate(duplicate = n() > 1) %>%
  ungroup() %>%
  # Remove whitespace from east and convert to numeric (values of "absent" become NA)
  mutate(east = as.numeric(str_remove_all(east, "\\s+"))) %>%
  # Anonymise names
  mutate(student_name = sapply(student_name, anonymise_name)) %>%
  # Identify rows that appear in combined_data using student_name
  mutate(in_google_drive_data = ifelse(student_name %in% combined_data$student_details_Student_Name, "Yes", "No")) %>%
  select(-firstname, -surname)

# Combine the two e-asTTLe dataframes and correct errors in year_level per email from Greg McNeil on 09/12/24
eas <-  full_join(eas_beg, eas_end, by = "student_name", suffix = c("_beg", "_end")) %>%
  # Remove white space from columns beginning with year_level
  mutate_at(vars(starts_with("year_level")), ~str_remove_all(., "\\s+")) %>%
  mutate(yr_agree = year_level_beg == year_level_end) %>%
  mutate(year_level = year_level_beg) %>%
  mutate(year_level = case_when(student_name == "919152512 81252315154" ~ 8,
                                student_name == "151292291 12523919" ~ 8,
                                student_name == "195114141 1313125114" ~ 7,
                                student_name == "1317795 1336118121144" ~ 6,
                                student_name == "381181295 13314591212" ~ 7,
                                student_name == "42512114 1651115" ~ 7,
                                student_name == "85265119512 2021162115121" ~ 7,
                                .default = as.numeric(year_level)))

# Write .xlsx of data in eas where yr_agree == FALSE
# write_xlsx(eas %>% filter(!yr_agree), "eas_disagreements.xlsx")

# Clean up eas by removing unnecessary columns
eas <- eas %>%
  select(student_name, year_level, east_beg, east_end) 





# table(eas$in_google_drive_data)
# Quantify number without a value for east
# eas %>%
# select(east) %>%
#   filter(is.na(east)) %>%
#   nrow() 

# 1+96+16-2+11-8


# Export .xlsx for LM to track down missing students
# write_xlsx(eas, "eas.xlsx")

# Remove the duplicated row for 13251 4122919-201212381185 by taking only the first iteration (not applicable after LM tidied this up)
# eas <- eas %>% 
#   group_by(Firstname, Surname) %>%
#   slice_head(n=1) %>%
#   ungroup() %>%
#   select(-duplicate)


# Merge the e-asTTLe data with the combined data
combined_data <- left_join(combined_data, eas %>% select(student_name, year_level, east_beg, east_end), by = c("student_details_Student_Name" = "student_name"))

# Check student_details_Year_Level and year_level match
combined_data <- combined_data %>%
  mutate(year_level_match = student_details_Year_level == year_level) 

View(combined_data %>% filter(year_level_match==F))

# Identify rows of eas that did not merge with combined_data
eas_not_merged <- eas %>%
  anti_join(combined_data, by = c("student_name" = "student_details_Student_Name"))
# None!

# Export .xlsx of students with no gender or year level
# write_xlsx(combined_data %>% filter(is.na(student_details_Gender) | is.na(student_details_Year_level)), "missing_gender_year.xlsx")


###########################################################################################################################################
###########################################################################################################################################

# Create the final dataframe needed by LM
DATA <- combined_data %>%
  dplyr::select(Firstname, Surname, student_details_Student_Name, Region, School, Teacher, student_details_Year_level, student_details_Gender, student_details_Ethnicity_1, student_details_Ethnicity_2,
                neurodiverse_diagnosis, neurodiverse_suspicion, `Total available`, `Total attended`, `N weeks`,
                `pre_adapted_bryant_assessment_Overall_Score_/50`, `mid_adapted_bryant_assessment_Overall_Score_/50`, 
                `final_adapted_bryant_assessment_Overall_Score_/50`, east_beg, east_end) %>%
  dplyr::rename(`Year level` = student_details_Year_level, `Ethnicity 1` = student_details_Ethnicity_1, `Ethnicity 2` = student_details_Ethnicity_2, 
                Gender = student_details_Gender,
                `Neurodiversity diagnosis` = neurodiverse_diagnosis, `Neurodiversity suspicion` = neurodiverse_suspicion,
                `Pre AB score` = `pre_adapted_bryant_assessment_Overall_Score_/50`, `Mid AB score` = `mid_adapted_bryant_assessment_Overall_Score_/50`, 
                `Final AB score` = `final_adapted_bryant_assessment_Overall_Score_/50`,
                `Beginning e-asTTle overall scaled score` = east_beg, `End e-asTTle overall scaled score` = east_end) %>%
  # If Gender == "M (F)", replace with "M" 
  mutate(Gender = ifelse(Gender=="M (F)", "M", Gender))

# sum(!is.na(DATA$`e-asTTle overall scaled score`))
# length(unique(paste0(DATA$Firstname, DATA$Surname)))

table(DATA$`Year level`, useNA = "always")
table(DATA$Gender, useNA = "always")

###########################################################################################################################################
###########################################################################################################################################

# ETHNICITY SHEET

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
  # If Gender == "M (F)", replace with "M" 
  #mutate(Gender = ifelse(Gender=="M (F)", "M", Gender)) %>%
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


################################################################################################################################################################
################################################################################################################################################################

# LEARNING DIFFERENCES SHEET

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


################################################################################################################################################################
################################################################################################################################################################

# Attendance sheet

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

################################################################################################################################################################
################################################################################################################################################################

# ADPATIVE BRYANT SHEET
# Need beginning and end (pre and final) AB results only.

# Create function to save repetitive code

create_region_assessment_df <- function(region) {
  
  # Filter and preprocess data
  ad_data <- DATA %>%
    filter(!is.na(`Pre AB score`) & !is.na(`Final AB score`)) %>%
    dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `Pre AB score`, `Final AB score`, `N weeks`) %>%
    { if (region != "National") filter(., Region == region) else . } %>% # Region filter if not National
    dplyr::select(-Region) %>%
    dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
    mutate(Ethnicity = as.factor(case_when(
      grepl("Maori|Māori", Ethnicity) ~ "Māori",
      grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
      grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
      #is.na(Ethnicity) ~ "Unknown",
      .default = "Other"  # I.e., merge unknown and other together as "other"
    ))) %>%
    mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other"))) %>%
    mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
    mutate(`Year level` = factor(`Year level`, levels = c(10:1, "Unknown"))) %>%
    mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
    mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown"))) %>%
    mutate(across(contains("score"), as.numeric))
  
  # Overall column
  col <- ad_data %>%
    group_by(`Year level`, .drop = F) %>%
    dplyr::summarise(Overall_N = n(), 
                     Overall_Beginning = round(mean(`Pre AB score`), 1), 
                     Overall_End = round(mean(`Final AB score`), 1),
                     N_weeks = round(mean(`N weeks`, na.rm = T), 1)) %>%
    # Set Year level as row names
    column_to_rownames(var = "Year level") %>%
    ungroup()
  
  # Pivot the data wider by Ethnicity and Gender
  ad_wide <- ad_data %>%
    group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
    dplyr::summarise(N = n(), Beginning = round(mean(`Pre AB score`), 1), End = round(mean(`Final AB score`), 1)) %>%
    pivot_longer(cols = c(N, Beginning, End), 
                 names_to = "metric", 
                 values_to = "value") %>%
    unite("ethnicity_gender", Ethnicity, Gender, sep = "_") %>%
    unite("column_name", ethnicity_gender, metric, sep = "_") %>%
    pivot_wider(names_from = column_name, values_from = value) %>%
    ungroup()
    
  # Overall row
  row <- ad_data %>%
    group_by(Gender, Ethnicity, .drop = F) %>%
    dplyr::summarise(N = n(), 
                     Beginning = round(mean(`Pre AB score`), 1), 
                     End = round(mean(`Final AB score`), 1)) %>%
    pivot_longer(cols = c(N, Beginning, End), 
                 names_to = "metric", 
                 values_to = "value") %>%
    unite("ethnicity_gender", Ethnicity, Gender, sep = "_") %>%
    unite("column_name", ethnicity_gender, metric, sep = "_") %>%
    pivot_wider(names_from = column_name, values_from = value) %>%
    mutate(`Year level` = "Overall") %>%
    relocate(`Year level`) %>%
    #select(all_of(desired_columns)) %>%
    mutate(Overall_N = nrow(ad_data), 
           Overall_Beginning = round(mean(ad_data$`Pre AB score`), 1), 
           Overall_End       = round(mean(ad_data$`Final AB score`), 1),
           N_weeks           = round(mean(ad_data$`N weeks`, na.rm = T), 1)) %>%
    ungroup()
  
  
  # Combine the wide data and overall column and row, and add region information
  ad_wide <- cbind(ad_wide, col) %>%
    rbind(row) %>%
    mutate(Region = region) %>%
    relocate(Region) %>%
    mutate(`Year level` = as.character(`Year level`)) %>%
    # Replace "NaN" with NA
    mutate(across(everything(), ~ifelse(is.numeric(.) & . == "NaN", NA, .)))
  
  # Rename columns
  # colnames(ad_wide) <- c("Region", "Year level", "N Māori F", "Median Māori F", "N Māori M", "Median Māori M", "N Māori U", "Median Māori U", 
  #                        "N Pasifika F", "Median Pasifika F", "N Pasifika M", "Median Pasifika M", "N Pasifika U", "Median Pasifika U", 
  #                        "N NZ European F", "Median NZ European F", "N NZ European M", "Median NZ European M", "N NZ European U", "Median NZ European U", 
  #                        "N Other F", "Median Other F", "N Other M", "Median Other M", "N Other U", "Median Other U",
  #                        "N Unknown F", "Median Unknown F", "N Unknown M", "Median Unknown M", "N Unknown U", "Median Unknown U",
  #                        "N overall", "Median overall")
  

  return(ad_wide)
}

# Dataframes
ad_national <- create_region_assessment_df("National")
ad_can      <- create_region_assessment_df("Canterbury")
ad_HB       <- create_region_assessment_df("Hawkes Bay")
ad_man      <- create_region_assessment_df("Manawatu/Rangitikei/Whanganui")
ad_nel      <- create_region_assessment_df("Nelson/Marlborough/West Coast")



##################################################################################################################
##################################################################################################################

# ADD E-ASTTLE DATA SHEET

# Function to save repetitive code
easttle_df <- function(region) {
  
  easttle <- DATA %>% 
    filter(`Beginning e-asTTle overall scaled score` != 0 & `End e-asTTle overall scaled score` != 0) %>%
    dplyr::select(Region, Gender, `Year level`, `Ethnicity 1`, `Beginning e-asTTle overall scaled score`, `End e-asTTle overall scaled score`, `N weeks`) %>%
    { if (region != "National") filter(., Region == region) else . } %>% # Region filter if not National
    dplyr::select(-Region) %>%
    dplyr::rename(Ethnicity = `Ethnicity 1`) %>%
    mutate(Ethnicity = as.factor(case_when(
      grepl("Maori|Māori", Ethnicity) ~ "Māori",
      grepl("Pakeha|Pākehā|NZ European", Ethnicity) ~ "NZ European",
      grepl("Pacific|Pascifica|Pasifika", Ethnicity) ~ "Pasifika",
      #is.na(Ethnicity) ~ "Unknown",  # I.e., merge unknown and other together as "other"
      .default = "Other"
    ))) %>%
    mutate(Ethnicity = factor(Ethnicity, levels = c("Māori", "Pasifika", "NZ European", "Other"))) %>%
    mutate(`Year level` = ifelse(is.na(`Year level`), "Unknown", `Year level`)) %>%
    mutate(`Year level` = factor(`Year level`, levels = c(10:min(eas$year_level), "Unknown"))) %>%
    mutate(Gender = ifelse(is.na(Gender), "Unknown", Gender)) %>%
    mutate(Gender = factor(Gender, levels = c("F", "M", "Unknown")))
  
  # Overall column
  col <- easttle %>%
    group_by(`Year level`, .drop = F) %>%
    dplyr::summarise(Overall_N = n(), 
                     Overall_Beginning = round(mean(`Beginning e-asTTle overall scaled score`), 1), 
                     Overall_End = round(mean(`End e-asTTle overall scaled score`), 1),
                     N_weeks = round(mean(`N weeks`, na.rm = T), 1)) %>%
    # Set Year level as row names
    column_to_rownames(var = "Year level") %>%
    ungroup()
  
  # Pivot the data wider by Ethnicity and Gender
  easttle_wide <- easttle %>%
    group_by(`Year level`, Ethnicity, Gender, .drop = F) %>%
    dplyr::summarise(N = n(), Beginning = round(mean(`Beginning e-asTTle overall scaled score`), 1), End = round(mean(`End e-asTTle overall scaled score`), 1)) %>%
    pivot_longer(cols = c(N, Beginning, End), 
                 names_to = "metric", 
                 values_to = "value") %>%
    unite("ethnicity_gender", Ethnicity, Gender, sep = "_") %>%
    unite("column_name", ethnicity_gender, metric, sep = "_") %>%
    pivot_wider(names_from = column_name, values_from = value) %>%
    ungroup()
  
  # Overall row
  row <- easttle %>%
    group_by(Gender, Ethnicity, .drop = F) %>%
    dplyr::summarise(N = n(), 
                     Beginning = round(mean(`Beginning e-asTTle overall scaled score`), 1), 
                     End = round(mean(`End e-asTTle overall scaled score`), 1)) %>%
    pivot_longer(cols = c(N, Beginning, End), 
                 names_to = "metric", 
                 values_to = "value") %>%
    unite("ethnicity_gender", Ethnicity, Gender, sep = "_") %>%
    unite("column_name", ethnicity_gender, metric, sep = "_") %>%
    pivot_wider(names_from = column_name, values_from = value) %>%
    mutate(`Year level` = "Overall") %>%
    relocate(`Year level`) %>%
    #select(all_of(desired_columns)) %>%
    mutate(Overall_N = nrow(easttle), 
           Overall_Beginning = round(mean(easttle$`Beginning e-asTTle overall scaled score`), 1), 
           Overall_End       = round(mean(easttle$`End e-asTTle overall scaled score`), 1),
           N_weeks           = round(mean(easttle$`N weeks`, na.rm = T), 1)) %>%
    ungroup()
  
  
  # Combine the wide data and overall column and row, and add region information
  easttle_wide <- cbind(easttle_wide, col) %>%
    rbind(row) %>%
    mutate(Region = region) %>%
    relocate(Region) %>%
    mutate(`Year level` = as.character(`Year level`)) %>%
    # Replace "NaN" with NA
    mutate(across(everything(), ~ifelse(is.numeric(.) & . == "NaN", NA, .)))
  
  return(easttle_wide)
}


# Make the dataframes
easttle_all <- easttle_df("National")
easttle_can <- easttle_df("Canterbury")
easttle_HB  <- easttle_df("Hawkes Bay")
easttle_man <- easttle_df("Manawatu/Rangitikei/Whanganui")
easttle_nel <- easttle_df("Nelson/Marlborough/West Coast")


##################################################################################################################
##################################################################################################################

# IDENTIFY TEACHERS WHOSE NEURODIVERSITY FIELDS ARE ALL EMPTY

# Use combined_data, group by Teacher, and make a list of teachers who have all NA for columns starting with "Neurodiversity"
empty_neurodiversity <- DATA %>%
  group_by(Teacher) %>%
  filter(all(is.na(across(starts_with("Neurodiversity"))))) %>%
  arrange(Teacher) %>%
  pull(Teacher) %>%
  unique() 

# Print as a vector I can paste into Excel
cat(paste(empty_neurodiversity, collapse = ", "))

##################################################################################################################
##################################################################################################################


# CREATE WORKBOOK

wb <- createWorkbook()

##################################################################################################################
# Add ethnicity worksheet to the workbook
addWorksheet(wb, sheetName = "Ethnicity")

# Write each dataframe to the worksheet with spacing
writeData(wb, sheet = "Ethnicity", x = ethn_all, startRow = 1, startCol = 1)
writeData(wb, sheet = "Ethnicity", x = ethn_cant, startRow = nrow(df_combined) + 4, startCol = 1)
writeData(wb, sheet = "Ethnicity", x = ethn_HB, startRow = nrow(df_combined) + nrow(ethn_cant) + 6, startCol = 1)
writeData(wb, sheet = "Ethnicity", x = ethn_man, startRow = nrow(df_combined) + nrow(ethn_cant) + nrow(ethn_HB) + 9, startCol = 1)
writeData(wb, sheet = "Ethnicity", x = ethn_nel, startRow = nrow(df_combined) + nrow(ethn_cant) + nrow(ethn_HB) + nrow(ethn_man) + 12, startCol = 1)


##################################################################################################################
# Add learning differences worksheet to the workbook
addWorksheet(wb, sheetName = "Learning differences")

# Write each dataframe to the worksheet with spacing
writeData(wb, sheet = "Learning differences", x = ld_all, startRow = 1, startCol = 1)
writeData(wb, sheet = "Learning differences", x = ld_can, startRow = nrow(ld_all) + 4, startCol = 1)
writeData(wb, sheet = "Learning differences", x = ld_HB, startRow = nrow(ld_all) + nrow(ld_can) + 6, startCol = 1)
writeData(wb, sheet = "Learning differences", x = ld_man, startRow = nrow(ld_all) + nrow(ld_can) + nrow(ld_HB) + 9, startCol = 1)
writeData(wb, sheet = "Learning differences", x = ld_nel, startRow = nrow(ld_all) + nrow(ld_can) + nrow(ld_HB) + nrow(ld_man) + 12, startCol = 1)


##################################################################################################################
# Add adaptive Bryant worksheet to the workbook
addWorksheet(wb, sheetName = "Adaptive Bryant")

# Function to create table header rows
add_table_headers <- function(wb, sheet_name, start_row = 1) {
  # Prepare header rows
  ethnicity_row <- c(rep("", 2), 
                     rep("Māori", 9), 
                     rep("Pasifika", 9), 
                     rep("NZ European", 9), 
                     rep("Other", 9), 
                     rep("Total", 4))
  
  gender_row <- c(rep("", 2),
                  rep(c("F", "M", "Unknown"), each = 3),  # Māori
                  rep(c("F", "M", "Unknown"), each = 3),  # Pasifika
                  rep(c("F", "M", "Unknown"), each = 3),  # NZ European
                  rep(c("F", "M", "Unknown"), each = 3),  # Other
                  rep("", 4))  # Total (no gender breakdown)
  
  metric_row <- c("Region", "Year level", 
                  rep(c("N", "Beginning", "End"), times = 4 * 3),  # 4 ethnicities, 3 genders
                  c("N", "Beginning", "End", "N weeks"))  # Total
  
  # Add the header rows to the sheet
  writeData(wb, sheet = sheet_name, x = rbind(ethnicity_row, gender_row, metric_row), 
            startRow = start_row, colNames = FALSE)
  
  # Merge cells for Ethnicity
  mergeCells(wb, sheet = sheet_name, cols = 3:11, rows = start_row)   # Māori
  mergeCells(wb, sheet = sheet_name, cols = 12:20, rows = start_row)  # Pasifika
  mergeCells(wb, sheet = sheet_name, cols = 21:29, rows = start_row)  # NZ European
  mergeCells(wb, sheet = sheet_name, cols = 30:38, rows = start_row)  # Other
  mergeCells(wb, sheet = sheet_name, cols = 39:42, rows = start_row)  # Total
  
  # Merge cells for Gender (within each Ethnicity group)
  mergeCells(wb, sheet = sheet_name, cols = 3:5, rows = start_row + 1)    # Māori F
  mergeCells(wb, sheet = sheet_name, cols = 6:8, rows = start_row + 1)    # Māori M
  mergeCells(wb, sheet = sheet_name, cols = 9:11, rows = start_row + 1)   # Māori Unknown
  
  mergeCells(wb, sheet = sheet_name, cols = 12:14, rows = start_row + 1)  # Pasifika F
  mergeCells(wb, sheet = sheet_name, cols = 15:17, rows = start_row + 1)  # Pasifika M
  mergeCells(wb, sheet = sheet_name, cols = 18:20, rows = start_row + 1)  # Pasifika Unknown
  
  mergeCells(wb, sheet = sheet_name, cols = 21:23, rows = start_row + 1)  # NZ European F
  mergeCells(wb, sheet = sheet_name, cols = 24:26, rows = start_row + 1)  # NZ European M
  mergeCells(wb, sheet = sheet_name, cols = 27:29, rows = start_row + 1)  # NZ European Unknown
  
  mergeCells(wb, sheet = sheet_name, cols = 30:32, rows = start_row + 1)  # Other F
  mergeCells(wb, sheet = sheet_name, cols = 33:35, rows = start_row + 1)  # Other M
  mergeCells(wb, sheet = sheet_name, cols = 36:38, rows = start_row + 1)  # Other Unknown
  
  # Merge cells for Total (no gender breakdown)
  mergeCells(wb, sheet = sheet_name, cols = 39:42, rows = start_row + 1)  # Total
}



# Write the data
# Add the headers
# Add the first table
add_table_headers(wb, sheet_name = "Adaptive Bryant", start_row = 1)
writeData(wb, sheet = "Adaptive Bryant", x = ad_national, startRow = 4, headerStyle = NULL, colNames = F)

# Add the second table
add_table_headers(wb, sheet_name = "Adaptive Bryant", start_row = 17)
writeData(wb, sheet = "Adaptive Bryant", x = ad_can, startRow = 20, headerStyle = NULL, colNames = F)

# Add the third table
add_table_headers(wb, sheet_name = "Adaptive Bryant", start_row = 33)
writeData(wb, sheet = "Adaptive Bryant", x = ad_HB, startRow = 36, headerStyle = NULL, colNames = F)

# Add the fourth table
add_table_headers(wb, sheet_name = "Adaptive Bryant", start_row = 49)
writeData(wb, sheet = "Adaptive Bryant", x = ad_man, startRow = 52, headerStyle = NULL, colNames = F)

# Add the fifth table
add_table_headers(wb, sheet_name = "Adaptive Bryant", start_row = 65)
writeData(wb, sheet = "Adaptive Bryant", x = ad_nel, startRow = 68, headerStyle = NULL, colNames = F)


##################################################################################################################
# Add e-asTTle worksheet to the workbook
addWorksheet(wb, sheetName = "e-asTTle")

# Write the data
# Add the first table
add_table_headers(wb, sheet_name = "e-asTTle", start_row = 1)
writeData(wb, sheet = "e-asTTle", x = easttle_all, startRow = 4, headerStyle = NULL, colNames = F)

# Add the second table
add_table_headers(wb, sheet_name = "e-asTTle", start_row = 17)
writeData(wb, sheet = "e-asTTle", x = easttle_can, startRow = 20, headerStyle = NULL, colNames = F)

# Add the third table
add_table_headers(wb, sheet_name = "e-asTTle", start_row = 33)
writeData(wb, sheet = "e-asTTle", x = easttle_HB, startRow = 36, headerStyle = NULL, colNames = F)

# Add the fourth table
add_table_headers(wb, sheet_name = "e-asTTle", start_row = 49)
writeData(wb, sheet = "e-asTTle", x = easttle_man, startRow = 52, headerStyle = NULL, colNames = F)

# Add the fifth table
add_table_headers(wb, sheet_name = "e-asTTle", start_row = 65)
writeData(wb, sheet = "e-asTTle", x = easttle_nel, startRow = 68, headerStyle = NULL, colNames = F)


##################################################################################################################
# Add DATA and full sheets

addWorksheet(wb, sheetName = "DATA")
writeData(wb, sheet = "DATA", x = DATA, startRow = 1, startCol = 1)

addWorksheet(wb, sheetName = "Raw data")
writeData(wb, sheet = "Raw data", x = combined_data, startRow = 1, startCol = 1)


##################################################################################################################
# Add an excluded files sheet, using the excluded_files dataframe
addWorksheet(wb, sheetName = "Excluded .xlsx files")
writeData(wb, sheet = "Excluded .xlsx files", x = excluded_file_paths_df, startRow = 1, startCol = 1)


##################################################################################################################
# Add a notes sheet
# Number of children in AB and e-asTTle tables that had no recorded ethnicity
DATA %>%
  filter(!is.na(`Pre AB score`) & !is.na(`Final AB score`) & is.na(`Ethnicity 1`)) %>%
  nrow()
# 10 children in the AB data had no recorded ethnicity

DATA %>%
  filter(`Beginning e-asTTle overall scaled score` != 0 & `End e-asTTle overall scaled score` != 0 & is.na(`Ethnicity 1`)) %>%
  nrow()
  

notes <- data.frame("NOTES" = c("Gender is expressed as F (female), M (male) and U (unknown - gender was not entered).",
                                "Year level contains all years entered by the teachers (i.e., some may be out of scope). Students with no year are set as 'unknown'.",
                                "A value of 'NaN' appears in some percentages because there were no students in that category (0 divided by 0 produces an error).",
                                "For the Ethnicity sheet, percentages are calculated row-wise, except for the total column. Only the first ethnicity value was used.",
                                "Where there were two students with the same first name and teacher, but one student was missing the surname, the student missing the surname was assumed to be the same as the student with the surname.",
                                "Mean values are presented for summaries of adaptive Bryant and e-asTTle scores across groups.",
                                "For summary Adaptive Bryant and e-asTTle tables, the ethnicity 'other' includes children with no recorded ethnicity (n=10 for Adaptive Bryant, n=0 for e-asTTle).",
                                "Excluded files are listed in the excluded files sheet. They were excluded due to being copies."))

addWorksheet(wb, sheetName = "Notes")
writeData(wb, sheet = "Notes", x = notes, startRow = 1, startCol = 1)


##################################################################################################################
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

# for (sheet in sheets) {
#   # Get the number of rows and columns used in the sheet
#   usedRows <- getRows(wb, sheet) + 10
#   usedCols <- getCols(wb, sheet)
#   
#   addStyle(wb, sheet = sheet, style = style, rows = 1:usedRows, cols = 1:usedCols, gridExpand = TRUE)
#   setColWidths(wb, sheet = sheet, cols = 1:usedCols, widths = "auto")
#   
# }

# Styles specific to AB and e-asTTle sheets
# Create a style for shading the first three rows grey and centering text
grey_style <- createStyle(
  fontColour = "black",       # Optional: make text stand out on grey
  fgFill = "grey",            # Grey fill
  halign = "center",          # Center-align text
  valign = "center",          # Vertical alignment
  textDecoration = "bold"     # Optional: make text bold
)

border_style <- createStyle(
  border = "TopBottomLeftRight",  # Add borders to all sides
  borderStyle = "thin"           # Thin border style
)

# Apply the grey style to the first 3 rows and the border style to each table of Adaptive Bryant sheet
addStyle(wb, sheet = "Adaptive Bryant", style = grey_style, rows = c(1:3, 17:19, 33:35, 49:51, 65:67), cols = 1:ncol(ad_national), gridExpand = TRUE)
addStyle(wb, sheet = "Adaptive Bryant", style = border_style, rows = c(1:15, 17:31, 33:47, 49:63, 65:79), cols = 1:ncol(ad_national), gridExpand = TRUE, stack = T)

# Apply the grey style to the first 3 rows and the border style to each table of e-asTTle sheet
addStyle(wb, sheet = "e-asTTle", style = grey_style, rows = c(1:3, 17:19, 33:35, 49:51, 65:67), cols = 1:ncol(easttle_all), gridExpand = TRUE)
addStyle(wb, sheet = "e-asTTle", style = border_style, rows = c(1:12, 17:28, 33:44, 49:60, 65:76), cols = 1:ncol(easttle_all), gridExpand = TRUE, stack = T)


##################################################################################################################
# Save the workbook to a file
saveWorkbook(wb, file = "test_report.xlsx", overwrite = TRUE)

length(unique(DATA$Teacher))

##################################################################################################################
##################################################################################################################

