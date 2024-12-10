This project is for the regular reports Learning Matters must provide to the Ministry of Education.

data_importation.R is for selecting, downloading and processing files on Google Drive and saving them as .xlsx files in the "Downloaded_xlsx_files" folder.

standard_report.R is for the usual recurring report. It takes the imported data in "Downloaded_xlsx_files", created by data_importation.R, and produces an xlsx report.

easTTle_beg_end_report.R is for when beginning and ending e-asTTle data need to be provided and therefore has a different format to the usual report.