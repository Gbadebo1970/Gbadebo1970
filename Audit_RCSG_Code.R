# Code to read in or import an excel sheet
# Debo's file
library(readxl)
Audit_RGCS_Debo <- readxl::read_excel("Debo.xlsx")
#Viewing the the dataset imported
View(Audit_RGCS_Debo)
# Ashley's file
library(readxl)
library(writexl)
Audit_RGCS_Ashley <- readxl::read_excel("Ashley_2021_09_29_FTE_per_entity_OFFICIAL.xlsx")
#Viewing the the dataset imported
View(Audit_RGCS_Ashley)
#Merging Debo and Ashley files
Debo_Ashley_latest <- merge(Audit_RGCS_Debo,Audit_RGCS_Ashley, by = "entity_code")
#Rows that exist in Debo but not in Ashley
Debo_Ashley_Difference_latest <- Audit_RGCS_Debo[is.na(match(Audit_RGCS_Debo$entity_code, Audit_RGCS_Ashley$entity_code)),]
View(Debo_Ashley_Difference_latest)
#Rows that exist in Ashley but not in Debo
Ashley_Debo_Difference_latest <- Audit_RGCS_Ashley[is.na(match(Audit_RGCS_Ashley$entity_code, Audit_RGCS_Ashley$entity_code)),]
View(Ashley_Debo_Difference_latest)

#compare Debo and Ashley FTE values
Debo_Ashley_latest$FTE_Result <- ifelse(Debo_Ashley$FTE_sum ==Debo_Ashley$fte, 'True', 'False')

#compare Debo and Ashley Salary Mean values
Debo_Ashley_latest$Salary_Mean_Result <- ifelse(Debo_Ashley$Salary_mean ==Debo_Ashley$mean_salary_mean, 'True', 'False')

#Rows that exist in Audit_RGCS_Ashley but not in Debo_Ashley
Ashley_Mani_Difference_latest <- Audit_RGCS_Ashley[is.na(match(Audit_RGCS_Ashley$entity_code, Debo_Ashley$entity_code)),]
View(Ashley_Debo_Difference_latest)

# Export the dataframe into excel: Install the writexl package, You may type the following command in the R console in order to install the writexl package:
install.packages("writexl")

write_xlsx(Debo_Ashley_latest, path = "C:/Audit/Result1_latest.xlsx")

write_xlsx(Ashley_Mani_Difference_latest, path = "C:/Audit/Ashley_Difference_latest.xlsx")




