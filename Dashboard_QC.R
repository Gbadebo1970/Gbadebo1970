#Loading the tables 
library(readxl)
library(writexl)
library(flexdashboard)
library(magrittr)
library(dplyr)
library(tibble)
setwd("C:/Audit/Dashboard_QC")



#DISCIPLINE QC

AGO_AGO <- readxl::read_excel("2021-10-15 AGO_AGO - 1 of 1 - Manual Edit - OFFSEN.xlsx")
AGO_CPS <- readxl::read_excel("2021-10-15 AGO_CPS - 1 of 1 - Raw - OFFSEN.xlsx")
DCMS_DCMS <- readxl::read_excel("2021-10-15 DCMS_DCMS - 1 of 1 - Raw - OFFSEN.xlsx")
BEIS_BBB <- readxl::read_excel("2021-10-15 BEIS_BBB - 1 of 1 - Raw - OFFSEN.xlsx")
BEIS_CCC <- readxl::read_excel("2021-10-15 BEIS_CCC - 1 of 1 - Raw - OFFSEN.xlsx")
BEIS_SBC <- readxl::read_excel("2021-10-15 BEIS_SBC - 1 of 1 - Raw - OFFSEN.xlsx")
DCMS_ACE <- readxl::read_excel("2021-10-15 DCMS_ACE - 1 of 1 - Manual Edit - OFFSEN.xlsx")
CO_EHRC  <- readxl::read_excel("2021-10-15 CO_EHRC - 1 of 1 - Raw - OFFSEN.xlsx")
CO_GCF <- readxl::read_excel("2021-10-15 CO_GCF - 1 of 1 - Raw - OFFSEN.xlsx")
CO_GDS <- readxl::read_excel("2021-10-15 CO_GDS - 1 of 1 - Raw - OFFSEN.xlsx")

AGO_AGO <- AGO_AGO  %>% 
  dplyr::group_by(Dept_Discipline)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

AGO_CPS <- AGO_CPS  %>% 
  dplyr::group_by(Dept_Discipline)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

DCMS_DCMS <- DCMS_DCMS  %>% 
  dplyr::group_by(Dept_Discipline)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

DCMS_ACE <- DCMS_ACE  %>% 
  dplyr::group_by(Dept_Discipline)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )



BEIS_BBB <- BEIS_BBB  %>% 
  dplyr::group_by(Dept_Discipline)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )


BEIS_CCC <- BEIS_CCC  %>% 
  dplyr::group_by(Dept_Discipline)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

BEIS_SBC <- BEIS_SBC  %>% 
  dplyr::group_by(Dept_Discipline)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )


BEIS_SBC <- BEIS_SBC  %>% 
  dplyr::group_by(Dept_Discipline)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

CO_EH <- CO_EH  %>% 
  dplyr::group_by(Dept_Discipline)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

CO_EHRC <- CO_EHRC  %>% 
  dplyr::group_by(Dept_Discipline)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

CO_GCF <- CO_GCF  %>% 
  dplyr::group_by(Dept_Discipline)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

CO_GDS <- CO_GDS  %>% 
  dplyr::group_by(Dept_Discipline)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

##Writing the extracting file to excel
writexl::write_xlsx(AGO_AGO, path = "C:/Audit/Dashboard_QC/Extraction to Excel/AGO_AGO.xlsx")
writexl::write_xlsx(AGO_CPS, path = "C:/Audit/Dashboard_QC/Extraction to Excel/AGO_CPS.xlsx")
writexl::write_xlsx(DCMS_ACE, path = "C:/Audit/Dashboard_QC/Extraction to Excel/DCMS_ACE.xlsx")
writexl::write_xlsx(BEIS_BBB, path = "C:/Audit/Dashboard_QC/Extraction to Excel/BEIS_BBB.xlsx")
writexl::write_xlsx(BEIS_CCC, path = "C:/Audit/Dashboard_QC/Extraction to Excel/BEIS_CCC.xlsx")
writexl::write_xlsx(BEIS_SBC, path = "C:/Audit/Dashboard_QC/Extraction to Excel/BEIS_SBC.xlsx")
writexl::write_xlsx(DCMS_DCMS, path = "C:/Audit/Dashboard_QC/Extraction to Excel/DCMS_DCMS.xlsx")
writexl::write_xlsx(CO_EHRC, path = "C:/Audit/Dashboard_QC/Extraction to Excel/CO_EHRC.xlsx")
writexl::write_xlsx(CO_GCF, path = "C:/Audit/Dashboard_QC/Extraction to Excel/CO_GCF.xlsx")
writexl::write_xlsx(CO_GDS, path = "C:/Audit/Dashboard_QC/Extraction to Excel/CO_GDS.xlsx")




#LOCATION QC

DFE_OS <- readxl::read_excel("2021-10-15 DFE_OS - 1 of 1 - Raw - OFFSEN.xlsx")
DFE_STA <- readxl::read_excel("2021-10-15 DFE_STA - 1 of 1 - Raw - OFFSEN.xlsx")
DFT_DVSA <- readxl::read_excel("2021-10-15 DFT_DVSA - 1 of 1 - Raw - OFFSEN.xlsx")
DFT_MCA <- readxl::read_excel("2021-10-15 DFT_MCA - 1 of 1 - Raw - OFFSEN.xlsx")
DHSC_DHSC <- readxl::read_excel("2021-10-15 DHSC_DHSC - 1 of 1 - Raw - OFFSEN.xlsx")
DHSC_NHSBT <- readxl::read_excel("2021-10-15 DHSC_NHSBT - 1 of 1 - Raw - OFFSEN.xlsx")
FCDO_WFD  <- readxl::read_excel("2021-10-15 FCDO_WFD - 1 of 1 - Raw - OFFSEN.xlsx")
FCDO_WP <- readxl::read_excel("2021-10-15 FCDO_WP - 1 of 1 - Raw - OFFSEN.xlsx")
DCMS_BFI <- readxl::read_excel("2021-10-15 DCMS_BFI - 1 of 1 - Manual Edit - OFFSEN.xlsx")
DCMS_ACE <- readxl::read_excel("2021-10-15 DCMS_ACE - 1 of 1 - Manual Edit - OFFSEN.xlsx")


DFE_OS <- DFE_OS  %>% 
  dplyr::group_by(Job_Location)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

DFE_STA <- DFE_STA %>% 
  dplyr::group_by(Job_Location)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

DFT_DVSA <- DFT_DVSA  %>% 
  dplyr::group_by(Job_Location)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

DFT_MCA <- DFT_MCA  %>% 
  dplyr::group_by(Job_Location)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )


DHSC_DHSC <- DHSC_DHSC  %>% 
  dplyr::group_by(Job_Location)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

DHSC_NHSBT <- DHSC_NHSBT  %>% 
  dplyr::group_by(Job_Location)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )


FCDO_WFD <- FCDO_WFD  %>% 
  dplyr::group_by(Job_Location)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

FCDO_WP <- FCDO_WP  %>% 
  dplyr::group_by(Job_Location)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

DCMS_BFI <- DCMS_BFI  %>% 
  dplyr::group_by(Job_Location)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )

DCMS_ACE <- DCMS_ACE  %>% 
  dplyr::group_by(Job_Location)  %>%
  dplyr::summarise(
    Comms_FTEs = sum(FTE)
  )


##Writing the extracting file to excel
writexl::write_xlsx(DFE_OS, path = "C:/Audit/Dashboard_QC/Location_QC/DFE_OS.xlsx")
writexl::write_xlsx(DFE_STA, path = "C:/Audit/Dashboard_QC/Location_QC/DFE_STA.xlsx")
writexl::write_xlsx(DFT_DVSA, path = "C:/Audit/Dashboard_QC/Location_QC/DFT_DVSA.xlsx")
writexl::write_xlsx(DFT_MCA, path = "C:/Audit/Dashboard_QC/Location_QC/DFT_MCA.xlsx")
writexl::write_xlsx(DHSC_DHSC, path = "C:/Audit/Dashboard_QC/Location_QC/DHSC_DHSC.xlsx")
writexl::write_xlsx(DHSC_NHSBT, path = "C:/Audit/Dashboard_QC/Location_QC/DHSC_NHSBT.xlsx")
writexl::write_xlsx(FCDO_WFD, path = "C:/Audit/Dashboard_QC/Location_QC/FCDO_WFD.xlsx")
writexl::write_xlsx(FCDO_WP, path = "C:/Audit/Dashboard_QC/Location_QC/FCDO_WP.xlsx")
writexl::write_xlsx(DCMS_BFI, path = "C:/Audit/Dashboard_QC/Location_QC/DCMS_BFI.xlsx")
writexl::write_xlsx(DCMS_ACE, path = "C:/Audit/Dashboard_QC/Location_QC/DCMS_ACE.xlsx")

