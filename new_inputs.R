library(openxlsx2)
library(data.table)
library(lubridate)
library(readxl)
library(dplyr)
library(tidyr)
library(reticulate)

setwd("C:/Users/Michael Smith/OneDrive - University of California, Davis/Documents/Research")

# Energy ------------------------------------------------------------------
pge_wide_LMPs <- fread("Code Base/temp_data/temptemp/pge_LMPs.csv",header = TRUE)
sce_wide_LMPs <- fread("Code Base/temp_data/temptemp/sce_LMPs.csv",header = TRUE)
sdge_wide_LMPs <- fread("Code Base/temp_data/temptemp/sdge_LMPs.csv",header = TRUE)

# Generation --------------------------------------------------------------
hourly_allocation_weights <- fread("Code Base/temp_data/temptemp/hourly_ra_allocation_2019_2025.csv")[,-1] #ADD THIS IN AT C11, sheet = "Generation"
cap_summary_values <- fread("Code Base/temp_data/temptemp/annual_system_ra_summary.csv")

yrly_cap_values <- dcast(
  melt(cap_summary_values, id.vars = "Year"), 
  variable ~ Year, 
  value.var = "value"
)

yrly_cap_values <- yrly_cap_values[variable == "Annual_Wtd_Avg_Price", !"variable"]
#ADD THIS IN AT C4, sheet = "Generation"

# Transmission --------------------------------------------------------------
PGE_PCAFs <- fread("Code Base/temp_data/temptemp/PGE_trans_PCAFs_2019_2025.csv")
SCE_PCAFs <- fread("Code Base/temp_data/temptemp/SCE_trans_PCAFs_2019_2025.csv")
SDGE_PCAFs <- fread("Code Base/temp_data/temptemp/SDGE_trans_PCAFs_2019_2025.csv")

comb_PCAFs <- cbind(PGE_PCAFs[,-1],SCE_PCAFs[,-1],SDGE_PCAFs[,-1]) #ADD THIS AT X17, sheet = "Transmission"

# Distribution --------------------------------------------------------------
dist_PCAFs_2019 <- fread("Code Base/temp_data/temptemp/dist_PCAFs_2019.csv")[,-c(1,2)] #ADD THIS AT X25, sheet = "Distribution"
dist_PCAFs_2020 <- fread("Code Base/temp_data/temptemp/dist_PCAFs_2020.csv")[,-c(1,2)] #ADD THIS AT AU25, sheet = "Distribution"
dist_PCAFs_2021 <- fread("Code Base/temp_data/temptemp/dist_PCAFs_2021.csv")[,-c(1,2)] #ADD THIS AT BR25, sheet = "Distribution"
dist_PCAFs_2022 <- fread("Code Base/temp_data/temptemp/dist_PCAFs_2022.csv")[,-c(1,2)] #ADD THIS AT CO25, sheet = "Distribution"
dist_PCAFs_2023 <- fread("Code Base/temp_data/temptemp/dist_PCAFs_2023.csv")[,-c(1,2)] #ADD THIS AT DL25, sheet = "Distribution"
dist_PCAFs_2024 <- fread("Code Base/temp_data/temptemp/dist_PCAFs_2024.csv")[,-c(1,2)] #ADD THIS AT EI25, sheet = "Distribution"
dist_PCAFs_2025 <- fread("Code Base/temp_data/temptemp/dist_PCAFs_2025.csv")[,-c(1,2)] #ADD THIS AT FF25, sheet = "Distribution"

# Emissions --------------------------------------------------------------
CAT_prices <- fread("Code Base/Supplementary Data/nc-allowance_prices.csv")[,-c(1,4)] %>%
  mutate(
    quarter = substr(`Quarter Year`, 1, 2),
    year    = substr(`Quarter Year`, nchar(`Quarter Year`)-3, nchar(`Quarter Year`)),
    Price_num = as.numeric(gsub("[\\$,]", "", `Current Auction Settlement Price`)) 
  ) %>%
  group_by(year) %>%
  summarize(average = mean(Price_num, na.rm = TRUE)) %>% 
  pivot_wider(names_from = year, values_from = average) %>% 
  dplyr::select(`2019`:`2025`) #ADD THIS AT S4, sheet = "Emissions"

fwrite(PGE_MOER_wider, "Code Base/temp_data/temptemp/pge_moer.csv")
fwrite(SCE_MOER_wider, "Code Base/temp_data/temptemp/sce_moer.csv")
fwrite(SDGE_MOER_wider, "Code Base/temp_data/temptemp/sdge_moer.csv")

# AS --------------------------------------------------------------
AS_data <- read_excel("Code Base/Supplementary Data/AS_data_plusproj.xlsx") 
AS_transposed <- as.data.table(t(AS_data[,-1])) 
setnames(AS_transposed, as.character(AS_data$year))
AS_data_filtered <- AS_transposed %>% 
  dplyr::select(`2019`:`2025`)
#ADD THIS AT C3, sheet = "AS Procurement"

# Losses --------------------------------------------------------------
losses_yearly <- fread("Code Base/temp_data/temptemp/losses_yearly.csv") 
#ADD THIS AT M7, sheet = "Losses"

# ADD TO ACC spreadsheet --------------------------------------------------------------
py_install(c("numpy", "pandas", "openpyxl"))
py_config()

py_run_string("
import openpyxl

path = 'Code Base/Supplementary Data/ACC/2024 ACC Electric Model_NewTemplate.xlsx'
wb = openpyxl.load_workbook(path)

def write_to_excel(ws_name, data, start_col, start_row, skip_first_cols=0):
    ws = wb[ws_name]
    
    if hasattr(data, 'iloc'): 
        data_values = data.iloc[:, skip_first_cols:].values.tolist()
    elif isinstance(data, list):
        if not isinstance(data[0], list):
            data_values = [data[skip_first_cols:]]
        else:
            data_values = [row[skip_first_cols:] for row in data]
    else:
        data_values = data

    for r_idx, row_data in enumerate(data_values):
        for c_idx, value in enumerate(row_data):
            ws.cell(row=start_row + r_idx, column=start_col + c_idx).value = value

# --- Energy Section --- 
write_to_excel('Energy', r.pge_wide_LMPs, 3, 6)   # C6
write_to_excel('Energy', r.sce_wide_LMPs, 35, 6)  # AI6
write_to_excel('Energy', r.sdge_wide_LMPs, 67, 6) # BO6

# --- Generation Section ---
write_to_excel('Generation Capacity', r.yrly_cap_values, 3, 4)           # C4
write_to_excel('Generation Capacity', r.hourly_allocation_weights, 3, 11) # C11

# --- Transmission Section ---
# X is 24
write_to_excel('Transmission', r.comb_PCAFs, 24, 17)            # X17

# --- Distribution Section ---
# AU=47, BR=70, CO=93, DL=116, EI=139, FF=162
write_to_excel('Distribution', r.dist_PCAFs_2019, 24, 25)      # X25
write_to_excel('Distribution', r.dist_PCAFs_2020, 47, 25)      # AU25
write_to_excel('Distribution', r.dist_PCAFs_2021, 70, 25)      # BR25
write_to_excel('Distribution', r.dist_PCAFs_2022, 93, 25)      # CO25
write_to_excel('Distribution', r.dist_PCAFs_2023, 116, 25)     # DL25
write_to_excel('Distribution', r.dist_PCAFs_2024, 139, 25)     # EI25
write_to_excel('Distribution', r.dist_PCAFs_2025, 162, 25)     # FF25

# --- Emissions Section ---
# S=19, C=3, AJ=36, BQ=69
write_to_excel('Emissions', r.CAT_prices, 19, 4)               # S4
write_to_excel('Emissions', r.PGE_MOER_wider, 3, 48, skip_first_cols=1)  # C48
write_to_excel('Emissions', r.SCE_MOER_wider, 36, 48, skip_first_cols=1) # AJ48
write_to_excel('Emissions', r.SDGE_MOER_wider, 69, 48, skip_first_cols=1) # BQ48

# --- AS Procurement Section ---
write_to_excel('AS Procurement', r.AS_data_filtered, 3, 3)     # C3

# --- Losses ---
write_to_excel('Losses', r.losses_yearly, 13, 7)     # M7

# Finalize: Set to Auto-Calc and Save
wb.calculation.calcMode = 'auto'
wb.properties.calcId = '0'
wb.save('Peak Hours Report/2024_ACC_Updated.xlsx')
")

# recalc --------------------------------------------------------------
system('powershell -Command "$xl = New-Object -ComObject Excel.Application; $wb = $xl.Workbooks.Open(\'C:/Users/Michael Smith/OneDrive - University of California, Davis/Documents/Research/Peak Hours Report/2024_ACC_Updated.xlsx\'); $xl.Calculate(); $wb.Save(); $xl.Quit()"')
