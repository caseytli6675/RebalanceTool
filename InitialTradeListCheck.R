#
# Look through "Rebalance Tool Manual" word file
#

# LOAD PACKAGES
libraries <- c("tidyverse","readxl","timeDate","RODBC","mailR","xlsx")
lapply(libraries, function(lib){ library(lib, character.only=TRUE) })

# READ FMC HOLDINGS
fmc_wd <- "S:/Quantitative Investing & Applied Research/QIAR/Downloads/"
fmc_fileN <- paste0("EQUAB_",format(Sys.Date(), "%Y%m%d"),".LOG")                                         

# READ PM INITIAL HOLDINGS
pmHolding_wd <- "S:/Quantitative Investing & Applied Research/QIAR/TradingTool/Input/S/InitialHolding/"     
pm_fileN <- "HOLDING_LVUSP_20171108.xlsx"     # change 

# Output Check Summary
holdck_wd <- "S:/Quantitative Investing & Applied Research/QIAR/TradingTool/Output/InitialHoldingCheck/"

# Email list
em_to <- c("Casey.Li@VestcorInvestments.com")

# items need to be removed from FMC holding
rm_items <- c("MMF", "CFX", "FS")

# items need to be replaced from FMC holding
GTE_old <- "38500T101U"
GTE_new <- gsub("U","",GTE_old)

WIR_u <- "92937G109"
WIR_u_cny <- "CAD"
#
# Functions for Pension portfolio
#

CkInitialHoldings <- function(fmc_fn, holding_fn, port, curcy = "CAD", receipt_emails){
  
  # Assign port name
  port_name <- port
  
  # Get FMC input
  fmc_inputN <- paste0(fmc_wd, fmc_fn)
  fmc_input <- read_csv(fmc_inputN, col_names = TRUE, col_types = "ccddd?c")
  colnames(fmc_input)[1] <- "ID"
  
  # Adjust anomolies
  fmc_input_cln[fmc_input_cln$ID==WIR_u,"CNATCUR"] <- WIR_u_cny
  fmc_input_cln[fmc_input_cln$ID==GTE_old,"ID"] <- GTE_new
  
  # Change currency NA to CAD, clean up FMC holdings
  fmc_input[is.na(fmc_input$CNATCUR), "CNATCUR"] <- "CAD"         # if it is NA, it is CAD by default
  ifelse(port == "QESBP" | port == "QESB1", fmc_input <- fmc_input[fmc_input$CNATCUR == curcy, ], fmc_input <- fmc_input)

  fmc_input_cln <- fmc_input[fmc_input$PORTCOD == port_name, c("ID", "HOLDINGS")]
  for(i in 1:length(rm_items)){
    fmc_input_cln <- fmc_input_cln[!grepl(rm_items[i], fmc_input_cln$ID, ignore.case = TRUE),]
  }

 
  # Get PM holding input
  pm_inputN <- paste0(pmHolding_wd, holding_fn)
  cl_type <- c("text","text","numeric")
  pm_input <- read_excel(pm_inputN, sheet = 1, col_types = cl_type)
  colnames(pm_input)[1] <- "ID"
  
  # remove NA or N/A ID
  pm_input_cln <- filter(pm_input, ID != "N/A")
  pm_input_cln <- pm_input_cln[!is.na(pm_input_cln$ID),]
  
  # Check PM list to see if all rows have ID
  ifelse(length(pm_input_cln$ID) == nrow(pm_input_cln), "OK", paste0((nrow(pm_input_cln)-length(pm_input_cln$ID)),"ID is Missing"))
  
  # Join and compare the two
  uniqueId <- unique(c(fmc_input_cln$ID,pm_input_cln$ID))
  masterList <- tibble(ID = uniqueId)
  
  # join masterID and fmc holdings
  myjoin1 <- left_join(masterList, fmc_input_cln, by = "ID")
  myjoin1[is.na(myjoin1$HOLDINGS),"HOLDINGS"] <- 0

  # join myjoin1 with PM holdings
  myjoin2 <- left_join(myjoin1, pm_input_cln, by = "ID") 
  myjoin2[is.na(myjoin2$SHARES),"SHARES"] <- 0
  
  # Compare the share difference
  myjoin2$Diff <- myjoin2$HOLDINGS - myjoin2$SHARES
  myjoin2$Msg <- "Good"
  
  myjoin2[myjoin2$Diff > 0,"Msg"] <- "More Shares in FMC"
  myjoin2[myjoin2$Diff < 0,"Msg"] <- "Warning: Short shares"
  myjoin2 <- myjoin2[order(myjoin2$Msg, decreasing = TRUE),]
  
  # Return the Initial holdings with ID, TICKER, FMC_HOLDINGS, SHARES, Msg
  res <- myjoin2 %>% rename(FMC_HOLDINGS = HOLDINGS)
  res <- res[,c("ID", "TICKER", "FMC_HOLDINGS", "SHARES", "Diff", "Msg")]

  
  # Output files to excel
  report_excel <- paste0(holdck_wd, "HoldingCheck (", port_name, ") ", Sys.Date(), ".xlsx")
  
  ifelse(file.exists(report_excel), file.remove(report_excel), "File created.")
  
  write.xlsx(res, report_excel, sheetName = port_name, row.names = TRUE)
  
  # Email trade (Not sent)
  # send.mail(from = "Casey.Li@VestcorInvestments.com", to = receipt_emails, subject = paste0("Tradelist for ", port_name),
  #           body = "", html = TRUE, send = FALSE, 
  #           smtp = list(host.name = "EXCHANGE.NBInvestment.com"),
  #           attach.files = report_excel)
  # 
  return(res)
}


#
# Trade list generation process
#
lvusp_holdingChk <- CkInitialHoldings(fmc_fileN, pm_fileN, "LVUSP", receipt_emails = em_to)
