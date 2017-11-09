#
# Modification to FMC input
#
# Look through "Rebalance Tool Manual" word file


# LOAD PACKAGES
libraries <- c("tidyverse","readxl","timeDate","RODBC","mailR","xlsx")
lapply(libraries, function(lib){ library(lib, character.only=TRUE) })

# READ FMC HOLDINGS
fmc_wd <- "S:/Quantitative Investing & Applied Research/QIAR/Downloads/"
fmc_fileN <- paste0("EQUAB_",format(Sys.Date(), "%Y%m%d"),".LOG")   

# READ TRADELIST
tradelist_wd <- "S:/Quantitative Investing & Applied Research/QIAR/TradingTool/Input/S/TradeList/"    # change 
tl_fileN <- "TL_QESBP_20171104.xlsx"                                                                  # change 

# Output Trader trade list
tgtreport_wd <- "S:/Quantitative Investing & Applied Research/QIAR/TradingTool/Output/TradeReport/"
tl2trader_wd <- "S:/Quantitative Investing & Applied Research/QIAR/TradingTool/Output/FinalList2Trader/"

# Email list
em_to <- c("Casey.Li@VestcorInvestments.com")

# items need to be removed from FMC holding
rm_items <- c("MMF", "CFX", "FS")

# items need to be replaced from FMC holding
GTE_old <- "38500T101U"
GTE_new <- gsub("U","",GTE_old)

WIR_u <- "92937G109"
WIR_u_cny <- "CAD"

# Shares threshold for target portfolio (Number less than threshold will be removed from target portfolio)
shares_threshold <- 0

#
# Helper function
#
AddLastPriceFromXpressfeed <- function(fmc_input_cln, curcy){
  # CUSIPS
  ids <- paste0("'", paste(fmc_input_cln$ID, collapse = "','"),"'")
  
  # Lastest date
  dts <- timeSequence(Sys.Date() - 5, Sys.Date() - 1)
  bizdts <- dts[isBizday(dts, holidayNYSE())]
  lastest_bizdt <- format(as.Date(bizdts[length(bizdts)]), "%Y-%m-%d")
  
  sql <- "SELECT  sec.cusip AS ID, prc.prccd AS PRICE, prc.datadate AS PrcDate
            FROM  Xpressfeed.dbo.sec_dprc AS prc INNER JOIN
            Xpressfeed.dbo.security AS sec ON prc.gvkey = sec.gvkey AND prc.iid = sec.iid
            WHERE (sec.cusip IN (MYIDS))
            AND (prc.datadate = CONVERT(DATETIME, 'MYDATE 00:00:00', 102))
            AND (sec.secstat = 'A')
            AND (prc.curcdd = 'MYCURRENCY')"
  
  new_sql <- gsub("MYIDS", ids, sql)
  new_sql <- gsub("MYDATE", lastest_bizdt, new_sql)
  new_sql <- gsub("MYCURRENCY", curcy, new_sql)
  
  conn <- odbcConnect('Capital-IQ')
  prc <- sqlQuery(conn, new_sql, stringsAsFactors = FALSE)
  
  res <- left_join(fmc_input_cln, prc, by = "ID")
  
  # Fill missing price
  if(nrow(price_lookup) > 0){
    comm_ids <- intersect(res$ID, price_lookup$ID)
    res[res$ID %in% comm_ids, "PRICE"] <- price_lookup[price_lookup$ID %in% comm_ids, "Price"]
  }
  
  res[is.na(res$PRICE), "PRICE"] <- 0
  ifelse(sum(is.na(res$PRICE)) != 0, 
         paste0("No price obtained for ", paste0(res[is.na(res$PRICE), "ID"], collapse = ","), " | Price changed to 0"), 
         "Good: No NA price")
  
  return(res)
}

#
# Functions for Pension portfolio
#

PTgtportGenerator <- function(fmc_input_cln, tl_input){
  uniqueId <- unique(c(fmc_input_cln$ID,tl_input$ID))
  masterList <- tibble(ID = uniqueId)
  
  # join masterID and fmc holdings
  myjoin1 <- left_join(masterList, fmc_input_cln, by = "ID")
  myjoin1[is.na(myjoin1$HOLDINGS),"HOLDINGS"] <- 0
  
  # melt tradeList by trade side
  tl_input_cln <- tl_input[,c("ID","TRADE","SIDE")] %>% 
    spread(SIDE, TRADE) 
  ifelse(length(intersect("B",colnames(tl_input_cln))) == 0, tl_input_cln$B <- 0, "B exist")
  ifelse(length(intersect("CB",colnames(tl_input_cln))) == 0, tl_input_cln$CB <- 0, "CB exist")
  ifelse(length(intersect("S",colnames(tl_input_cln))) == 0, tl_input_cln$S <- 0, "S exist")
  ifelse(length(intersect("SS",colnames(tl_input_cln))) == 0, tl_input_cln$SS <- 0, "SS exist")
  
  # check shares remained same
  ifelse(sum(tl_input$TRADE) == sum(tl_input_cln[,-1], na.rm = TRUE), "OK", "Error")
  
  # join masterID, fmc holdings with melt tradeList
  myjoin2 <- full_join(myjoin1, tl_input_cln, by = "ID") %>% 
    mutate(B = ifelse(is.na(B), 0, abs(B))) %>% 
    mutate(CB = ifelse(is.na(CB), 0, abs(CB))) %>%
    mutate(S = ifelse(is.na(S), 0, abs(S))) %>%
    mutate(SS = ifelse(is.na(SS), 0, abs(SS)))
  
  # check shares remained same after transfer to absulute value
  ifelse(sum(abs(tl_input$TRADE)) == sum(myjoin2[,-c(1,2)], na.rm = TRUE), "OK", "Error")
  
  # compare each trading record with holdings
  myjoin2$AdjS <- ""
  myjoin2$NewSS <- ""
  myjoin2$AdjCB <- ""
  myjoin2$NewB <- ""
  myjoin2$TgtShares <- ""
  myjoin2$Msg <- ""
  
  for(i in 1:nrow(myjoin2)){
    # Set default value
    s_adj <- 0
    ss_add <- 0
    cb_adj <- 0
    b_add <- 0
    msg <- "Good"
    
    rec <- myjoin2[i,]
    if(rec$S != 0){
      if(rec$HOLDINGS > 0){
        # Long
        if(rec$S - rec$HOLDINGS > 0){
          s_adj <- rec$HOLDINGS
          ss_add <- rec$S - rec$HOLDINGS
          msg <- paste0("Warning: Leveraged. S side capped by holdings at ", 
                        rec$HOLDINGS, " | New SS generated: ", ss_add)
        } else {
          s_adj <- rec$S
          msg <- "Good"
        }
      } else { 
        # Short
        s_adj <- 0
        ss_add <- rec$S
        msg <- paste0("Warning: No S due to negative holding. Change to SS. Check with PM")
      }
    } else if(rec$CB !=0) {
      if(rec$HOLDINGS > 0){
        # long
        cb_adj <- 0
        b_add <- rec$CB
        msg <- paste0("Warning: B instead of CB")
      } else { 
        # Short
        if(rec$CB > abs(rec$HOLDINGS)){
          cb_adj <- abs(rec$HOLDINGS)
          b_add <- rec$CB - abs(rec$HOLDINGS)
          msg <- paste0("Warning: CB side capped by holdings at ", 
                        rec$HOLDINGS, " | New B generated: ", b_add) 
        } else {
          cb_adj <- rec$CB
          msg <- "Good"
        }
      }
    } else {
      # no check
    }
    
    # Add result to target list
    myjoin2[i, "AdjS"] <- s_adj
    myjoin2[i, "NewSS"] <- ss_add
    myjoin2[i, "AdjCB"] <- cb_adj
    myjoin2[i, "NewB"] <- b_add
    myjoin2[i, "TgtShares"] <- rec$HOLDINGS + rec$B + cb_adj + b_add - s_adj - ss_add - rec$SS
    myjoin2[i, "Msg"] <- msg
  }
  
  myjoin2 <- myjoin2 %>%
    mutate(NewPosition = ifelse(HOLDINGS == 0 & TgtShares != 0, "Yes", "No")) %>% 
    mutate(TgtShares = as.numeric(TgtShares))
  
  
  return(myjoin2)
}

TradelistGenerator <- function(fmc_input_cln, tgt_port){
  
  # Master list
  uniqueId <- unique(c(fmc_input_cln$ID, tgt_port$ID))
  masterList <- tibble(ID = uniqueId)
  
  tmp1 <- left_join(masterList, fmc_input_cln, by = "ID")
  tmp2 <- left_join(tmp1, tgt_port, by = "ID")
  
  # Generate primary trade list
  tmp2$TradeShares <- 0
  tmp2[is.na(tmp2$HOLDINGS), "TradeShares"] <- tmp2[is.na(tmp2$HOLDINGS), "TgtShares"]
  tmp2[is.na(tmp2$TgtShares), "TradeShares"] <- -tmp2[is.na(tmp2$TgtShares), "HOLDINGS"]
  tmp2[!is.na(tmp2$TgtShares) & !is.na(tmp2$HOLDINGS), "TradeShares"] <- 
    tmp2[!is.na(tmp2$TgtShares) & !is.na(tmp2$HOLDINGS), "TgtShares"] -
    tmp2[!is.na(tmp2$TgtShares) & !is.na(tmp2$HOLDINGS), "HOLDINGS"]
  
  # Generate primary trade list
  tmp2$NewSide <- ""
  tmp2[is.na(tmp2$HOLDINGS) & !is.na(tmp2$TgtShares) & tmp2$TgtShares > 0, "NewSide"] <- "B"
  tmp2[is.na(tmp2$HOLDINGS) & !is.na(tmp2$TgtShares) & tmp2$TgtShares < 0, "NewSide"] <- "SS"
  tmp2[is.na(tmp2$TgtShares) & !is.na(tmp2$HOLDINGS) & tmp2$HOLDINGS > 0, "NewSide"] <- "S"
  tmp2[is.na(tmp2$TgtShares) & !is.na(tmp2$HOLDINGS) & tmp2$HOLDINGS < 0, "NewSide"] <- "B"
  tmp2[!is.na(tmp2$TgtShares) & !is.na(tmp2$HOLDINGS) &
        (tmp2$TgtShares > tmp2$HOLDINGS), "NewSide"] <- "B"
  tmp2[!is.na(tmp2$TgtShares) & !is.na(tmp2$HOLDINGS) &
         (tmp2$TgtShares < tmp2$HOLDINGS) &
         (tmp2$TgtShares < 0), "NewSide"] <- "SS"
  tmp2[!is.na(tmp2$TgtShares) & !is.na(tmp2$HOLDINGS) &
         (tmp2$TgtShares < tmp2$HOLDINGS) &
         (tmp2$TgtShares >= 0), "NewSide"] <- "S"
  
  # Check if new position
  tmp2$NewPosition <- "No"
  tmp2[is.na(tmp2$HOLDINGS), "NewPosition"] <- "Yes"
         
  return(tmp2[tmp2$TradeShares != 0, c("ID", "TradeShares", "NewSide", "NewPosition", "Msg")])
}

TradelistComp <- function(new_tl, org_tl){
  uniqueId <- unique(c(org_tl$ID, new_tl$ID))
  masterList <- data.frame(ID = uniqueId, stringsAsFactors = FALSE)  # Use data.frame on purpose
  
  # tmp1: ID, TICKER, PRICE, TRADE, SIDE, COUNTRY
  tmp1 <- left_join(masterList, org_tl, by = "ID")
  
  # tmp2: ID, TICKER, PRICE, TRADE, SIDE, TradeShares, NewSide, Msg 
  tmp2 <- left_join(tmp1, new_tl, by = c("ID"))
  
  # Generate primary trade list
  tmp2$Msg2 <- "Good"
  tmp2[is.na(tmp2$TRADE), "Msg2"] <- paste0("Warning: Trade added", tmp2[is.na(tmp2$TRADE), "TradeShares"])                # Should never happen theoretically and logically with the exception of programming error
  tmp2[is.na(tmp2$TradeShares), "Msg2"] <- paste0("Warnng: Trade deleted", tmp2[is.na(tmp2$TradeShares), "TRADE"])
  tmp2[!is.na(tmp2$TRADE) & !is.na(tmp2$TradeShares) & tmp2$TRADE != tmp2$TradeShares, "Msg2"] <- 
    paste0("Warnings: Trade modified from ", 
           tmp2[!is.na(tmp2$TRADE) & !is.na(tmp2$TradeShares) & tmp2$TRADE != tmp2$TradeShares, "TRADE"], 
           " to ", 
           tmp2[!is.na(tmp2$TRADE) & !is.na(tmp2$TradeShares) & tmp2$TRADE != tmp2$TradeShares, "TradeShares"])
  
  # tmp2: ID, TICKER, PRICE, TRADE, COUNTRY, TradeShares, NewSide, Msg, Msg2
  res <- tmp2[, c("ID", "TICKER", "PRICE", "COUNTRY", "TradeShares", "NewSide", "Msg", "Msg2")]
  res <- as_tibble(res)
  return(res)
}

TradeSummary <- function(tl, portName){
  #
  tl <- tl %>% 
    mutate(Side = ifelse(NewSide == "S" | NewSide == "SS", "S", "B")) %>% 
    mutate(TGTSIZE = abs(TradeShares)) %>% 
    mutate(TRADEDOLLARAMT = TGTSIZE * PRICE)
  
  sum_tbl <- tl %>% 
    group_by(Side) %>% 
    summarise(NoOfEquities = n(), 
              TradeShares = sum(TGTSIZE), 
              TradeDollarAmount = sum(TRADEDOLLARAMT))
  
  colnames(sum_tbl)[-1] <- paste0(portName, " - ", colnames(sum_tbl)[-1])
  return(sum_tbl)
}

TL2Trader <- function(tl, portName){
  
  if(portName == "QESBP" | portName == "QESB1" | portName == "ACES" | 
     portName == "LVCA" | portName == "LVUSP" | portName == "LVUS" ){
    # Use ticker
    tl_fmtd <- tl %>% 
      mutate(`Trg Size` = abs(TradeShares)) %>% 
      mutate(SIDE = NewSide)
    res <- tl_fmtd[,c("TICKER", "Trg Size", "SIDE", "COUNTRY")]
  } else {
    # Use SEDOL
    tl_fmtd <- tl %>% 
      mutate(`Trg Size` = abs(TradeShares)) %>% 
      mutate(SIDE = NewSide) %>% 
      mutate(SEDOL = ID)
    res <- tl_fmtd[,c("SEDOL", "Trg Size", "SIDE", "COUNTRY")]
  }
  
  return(res)
}

PTradeList <- function(fmc_fn, tl_fn, port, curcy = "CAD", receipt_emails){
  
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
  fmc_input[is.na(fmc_input$CNATCUR), "CNATCUR"] <- "CAD"
  ifelse(port == "QESBP", fmc_input <- fmc_input[fmc_input$CNATCUR == curcy, ], fmc_input <- fmc_input)
  fmc_input_cln <- fmc_input[fmc_input$PORTCOD == port_name, c("ID", "HOLDINGS")]
  for(i in 1:length(rm_items)){
    fmc_input_cln <- fmc_input_cln[!grepl(rm_items[i], fmc_input_cln$ID, ignore.case = TRUE),]
  }
  
  # Get Tradelist input
  tl_inputN <- paste0(tradelist_wd, tl_fn)
  cl_type <- c("text","text","numeric","numeric","text","text")
  tl_input <- read_excel(tl_inputN, sheet = 1, col_types = cl_type)
  colnames(tl_input)[1] <- "ID"
  
  # Full target portfolio with all info
  tgt_port_full <- PTgtportGenerator(fmc_input_cln, tl_input)
  
  # Standardized target portfolio with ID, TgtShares and Message only
  tgt_port_std <- tgt_port_full[abs(tgt_port_full$TgtShares) > shares_threshold, c("ID", "TgtShares", "Msg")]
  
  # Generate Primiary trade list with ID, TradeShares, NewSide, Message
  trade_list1 <- TradelistGenerator(fmc_input_cln, tgt_port_std)
  
  # Compare generated tradelist with original tradelist
  # trade_list2: ID, TICKER, PRICE, COUNTRY, TradeShares, NewSide, Msg1, Msg2
  tl2casey <- TradelistComp(trade_list1, tl_input)
  
  # Get summary information
  trade_summary <- TradeSummary(tl2casey, port_name)
  
  # Format final trader list to trader
  tl2trader <- TL2Trader(tl2casey, port_name)
  
  # Output files to excel
  trader_excel <- paste0(tl2trader_wd, "Tradelist (", port_name, ") ", Sys.Date(), ".xlsx")
  report_excel <- paste0(tgtreport_wd, "Report (", port_name, ") ", Sys.Date(), ".xlsx")
  
  ifelse(file.exists(trader_excel), file.remove(trader_excel), "File created.")
  ifelse(file.exists(report_excel), file.remove(report_excel), "File created.")
  
  write.xlsx(tl2trader, trader_excel, sheetName = port_name, row.names = TRUE)
  write.xlsx(tgt_port_std, report_excel, sheetName = "Target Portfolio", append = TRUE, row.names = TRUE)
  write.xlsx(tl2casey, report_excel, sheetName = "Trade list", append = TRUE, row.names = TRUE)
  write.xlsx(trade_summary, report_excel, sheetName = "Trade Summary", append = TRUE, row.names = TRUE)

  # Email trade (Not sent)
  send.mail(from = "Casey.Li@VestcorInvestments.com", to = receipt_emails, subject = paste0("Tradelist for ", port_name),
            body = paste0(trade_summary,collapse = ","), html = TRUE, send = FALSE, 
            smtp = list(host.name = "EXCHANGE.NBInvestment.com"),
            attach.files = c(trader_excel, report_excel))
  
  return(list(TargetPortfolio = tgt_port_std, 
              CaseyTradeList = tl2casey,
              TradeList2Trader = tl2trader,
              TradeSummary = trade_summary))
}

#
# Functions for Non pension portfolio
#

NPTgtportGenerator <- function(p_tgtport, fmc_input_cln, port_name, curcy = "CAD"){
  
  # Get pension target portfolio weight
  p_tgtport_wprc <- AddLastPriceFromXpressfeed(p_tgtport, curcy)
  p_base_value <- sum(abs(p_tgtport_wprc$TgtShares * p_tgtport_wprc$PRICE), na.rm = TRUE)
  p_tgtport_wprc$Wgt <- p_tgtport_wprc$TgtShares * p_tgtport_wprc$PRICE / p_base_value
  
  # This will add a column. Price might be NA if no data in xpressfeed
  fmc_input_wprc <- AddLastPriceFromXpressfeed(fmc_input_cln, curcy)
  np_base_value <- sum(abs(fmc_input_wprc$HOLDINGS * fmc_input_wprc$PRICE), na.rm = TRUE)     
  
  # Get non-pension target portfolio shares
  np_tgtport <- tibble(ID = p_tgtport_wprc$ID,
                       TgtShares = p_tgtport_wprc$Wgt * np_base_value / p_tgtport_wprc$PRICE,
                       Msg = paste0("QESBP Legacy msg: ", p_tgtport_wprc$Msg))
  return(np_tgtport)
}

NPTradeSummary <- function(tl, port_name, p_trd_sum){
  
  np_trd_sum <- TradeSummary(tl, port_name)
  fnl_trd_sum <- full_join(p_trd_sum, np_trd_sum, by = "Side")
  fnl_trd_sum[is.na(fnl_trd_sum)] <- 0
  
  ratio <- fnl_trd_sum %>% filter(Side == B)/fnl_trd_sum %>% filter(Side == B)
  fnl_trd_sum <- bind_rows(fnl_trd_sum, ratio)
  
  return(fnl_trd_sum)
}

NPTradeList <- function(fmc_fn, tl_fn, port, p_tgt_port, p_trd_summary, curcy = "CAD", receipt_emails){
  
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
  fmc_input[is.na(fmc_input$CNATCUR), "CNATCUR"] <- "CAD"
  ifelse(port == "QESB1", fmc_input <- fmc_input[fmc_input$CNATCUR == curcy, ], fmc_input <- fmc_input)
  fmc_input_cln <- fmc_input[fmc_input$PORTCOD == port_name, c("ID", "HOLDINGS")]
  for(i in 1:length(rm_items)){
    fmc_input_cln <- fmc_input_cln[!grepl(rm_items[i], fmc_input_cln$ID, ignore.case = TRUE),]
  }

  # Get Tradelist input
  tl_inputN <- paste0(tradelist_wd, tl_fn)
  cl_type <- c("text","text","numeric","numeric","text","text")
  tl_input <- read_excel(tl_inputN, sheet = 1, col_types = cl_type)
  colnames(tl_input)[1] <- "ID"
  
  tgt_port_np <- NPTgtportGenerator(p_tgt_port, fmc_input_cln, port_name, curcy)
  tgt_port_std <- tgt_port_np[abs(tgt_port_full$TgtShares) >= shares_threshold, c("ID", "TgtShares", "Msg")]
  
  # Trade list
  tradelist1 <- TradelistGenerator(fmc_input_cln, tgt_port_std)
  
  # Compare generated tradelist with original tradelist
  # trade_list1: ID, TICKER, PRICE, COUNTRY, TradeShares, NewSide, Msg1, Msg2
  tl2casey <- TradelistComp(tradelist1, tl_input)
  
  # Summary
  trade_summary <- NPTradeSummary(tl2casey, port_name, p_trd_summary)
  
  # Format final trader list to trader
  tl2trader <- TL2Trader(tl2casey, port_name)
  
  # Output files to excel
  trader_excel <- paste0(tl2trader_wd, "Tradelist (", port_name, ") ", Sys.Date(), ".xlsx")
  report_excel <- paste0(tgtreport_wd, "Report (", port_name, ") ", Sys.Date(), ".xlsx")
  
  ifelse(file.exists(trader_excel), file.remove(trader_excel), "File created.")
  ifelse(file.exists(report_excel), file.remove(report_excel), "File created.")
  
  write.xlsx(tl2trader, trader_excel, sheetName = port_name, row.names = TRUE)
  write.xlsx(tgt_port_std, report_excel, sheetName = "Target Portfolio", append = TRUE, row.names = TRUE)
  write.xlsx(tl2casey, report_excel, sheetName = "Trade list", append = TRUE, row.names = TRUE)
  write.xlsx(trade_summary, report_excel, sheetName = "Trade Summary", append = TRUE, row.names = TRUE)

  # Email trade (Not sent)
  send.mail(from = "Casey.Li@VestcorInvestments.com", to = receipt_emails, subject = paste0("Tradelist for ", port_name),
            body = trade_summary, html = TRUE, send = FALSE, 
            attach.files = c(trader_excel, report_excel))
  
  return(list(TargetPortfolio = tgt_port_std, 
              CaseyTradeList = tl2casey,
              TradeList2Trader = tl2trader,
              TradeSummary = trade_summary))
}

#
# Trade list generation process
#
qesbp_tl <- PTradeList(fmc_fileN, tl_fileN, "QESBP", "CAD", em_to)
qesb1_tl <- NPTradeList(fmc_fileN, tl_fileN, "QESB1", qesbp$TargetPortfolio, qesbp$TradeSummary, "CAD", em_to)