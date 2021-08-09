#Loading the library's 
source("library.R")


lapply(list.files("outputs",full.names = T), unlink)# Removes any previous outputs


filenames <- list.files("inputs", full.names = TRUE, recursive = TRUE)
# Extracting out parameters/meta data from the files names to be later added into the datasets
sites <- str_extract(filenames, "BP|PIB|PCB|APCB|BOH|BOG")  # same length as filenames
years <- str_extract(filenames, "\\d{4}")  # extracting the years from the files
quarters <- str_extract(filenames, "\\d{1}+(?= Prudential|_)")  # extracting the quarter from the files


#function for combining file meta data information with file created
FUN_metainfo_combine <- function(filestoread=filenames,sheetname,skipnumber){
  stopifnot(length(filestoread)==length(sites))  # returns error if not the same length
  ans <- map2(filestoread, sites, ~read_excel(.x,sheetname,skip=skipnumber) %>% mutate(id = .y))  # .x is element in filenames, and .y is element in sites
  #adding year as a column
  ans_withyear<- Map(cbind, ans, year = years)
  #Adding the quarter as a column
  ans_everything<- Map(cbind, ans_withyear, quarter = quarters)
  return(ans_everything)
}

#-----------------------------------Balance Sheet---------------------S-------------------------
bs_df.list<- FUN_metainfo_combine(filestoread=filenames,
                                  sheetname = "Form 1 Bal Sheet",skipnumber=8)

#Cleaning up each excel file by removing columns,using row names as columns, changing column datatypes
df.list <- lapply(bs_df.list, function(x) {
  x <- x%>%select(c(-1,-6))         # Selecting specific columns
  x<- x[-1:-2,]
  x[2:4] <- lapply(x[2:4],as.integer) # making column specific columns to be integer
  x <- x[!apply(is.na(x[,2:4]), 1, all),] # removing NA values from specific columns
  return(x)})

listtotake<- df.list[sapply(df.list, ncol)==7]
balance_sheetmigrate<- do.call("rbind", listtotake)
fwrite(balance_sheetmigrate,"outputs/balancesheet.csv",col.names = FALSE)
#-----------------------------------Interest Rates-------------------------------------------
interestrates_df.list<- FUN_metainfo_combine(sheetname = "Sch 2 Interest Rates",skipnumber=9)

df.list <- lapply(interestrates_df.list, function(d) {
  d<- d%>%select(c(-1))
  d<- d[-1:-2,]
  d[2:9] <- lapply(d[2:9],as.numeric) # making column specific columns to be integer
  d[3:9]<- d[3:9]*100
  #d[14:16]<- d[14:16]*100
  return(d)})

interestrates_sheetmigrate<- do.call("rbind", df.list)
fwrite(interestrates_sheetmigrate,"outputs/interestrate.csv",col.names = FALSE)



#------------------------------consolidated P and L-----------------------------------------
pandl_df.list<- FUN_metainfo_combine(sheetname = "Form 2 - Consolidated P&L",skipnumber=9)

df.list <- lapply(pandl_df.list, function(d) {
  d<- d%>%select(c(2:5,16:18))
  d <- d[-1,]
  d[2:4] <- lapply(d[2:4],as.numeric) 
  return(d)
})
PandL_migrate<- do.call("rbind", df.list)
fwrite(PandL_migrate,"outputs/PandL.csv",col.names = FALSE)


#-------------------------------FX exposure-----------------------------------
fxexposure_df.list<- FUN_metainfo_combine(sheetname = "Sch 4 FX Exposure",skipnumber=9)

df.list <- lapply(fxexposure_df.list, function(d) {
  d<- d%>%select(c(-1))
  d[2:9] <- lapply(d[2:9],as.numeric) 
  return(d)
  
})

nocolnamelist<- lapply(df.list, function(x){
  colnames(x) <- NULL
  return(x)
})

nocolnamelist<- nocolnamelist[sapply(nocolnamelist, ncol)==13]

for (i in seq_along(nocolnamelist)) {
  colnames(nocolnamelist[[i]]) = paste0('V', seq_len(ncol(nocolnamelist[[i]])))
}


FXexposure_migrate<- do.call("rbind", nocolnamelist)
fwrite(FXexposure_migrate,"outputs/fxexposure.csv",col.names = FALSE)



#------------------------Remittances-----------------------------------
remmitances.list <- 
  FUN_metainfo_combine(sheetname = "Sch 4.1 Remittances",skipnumber=8)

test<- data.frame(remmitances.list[10])

df.list <- lapply(remmitances.list, function(d) {
  d<- d%>%select(c(-1))
  d[2:9] <- lapply(d[2:9],as.numeric) 
  return(d)
  
})

nocolnamelist<- lapply(df.list, function(x){
  colnames(x) <- NULL
  return(x)
})


nocolnamelist<- nocolnamelist[sapply(nocolnamelist, ncol)==13]

for (i in seq_along(nocolnamelist)) {
  colnames(nocolnamelist[[i]]) = paste0('V', seq_len(ncol(nocolnamelist[[i]])))
}


remitanceto_migrate<- do.call("rbind", nocolnamelist)
fwrite(remitanceto_migrate,"outputs/remittance.csv",col.names = FALSE)





#-----------------------------Large Borrowers---------------------------
largeborrowers_df.list<- FUN_metainfo_combine(sheetname = "Sch 7 - Large Borrowers",skipnumber=8)

df.list <- lapply(largeborrowers_df.list, function(d) {
  d<- d[-1,]
  d<- d%>%select(c(-1))
  d[2:8] <- lapply(d[2:8],as.numeric) # making column specific columns to be integer
  return(d)})

largeborrowers_sheetmigrate<- do.call("rbind", df.list)
fwrite(largeborrowers_sheetmigrate,"outputs/largeborrowers.csv",col.names = FALSE)
#---------------------------Large Depositors---------------------------
largedepositors_df.list<- FUN_metainfo_combine(sheetname = "Sch 6 - Large Depositors",skipnumber=8)

df.list <- lapply(largedepositors_df.list, function(d) {
  d<- d[-1,]
  d<- d%>%select(c(-1))
  d[2:10] <- lapply(d[2:10],as.numeric) # making column specific columns to be integer
  return(d)})

largedepositors_sheetmigrate<- do.call("rbind", df.list)
fwrite(largedepositors_sheetmigrate,"outputs/largedepositors.csv",col.names = FALSE)
#---------------------------Loans and Advances------------------------------
loansandadvances_df.list<- FUN_metainfo_combine(sheetname = "Schedule 1 Loans & Advance",skipnumber=9)
df.list <- lapply(loansandadvances_df.list, function(d) {
  d<- d[-1,]
  d<- d%>%select(c(-1))
  d[2:14] <- lapply(d[2:14],as.numeric) # making column specific columns to be integer
  return(d)})


loansadvances_sheetmigrate<- do.call("rbind", df.list)
fwrite(loansadvances_sheetmigrate,"outputs/loansandadvances.csv",col.names = FALSE)

#------------------------Maturity Gap------------------------------------------
maturitygap_df.list<- FUN_metainfo_combine(sheetname = "Sch 5 - Maturity GAP",skipnumber=8)

df.list <- lapply(maturitygap_df.list, function(d) {
  d<- d%>%select(c(-1))
  d[2:11] <- lapply(d[2:11],as.numeric) # making column specific columns to be integer
  return(d)})



maturitygap_sheetmigrate<- do.call("rbind", df.list)
fwrite(maturitygap_sheetmigrate,"outputs/maturitygap.csv",col.names = FALSE)

#----------------------Rate Sensitivity--------------------------------------
ratesensitivity_df.list<- FUN_metainfo_combine(sheetname = "Sch 8 - Rate Sensitivity",skipnumber=8)
df.list <- lapply(ratesensitivity_df.list, function(d) {
  d<- d%>%select(c(-1))
  d[2:6] <- lapply(d[2:6],as.numeric) # making column specific columns to be integer
  return(d)})

ratesensitivity_sheetmigrate<- do.call("rbind", df.list)

fwrite(ratesensitivity_sheetmigrate,"outputs/ratesensitivity.csv",col.names = FALSE)


#-----------------Related Parties--------------------------------------------
relatedparties_df.list<- FUN_metainfo_combine(sheetname = "Sch 3 - Related Parties",skipnumber=9)

df.list <- lapply(relatedparties_df.list, function(d) {
  d<- d%>%select(c(-1))
  d[2:9] <- lapply(d[2:9],as.numeric) # making column specific columns to be integer
  return(d)})

relatedparties_sheetmigrate<- do.call("rbind", df.list)
fwrite(relatedparties_sheetmigrate,"outputs/relatedparties.csv",col.names = FALSE)

#----------------Risk Based Capital------------------------------------------
riskbasedcapital_df.list<-
  FUN_metainfo_combine(sheetname = "Form 3 - Risk Based Capital",skipnumber=8)

df.list.step1 <- lapply(riskbasedcapital_df.list, function(d) {
  d<- d[1:64,2:10]
  d <- d[c(-5,-6)]
  d[2:4] <- lapply(d[2:4],as.numeric) # making column specific columns to be integer
  return(d)})



df.list.step2 <- lapply(riskbasedcapital_df.list, function(d) {
  d<- d[67:140,2:10]
  d[2:6] <- lapply(d[2:6],as.numeric) # making column specific columns to be integer
  return(d)})


test <- data.frame(df.list.step2[10])


df.list.step3 <- lapply(riskbasedcapital_df.list, function(d) {
  d<- d[144:180,2:10]
  d <- d[c(-5,-6)]
  d[2:4] <- lapply(d[2:4],as.numeric) # making column specific columns to be integer
  return(d)})

test <- data.frame(df.list.step3[10])



riskbasedcapital_step1_sheetmigrate<- do.call("rbind", df.list.step1)
riskbasedcapital_step2_sheetmigrate<- do.call("rbind", df.list.step2)
riskbasedcapital_step3_sheetmigrate<- do.call("rbind", df.list.step3)

fwrite(riskbasedcapital_step1_sheetmigrate,"outputs/riskbasedcapital_step1.csv",col.names = FALSE)
fwrite(riskbasedcapital_step2_sheetmigrate,"outputs/riskbasedcapital_step2.csv",col.names = FALSE)
fwrite(riskbasedcapital_step3_sheetmigrate,"outputs/riskbasedcapital_step3.csv",col.names = FALSE)



