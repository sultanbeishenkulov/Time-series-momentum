rm(list=ls())
#downloading the whole df function
read_excel_allsheets <- function(filename, tibble = FALSE) {
  sheets <- readxl::excel_sheets(filename)
  x <- lapply(sheets, function(X) readxl::read_excel(filename, sheet = X))
  if(!tibble) x <- lapply(x, as.data.frame)
  names(x) <- sheets
  x
}

#spots and forward have the same names, we need to have ones for both
common_column_names <- c('DATE', 'GBP(UK)', 'CHF(SWITZERLAND)', 'JPY(JAPAN)', 'CAD(CANADA)','AUD(AUSTRALIA)','NZD(NEW_ZEALAND)', 'SEK(SWEDEN)', 'NOK(NORWAY)', 'DKK(DANIA)', 'EUR', 'DM(GERMANY)', 'ITL(ITALY)','FRF(FRANCE)','NLG(NETHERLANDS)', 'BEF(BELGIUM)', 
                        'FIM(FINLAND)', 'IEP(IRELAND)', 'ZAR(SOUTH_AFRICA)', 'SGD(SINGAPORE)', 'ATS(AUSTRALIA)','CZK(ZHECH_REPUBLIC)', 'GRD(GREECE)', 'HUF(HUNGARY)', 'INR(INDIA)', 'IDR(INDONESIA)', 'KWD(KUWAIT)', 'MYR(MALAYSIA)', 'MXN(MEXICO)',
                        'PHP(PHILIPPINES)', 'PLN(POLAND)', 'PTE(PORTUGAL)', 'KRW(SOUTH_KOREA)', 'ESP(SPAIN)', 'TWD(TAIWAN)', 'THB(THAILAND)', 'BRL(BRAZIL)','EGP(EGYPT)', 'RUB(RUSSIA)', 'SKK(SLOVAKIA)', 'HRK(CROATIA)', 'CYP(CYPRUS)', 'ILS(ISRAEL)', 
                        'ISK(ICELAND)', 'SIT(SLOVENIA)','BGN(BULGARIA)', 'UAH(UKRAINE)'
                        )

# Packages -------------------------------------------------------------------------------------
library("readxl")
library("quantmod")
library("tseries")
library("PerformanceAnalytics")
library("ggplot2")
library("dplyr")
library("tsibble") 
library("feasts") # feature extraction and statistics
library("fable") # time series modeling and forecasting
library("forecast")
library("lubridate") # date, time and frequency
library("gridExtra")
library('xts')
library("PerformanceAnalytics")
library("matrixStats")


#nice colour for plot
UTblue <- rgb(5/255, 167/255, 1)

#downloading the dataframe
df <- read_excel_allsheets("/Users/sultanbeishenkulov/Master degree in Data Science/fourth semester/quant/paper/ERdata.xlsx")

#dividing by sheets
Bids1 <-df$Bids1
Bids2 <-df$Bids2
Offers1 <- df$Offers1
Offers2 <- df$Offers2
FBids1 <- df$FBids1
FBids2 <- df$FBids2
FOffers1 <- df$FOffers1
FOffers2 <- df$FOffers2

#merging the frames
Bids <- merge(df$Bids1, df$Bids2)
Offers <- merge(df$Offers1, df$Offers2)
FBids<- merge(df$FBids1, df$FBids2)
FOffers <- merge(df$FOffers1, df$FOffers2)


#indexing the dataframes with dates (NAs in frames are strings, so we need to convert)
#deleting the name column
Bids <- mutate_all(Bids, function(x) as.numeric(as.character(x)))
Bids <- Bids[,-1]

Offers <- mutate_all(Offers, function(x) as.numeric(as.character(x)))
Offers <- Offers[,-1]

FBids <- mutate_all(FBids, function(x) as.numeric(as.character(x)))
FBids <- FBids[,-1]

FOffers <- mutate_all(FOffers, function(x) as.numeric(as.character(x)))
FOffers <- FOffers[,-1]


#We have two exchange rates inversed, so here we gonna convert it in all 4 dataframes
Bids$`US $ TO UK £ (BBI) - BID SPOT` <- lapply(Bids$`US $ TO UK £ (BBI) - BID SPOT`, function(x) round((1/x), digits = 4))
Bids$`US $ TO IRISH PUNT (BBI) - BID SPOT` <- lapply(Bids$`US $ TO IRISH PUNT (BBI) - BID SPOT`, function(x) round((1/x), digits = 4))
names(Bids)[names(Bids) == 'US $ TO UK £ (BBI) - BID SPOT'] <- 'UK £ TO US $ (BBI) - BID SPOT'
names(Bids)[names(Bids) == 'US $ TO IRISH PUNT (BBI) - BID SPOT'] <- 'IRISH PUNT TO US $ (BBI) - BID SPOT'


Offers$`US $ TO UK £ (BBI) - SPOT OFFERED` <- lapply(Offers$`US $ TO UK £ (BBI) - SPOT OFFERED`, function(x) round((1/x), digits = 4))
Offers$`US $ TO IRISH PUNT (BBI) - SPOT OFFERED` <- lapply(Offers$`US $ TO IRISH PUNT (BBI) - SPOT OFFERED`, function(x) round((1/x), digits = 4))
names(Offers)[names(Offers) == 'US $ TO UK £ (BBI) - SPOT OFFERED'] <- 'UK £ TO US $ (BBI) - SPOT OFFERED'
names(Offers)[names(Offers) == 'US $ TO IRISH PUNT (BBI) - SPOT OFFERED'] <- 'IRISH PUNT TO US $ (BBI) - SPOT OFFERED'

FBids$`US $ TO UK £ 1M FWD (BBI) - BID SPOT` <- lapply(FBids$`US $ TO UK £ 1M FWD (BBI) - BID SPOT`, function(x) round((1/x), digits = 4))
FBids$`US $ TO IRISH PUNT 1M FWD(BBI) DISC - BID SPOT` <- lapply(FBids$`US $ TO IRISH PUNT 1M FWD(BBI) DISC - BID SPOT`, function(x) round((1/x), digits = 4))
names(FBids)[names(FBids) == 'US $ TO UK £ 1M FWD (BBI) - BID SPOT'] <- 'UK £ TO US $ 1M FWD (BBI) - BID SPOT'
names(FBids)[names(FBids) == 'US $ TO IRISH PUNT 1M FWD(BBI) DISC - BID SPOT'] <- 'IRISH PUNT TO US $ 1M FWD(BBI) DISC - BID SPOT'

FOffers$`US $ TO UK £ 1M FWD (BBI) - SPOT OFFERED` <- lapply(FOffers$`US $ TO UK £ 1M FWD (BBI) - SPOT OFFERED`, function(x) round((1/x), digits = 4))
FOffers$`US $ TO IRISH PUNT 1M FWD(BBI) DISC - SPOT OFFERED` <- lapply(FOffers$`US $ TO IRISH PUNT 1M FWD(BBI) DISC - SPOT OFFERED`, function(x) round((1/x), digits = 4))
names(FOffers)[names(FOffers) == 'US $ TO UK £ 1M FWD (BBI) - SPOT OFFERED'] <- 'UK £ TO US $ 1M FWD (BBI) - SPOT OFFERED'
names(FOffers)[names(FOffers) == 'US $ TO IRISH PUNT 1M FWD(BBI) DISC - SPOT OFFERED'] <- 'IRISH PUNT TO US $ 1M FWD(BBI) DISC - SPOT OFFERED'

#Dropping pegged currencies:
Bids <- Bids[ , -which(names(Bids) %in% c("HONG KONG $ TO US $ (BBI) - BID SPOT","SAUDI RIYAL TO US $ (WMR) - BID SPOT"))]
Offers <- Offers[ , -which(names(Offers) %in% c("HONG KONG $ TO US $ (BBI) - SPOT OFFERED","SAUDI RIYAL TO US $ (WMR) - SPOT OFFERED"))]
FBids <- FBids[ , -which(names(FBids) %in% c("HONG KONG $ TO US $ 1M FWD (BBI) - BID SPOT","SAUDI RIYAL TO US $ 1M FWD (WMR) - BID SPOT"))]
FOffers <- FOffers[ , -which(names(FOffers) %in% c("HONG KONG $ TO US $ 1M FWD (BBI) - SPOT OFFERED","SAUDI RIYAL TO US $ 1M FWD (WMR) - SPOT OFFERED"))]


#transforming to numeric again because of UK and Ireland transformations
Bids <- mutate_all(Bids, function(x) as.numeric(as.character(x)))
Offers <- mutate_all(Offers, function(x) as.numeric(as.character(x)))
FBids <- mutate_all(FBids, function(x) as.numeric(as.character(x)))
FOffers <- mutate_all(FOffers, function(x) as.numeric(as.character(x)))

#data column
Bids <- cbind(as.Date(Bids1$Name), Bids)
names(Bids)[names(Bids) == 'as.Date(Bids1$Name)'] <- 'Date'
Offers <- cbind(as.Date(Offers1$Name), Offers)
names(Offers)[names(Offers) == 'as.Date(Offers1$Name)'] <- 'Date'
FBids <- cbind(as.Date( FBids1$Name), FBids)
names(FBids)[names(FBids) == 'as.Date(FBids1$Name)'] <- 'Date'
FOffers <- cbind(as.Date( FOffers1$Name), FOffers)
names(FOffers)[names(FOffers) == 'as.Date(FOffers1$Name)'] <- 'Date'

#renaming the columns with the common ones
colnames(Bids) <- common_column_names
colnames(Offers) <- common_column_names
colnames(FBids) <- common_column_names
colnames(FOffers) <- common_column_names

#Putting NAs for European Countries after 2001
#Let us have the list of related countries:
#1.German Mark at January of 1999 (index: 12)
#2.Italian lira at January of 1999 (index: 13)
#3.French Franc at January of 1999 (index: 14)
#4.Neth Guilder at January of 1999 (index: 15)
#5.Belgian Franc at January of 1999 (index: 16)
#6.Finish Markka at January of 1999 (index: 17)
#7. Irish Punt at January of 1999 (index: 18)
#8. Greek Drachma at January of 2001 (index: 23)
#9.Portuguese Escudo at of January 1999 (index: 32)
#10.Spanish Paseta at January of 1999 (index: 34)
#11.Slovak Koruna at January of 2009 (index: 40)
#12.Cyprus Lira at January of 2008 (index: 42)
#13.Slovenian Tolar at January of 2007 (index: 45)

#For Bid
#putting NA for countries that yielded Euro starting at January of 1999
Bids[183:dim(Bids)[1],c(12,13,14,15,16,17,18,23,32,34)] <- NA
#putting NA for Greece that yielded Euro starting at January of 2001
Bids[219:dim(Bids)[1],23] <- NA
#putting NA for Slovakia that yielded Euro starting at January of 2001
Bids[303:dim(Bids)[1],40] <- NA
#putting NA for Cyprus that yielded Euro starting a  January of 2008
Bids[291:dim(Bids)[1],42] <- NA
#putting NA for Slovenia that yielded Euro starting at January of 2007
Bids[279:dim(Bids)[1],45] <- NA

#for Offer
#putting NA for countries that yielded Euro starting at January of 1999
Offers[183:dim(Offers)[1],c(12,13,14,15,16,17,18,23,32,34)] <- NA
#putting NA for Greece that yielded Euro starting at January of 2001
Offers[219:dim(Offers)[1],23] <- NA
#putting NA for Slovakia that yielded Euro starting at January of 2001
Offers[303:dim(Offers)[1],40] <- NA
#putting NA for Cyprus that yielded Euro starting at January of 2008
Offers[291:dim(Offers)[1],42] <- NA
#putting NA for Slovenia that yielded Euro starting at January of 2007
Offers[279:dim(Offers)[1],45] <- NA

#for Forward bids
#putting NA for countries that yielded Euro starting at January of 1999
FBids[183:dim(FBids)[1],c(12,13,14,15,16,17,18,23,32,34)] <- NA
#putting NA for Greece that yielded Euro starting at January of 2001
FBids[219:dim(FBids)[1],23] <- NA
#putting NA for Slovakia that yielded Euro starting at January of 2001
FBids[303:dim(FBids)[1],40] <- NA
#putting NA for Cyprus that yielded Euro starting at January of 2008
FBids[291:dim(FBids)[1],42] <- NA
#putting NA for Slovenia that yielded Euro starting at January of 2007
FBids[279:dim(FBids)[1],45] <- NA

#for Forward offers
#putting NA for countries that yielded Euro starting at January of 1999
FOffers[183:dim(FOffers)[1],c(12,13,14,15,16,17,18,23,32,34)] <- NA
#putting NA for Greece that yielded Euro starting at January of 2001
FOffers[219:dim(FOffers)[1],23] <- NA
#putting NA for Slovakia that yielded Euro starting at January of 2001
FOffers[303:dim(FOffers)[1],40] <- NA
#putting NA for Cyprus that yielded Euro starting at January of 2008
FOffers[291:dim(FOffers)[1],42] <- NA
#putting NA for Slovenia that yielded Euro starting at January of 2007
FOffers[279:dim(FOffers)[1],45] <- NA


#Calculate midrates
Midrates <- (Bids[,-1] + Offers[,-1])/2
FMidRates <- (FBids[,-1] + FOffers[,-1])/2

#logs
LogBid <- log(Bids[,-1])
LogOffer <- log(Offers[,-1])
LogFBid <- log(FBids[,-1])
LogFOffer <- log(FOffers[,-1])

#mid-rates logs
MidRateLog <- (LogBid + LogOffer)/2
FMidRateLog <- (LogFBid + LogFOffer)/2

#midrates excess returns
Midratesexreturns <- FMidRateLog[-427,]-MidRateLog[-1,]

#creating new variable in order to work with it later in the loop
#assing 0 for NAs
Nod <- Midratesexreturns
Nod[is.na(Nod)] <- 0

##equally weighted portfolio (ev)
Nod.ev <- as.data.frame(as.numeric(rowMeans(Nod, na.rm = TRUE)))
View(Nod.ev)

#initializing momentum column (with NA)
Nod.mom <- matrix(NA,dim(Nod)[1],1)

#sorting for momentum:
#print(dim(Nod)[1])
for(i in 13:dim(Nod.mom)[1]){ #13
  #print(i)
  #Column products 
  Nodx <- colProds(as.matrix(Nod[(i-12),])+1) - 1 
  #print(Nodx)
  #Sorted quantile for the portfolio
  Nodx1 <- as.numeric(Nodx>=rep(quantile(Nodx,0.8,na.rm=TRUE),12))
  #print(Nodx1)
  #Sum of Sorted quartile
  Nodx2 <- sum(as.numeric(Nodx>=rep(quantile(Nodx,0.8,na.rm=TRUE),12)))
  
  #Multiply with matrix
  Nodx3<- Nodx1%*%t(Nod[(i),]) 
  
  Nod.mom[i] <- Nodx3/Nodx2
  #print(Nod.x4)
}

#Nodx1 = 46x1
#Nod = 426x46

#creating dataframe for both portfolios
LSPort <- cbind(Nod.ev,Nod.mom)[-(1:12),]
colnames(LSPort) <- c("EW","Momentum")
View(LSPort)

##################################################################


#sending estimates to txt file
sink("results_of_estimation_NEW.txt")

#preventing scientific notions
options(scipen=999)

#Creating Date column
Date1 <- ts(data=LSPort , start=c(1984, 12) ,end= c(2019,05) , frequency = 12)

#Convert in to timeseries data
LSPort_ts <- xts(x = LSPort, order.by =as.Date(Date1))


#estimations, sharpe (annual as well)
sharpe1 <- lapply(LSPort_ts, function(x) SharpeRatio(x*12))
sharpe2 <- data.frame(sharpe1)
sharpe2


#Value at Risk, "Cornish-Fisher", "gaussian", "historical", 
VaR_mod <- lapply(LSPort_ts, function(x) VaR(x,method = c("modified")) )
VaR_mod1 <- data.frame(VaR_mod)
VaR_mod1

VaR_gau <- lapply(LSPort_ts, function(x) VaR(x,method = c("gaussian")) )
VaR_mod1 <- data.frame(VaR_gau)
VaR_mod1

VaR_his <- lapply(LSPort_ts, function(x) VaR(x,method = c("historical")) )
VaR_his1 <- data.frame(VaR_his)
VaR_his1
sink()

#checking
SharpeRatio(LSPort_ts)
SharpeRatio.annualized(LSPort_ts)
StdDev.annualized(LSPort_ts, scale = 12)
Return.annualized(LSPort_ts, scale = 12, geometric = TRUE)

#Converting to dataframe
LSPort_df <- as.data.frame(LSPort_ts)
Date3 = seq.Date(as.Date("1984-12-01"),as.Date("2019-05-01"),by="month")
LSPort_ts2 <- cbind(Date3,LSPort)


Return.cumulative(LSPort_ts2$Momentum, geometric = TRUE)
#plot cumularive returns of momentum
plot(ts(data = cumsum(LSPort$Momentum), end = c(2019, 5), frequency = 12),type = 'l', main = 'cumulative returns', xlab = 'dates' ,ylab = 'returns')
lines(ts(data = cumsum(LSPort$Momentum), end = c(2019, 5), frequency = 12), type = 'l', col = 'red', legend = 'asd')
lines(ts(data = cumsum(LSPort$EW), end = c(2019, 5), frequency = 12), type = 'l', col = UTblue)
legend("topleft",
       c("momentum","EW"),
       fill=c("red",UTblue),
       cex=0.7
)
##Plot for the GBP currency, Examples of result 
ggplot(data = LSPort_ts2) + 
  geom_line(aes(x= Date3, y = LSPort_ts2$Momentum, group = 1, color = 'mom')) +
  labs(color="", title = "Momentum Porfolio") +
  xlab('Years') + ylab('Returns')+
  theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust=1))



