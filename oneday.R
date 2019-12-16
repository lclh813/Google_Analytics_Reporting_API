# Install packages.
install.packages("googleAnalyticsR")
install.packages("magrittr")
install.packages("dplyr")
install.packages("xlsx")

library(googleAnalyticsR)
library(magrittr)
library(dplyr)
library(xlsx)

# Email authentication.
ga_auth(email="")

# Lists all accounts to which the user has access.
account_list <- ga_account_list() %>% View()

# Specify which account to access.
viewId_web = ""

# check general log info by device.
get_overview = function(s) {
  overview = google_analytics(viewId_web,
                              date_range = c(s, s),
                              metrics = c("ga:users", "ga:sessions", "ga:pageviews", "searchSessions"),
                              dimensions = c("ga:deviceCategory"))
  return (overview)
}

# Check add-to-cart action by device.
get_cart = function(s) {
  filter_includeCart = dim_filter(dimension = "ga:eventCategory",
                                  operator = "REGEXP",
                                  expressions = "ADD_TO_CART",
                                  not = FALSE)
  filter_excludeCart = dim_filter(dimension = "ga:eventAction",
                                  operator = "REGEXP",
                                  expressions = "lock",
                                  not = TRUE)
  
  filter_main = filter_clause_ga4(list(filter_includeCart, filter_excludeCart), operator = "AND")
  
  cart = google_analytics(viewId_web,
                          date_range =c(s, s),
                          metrics = c("ga:totalEvents"),
                          dimensions = c("ga:deviceCategory"),
                          dim_filters = filter_main,
                          max = -1,
                          useResourceQuotas = TRUE)
  
  return (cart)
}

# Check campaign pageviews by device.
get_campaign = function(s){
  filter_campaign = dim_filter(dimension = "ga:pagePath",
                               operator = "REGEXP",
                               expressions = "/campaign/",
                               not = FALSE)
  
  filter_main = filter_clause_ga4(list(filter_campaign))
  
  campaign = google_analytics(viewId_web,
                              date_range =c(s, s),
                              metrics = c("ga:pageviews"),
                              dimensions = c("ga:deviceCategory"),
                              dim_filters = filter_main,
                              max = -1,
                              useResourceQuotas = TRUE)

  return (campaign)
}

# Deal with missing values of campaign pageviews.
set_campaign = function(campaign) {
  temp = data.frame("deviceCategory"=c("desktop","mobile","tablet"))
  
  if (is.null(campaign)) {
    campaign_null = data.frame("deviceCategory"=c("desktop","mobile","tablet"), "pageviews"=c(0,0,0))
    campaign_corr = left_join(temp, campaign_null, by=c("deviceCategory"="deviceCategory"))
  } else {
    campaign_corr = left_join(temp, campaign, by=c("deviceCategory"="deviceCategory"))
    campaign_corr[is.na(campaign_corr)] = 0
  }
  
  return(campaign_corr)
}

# Declare empty vectors to store responding data. 
desktop = c()
mobile = c()
tablet = c()

# Get current date.
today = Sys.Date()
today = as.character(today)
# Split current date into year, month, and day.
today = strsplit(today, "-")
today[[1]]
year = today[[1]][1]
month = today[[1]][2]
day = as.numeric(today[[1]][3]) - 1
# Select year and month of the date range.
year_month = paste(year, "-", month, "-%02d", sep = "")
# Select day of the date range.
days = 1:day

for (d in days) {
  s = sprintf(year_month, d)
  overview = get_overview(s)
  campaign = set_campaign(get_campaign(s))
  cart = get_cart(s)
  names(campaign)[names(campaign) == "pageviews"] = "campaignPageviews"
  temp = Reduce(function(x, y) merge(x, y, by="deviceCategory"), list(overview, campaign, cart))
  temp["date"] = s
  desktop = rbind(desktop, temp[1,])
  mobile = rbind(mobile, temp[2,])
  tablet = rbind(tablet, temp[3,])
}

# Transpose dataframes.
dfList = list(df1=desktop, df2=mobile, df3=tablet)
dfList = lapply(dfList, function(df) {
  df = t(df) # Transpose
  colnames(df) = df[nrow(df), ] # Make the last row be the header
  df = df[-nrow(df), ] # Remove the last row
  df = df[c(1,2,3,4,7,5,6),] # Reorder
  df = apply(df,2,trimws) # Delete blanks
})

# Export Excel file.
write.xlsx(dfList$df1, file="oneday.xlsx", sheetName="desktop")
write.xlsx(dfList$df2, file="oneday.xlsx", sheetName="mobile", append=TRUE)
write.xlsx(dfList$df3, file="oneday.xlsx", sheetName="tablet", append=TRUE)
