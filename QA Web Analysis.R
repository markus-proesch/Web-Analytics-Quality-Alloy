------------------------
#Assignment 2: Quality Alloys, Inc. Website Analysis
#Team 4
#Date: 12/04/2019
------------------------

library(readxl)
library(ggplot2)
library(moments)
library(plotly)

# Function to imports tables from specific excel sheets and range
sheet_import <- function(sheet, range) {
  read_excel('Web Analytics QA.xls', sheet = sheet, range = range, col_names = TRUE)
}

# Import all tables needed for the basic analysis
weekly_visits  <- sheet_import(2, 'A5:H71')
financials     <- sheet_import(3, 'A5:E71')
daily_visits   <- sheet_import(5, 'A5:B467')
pounds_sold    <- sheet_import(4, 'A5:B295')
demo           <- sheet_import(6, 'B6:C78')
base_df        <- merge(weekly_visits, financials, sort=FALSE) #merging tables to get a base table with all data

# Divide the table into the 4 periodes of "Initial", "Pre-Promotion", "Promotion", "Post-Promotion"
all_periods  <- base_df
init_main    <- base_df[1:14,]
pre_main     <- base_df[15:35,]
prom_main    <- base_df[36:52,]
post_main    <- base_df[53:66,]

# Add a column to indicate the period
period_vector <- c(rep('Initial',14), rep('Pre-Promotion', 21), 
                   rep('Promotion', 17), rep('Post-Promotion', 14))
all_periods$Period
all_periods$Period <- as.factor(period_vector)

# Question 1
# Plot the variables with bar graphs 
ggplot(all_periods, aes(x = `Week (2008-2009)`, y = Visits, fill = Period))+
  geom_bar(stat = 'identity', width = 1) +
  theme(axis.text.x = element_text(angle = 90)) +
  labs(title = 'Visits per week (2008-2009)') +
  scale_x_discrete(limits = all_periods$`Week (2008-2009)`)
  
ggplot(all_periods, aes(x = `Week (2008-2009)`, y = `Unique Visits`, fill = Period))+
  geom_bar(stat = 'identity', width = 1) +
  theme(axis.text.x = element_text(angle = 90)) +
  labs(title = 'Unique Visits per week (2008-2009)') +
  scale_x_discrete(limits = all_periods$`Week (2008-2009)`)

ggplot(all_periods, aes(x = `Week (2008-2009)`, y = Revenue, fill = Period))+
  geom_bar(stat = 'identity', width = 1) +
  theme(axis.text.x = element_text(angle = 90)) +
  labs(title = 'Revenue per week (2008-2009)') +
  scale_x_discrete(limits = all_periods$`Week (2008-2009)`)

ggplot(all_periods, aes(x = `Week (2008-2009)`, y = Profit)) +
  geom_bar(stat = 'identity', width = 1, fill = 'blue') +
  theme(axis.text.x = element_text(angle = 90)) +
  labs(title = 'Profit per week (2008-2009)') +
  scale_x_discrete(limits = all_periods$`Week (2008-2009)`)

ggplot(all_periods, aes(x = `Week (2008-2009)`, y = `Lbs. Sold`, fill = Period))+
  geom_bar(stat = 'identity', width = 1) +
  theme(axis.text.x = element_text(angle = 90)) +
  labs(title = 'Lbs. Sold per week (2008-2009)') +
  scale_x_discrete(limits = all_periods$`Week (2008-2009)`)

ggplot(all_periods, aes(x = `Week (2008-2009)`, y = Revenue, group=1))+
  geom_point() +
  geom_smooth(method = 'lm', se = F) +
  geom_line() +
  theme(axis.text.x = element_text(angle = 90))+
  labs(title = 'Revenue per week (2008-2009)')+
  scale_x_discrete(limits = all_periods$`Week (2008-2009)`)

# Question 2
# Function to create a summary table
sum_meas <- function(column) {
  a <- mean(column)
  b <- median(column)
  c <- sd(column)
  d <- min(column)
  e <- max(column)
  return(c(a,b,c,d,e))
}

# Function creating a summary table based on period 
period_table  <- function(period){
  data.frame(Metrics = c('mean','median','st.dev','minimum','maximum'),
             Visits = sum_meas(period$Visits),
             Unique_Visits = sum_meas(period$`Unique Visits`),
             Revenue = sum_meas(period$Revenue),
             Profit = sum_meas(period$Profit),
             Lbs_sold = sum_meas(period$`Lbs. Sold`))
}

all_period_table  <- period_table(all_periods)
intital_table     <- period_table(init_main)
pre_period_table  <- period_table(pre_main)
promotion_table   <- period_table(prom_main)
post_period_table <- period_table(post_main)


# Question 3
# Calculation of mean values for different variables in different periods
mean_visits <- c(mean(init_main$Visits),
                 mean(pre_main$Visits),
                 mean(prom_main$Visits),
                 mean(post_main$Visits))

mean_unique <- c(mean(init_main$`Unique Visits`),
                 mean(pre_main$`Unique Visits`),
                 mean(prom_main$`Unique Visits`),
                 mean(post_main$`Unique Visits`))

mean_revenue <- c(mean(init_main$Revenue),
                  mean(pre_main$Revenue),
                  mean(prom_main$Revenue),
                  mean(post_main$Revenue))

mean_profit <-  c(mean(init_main$Profit),
                  mean(pre_main$Profit),
                  mean(prom_main$Profit),
                  mean(post_main$Profit))

mean_lbs <-  c(mean(init_main$`Lbs. Sold`),
                  mean(pre_main$`Lbs. Sold`),
                  mean(prom_main$`Lbs. Sold`),
                  mean(post_main$`Lbs. Sold`))

# DataFrame of mean per period per variable 
mean_table <- data.frame(Metrics = c('Initial', 'Pre-Promotion', 'Promotion', 'Post-Promotion'),
                         Visits = mean_visits,
                         Unique_Visits = mean_unique,
                         Revenue = mean_revenue,
                         Profit = mean_profit,
                         Lbs_sold = mean_lbs)

# DataFrame of mean per period with pecentage change
mean_table_percent <- data.frame(Metrics = c('Initial', 'Pre-Promotion', 'Promotion', 'Post-Promotion'),
                         Visits = mean_visits,
                         Percentage_Change = round(((mean_visits - lag(mean_visits))/lag(mean_visits)) * 100),
                         Revenue = mean_revenue,
                         Percentage_Change1 = round(((mean_revenue - lag(mean_revenue))/lag(mean_revenue)) * 100),
                         Lbs_sold = mean_lbs,
                         Percentage_Change2 = round(((mean_lbs - lag(mean_lbs))/lag(mean_lbs)) * 100))

# Question 5
# Checking correlation and relationship between financial variable: REVENUE and LBS. SOLD
rev_lbs_sold <- ggplot(all_periods, aes(x = `Lbs. Sold`, y = Revenue ))+
  geom_point()+
  geom_smooth(method = 'lm', color='orange')+
  ggtitle('Revenue vs Lbs. Sold for the entire period')

# Linear regression
lm_rev_lbs  <- lm(Revenue ~ `Lbs. Sold`, data = all_periods)
summary(lm_rev_lbs)
# Correlation coefficient
rev_lbs_cor <- cor(all_periods$`Lbs. Sold`, all_periods$Revenue)


# Question 6
# Checking correlation and relationship between financial variables and website data
vis_rev <- ggplot(all_periods, aes(x = Visits, y = Revenue))+
  geom_point()+
  geom_smooth(method = 'lm')

# Linear regression 
lm_vis_rev  <- lm(Revenue ~ Visits, data = all_periods)
summary(lm_vis_rev)
# Correlation coefficient
rev_vis_cor <- cor(all_periods$Visits,all_periods$Revenue)


# Question 8
# Lbs. sold over an extended period January 2, 2005 to July 19, 2010
lbs_sum_meas(pounds_sold$`Lbs. Sold`)

# Summary table with the same function as before
lbs_sold_table <- data.frame(Metrics = c('mean','median','st.dev','minimum','maximum'),
           Pounds_Sold = sum_meas(pounds_sold$`Lbs. Sold`))

# Histogram of distribution of Lbs. sold
hist_lbs_sold <- ggplot(pounds_sold, aes(`Lbs. Sold`))+
  geom_histogram(bins = 52, fill = 'navy')

lbs_len <- length(pounds_sold$`Lbs. Sold`)

mean_lbs <- mean(pounds_sold$`Lbs. Sold`)
sd_lbs   <- sd(pounds_sold$`Lbs. Sold`)

# Develop a table with theoretical and actual observations 1,2 and 3 standard deviations from the mean

# Calculate theoretical No. Obs in the Lbs sold sheet
theoretical_1 <- round(lbs_len * 0.68)
theoretical_2 <- round(lbs_len * 0.95)
theoretical_3 <- round(lbs_len * 0.99)

# Calculate the actual No. Obs in the Lbs sold
actual_1 <- length(pounds_sold$`Lbs. Sold`[which(pounds_sold$`Lbs. Sold` < mean_lbs + sd_lbs & 
                                                   pounds_sold$`Lbs. Sold` > mean_lbs - sd_lbs)])
actual_2 <- length(pounds_sold$`Lbs. Sold`[which(pounds_sold$`Lbs. Sold` < mean_lbs + 2*sd_lbs & 
                                                   pounds_sold$`Lbs. Sold` > mean_lbs - 2*sd_lbs)])
actual_3 <- length(pounds_sold$`Lbs. Sold`[which(pounds_sold$`Lbs. Sold` < mean_lbs + 3*sd_lbs & 
                                                   pounds_sold$`Lbs. Sold` > mean_lbs - 3*sd_lbs)])

theoretical_actual_table <- data.frame(Mean_Interval = c('mean ± 1 std. dev.','mean ± 2 std. dev.', 'mean ± 3 std. dev.'),
                                       Theoretical_percent_of_data = c('68','95','99'),
                                       Theoretical_No_obs = c(theoretical_1, theoretical_2, theoretical_3),
                                       Actual_No_Obs = c(actual_1, actual_2, actual_3))

# Calculate skewness and kurtosis of distribution of lbs sold
lbs_sold_skew <- skewness(pounds_sold$`Lbs. Sold`)
lbs_sold_kurt <- kurtosis(pounds_sold$`Lbs. Sold`)


#######################
# Additional coding for presentation
#######################

# Descriptive analysis of the variables produced by the website

# Traffic sources
all_traffic           <- as.data.frame(sheet_import(6, 'B7:C11'))
colnames(all_traffic) <- c('Traffic Source', 'Visits')

ggplot(all_traffic, aes(x = `Traffic Source`, y = Visits, fill = `Traffic Source`))+
  geom_bar(stat = 'identity')+
  scale_x_discrete(limits = all_traffic$`Traffic Source`)


# Top referring websites 
top_ref           <- as.data.frame(sheet_import(6, 'B14:C24'))
colnames(top_ref) <- c('Referring Sites', 'Visits')

ggplot(top_ref, aes(x = `Referring Sites`, y = Visits, fill = `Referring Sites`))+
  geom_bar(stat = 'identity')+
  coord_flip()+
  scale_x_discrete(limits = top_ref$`Referring Sites`)


# Search Engine 
search_eng           <- as.data.frame(sheet_import(6, 'B27:C37'))
colnames(search_eng) <- c('Search Engine', 'Visits')

ggplot(search_eng, aes(x = `Search Engine`, y = Visits, fill = `Search Engine`))+
  geom_bar(stat = 'identity')+
  scale_x_discrete(limits = search_eng$`Search Engine`)


# Geographic location
top_geo           <- as.data.frame(sheet_import(6, 'B40:C50'))
colnames(top_geo) <- c('Geographic Location', 'Visits')

color_list <- c('red', 'red', rep('grey', 8))

ggplot(top_geo, aes(x = `Geographic Location`, y = Visits)) +
  geom_bar(stat = 'identity', fill = color_list) +
  scale_x_discrete(limits = top_geo$`Geographic Location`) +
  coord_flip()


# bar chart of lbs sold mean per period
ggplot(mean_table, aes(x=Metrics, y= Lbs_sold))+
  geom_bar(stat = 'identity', fill='orange')+
  scale_x_discrete(limits = mean_table$Metrics)+
  ggtitle('Pounds sold per period')+
  xlab('')+
  ylab('Avg. lbs sold')

# Two plots on top of each other. Bar chart of REVENUE with line chart of VISITES over it
rev_vis <- plot_ly() %>%
  add_lines(x = ~all_periods, y = ~mean_visits, name = "Visitors", yaxis = "y2", color=I('blue')) %>%
  add_markers(x = ~all_periods, y = ~mean_visits, name = "Visitors", yaxis = "y2", color=I('blue'), size = 6) %>%
  add_bars(x = ~all_periods, y = ~mean_revenue, name = "Revenue", color=I('orange')) %>%
  layout(
    title = "Revenues VS Visits", yaxis2 = p,
    xaxis = list(title="Promotion periods"),
    yaxis= list(title="Revenue")
  )
p <- list(tickfont = list(color = "red"),
          overlaying = "y",
          side = "right",
          title = "Visitors")


# Relationships and Visuals between website data and financial data
# Revenue and Inquiries
rev_inq <- ggplot(all_periods, aes(x = Inquiries, y = Revenue))+
  geom_point()+
  geom_smooth(method = 'lm')

lm_rev_inq <- lm(Revenue ~ Inquiries, data = all_periods)
summary(lm_rev_inq)

# Revenue and Avg Time Spent
rev_avg_time <- ggplot(all_periods, aes(x = `Avg. Time on Site (secs.)`, y = Revenue))+
  geom_point()+
  geom_smooth(method = 'lm')

lm_rev_avg_time <- lm(Revenue ~ `Avg. Time on Site (secs.)`, data = all_periods)
summary(lm_rev_avg_time)

# Revenue and bounce %
rev_bounc <- ggplot(all_periods, aes(x = `Bounce Rate`, y = Revenue))+
  geom_point()+
  geom_smooth(method = 'lm')

lm_rev_bounc <- lm(Revenue ~ `Bounce Rate`, data = all_periods)
summary(lm_rev_bounc)

# Revenue and New visits
rev_new_vis <- ggplot(all_periods, aes(x = `% New Visits`, y = Revenue))+
  geom_point()+
  geom_smooth(method = 'lm')

lm_rev_new_vis <- lm(Revenue ~ `% New Visits`, data = all_periods)
summary(lm_rev_new_vis)




