#####################################################
### Created: 12/01/2018
### Author : Arsen Oz
### 
### Web Analytics at Quality Alloys, Inc.
#####################################################

install.packages("moments")
install.packages("corrplot")
install.packages("PerformanceAnalytics")
install.packages("psych")
install.packages("rsq")

library(readxl)
library(data.table)
library(ggplot2)
library(plotly)
library(moments)
library(corrplot)
library(PerformanceAnalytics)
library(psych)
library(lmtest)
library(rsq)

source("http://www.sthda.com/upload/rquery_cormat.r")

#####################################################
### 1. Importing Datasets
#####################################################

Weekly_Visits <- read_excel("C:/Users/earse/Desktop/MBAN/R/Project/Web Analytics Case Student Spreadsheet.xls", 
                            sheet = "Weekly Visits", range = "A5:H71")
Financials <- read_excel("C:/Users/earse/Desktop/MBAN/R/Project/Web Analytics Case Student Spreadsheet.xls", 
                         sheet = "Financials", range = "A5:E71")
Lbs_Sold <- read_excel("C:/Users/earse/Desktop/MBAN/R/Project/Web Analytics Case Student Spreadsheet.xls", 
                       sheet = "Lbs. Sold", range = "A5:B295")
Daily_Visits <- read_excel("C:/Users/earse/Desktop/MBAN/R/Project/Web Analytics Case Student Spreadsheet.xls", 
                           sheet = "Daily Visits", range = "A5:B467")
main_table <- read_excel("C:/Users/earse/Desktop/MBAN/R/Project/Web Analytics Case Student Spreadsheet.xls", 
                         sheet = "Final", range = "A1:L67")

#####################################################
### 1. Case Question Answers
#####################################################

#### Question 1

## Creating plots for different variables

# Unique Visits
chart1 <- ggplot(Weekly_Visits, aes(`Week (2008-2009)`, `Unique Visits`)) +
  geom_bar(stat="identity", width = 0.5, fill="tomato1") +
  labs(title="Unique Per Week") +
  theme(axis.text.x = element_text(angle=90, vjust=0.8))

# Revenue
chart2 <- ggplot(Financials, aes(`Week (2008-2009)`, `Revenue`))+ 
  geom_bar(stat="identity", width = 0.5, fill="tomato2") +
  labs(title="Revenue Per Week") +
  theme(axis.text.x = element_text(angle=90, vjust=0.8))

# Profit
chart3 <- ggplot(Financials, aes(`Week (2008-2009)`, `Profit`)) + 
  geom_bar(stat="identity", width = 0.5, fill="tomato2") +
  labs(title="Profit Per Week") +
  theme(axis.text.x = element_text(angle=90, vjust=0.8))

#Lbs. Sold
chart4 <- ggplot(Financials, aes(`Week (2008-2009)`, `Lbs. Sold`)) + 
  geom_bar(stat="identity", width = 0.5, fill="tomato2") +
  labs(title="Revenue Per Week") +
  theme(axis.text.x = element_text(angle=90, vjust=0.8))


#### Question 2

# Creating different tables for periods
Initial_Period <- main_table[1:14,]
Pre_Promotion_Period <- main_table[15:35,]
Promotion_Period <- main_table[36:52,]
Post_Promotion_Period <- main_table[53:66,]

# Function for calculating different metrics
metric_cal <- function (x) {
  y1 <- mean(x)
  y2 <- median(x)
  y3 <- sd(x)
  y4 <- min(x)
  y5 <- max(x)
  return(c(y1, y2, y3, y4, y5))
}


## Data tables with different metrics for each period

# Initial Period
Q2_Initial <- data.table(Metrics = c("Mean","Median","Std. Dev","Minimum","Maximum"),
                         Visits = metric_cal(Initial_Period$Visits),
                         Unique_Visits = metric_cal(Initial_Period$`Unique Visits`),
                         Revenue = metric_cal(Initial_Period$Revenue),
                         Profit = metric_cal(Initial_Period$Profit),
                         Lbs_Sold = metric_cal(Initial_Period$`Lbs. Sold`))

# Pre Promotion Period
Q2_Pre_Prom <- data.table(Metrics = c("Mean","Median","Std. Dev","Minimum","Maximum"),
                          Visits = metric_cal(Pre_Promotion_Period$Visits),
                          Unique_Visits = metric_cal(Pre_Promotion_Period$`Unique Visits`),
                          Revenue = metric_cal(Pre_Promotion_Period$Revenue),
                          Profit = metric_cal(Pre_Promotion_Period$Profit),
                          Lbs_Sold = metric_cal(Pre_Promotion_Period$`Lbs. Sold`))

# Promotion Period
Q2_Prom <- data.table(Metrics = c("Mean","Median","Std. Dev","Minimum","Maximum"),
                      Visits = metric_cal(Promotion_Period$Visits),
                      Unique_Visits = metric_cal(Promotion_Period$`Unique Visits`),
                      Revenue = metric_cal(Promotion_Period$Revenue),
                      Profit = metric_cal(Promotion_Period$Profit),
                      Lbs_Sold = metric_cal(Promotion_Period$`Lbs. Sold`))

# Post Promotion Period
Q2_Post_Prom <- data.table(Metrics = c("Mean","Median","Std. Dev","Minimum","Maximum"),
                           Visits = metric_cal(Post_Promotion_Period$Visits),
                           Unique_Visits = metric_cal(Post_Promotion_Period$`Unique Visits`),
                           Revenue = metric_cal(Post_Promotion_Period$Revenue),
                           Profit = metric_cal(Post_Promotion_Period$Profit),
                           Lbs_Sold = metric_cal(Post_Promotion_Period$`Lbs. Sold`))


#### Question 3

## Calculating mean values for different values in each period

# Visits
Q3_Visits <- c(mean(Initial_Period$Visits),
               mean(Pre_Promotion_Period$Visits),
               mean(Promotion_Period$Visits),
               mean(Post_Promotion_Period$Visits))

# Unique Visits
Q3_Unique_Visits <- c(mean(Initial_Period$`Unique Visits`),
                      mean(Pre_Promotion_Period$`Unique Visits`),
                      mean(Promotion_Period$`Unique Visits`),
                      mean(Post_Promotion_Period$`Unique Visits`))

# Revenue
Q3_Revenue <- c(mean(Initial_Period$Revenue),
                mean(Pre_Promotion_Period$Revenue),
                mean(Promotion_Period$Revenue),
                mean(Post_Promotion_Period$Revenue))

# Profit
Q3_Profit <- c(mean(Initial_Period$Profit),
               mean(Pre_Promotion_Period$Profit),
               mean(Promotion_Period$Profit),
               mean(Post_Promotion_Period$Profit))

# Lbs. Sold
Q3_Lbs_Sold <- c(mean(Initial_Period$`Lbs. Sold`),
                 mean(Pre_Promotion_Period$`Lbs. Sold`),
                 mean(Promotion_Period$`Lbs. Sold`),
                 mean(Post_Promotion_Period$`Lbs. Sold`))


# Data Table for comparing each mean value for different periods
Q3_table <- data.table(Metrics = c("Initial", "Pre-Promo", "Promotion", "Post-Promo"),
                       Visits = Q3_Visits,
                       Unique_Visits = Q3_Unique_Visits,
                       Revenue = Q3_Revenue,
                       Profit = Q3_Profit,
                       Lbs_Sold = Q3_Lbs_Sold)


#### Question 4

# Plot for Lbs. Sold vs Revenue
Lbs_Rev_plot <- ggplot(Financials, aes(x=Financials$`Lbs. Sold`, y=Financials$Revenue)) +
  geom_jitter() +
  ggtitle("Revenue vs Lbs Sold") + 
  xlab("Lbs Sold") + 
  ylab("Revenue")

# Correlation between Lbs.Sold and Revenue
Lbs_Rev_cor <- cor(Financials$`Lbs. Sold`,Financials$Revenue)


#### Question 5

# Plot for Visits vs Revenue
Visit_Rev_plot <-ggplot(Financials, aes(x=Weekly_Visits$Visits, y=Financials$Revenue)) + 
  geom_jitter() + 
  ggtitle("Revenue vs Visits") + 
  xlab("Visits") + 
  ylab("Revenue")

# Correlation between Visits and Revenue
Visit_Rev_cor <- cor(Weekly_Visits$Visits,Financials$Revenue)


#### Question 6

# Using metric_cal function to calculate different metrics for Lbs. Sold

Lbs_sold_metrics <- metric_cal(Lbs_Sold$`Lbs. Sold`)

# Basic histogram for Lbs Sold
lbs_histogram <- hist(Lbs_Sold$`Lbs. Sold`)
lbs_count <- length(Lbs_Sold$`Lbs. Sold`)

# Theoretical number of values that are between different standard deviation ranges
theory_1sd <- round(lbs_count * 0.68)
theory_2sd <- round(lbs_count * 0.95)
theory_3sd <- round(lbs_count * 0.99)

lbs_std <- sd(Lbs_Sold$`Lbs. Sold`)

# Actual number of values that are between different standard deviation ranges
actual_1sd <- length(Lbs_Sold$`Lbs. Sold`[which(Lbs_Sold$`Lbs. Sold` < lbs_mean + lbs_std
                                                & Lbs_Sold$`Lbs. Sold` > lbs_mean - lbs_std)])
actual_2sd <- length(Lbs_Sold$`Lbs. Sold`[which(Lbs_Sold$`Lbs. Sold` < lbs_mean + 2*lbs_std
                                                & Lbs_Sold$`Lbs. Sold` > lbs_mean - 2*lbs_std)])
actual_3sd <- length(Lbs_Sold$`Lbs. Sold`[which(Lbs_Sold$`Lbs. Sold` < lbs_mean + 3*lbs_std
                                                & Lbs_Sold$`Lbs. Sold` > lbs_mean - 3*lbs_std)])

# Comparison table for Theoretical vs. Actual
Q6_1_table <- data.table(interval = c("mean ± 1 std. dev","mean ± 2 std. dev","mean ± 3 std. dev"),
                        Theoretical_Perc_of_Data = c("68", "95","99"),
                        Theoretical_No_Obs = c(theory_1sd, theory_2sd, theory_3sd),
                        Actual_No_Obs = c(actual_1sd, actual_2sd, actual_3sd))


# Actual number of values that are greater/smaller than different standard deviation ranges
actual_1sd_up <- length(Lbs_Sold$`Lbs. Sold`[which(Lbs_Sold$`Lbs. Sold` < lbs_mean + lbs_std
                                                   & Lbs_Sold$`Lbs. Sold` > lbs_mean)])
actual_1sd_low <- length(Lbs_Sold$`Lbs. Sold`[which(Lbs_Sold$`Lbs. Sold` > lbs_mean - lbs_std
                                                    & Lbs_Sold$`Lbs. Sold` < lbs_mean)])
actual_2sd_up <- length(Lbs_Sold$`Lbs. Sold`[which(Lbs_Sold$`Lbs. Sold` < lbs_mean + 2*lbs_std
                                                   & Lbs_Sold$`Lbs. Sold` > lbs_mean + lbs_std)])
actual_2sd_low <- length(Lbs_Sold$`Lbs. Sold`[which(Lbs_Sold$`Lbs. Sold` > lbs_mean - 2*lbs_std
                                                    & Lbs_Sold$`Lbs. Sold` < lbs_mean - lbs_std)])
actual_3sd_up <- length(Lbs_Sold$`Lbs. Sold`[which(Lbs_Sold$`Lbs. Sold` < lbs_mean + 3*lbs_std
                                                   & Lbs_Sold$`Lbs. Sold` > lbs_mean + 2*lbs_std)])
actual_3sd_low <- length(Lbs_Sold$`Lbs. Sold`[which(Lbs_Sold$`Lbs. Sold` > lbs_mean - 3*lbs_std
                                                    & Lbs_Sold$`Lbs. Sold` < lbs_mean - 2*lbs_std)])


# Data table for the values above
Q6_2_table <- data.table(interval = c("mean + 1 std. dev","mean - 1 std. dev",
                                      "mean + 2 std. dev","mean - 2 std. dev",
                                      "mean + 3 std. dev","mean - 3 std. dev"),
                        Theoretical_Perc_of_Data = c("34","34", "13.5","13.5","2","2"),
                        Theoretical_No_Obs = c(theory_1sd/2, theory_1sd/2,
                                               (theory_2sd-theory_1sd)/2, (theory_2sd-theory_1sd)/2,
                                               (theory_3sd-theory_2sd)/2, (theory_3sd-theory_2sd)/2),
                        Actual_No_Obs = c(actual_1sd_up, actual_1sd_low,
                                          actual_2sd_up, actual_2sd_low,
                                          actual_3sd_up, actual_3sd_low))

# Checking for skewness and kurtosis

skewness(Lbs_Sold$`Lbs. Sold`) 
kurtosis(Lbs_Sold$`Lbs. Sold`)


#####################################################
### 2. Correlation Matrices and Heatmaps
#####################################################

# Main Correlation Matrix
correl_main <- round(cor(main_table[,c(2,3,4,5,6,7,8,9,10,11,12)]),2)

# Visuals for correlation
rquery.cormat(main_table[,c(2,3,4,5,6,7,8,9,10,11,12)], type = "full")

# Correlation matrices for each period
par(mfrow=c(2,2))
rquery.cormat(Initial_Period[,c(2,3,4,5,6,7,8,9,10,11,12)], type = "full")
rquery.cormat(Pre_Promotion_Period[,c(2,3,4,5,6,7,8,9,10,11,12)], type = "full")
rquery.cormat(Promotion_Period[,c(2,3,4,5,6,7,8,9,10,11,12)], type = "full")
rquery.cormat(Post_Promotion_Period[,c(2,3,4,5,6,7,8,9,10,11,12)], type = "full")
dev.off()

# Additional visuals for correlation
chart.Correlation(main_table[,c(2,3,4,5,6,7,8,9,10,11,12)], histogram=TRUE, pch=19)
chart.Correlation(Initial_Period[,c(2,3,4,5,6,7,8,9,10,11,12)], histogram=TRUE, pch=19)
chart.Correlation(Pre_Promotion_Period[,c(2,3,4,5,6,7,8,9,10,11,12)], histogram=TRUE, pch=19)
chart.Correlation(Promotion_Period[,c(2,3,4,5,6,7,8,9,10,11,12)], histogram=TRUE, pch=19)
chart.Correlation(Post_Promotion_Period[,c(2,3,4,5,6,7,8,9,10,11,12)], histogram=TRUE, pch=19)

pairs.panels(main_table[,c(2,3,4,5,6,7,8,9,10,11,12)])

# Heatmaps for different periods
heatmap(x = cor(main_table[,c(2,3,4,5,6,7,8,9,10,11,12)]), symm = TRUE)
heatmap(x = cor(Initial_Period[,c(2,3,4,5,6,7,8,9,10,11,12)]), symm = TRUE)
heatmap(x = cor(Pre_Promotion_Period[,c(2,3,4,5,6,7,8,9,10,11,12)]), symm = TRUE)
heatmap(x = cor(Promotion_Period[,c(2,3,4,5,6,7,8,9,10,11,12)]), symm = TRUE)
heatmap(x = cor(Post_Promotion_Period[,c(2,3,4,5,6,7,8,9,10,11,12)]), symm = TRUE)


#####################################################
### 3. Additional Visuals
#####################################################

### Each bar in the graphsb below represents 5 weeks period. ###


## Profits per period
`5_weeks_prof` <- c(sum(main_table$Profit[1:5]),
                    sum(main_table$Profit[6:10]),
                    sum(main_table$Profit[11:15]),
                    sum(main_table$Profit[16:20]),
                    sum(main_table$Profit[21:25]),
                    sum(main_table$Profit[26:30]),
                    sum(main_table$Profit[31:35]),
                    sum(main_table$Profit[36:40]),
                    sum(main_table$Profit[41:45]),
                    sum(main_table$Profit[46:50]),
                    sum(main_table$Profit[51:55]),
                    sum(main_table$Profit[56:60]),
                    sum(main_table$Profit[61:65]))

prof_table <- data.table(`5_weeks` = seq(1,13,1), Profit = `5_weeks_prof`, Group = c(rep("Initial",3),
                                                                                     rep("Pre",4),
                                                                                     rep("Promo",3),
                                                                                     rep("Post",3)))

ggplot(data = prof_table, aes(x=prof_table$`5_weeks`, y=prof_table$Profit/1000, fill = Group))  + 
  geom_bar(stat="identity", colour="black") + 
  scale_x_discrete("5 Weeks Period") + 
  scale_y_continuous("Profit(in thousand)") + 
  scale_fill_discrete(breaks=c("Initial","Pre","Promo", "Post")) +
  ggtitle(label = "Profits") +
  theme(text = element_text(size=30),
        legend.title = element_text(size=30),
        legend.text = element_text(size=30),
        legend.position = "bottom")


## Unique Visits per period

`5_weeks_uni_v` <- c(sum(main_table$`Unique Visits`[1:5]),
                     sum(main_table$`Unique Visits`[6:10]),
                     sum(main_table$`Unique Visits`[11:15]),
                     sum(main_table$`Unique Visits`[16:20]),
                     sum(main_table$`Unique Visits`[21:25]),
                     sum(main_table$`Unique Visits`[26:30]),
                     sum(main_table$`Unique Visits`[31:35]),
                     sum(main_table$`Unique Visits`[36:40]),
                     sum(main_table$`Unique Visits`[41:45]),
                     sum(main_table$`Unique Visits`[46:50]),
                     sum(main_table$`Unique Visits`[51:55]),
                     sum(main_table$`Unique Visits`[56:60]),
                     sum(main_table$`Unique Visits`[61:65]))

Uniq_v_table <- data.table(`5_weeks` = seq(1,13,1), Unique_Visits = `5_weeks_uni_v`, Group = c(rep("Initial",3),
                                                                                               rep("Pre",4),
                                                                                               rep("Promo",3),
                                                                                               rep("Post",3)))

ggplot(data = Uniq_v_table, aes(x=Uniq_v_table$`5_weeks`, y=Uniq_v_table$Unique_Visits, fill = Group))  + 
  geom_bar(stat="identity", colour="black") + 
  scale_x_discrete("5 Weeks Period") + 
  scale_y_continuous("Unique Visits") +
  scale_fill_discrete(breaks=c("Initial","Pre","Promo", "Post")) +
  ggtitle(label = "Unique Visits") +
  theme(text = element_text(size=30),
        legend.title = element_text(size=30),
        legend.text = element_text(size=30),
        legend.position = "bottom")


## Bounce Rate per period

`5_weeks_bounce` <- c(sum(main_table$`Bounce Rate`[1:5])/5,
                      sum(main_table$`Bounce Rate`[6:10])/5,
                      sum(main_table$`Bounce Rate`[11:15])/5,
                      sum(main_table$`Bounce Rate`[16:20])/5,
                      sum(main_table$`Bounce Rate`[21:25])/5,
                      sum(main_table$`Bounce Rate`[26:30])/5,
                      sum(main_table$`Bounce Rate`[31:35])/5,
                      sum(main_table$`Bounce Rate`[36:40])/5,
                      sum(main_table$`Bounce Rate`[41:45])/5,
                      sum(main_table$`Bounce Rate`[46:50])/5,
                      sum(main_table$`Bounce Rate`[51:55])/5,
                      sum(main_table$`Bounce Rate`[56:60])/5,
                      sum(main_table$`Bounce Rate`[61:65]/5))

bounce_table <- data.table(`5_weeks` = seq(1,13,1), Bounce_Rate = `5_weeks_bounce`, Group = c(rep("Initial",3),
                                                                                              rep("Pre",4),
                                                                                              rep("Promo",3),
                                                                                              rep("Post",3)))

ggplot(data = bounce_table, aes(x=bounce_table$`5_weeks`, y=bounce_table$Bounce_Rate, fill = Group))  + 
  geom_bar(stat="identity", colour="black") + 
  scale_x_discrete("5 Weeks Period") + 
  scale_y_continuous("Bounce Rate") +
  scale_fill_discrete(breaks=c("Initial","Pre","Promo", "Post")) +
  ggtitle(label = "Bounce Rate") +
  theme(text = element_text(size=30),
        legend.title = element_text(size=30),
        legend.text = element_text(size=30),
        legend.position = "bottom")

## Inquiries per period

`5_weeks_inquiries` <- c(sum(main_table$Inquiries[1:5]),
                         sum(main_table$Inquiries[6:10]),
                         sum(main_table$Inquiries[11:15]),
                         sum(main_table$Inquiries[16:20]),
                         sum(main_table$Inquiries[21:25]),
                         sum(main_table$Inquiries[26:30]),
                         sum(main_table$Inquiries[31:35]),
                         sum(main_table$Inquiries[36:40]),
                         sum(main_table$Inquiries[41:45]),
                         sum(main_table$Inquiries[46:50]),
                         sum(main_table$Inquiries[51:55]),
                         sum(main_table$Inquiries[56:60]),
                         sum(main_table$Inquiries[61:65]))

Inquiries_table <- data.table(`5_weeks` = seq(1,13,1), Inquiries = `5_weeks_inquiries`, Group = c(rep("Initial",3),
                                                                                                  rep("Pre",4),
                                                                                                  rep("Promo",3),
                                                                                                  rep("Post",3)))

ggplot(data = Inquiries_table, aes(x=Inquiries_table$`5_weeks`, y= Inquiries_table$Inquiries, fill = Group))  + 
  geom_bar(stat="identity", colour="black") + 
  scale_x_discrete("5 Weeks Period") + 
  scale_y_continuous("Inquiries") +
  scale_fill_discrete(breaks=c("Initial","Pre","Promo", "Post")) +
  ggtitle(label = "Inquiries") +
  theme(text = element_text(size=30),
        legend.title = element_text(size=30),
        legend.text = element_text(size=30),
        legend.position = "bottom")


# Basic Linear Model Representation for Profit
reg_table <- data.table(`5_weeks` = seq(1,7,1), Profit = `5_weeks_prof`[1:7])

ggplot(reg_table, aes(x = `5_weeks`, y = Profit/1000)) + 
  stat_smooth(method = "glm", formula = y ~ x,alpha = 0.2, size = 2, fill = "tomato", colour = "red") +
  geom_point(size = 3) +
  scale_x_continuous("5 Weeks Period", breaks=c(1:7), labels=c(1:7),limits=c(1,7)) + 
  scale_y_continuous("Profit (in thousands)") +
  ggtitle("Profit Linear Model") +
  theme(text = element_text(size = 22),
        legend.title = element_text(size=30))


