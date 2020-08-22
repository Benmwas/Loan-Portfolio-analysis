library(readxl)
library(dplyr)
library(tidyr)
library(lubridate)
library(ggplot2)
library(scales)
install.packages("highcharter")
install.packages("ggpubr")


data <- read_excel("Data/Bidhaa_data.xlsx")
nrow(data)
str(data)
as.integer(64.000000+0.5) 

summary <- data %>% 
  filter(Cycle > 1) %>% 
 select(Product,`Disbursed Date`,Cycle)

Order[order(as.Date(Order$`Disbursed Date`, format "%Y-%m-%d")), ] 
  
hist(summary$Cycle)
 

plot_data <- summary %>% 
  group_by(Product, Cycle) %>% 
  summarise(count = n())

plot_data_theme<-  theme(legend.position = "right", 
                                      legend.direction = "vertical",
                                      #legend.title = element_blank(),
                                      plot.title = element_text( size = rel(1.5), hjust = 0.5),
                                      plot.subtitle = element_text(size = rel(1.5), hjust = 0.5),
                                      #axis.text = element_text( size = rel(1.5)),
                                      axis.text.x = element_text(size =rel(1.0),angle = 0),
                                      axis.text.y = element_text(size =rel(1.0),angle = 0),
                                      axis.title = element_text( size = rel(1.55)),
                                      axis.line.x = element_line(size = 1.5, colour = "#c94a38"),
                                      panel.background = element_rect(fill = NA))
writexl::write_xlsx(plot_data, "Q1_data.xlsx")
#plot for all products
ggplot(plot_data, aes(x=Cycle, y = count, fill = Product)) +
  geom_bar(stat = "identity", position = "dodge")+
  theme(panel.grid.minor.x=element_blank(),
        panel.grid.major.x=element_blank()) +theme(panel.background = element_blank()) +
  theme(panel.grid.major = element_line(colour = "light grey"))+
  theme(axis.line.x = element_line(size = 1.5, colour = "#c94a38"))+
  theme(legend.position = "right")+
  theme(plot.title = element_text(hjust = 0.5, colour = "dark green"))+
  theme(axis.text.x = element_text(face = "bold", size=8, angle=0, color="black"))+
  labs(title = "Distribution of all products sold per cycle", x = "Cycles", y = "No. of products bought")

#Filter top products
Top_pdct <- plot_data %>% 
  filter(Product=="BORA" | Product=="LPG2" | Product=="CNVS" | Product=="P400" | Product=="BOOM")

head(Top_pdct)

writexl::write_xlsx(Top_pdct, "Q1a_data.xlsx")

ggplot(Top_pdct, aes(x=Cycle, y = count, fill = Product)) +
  geom_bar(stat = "identity", position = "dodge")+
  scale_fill_manual(values = c("#377DA9","#E79435", "#228B22", "#FF9999", "#BB523A"))+
  labs(title = "Distribution of Top Products sold per cycle", x = "Cycles", y = "No. of products bought")+
  theme(panel.grid.minor.x=element_blank(),
        panel.grid.major.x=element_blank()) +theme(panel.background = element_blank()) +
  theme(panel.grid.major = element_line(colour = "light grey"))+
  theme(axis.line.x = element_line(size = 1.5, colour = "#c94a38"))+
  theme(legend.position = "bottom", legend.direction = "horizontal")+
  theme(plot.title = element_text(hjust = 0.5, colour = "dark green"))+
  theme(axis.text.x = element_text(face = "bold", size=8, angle=0, color="black"))

#Correlation/pattern between age and  products purchased
(Sys.Date()- as.Date("1990-01-29"))/365.25
 
Cor_data <- data %>% 
  select(DOB, Product)  

which(is.na(Cor_data$DOB))

Cor_data$DOB <- as.Date(Cor_data$DOB, format = "%Y-%m-%d")
Cor_clean <- Cor_data[-c(6079,6307,15336,17364,17378,18088,21726,22837,23095,23192,23457,23989,27506,27626), ]

#Add years column
Cor_clean$Age_days <- Sys.Date()- Cor_clean$DOB
Cor_clean$Age <- Cor_clean$Age_days/365
as.numeric(Cor_clean$Age)
write_xlsx(Cor_clean, "age.xlsx")
Age <- read_excel("Data/age.xlsx")
Age[,4] <- trunc(Age[,4])

hist(Age$Age)

Age_plot <- Age %>% 
  group_by(Age, Product) %>% 
  summarise(count = n())


 #Percentage of product purchace per age
Percentage <- group_by(Age_plot, Age) %>% 
  mutate(Percent = (count/sum(count))*100)

Percentage[,4] <- round(Percentage[,4], digits = 2)

head(Percentage)
writexl::write_xlsx(Percentage, "Q2_data.xlsx")

ggplot(Percentage, aes(x = Age, y = Percent, color = Product)) + geom_point()+
  labs(title = "Percentage distribution all demanded Products by Age", x = "Age in Years", y = "Percentage of Products bought(%)")+
  theme(plot.title = element_text(hjust = 0.5, colour = "dark green"))


ggplot(Percentage, aes(x = Age, y = Percent, group = Product)) + geom_line(aes(color = Product)) +
  geom_point(aes(color = Product))

#Top Percentage
Top_Percent <- Percentage %>% 
  filter(Product=="CNVS" | Product=="LPG2" | Product=="LPG3" | Product=="BORA" | Product=="P400")

ggplot(Top_Percent, aes(x = Age, y = Percent, group = Product)) + geom_line(aes(color = Product))+
  geom_hline(yintercept = 12.5,color="black", linetype = "dashed")+
  geom_vline(xintercept = 50,color="red", linetype = "dashed")+
  theme(panel.grid.minor.x=element_blank(),
       panel.grid.major.x=element_blank()) +theme(panel.background = element_blank()) +
  theme(panel.grid.major = element_line(colour = "light grey"))+
  theme(axis.line.x = element_line(size = 1.5, colour = "#c94a38"))+
  theme(legend.position = "top", legend.direction = "horizontal")+
  theme(plot.title = element_text(hjust = 0.5, colour = "dark green"))+
  theme(axis.text.x = element_text(face = "bold", size=8, angle=0, color="black"))+
  labs(title = "Percentage distribution of Top demanded Products by Age", x = "Age in Years", y = "Percentage of Products bought(%)")

#rep behaviour measured with TRP and choice of product boughtayment?
#TRP = {# timely installment/SUM(installments due)}*100

TRP_data <- data %>% 
  select(`Disbursed Date`,Product, TRP)
which(is.na(TRP_clean))

writexl::write_xlsx(TRP_data, "TRP_clean.xlsx")

#Wanted to remove NA rows
 TRP_clean <- TRP_data[-c(31302,31303,31304,31305,31306,31307,31308,31309,31310,31311,31312,31313,31314,31315,31316,31317,31318,31319,
                            313203,13213,13223,13233,13243,13253,13263,1327,31328,31329,31330,31331,31332,31333,31334,31335,31336,31337,
                            313383,1339,313403,13413,13423,13433,1344,31345,31346,31347,31348,31349,31356,31357,31358,313593,1360,31361,
                            31362,31363,31364,31365,31366,31367,31368,31375,31376,31377,31378,31379,31380,31381,31382,31383,31384,31385,
                            31386,31387,313883,1389,31390,31391,31392,31393,31394,31395,31396,31397,31398,31399,31400,31401,31402,31403,
                            31404,314053,1406,314073,1408,31409,31410,31411,31412,31413,31414,31415,31416,31417,31418,31419,31420,31421,
                            314223,1423,31424,31425,31426,31427,31428,31429,31430,31431,31432,31433,31434,31435,31436,31437,31438,31439,
                            314403,14413,1442,314433,1444,314453,1446,31447,314483,1449,31450,31451,31452,31453,31454,31455,31456,31457,
                            314583,14593,1460,314613,14623,14633,14643,14653,14663,14673,14683,1469,31470,31471,314723,14733,14743,1475,
                            314763,1477,31478,31479,31480,314813,1482,314833,14843,14853,1486,314873,1488,31489,314903,1491,314923,1493,
                            31494,31495,314963,14973,14983,14993,15003,15013,15023,15033,15043,1505,315063,15073,15083,15093,1510,31511,
                            315123,1513,315143,1515,31516), ]

 TRP_clean <- read_excel("Data/TRP_clean.xlsx")

 TRP_summary <-  group_by(TRP_clean, Product) %>%   
  summarise(mean(TRP)*100)
 names(TRP_summary) <- c("Product", "Average_TRP")
 TRP_summary[,2] <- round(TRP_summary[,2], digits = 2)
  
 head(TRP_summary)
 writexl::write_xlsx(TRP_summary, "Q3_data")
#Plot for repayment pattern
 Plot_TRP <- arrange(TRP_summary, Average_TRP)
 
 ggplot(TRP_summary, aes(x = reorder(Product, -Average_TRP), y = Average_TRP)) + 
   geom_bar(stat = "identity", position = "dodge", color='skyblue',fill='steel blue')+
   theme(panel.grid.minor.x=element_blank(),
         panel.grid.major.x=element_blank()) +theme(panel.background = element_blank()) +
   theme(panel.grid.major = element_line(colour = "light grey"))+
   theme(axis.line.x = element_line(size = 1.5, colour = "#c94a38"))+
   theme(plot.title = element_text(hjust = 0.5, colour = "dark green"))+
   theme(axis.text.x = element_text(face = "bold", size=8, angle=60, color="black"))+
   geom_text(aes(label = Average_TRP),vjust = -0.25, size = 2)+
   labs(title = "Average TRP for All Products bought\n (Repayment pattern per product)", x = "Products", y = "Mean TRP (%)")

 #any correlation/pattern between gender and choice of product bought?
 
 Gender_data <- data %>% 
   group_by(Gender, Product) %>% 
   summarize(count = n())
 
 Gender_summary <- group_by(Gender_data, Gender) %>% 
   mutate(Percentage_gender =  (count/sum(count))*100)
 Gender_summary[,4] <- round(Gender_summary[,4], digits = 1)
 
 #for stacked bar charts need to rearrange the data for stacking
 #Calculate the cumulative sum of Percentage_gender for each Gender category. Used as the y coordinates of labels. 
 #To put the label in the middle of the bars, we’ll use cumsum(Percentage_gender) - 0.5 * Percentage_gender
#lab_ypos-----Label_y position
 
 Gender_plot <- Gender_summary %>% 
   group_by(Product) %>% 
   arrange(Product, desc(Gender)) %>% 
   mutate(lab_ypos = cumsum(Percentage_gender)- 0.5 * Percentage_gender)
 
writexl::write_xlsx(Gender_plot, "Q4_data")

 ggplot(Gender_plot, aes(fill=Gender, y= Percentage_gender, x= Product)) + 
   geom_bar(position="stack", stat="identity")+
   geom_text(aes(y = lab_ypos, label = paste0(Percentage_gender,"%"), group = Gender), color = "black", size =2.5)+
   theme(panel.grid.minor.x=element_blank(),
         panel.grid.major.x=element_blank()) +theme(panel.background = element_blank()) +
   theme(panel.grid.major = element_line(colour = "light grey"))+
   theme(axis.line.x = element_line(size = 1.5, colour = "#c94a38"))+
   theme(plot.title = element_text(hjust = 0.5, colour = "dark green"))+
   theme(axis.text.x = element_text(face = "bold", size=8, angle=60, color="black"))+
   theme(legend.position = "top", legend.direction = "horizontal")+
   labs(title = "Percentage of gender per product bought", x = "Products", y = "Gender Percentage (%)")
 
 #how long do clients really take to fully pay their loans, what is the distribution? Any difference between products?
   
   Repayments <- data %>% 
     select(Product,`Disbursed Date`,Status,`Final Payment Date`, TRP) %>% 
     filter(Status == "Closed" & TRP == 1.000) 
   
 Repayments$`Disbursed Date` <- as.Date(Repayments$`Disbursed Date`, format ="%Y-%m-%d")  
 Repayments$`Final Payment Date` <- as.Date(Repayments$`Final Payment Date`, format = "%Y-%m-%d") 
 Repayments <-  mutate(Repayments, Days = as.numeric(`Final Payment Date` - `Disbursed Date`))
 
 hist(Repayments$Days)
 writexl::write_xlsx(Repayments, "Q5_data.xlsx")
 
 ggplot(Repayments, aes(x=Days)) +
   geom_histogram(color="dark blue", fill= "light blue",position="identity", alpha=0.5)+
   geom_density(alpha=0.6)+theme_classic()+
   labs(title = "Distribution of repayment days\n (Fully repaid loans)~ TRP=100%, Status=closed")+
   theme(plot.title = element_text(hjust = 0.5, colour = "dark green"))
 #distribution of payment days with products (mean repayment days)
 
 mean_Repayment <- group_by(Repayments, Product) %>% 
   mutate(Mean_days = mean(Days))
  mean_Repayment[, 7] <- trunc(mean_Repayment[, 7])
  
  writexl::write_xlsx(mean_Repayment, "Q5a_data.xlsx")
  
  ggplot(mean_Repayment, aes(x=Product, y=Mean_days))+ 
    geom_bar(position = "dodge",stat = "identity", fill ="light green", color = "black" )+
    theme(panel.grid.minor.x=element_blank(),
          panel.grid.major.x=element_blank()) +theme(panel.background = element_blank()) +
    theme(panel.grid.major = element_line(colour = "light grey"))+
    theme(axis.line.x = element_line(size = 1.5, colour = "#c94a38"))+
    theme(plot.title = element_text(hjust = 0.5, colour = "brown", face = "bold"))+
    theme(axis.text.x = element_text(face = "bold", size=8, angle=60, color="black"))+
    theme(legend.position = "right", legend.direction = "verticle")+
    geom_text(aes(label = Mean_days),vjust = -0.25, size = 3)+
    labs(title = "Distribution of Timely full loan Repayment days per Product\n (Days = Average TRP days per product)", x = "Products", y = "TRP days(mean)")
  
 #Historical demand and repayment behavior----which product to push more and why?
  
  Demand <- data %>% 
    group_by(Product,Status) %>% 
    summarise(count = n()/20)
  
  Demand_plot <- group_by(Demand, Product) %>% 
    mutate(Percentage = (count/sum(count))*100)
  
  Demand_plot[,4] <- round(Demand_plot[,4], digits = 1)
  
  Demand_plot_plotII <- Demand_plot %>% 
    group_by(Product) %>% 
    arrange(Product, desc(Status)) %>% 
    mutate(lab_ypos = cumsum(Percentage)- 0.5 * Percentage)
  options(scipen = 999)
  #stack for demand and status
  
 
  
  ggplot(Demand_plot_plotII, aes(fill=Status, y= Percentage, x= Product)) + 
    geom_bar(position="stack", stat="identity")+
    geom_line(aes(x = Product, y = count),stat = "identity",size = 1, color="maroon", group = 1)+
    geom_text(aes(y = lab_ypos, label = paste0(Percentage,"%"), group =Status), color = "black", size =2.5)+
    theme(panel.grid.minor.x=element_blank(),
          panel.grid.major.x=element_blank()) +theme(panel.background = element_blank()) +
    theme(panel.grid.major = element_line(colour = "light grey"))+
    theme(axis.line.x = element_line(size = 1.5, colour = "#c94a38"))+
    theme(plot.title = element_text(hjust = 0.5, colour = "dark green"))+
    theme(axis.text.x = element_text(face = "bold", size=8, angle=60, color="black"))+
    theme(legend.position = "top", legend.direction = "horizontal")+
    labs(title = "Products demand and Percentage loan status\n (Maroon line = Products demand)~ scale factor:n/20", x = "Products", y = "Percentage Loan Status (%)")

  ##Additional analysis
  
 #Time series Product demand
  data$`Disbursed Date` <- as.Date(data$`Disbursed Date`, format = "%Y-%m-%d")
  
  Time_series <- data %>% 
    group_by(Product, `Disbursed Date`) %>% 
  summarise(count = n())
  
  
  Time_series_plot <- Time_series %>% 
    group_by(Product, Month) %>% 
    mutate(Total = sum(count))
  
Time_series_plot$Month <- format(as.Date(Time_series_plot$`Disbursed Date`), "%b - %Y")

Time_series_plot$Month <- factor(Time_series_plot$Month, 
                            levels = c("Jan - 2019", "Feb - 2019", "Mar - 2019", "Apr - 2019", "May - 2019", "Jun - 2019",
                                       "Jul - 2019", "Aug - 2019", "Sep - 2019", "Oct - 2019", "Nov - 2019","Dec - 2019",
                                       "Jan - 2020", "Feb - 2020", "Mar - 2020"))
 
ggplot(Time_series_plot, aes(x= Month, y=Total, group=Product)) +
  geom_line(aes(color=Product))+
  geom_point(aes(color=Product))+ ylab("Total Product demanded")+
  xlab("Month-Year")+theme(axis.text.x = element_text( size=8, angle=60, color="black"))+
  ggtitle("Time series analysis for Products demanded")+
  theme(plot.title = element_text( colour = "dark green"))+
  scale_y_continuous(labels = label_comma(big.mark = ","))+
  theme(legend.position="right")+
  theme(axis.line.x = element_line(size = 1.5, colour = "#c94a38"))+
  theme(plot.title = element_text(hjust = 0.5, colour = "dark green"))+
  theme(axis.text.x = element_text(face = "bold", size=8, angle=60, color="black"))
  
#Top Products time series analysis

Time_series_tp <- Time_series_plot %>% 
  filter(Product == "LPG2" | Product == "LPG3" | Product == "BORA" | Product =="CNVS" | Product == "JIKO" | Product == "P400")
  
ggplot(Time_series_tp, aes(x= Month, y=Total, group=Product)) +
  geom_line(aes(color=Product))+
  geom_point(aes(color=Product))+ ylab("Total Product demanded")+
  xlab("Month-Year")+theme(axis.text.x = element_text( size=8, angle=60, color="black"))+
  ggtitle("Time series analysis for Top Products demanded per Month")+
  theme(plot.title = element_text( colour = "dark green"))+
  scale_y_continuous(labels = label_comma(big.mark = ","))+
  theme(legend.position="top")+
  theme(axis.line.x = element_line(size = 1.5, colour = "#c94a38"))+
  theme(plot.title = element_text(hjust = 0.5, colour = "dark green"))+
  theme(axis.text.x = element_text(face = "bold", size=8, angle=60, color="black"))

#Repayment day of the week

Week <- data %>% 
  group_by(Product, `Final Payment Date`) %>% 
  summarise(count = n())
 Week$`Final Payment Date` <-  as.Date(Week$`Final Payment Date`, format = "%Y-%m-%d")
 
 Week$day <- format(as.Date(Week$`Final Payment Date`), "%A")
 
 Week_plot <- group_by(Week, day) %>% 
   mutate(Total = sum(count))
   
 
Week_plot <- Week %>%
     group_by(day) %>% 
     summarise(sum(count))

Week_plot <- Week_plot[-c(8), ]

names(Week_plot) <- c("Day", "Total_Repayment")

Week_plot$Day <- factor(Week_plot$Day, 
                        levels = c("Monday", "Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"))
 
#Plot

ggplot(Week_plot, aes(Day, y = Total_Repayment)) + 
  geom_bar(stat = "identity", position = "dodge", color='brown',fill='tan1', width = 0.5)+
  theme(panel.grid.minor.x=element_blank(),
        panel.grid.major.x=element_blank()) +theme(panel.background = element_blank()) +
  theme(panel.grid.major = element_line(colour = "light grey"))+
  theme(axis.line.x = element_line(size = 1.5, colour = "#c94a38"))+
  theme(plot.title = element_text(hjust = 0.5, colour = "dark green"))+
  theme(axis.text.x = element_text(face = "bold", size=8, angle=0, color="black"))+
  geom_text(aes(label = Total_Repayment),vjust = -0.25, size = 3)+
  labs(title = "Total daily repayments", x = "Days", y = "Repayments")+
  scale_y_continuous(labels = label_comma(big.mark = ","))

render("Report.Rmd", "pdf_document")

hist(Demand_plot$count) 
 
 str(Week)
 nrow(Repayments)
tail(Repayments)
cyle <- filter(data, data$Cycle == 1.000)
