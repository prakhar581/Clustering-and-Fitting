#Does organisational innovation lead to higher sales?
library("xlsx")
library("openxlsx")
library(tidyverse)
library(ggplot2)
library(ggcorrplot)
library(corrplot)
df=read.xlsx('MIP_2013.xlsx',sheet=1)
#get the data
#Employee,Turnover,1st OI,2nd OI,3rd OI(any true innovation true),Total Inovation project
df1=df[c('q13a_2012','q15a_2012','q123a','q123b','q123c','q71a')]
#Renaming columns
df2 <-df1 %>% 
  rename(
    TotEmp = q13a_2012,Turnover = q15a_2012,OrgInv1 = q123a,OrgInv2 = q123b,OrgInv3 = q123c
    ,TotProj = q71a
  )
#data cleaning
colSums(is.na(df2))
#remove all nulls
df3 <- na.omit(df2)
#After remove nulls
colSums(is.na(df3))
#create activity column, If any innovation activity is activity is mar as true
df4=df3[df3$OrgInv1=='yes' | df3$OrgInv2=='yes' | df3$OrgInv3=='yes' ,]
df4$activity='yes'
df4 = subset(df4, select = -c(OrgInv1,OrgInv2,OrgInv3) )
df5=df3[df3$OrgInv1=='no' & df3$OrgInv2=='no' & df3$OrgInv3=='no' ,]
df5$activity='no'
df5 = subset(df5, select = -c(OrgInv1,OrgInv2,OrgInv3) )
#joining df4 and df5
dat=rbind(df4,df5)
#Employee cannot be in decimal
dat$TotEmp=round(dat$TotEmp)
dat$inovActivity<- ifelse( dat$activity =="yes", 1, 0)
dat<- subset(dat, select = -c(activity))
### Data manipulation and visualization
#How many activities are yes and how many are no
table(dat$inovActivity) #949 no's and 474 yes
barplot(table(dat$inovActivity),main="Innovation activity")
#Turnover vs activity
ggplot(dat, aes(x=inovActivity, y=Turnover)) +
  geom_point() 
ggplot(dat) +
  aes(x = inovActivity, fill = Turnover) +
  geom_bar() +
  scale_fill_hue() +
  theme_minimal()+ theme(axis.text.x = element_text(angle = 60, hjust = 1))
##Correlation polot
#corr plot
plot =cor(dat)
options(repr.plot.width=100, repr.plot.height=50)
ggcorrplot(plot,hc.order = TRUE, type = "lower",
           lab = TRUE)
#We can see from the above plots there is no relationship between innovation activity and sales of an organization

#regression analysis and visualization using ggplot
library(caTools)
set.seed(123)
split = sample.split(dat$Turnover, SplitRatio = 0.8)
training_set = subset(dat, split == TRUE)
test_set = subset(dat, split == FALSE)
# Fitting Multiple Linear Regression to the Training set
regressor = lm(formula = Turnover ~ .,
               data = training_set)
# Predicting the Test set results
library(Metrics)
y_pred = predict(regressor, newdata = test_set)
rmse(test_set$Turnover, y_pred)

#Cluster analysis
#As we saw that there is strong corrleation between number of employess and Turnover
# Using the elbow method to find the optimal number of clusters
dat2=subset(dat,select =c(Turnover,TotEmp))
set.seed(6)
wcss = vector()
for (i in 1:10) wcss[i] = sum(kmeans(dat2, i)$withinss)
plot(1:10,
     wcss,
     type = 'b',
     main = paste('The Elbow Method'),
     xlab = 'Number of clusters',
     ylab = 'WCSS')

# Fitting K-Means to the dataset
set.seed(29)
kmeans = kmeans(x = dat2, centers = 3)
y_kmeans = kmeans$cluster

# Visualising the clusters
library(cluster)
clusplot(dat2,
         y_kmeans,
         lines = 0,
         shade = TRUE,
         color = TRUE,
         labels = 2,
         plotchar = FALSE,
         span = TRUE,
         main = paste('Clusters of customers'),
         xlab = 'Total employees',
         ylab = 'Turnover')
#simple linear regression 
set.seed(123)
split = sample.split(dat2$Turnover, SplitRatio = 2/3)
training_set = subset(dat2, split == TRUE)
test_set = subset(dat2, split == FALSE)
# Fitting Simple Linear Regression to the Training set
regressor = lm(formula = Turnover ~ TotEmp,
               data = training_set)

# Predicting the Test set results
y_pred = predict(regressor, newdata = test_set)
rmse(test_set$Turnover, y_pred)

# Visualising the Test set results
library(ggplot2)
ggplot() +
  geom_point(aes(x = test_set$TotEmp, y = test_set$Turnover),
             colour = 'red') +
  geom_line(aes(x = training_set$TotEmp, y = predict(regressor, newdata = training_set)),
            colour = 'blue') +
  ggtitle('Net sales Turnover vs Employess (Test set)') +
  xlab('Employess') +
  ylab('Net sales Turnover')

#Simple linear regression wrt activity
dat3=subset(dat,select =c(Turnover,inovActivity))
set.seed(123)
split = sample.split(dat3$Turnover, SplitRatio = 2/3)
training_set = subset(dat3, split == TRUE)
test_set = subset(dat3, split == FALSE)
# Fitting Simple Linear Regression to the Training set
regressor = lm(formula = Turnover ~ inovActivity,
               data = training_set)
# Predicting the Test set results
y_pred = predict(regressor, newdata = test_set)
rmse(test_set$Turnover, y_pred)

