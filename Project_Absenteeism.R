# Clean current environment
rm(list=ls())
# Set work directory
setwd("C:/Users/Harshal/Desktop/Edwisor/Project/Project_Datafiles")

# Load require Packages
p <- c("xlsx","DMwR","corrgram","caret","usdm","rpart","DataCombine","randomForest",
       "e1071")
lapply(p, require, character.only=TRUE)
rm(p)

EmployeeData <- read.xlsx("Absenteeism_at_work_Project.xls", sheetIndex = 1)

###################### Exploratory Data Analysis ##################
# Check variable types
cnames <- colnames(EmployeeData)

# Check number of unique variables
for (i in cnames){
  print(i)
  print(aggregate(data.frame(count = EmployeeData[,i]), 
                  list(value = EmployeeData[,i]), length))
}

# Data Preprocessing 
preprocessing <- function(EmployeeData){
  EmployeeData$ID <- as.factor(EmployeeData$ID)
  for (i in (1:nrow(EmployeeData))){
    if (EmployeeData$Absenteeism.time.in.hours[i] != 0 || is.na(EmployeeData$Absenteeism.time.in.hours[i])){
      if(EmployeeData$Reason.for.absence[i] == 0 || is.na(EmployeeData$Reason.for.absence[i])){
        EmployeeData$Reason.for.absence[i] = NA
      }
      if(EmployeeData$Month.of.absence[i] == 0 || is.na(EmployeeData$Month.of.absence[i])){
        EmployeeData$Month.of.absence[i] = NA
      }
    }
  }
  EmployeeData$Reason.for.absence <- as.factor(EmployeeData$Reason.for.absence)
  EmployeeData$Month.of.absence <- as.factor(EmployeeData$Month.of.absence)
  EmployeeData$Day.of.the.week <- as.factor(EmployeeData$Day.of.the.week)
  EmployeeData$Seasons <- as.factor(EmployeeData$Seasons)
  EmployeeData$Disciplinary.failure <- as.factor(EmployeeData$Disciplinary.failure)
  EmployeeData$Education <- as.factor(EmployeeData$Education)
  EmployeeData$Son <- as.factor(EmployeeData$Son)
  EmployeeData$Social.drinker <- as.factor(EmployeeData$Social.drinker)
  EmployeeData$Social.smoker <- as.factor(EmployeeData$Social.smoker)
  EmployeeData$Pet <- as.factor(EmployeeData$Pet)
  return(EmployeeData)
}
EmployeeData <- preprocessing(EmployeeData)

#selecting only factor
get_cat <- function(data) {
  return(colnames(data[,sapply(data, is.factor)]))
}
cat_cnames <- get_cat(EmployeeData)

#selecting only numeric
get_num <- function(data) {
  return(colnames(data[,sapply(data, is.numeric)]))
}
num_cnames <- get_num(EmployeeData)

# Covert factor varaible values to labels 
for (i in cat_cnames){
  EmployeeData[,i] <- factor(EmployeeData[,i], 
                             labels = (1:length(levels(factor(EmployeeData[,i])))))
}

############################### Data PreProcessing #########################
# Missing Value Analysis
# Get Missing Values for all Variables
missingValueCheck <- function(data){
  for (i in colnames(data)){
    print(i)
    print(sum(is.na(data[i])))
  }
  print("Total")
  print(sum(is.na(EmployeeData)))
}
missingValueCheck(EmployeeData)

# Impute values related to ID
Depenedent_ID <- c("ID","Transportation.expense","Service.time","Age","Height",
                   "Distance.from.Residence.to.Work","Education","Son","Weight",
                   "Social.smoker","Social.drinker","Pet","Body.mass.index")
Depenedent_ID_data <- EmployeeData[,Depenedent_ID]
Depenedent_ID_data <- aggregate(. ~ ID, data = Depenedent_ID_data, 
                     FUN = function(e) c(x = mean(e)))
for (i in Depenedent_ID) {
  for (j in (1:nrow(EmployeeData))){
    ID <- EmployeeData[j,"ID"]
    if(is.na(EmployeeData[j,i])){
      EmployeeData[j,i] <- Depenedent_ID_data[ID,i]
    }
  }
}
# Impute values for other variables
EmployeeData = knnImputation(EmployeeData, k = 7)
missingValueCheck(EmployeeData)

# Outlier Analysis
# Histogram
for(i in num_cnames){
  hist(EmployeeData[,i], xlab=i, main=" ", col=(c("lightblue","darkgreen")))
}

# BoxPlot
num_cnames <- num_cnames[ num_cnames != "Absenteeism.time.in.hours"]
for(i in num_cnames){
  boxplot(EmployeeData[,i]~EmployeeData$Absenteeism.time.in.hours,
          data=EmployeeData, main=" ", ylab=i, xlab="Absenteeism.time.in.hours",  
          col=(c("lightblue","darkgreen")), outcol="red")
}

# Removing ID related varaibles from num_cnames
for (i in Depenedent_ID){
  num_cnames <- num_cnames[ num_cnames != i]
}
# Replace all outliers with NA and impute
for(i in num_cnames){
  val = EmployeeData[,i][EmployeeData[,i] %in% boxplot.stats(EmployeeData[,i])$out]
  EmployeeData[,i][EmployeeData[,i] %in% val] = NA
}

# Impute Missing Values
missingValueCheck(EmployeeData)
EmployeeData <- knnImputation(EmployeeData, k = 7)

# Copy Employee Data into DataSet for further analysis
DataSet <- EmployeeData

######################## Fature Selection #################
# Correlation Plot
corrgram(EmployeeData, upper.panel=panel.pie, text.panel=panel.txt, 
         main = "Correlation Plot")

# ANOVA
cnames <- colnames(EmployeeData)
for (i in cnames){
  print(i)
  print(summary(aov(EmployeeData$Absenteeism.time.in.hours~EmployeeData[,i], 
                    EmployeeData)))
}

# Dimensionality Reduction
EmployeeData <- subset(EmployeeData, select = -c(Weight,Education,Service.time,
                                                 Social.smoker,Body.mass.index,
                                                 Work.load.Average.day.,Seasons,
                                                 Transportation.expense,Pet,
                                                 Disciplinary.failure,Hit.target,
                                                 Month.of.absence,Social.drinker))

######################### Fature Normalization ######################
cat_cnames <- get_cat(EmployeeData)
num_cnames <- get_num(EmployeeData)

# Fature Scaling
for (i in num_cnames){
  EmployeeData[,i] <- (EmployeeData[,i] - min(EmployeeData[,i])) / 
                        (max(EmployeeData[,i]) - min(EmployeeData[,i]))
}

#################### Model Development ########################
# Data Divide
X_index <- sample(1:nrow(EmployeeData), 0.8 * nrow(EmployeeData))
X_train <- EmployeeData[X_index,-8]
X_test <- EmployeeData[-X_index,-8]
y_train <- EmployeeData[X_index,8]
y_test <- EmployeeData[-X_index,8]

train <- EmployeeData[X_index,]
test <- EmployeeData[-X_index,]

#calculate RMSE
RMSE <- function(y, yhat){
  sqrt(mean((y - yhat)^2))
}

#calculate MSE
MSE <- function(y, yhat){
  (mean((y - yhat)^2))
}

###################### Multiple Linear Regression ########################
num_data <- train[sapply(train, is.numeric)]
vifcor(num_data, th=0.9)

lm_regressor <- lm(Absenteeism.time.in.hours~.,data = train)
summary(lm_regressor)
#Predict for new test cases
for (i in cat_cnames){
  lm_regressor$xlevels[[i]] <- union(lm_regressor$xlevels[[i]], 
                                        levels(X_test[[i]]))
}
lm_predict <- predict(lm_regressor,newdata=X_test)
RMSE(lm_predict,y_test)
MSE(lm_predict,y_test)

###############################Decision Trees for regression ###################
DT_regressor <- rpart(Absenteeism.time.in.hours~.,data = train, method="anova")
#Predict for new test cases
DT_predict <- predict(DT_regressor, X_test)
RMSE(DT_predict,y_test)
MSE(DT_predict,y_test)

##################### Random Forest ########################
RF_regressor <- randomForest(x = X_train, y = y_train, ntree = 100)
#Predict for new test cases
RF_predict <- predict(RF_regressor, X_test)
RMSE(RF_predict,y_test)
MSE(RF_predict,y_test)

#################### Support Vector Regressor #####################
SVR_regressor <- svm(formula = Absenteeism.time.in.hours ~ ., 
                data = train, type = 'eps-regression')
#Predict for new test cases
SVR_predict <- predict(SVR_regressor, X_test)
RMSE(SVR_predict,y_test)
MSE(SVR_predict,y_test)

############################## Problems ##################################
# Suggesting the changes
lm_regressor_p1 <- lm(Absenteeism.time.in.hours~.,data = EmployeeData)
summary(lm_regressor_p1)

# Calculating Losses
p2_data = EmployeeData[,-8]
#Predict for new test cases
p2_predict <- predict(RF_regressor, p2_data)

# Convert predict values back to original scale
p2_predict <- (p2_predict * 120)

# Add predicted values to the DataSet
p2_dataSet <- merge(DataSet,p2_predict,by="row.names",all.x=TRUE)

# Calculate the total Loss 
Loss <- 0
for (i in 1:nrow(p2_dataSet)){
  if (p2_dataSet$Hit.target[i] != 100)
    if (p2_dataSet$Age[i] >= 25 &&  p2_dataSet$Age[i] <= 32){
      Loss = Loss + as.numeric(p2_dataSet$Disciplinary.failure[i]) * 2000 +
        (as.numeric(p2_dataSet$Education[i]) + 1) * 500 + p2_dataSet$y[i] * 1000
    }else if(p2_dataSet$Age[i] >= 33 &&  p2_dataSet$Age[i] <= 40){
      Loss = Loss + as.numeric(p2_dataSet$Disciplinary.failure[i]) * 2000 +
        (as.numeric(p2_dataSet$Education[i]) + 2) * 500 + p2_dataSet$y[i] * 1000
    }else if(p2_dataSet$Age[i] >= 41 &&  p2_dataSet$Age[i] <= 49){
      Loss = Loss + as.numeric(p2_dataSet$Disciplinary.failure[i]) * 2000 +
        (as.numeric(p2_dataSet$Education[i]) + 3) * 500 + p2_dataSet$y[i] * 1000
    }else if(p2_dataSet$Age[i] >= 50 &&  p2_dataSet$Age[i] <= 60){
      Loss = Loss + as.numeric(p2_dataSet$Disciplinary.failure[i]) * 2000 +
        (as.numeric(p2_dataSet$Education[i]) + 4) * 500 + p2_dataSet$y[i] * 1000
    } 
}
# To calculate loss per month
Loss <- Loss/12
Loss
