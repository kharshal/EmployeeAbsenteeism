# -*- coding: utf-8 -*-
"""
Created on Mon Aug 13 18:22:34 2018

@author: Harshal
"""

# Import Libraries
import os
import numpy as np
import pandas as pd
from fancyimpute import KNN

# Set active Work Directory
os.chdir("C:/Users/Harshal/Desktop/Edwisor/Project/Project_Datafiles")

# Load Data
xls = pd.ExcelFile("Absenteeism_at_work_Project.xls")
EmployeeData = xls.parse()

# Get Cnames
def get_cname(data):
    all_cnames = []
    num_cnames = []
    cat_cnames = []
    for i in data.columns:
        all_cnames.append(str(i))
        if(data[i].dtype == "object"):
            cat_cnames.append(str(i))
        else:
            num_cnames.append(str(i))
    cnames = [all_cnames, num_cnames, cat_cnames]    
    return(cnames)

# Get cnames - cnames[0] - all cnames; cnames[1] - numeric cnames; cnames[2] - categorical cnames 
cnames = get_cname(EmployeeData)

#rows, cols = EmployeeData.shape
rows = EmployeeData.shape[0] #gives number of row count
cols = EmployeeData.shape[1] #gives number of col count

# Exploratory Data Analysis 
for i in cnames[0]:
    print(str(i) + "    " + str(type(EmployeeData[i][1])))
    
# Change Data as per problem requirement  
    for i in range(0,rows):
        if EmployeeData["Absenteeism time in hours"][i] != 0:
            if EmployeeData["Reason for absence"][i] == 0:
                EmployeeData["Reason for absence"][i] = np.nan
            if EmployeeData["Month of absence"][i] == 0:
                EmployeeData["Month of absence"][i] = np.nan
        
def preprocessing(EmployeeData):
    # Change into require Data Types
    EmployeeData["ID"] = EmployeeData["ID"].astype(str)
    EmployeeData["Reason for absence"] = EmployeeData["Reason for absence"].astype(str)
    EmployeeData["Month of absence"] = EmployeeData["Month of absence"].astype(str)
    EmployeeData["Day of the week"] = EmployeeData["Day of the week"].astype(str)
    EmployeeData["Seasons"] = EmployeeData["Seasons"].astype(str)
    EmployeeData["Disciplinary failure"] = EmployeeData["Disciplinary failure"].astype(str)
    EmployeeData["Education"] = EmployeeData["Education"].astype(str)
    EmployeeData["Son"] = EmployeeData["Son"].astype(str)
    EmployeeData["Social drinker"] = EmployeeData["Social drinker"].astype(str)
    EmployeeData["Social smoker"] = EmployeeData["Social smoker"].astype(str)
    EmployeeData["Pet"] = EmployeeData["Pet"].astype(str)
    
    # Change NaN string values back to NaN 
    EmployeeData["ID"] = EmployeeData["ID"].replace("nan",np.nan)
    EmployeeData["Reason for absence"] = EmployeeData["Reason for absence"].replace("nan",np.nan)
    EmployeeData["Month of absence"] = EmployeeData["Month of absence"].replace("nan",np.nan)
    EmployeeData["Day of the week"] = EmployeeData["Day of the week"].replace("nan",np.nan)
    EmployeeData["Seasons"] = EmployeeData["Seasons"].replace("nan",np.nan)
    EmployeeData["Disciplinary failure"] = EmployeeData["Disciplinary failure"].replace("nan",np.nan)
    EmployeeData["Education"] = EmployeeData["Education"].replace("nan",np.nan)
    EmployeeData["Son"] = EmployeeData["Son"].replace("nan",np.nan)
    EmployeeData["Social drinker"] = EmployeeData["Social drinker"].replace("nan",np.nan)
    EmployeeData["Social smoker"] = EmployeeData["Social smoker"].replace("nan",np.nan)
    EmployeeData["Pet"] = EmployeeData["Pet"].replace("nan",np.nan)
    
    # Covert factor varaible values to labels 
    for i in range(0, len(EmployeeData.columns)):
        if(EmployeeData.iloc[:,i].dtypes == 'object'):
            EmployeeData.iloc[:,i] = pd.Categorical(EmployeeData.iloc[:,i])
            EmployeeData.iloc[:,i] = EmployeeData.iloc[:,i].cat.codes 
            EmployeeData.iloc[:,i] = EmployeeData.iloc[:,i].astype('object')
    
    # Convert -1 values back to NaN
    for i in cnames[0]:
        for j in range(0,rows):
            if EmployeeData.loc[j,i] == -1:
                EmployeeData.loc[j,i] = np.nan
    
    return EmployeeData

EmployeeData = preprocessing(EmployeeData)
cnames = get_cname(EmployeeData)

# Missing Value analysis
def missingValueCheck(data):
    print(data.isna().sum())
missingValueCheck(EmployeeData)

# Impute values related to ID
Depenedent_ID =  ["ID","Transportation expense","Service time","Age","Height",
                   "Distance from Residence to Work","Education","Son","Weight",
                   "Social smoker","Social drinker","Pet","Body mass index"]
Depenedent_ID_data = EmployeeData[Depenedent_ID].copy()
Depenedent_ID_data = Depenedent_ID_data.groupby("ID").max()
Depenedent_ID.remove('ID')

for i in Depenedent_ID:
    for j in range(0,rows):
        RI = EmployeeData["ID"][j]
        if np.isnan(EmployeeData.loc[j,i]):
            EmployeeData[i][j] = Depenedent_ID_data.loc[RI,i]
            
# Impute values for other variables with KNN            
EmployeeData = pd.DataFrame(KNN(k = 7).complete(EmployeeData), columns = EmployeeData.columns)
EmployeeData = preprocessing(EmployeeData)
missingValueCheck(EmployeeData)

# Outlier Analysis
# Get cnames of numeric varaibles not dependent on ID
numeric_cnames = ['Work load Average/day ','Hit target']

# Impute Outliers with NA
for i in numeric_cnames:
    q75, q25 = np.nanpercentile(EmployeeData.loc[:,i],[75, 25])
    iqr = q75 - q25
    min = q25 - (iqr*1.5)
    max = q75 + (iqr*1.5)
    EmployeeData.loc[EmployeeData[i] < min, i] = np.nan
    EmployeeData.loc[EmployeeData[i] > max, i] = np.nan
#Impute with KNN
EmployeeData = pd.DataFrame(KNN(k = 7).complete(EmployeeData), columns = EmployeeData.columns)
EmployeeData = preprocessing(EmployeeData)
missingValueCheck(EmployeeData)

# Feature Selection
from sklearn.feature_selection import SelectKBest
from sklearn.feature_selection import f_regression
X = EmployeeData.iloc[:,:-1].values
y = EmployeeData.iloc[:,20].values
# Create an SelectKBest object to select features with two best ANOVA F-Values
selector = SelectKBest(f_regression, k=7)
# Apply the SelectKBest object to the features and target
X = selector.fit_transform(X, y)
selector.scores_

# Feature Scaling
#Nomalisation
from sklearn.preprocessing import normalize
X = normalize(X, norm='l2')
#for i in range(0,(X_kbest.shape[1]-1)):
#    X_kbest[i] = (X_kbest[i] - np.min(X_kbest[i])) / (np.max(X_kbest[i]) - np.min(X_kbest[i]))

# Error Matrix
from sklearn.metrics import mean_squared_error
from math import sqrt
def RMSE(y, pred):
    print(sqrt(mean_squared_error(y, pred)))
    
def MSE(y, pred):
    print(mean_squared_error(y, pred))

from sklearn.model_selection import train_test_split
X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)

# Multiple Linear Regression
from sklearn.linear_model import LinearRegression
lm_regressor = LinearRegression()
lm_regressor.fit(X_train, y_train)
lm_predict = lm_regressor.predict(X_test)
RMSE(y_test, lm_predict)
MSE(y_test, lm_predict)
     
# Decision Tree Regressor
from sklearn.tree import DecisionTreeRegressor
DT_regressor = DecisionTreeRegressor()
DT_regressor.fit(X_train, y_train)
DT_predict = DT_regressor.predict(X_test)
RMSE(y_test, DT_predict)
MSE(y_test, DT_predict)

# Random Forest Regressor
from sklearn.ensemble import RandomForestRegressor
RF_regressor = RandomForestRegressor()
RF_regressor.fit(X_train, y_train)
RF_predict = RF_regressor.predict(X_test)
RMSE(y_test, RF_predict)
MSE(y_test, RF_predict)

# Support Vector Regressor
from sklearn.svm import SVR
SVR_regressor = SVR(kernel='rbf')
SVR_regressor.fit(X_train, y_train)
SVR_predict = SVR_regressor.predict(X_test)
RMSE(y_test, SVR_predict)
MSE(y_test, SVR_predict)

############################## Problems ##################################
# Suggesting the changes
RF_regressor_p1 = RandomForestRegressor().fit(X, y)
RF_regressor_p1.feature_importances_

# Calculating Losses
p2_data = X
#Predict for new test cases
p2_predict = RF_regressor.predict(p2_data)
p2_predict = pd.DataFrame(p2_predict)
p2_predict.columns = ['predictions']
# Add predicted values to the DataSet
p2_frames = [EmployeeData, p2_predict]
p2_dataSet = pd.concat(p2_frames, axis=1)

# Calculate the total Loss 
Loss = 0
for i in range(0,p2_dataSet.shape[0]):
    if (p2_dataSet['Hit target'][i] != 100):
        if (p2_dataSet['Age'][i] >= 25 and  p2_dataSet['Age'][i] <= 32):
            Loss += (p2_dataSet['Disciplinary failure'][i] + 1) * 2000 + (p2_dataSet['Education'][i] + 1 + 1) * 500 + p2_dataSet['predictions'][i] * 1000
        elif (p2_dataSet['Age'][i] >= 33 and  p2_dataSet['Age'][i] <= 40):
            Loss += (p2_dataSet['Disciplinary failure'][i] + 1) * 2000 + (p2_dataSet['Education'][i] + 1 + 2) * 500 + p2_dataSet['predictions'][i] * 1000
        elif (p2_dataSet['Age'][i] >= 41 and  p2_dataSet['Age'][i] <= 49):
            Loss += (p2_dataSet['Disciplinary failure'][i] + 1) * 2000 + (p2_dataSet['Education'][i] + 1 + 3) * 500 + p2_dataSet['predictions'][i] * 1000
        elif (p2_dataSet['Age'][i] >= 50 and  p2_dataSet['Age'][i] <= 60):
            Loss += (p2_dataSet['Disciplinary failure'][i] + 1) * 2000 + (p2_dataSet['Education'][i] + 1 + 4) * 500 + p2_dataSet['predictions'][i] * 1000

# To calculate loss per month
Loss = Loss/12
print(Loss)

