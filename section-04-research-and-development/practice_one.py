#%%
import pandas as pd
import numpy as np

import matplotlib.pyplot as plt
import seaborn as sns

import joblib

import scipy.stats as stats

from sklearn.linear_model import Lasso

from sklearn.metrics import mean_squared_error, r2_score

pd.set_option('display.max_columns', None)

#%%
# Load Data.
data = pd.read_csv('train.csv', encoding='utf-8')
#%% 觀察欄位
data.head()
#%% Data Cleaning and Analysis.  包含去掉不重要的資訊，及將資料的欄位屬性修正，以便往後Pipline.
# 像是如果這是分等級的欄位，卻是用數值紀錄，因此會變成int64，因此要轉換成‘Ｏ’。
data.drop('Id', axis=1, inplace=True)
data['MSSubClass'] = data['MSSubClass'].astype('O')
#%% 先來看看答案的分佈情形。
data['SalePrice'].hist(bins=50, density=True)
plt.yaxis= "Number of Hourse"
plt.xaxis = 'Sale Price'
# plt.title = "Price - Number"
# %% 因為發現答案是skew，所以取log。
np.log(data['SalePrice']).hist(bins=50, density=True)
plt.yaxis= "Number of Hourse"
plt.xaxis = 'Log of Sale Price'

# %% Categorical.
cat_vars = [var for var in data.columns if data[var].dtype == 'O']
print('Num of Cartegory: ', len(cat_vars))
# %% Numerical.
# num_vars = [var for var in data.columns if var not in cat_vars and var != 'SalePrice']
num_vars = [var for var in data.columns if var not in cat_vars and var != 'SalePrice']
print('Num of Numerical: ', len(num_vars))
# %% Missing Value.
var_with_na = [var for var in data.columns if data[var].isnull().sum()>0]
# 讓缺少資料的欄位按大小排列。
data[var_with_na].isnull().mean().sort_values(ascending=False)
# Visualization missing data.
data[var_with_na].isnull().mean().sort_values(ascending=False).plot.bar(figsize=(10, 4))
plt.axhline(y=0.8, color='r', linestyle='-')
plt.axhline(y=0.9, color='g', linestyle='-')

# %% Calculating how many number of Category and numerical Variable with missing data.
cat_na = [var for var in cat_vars if var in var_with_na]
num_na = [var for var in num_vars if var in var_with_na]

print('Number of Categorical variables with na:', len(cat_na))
print('Number of Numerical variables with na:', len(num_na))


# %% 對於缺失值與Ｙ之間的關係作評估。
# Data analysis 的好朋友就是 Visualization.
def analyse_na_value(df, var, Y):
    df = df.copy()
    df[var] = np.where(df[var].isnull(), 1, 0)
    tmp = df.groupby(var)[Y].agg(['mean', 'std'])
    tmp.plot(kind='barh', y='mean', legend=False, xerr='std', title="SalePrice", color='green')
    plt.show()

# Plot.
for var in var_with_na:
    analyse_na_value(data, var, 'SalePrice')

# %% Temporal data are in the Numerical.
year_vars = [var for var in num_vars if 'Yr' in var or 'Year' in var]
# 對數據的觀察： 1. 直接畫出數據。
data.groupby('YrSold')['SalePrice'].median().plot()
plt.ylabel('Median House Price')
plt.show()
data.groupby('YearBuilt')['SalePrice'].median().plot()
plt.ylabel('Median House Price')
plt.show()
# 對數據的觀察： 2. 利用var之間的關係，組合關係，做觀察。
def analyse_year_vars(df, var):
    df = df.copy()
    df[var] = df['YrSold'] - df[var]
    df.groupby('YrSold')[var].median().plot()
    plt.ylabel('Time from ' + var)
    plt.show()
for var in year_vars:
    if var != "YrSold":
        analyse_year_vars(data, var)
# 對數據的觀察： 3. 利用var之間的關係，組合關係，觀察與Ｙ的關係，驗證整組合。
def analyse_year_vars(df, var):
    
    df = df.copy()
    
    # capture difference between a year variable and year
    # in which the house was sold
    df[var] = df['YrSold'] - df[var]
    
    plt.scatter(df[var], df['SalePrice'])
    plt.ylabel('SalePrice')
    plt.xlabel(var)
    plt.show()
    
    
for var in year_vars:
    if var !='YrSold':
        analyse_year_vars(data, var)

#%% 離散型資料。  在此，作著將數值種類少於20種，就當作是一種離散型資料。
discrete_vars = [var for var in num_vars if var != year_vars and len(data[var].unique() < 20)]
print("Number of discrete variables:", len(discrete_vars))
# See the contribution of vars to the tendency.
for var in discrete_vars:
    # make boxplot with Catplot
    sns.catplot(x=var, y='SalePrice', data=data, kind='box', height=4, aspect=1.5)
    # add data points to boxplot with stripplot.
    sns.stripplot(x=var, y='SalePrice', data=data, jitter=0.1, alpha=0.3, color='k')
    plt.show()

#%% Continuous Variables.
cont_var = [var for var in num_var if var not in discrete_vars+year_vars]
print ('Number of continuous variables: ', len(cont_vars))

# 畫出所有連續型的資料，看分佈情形。
data[cont_var].hist(bins=30, figsize=(15,15))
# %% Skew and bad distribution -> nearly normal distribution.
# Sometimes, transforming the variables to improve the value spread, improves the model performance.
# But it is unlikely that a transformation will help change the distribution of the super skewed variables dramatically.

skewed = [
    'BsmtFinSF2', 'LowQualFinSF', 'EnclosedPorch',
    '3SsnPorch', 'ScreenPorch', 'MiscVal'
]

cont_vars = [
    'LotFrontage',
    'LotArea',
    'MasVnrArea',
    'BsmtFinSF1',
    'BsmtUnfSF',
    'TotalBsmtSF',
    '1stFlrSF',
    '2ndFlrSF',
    'GrLivArea',
    'GarageArea',
    'WoodDeckSF',
    'OpenPorchSF',
]

# Continue vars.
# Try Yeo-Johnson transformation first, because it will not to process 0 data.
tmp = data.copy()
for var in cont_vars:
    tmp[var], param =stats.yeojohnson(data[var])
tmp[cont_var].hist(bins=30, figsize=(15, 15))
plt.show()
# let's plot the original or transformed variables vs sale price, and see if there is a relationship.
for var in cont_vars:
    plt.figure(figsize=(12, 4))

    # Plot ther original variable vs sale price.
    plt.subplot(1, 2, 1)
    plt.scatter(data[var], np.log(data['SalePrice']))
    plt.ylabel('Sale Price')
    plt.xlabel('Original ' + var)

    # Plot ther original variable vs sale price.
    plt.subplot(1, 2, 2)
    plt.scatter(data[var], np.log(tmp['SalePrice']))
    plt.ylabel('Sale Price')
    plt.xlabel('Transformed ' + var)

#%% Log Transformation.
tmp = data.copy()
for var in ["LotFrontage", "1stFlrSF", "GrLivArea"]:
    tmp[var] = np.log(data[var])
tmp[["LotFrontage", "1stFlrSF", "GrLivArea"]].hist(bins=30)
plt.show()

# vs SalePrice.
for var in ["LotFrontage", "1stFlrSF", "GrLivArea"]:
    
    plt.figure(figsize=(12,4))
    
    # plot the original variable vs sale price    
    plt.subplot(1, 2, 1)
    plt.scatter(data[var], np.log(data['SalePrice']))
    plt.ylabel('Sale Price')
    plt.xlabel('Original ' + var)

    # plot transformed variable vs sale price
    plt.subplot(1, 2, 2)
    plt.scatter(tmp[var], np.log(tmp['SalePrice']))
    plt.ylabel('Sale Price')
    plt.xlabel('Transformed ' + var)
                
    plt.show()

# %% Skew.
# 利用 0, 1 的方式，暸解skew的資料，是否之間是有顯著性差異的<，初步看來是沒有的，因為error區間重疊。
for var in skewed:
    tmp = data.copy()
    tmp[var] = np.where(tmp[var]==0, 0, 1)
    tmp = tmp.groupby(var)['SalePrice'].agg(['mean', 'std'])
    tmp.plot(kind='hbar', y='mean', xerr='std', legend=False, title='Sale Price', color='green')
    plt.show()

#%% Categorical variables.
data[cat_vars].nunique().sort_values(ascending=False).plot.bar(figsize=(12,5))
# %%
# Categorical Variables.
data[cat_vars].nunique().sort_values(ascending=False).plot.bar(figsize=(12, 5))

# %%
# Quality variables.
data[cat_vars].head(5)
# Find the columns by hand.
qual_vars = ['ExterQual', 'ExterCond', 'BsmtQual', 'BsmtCond',
             'HeatingQC', 'KitchenQual', 'FireplaceQu',
             'GarageQual', 'GarageCond',
            ]
qual_mappings = {'Po': 1, 'Fa': 2, 'TA': 3, 'Gd': 4, 'Ex': 5, 'Missing': 0, 'NA': 0}
# print(data[qual_vars].head())
for var in qual_vars:
    data[var] = data[var].map(qual_mappings)
# print(data[qual_vars].head())
# %%
# --------------
var = 'BsmtExposure'
exposure_mappings = {'No': 1, 'Mn': 2, 'Av': 3, 'Gd': 4, 'Missing': 0, 'NA': 0}

data[var] = data[var].map(exposure_mappings)
# %%
# --------------
finish_mappings = {'Missing': 0, 'NA': 0, 'Unf': 1, 'LwQ': 2, 'Rec': 3, 'BLQ': 4, 'ALQ': 5, 'GLQ': 6}

finish_vars = ['BsmtFinType1', 'BsmtFinType2']

for var in finish_vars:
    data[var] = data[var].map(finish_mappings)
# --------------
garage_mappings = {'Missing': 0, 'NA': 0, 'Unf': 1, 'RFn': 2, 'Fin': 3}

var = 'GarageFinish'

data[var] = data[var].map(garage_mappings)
# --------------
fence_mappings = {'Missing': 0, 'NA': 0, 'MnWw': 1, 'GdWo': 2, 'MnPrv': 3, 'GdPrv': 4}

var = 'Fence'

data[var] = data[var].map(fence_mappings)
# ----Visualization.------
qual_vars  = qual_vars + finish_vars + ['BsmtExposure','GarageFinish','Fence']
for var in qual_vars:
    # make boxplot with Catplot
    sns.catplot(x=var, y='SalePrice', data=data, kind="box", height=4, aspect=1.5)
    # add data points to boxplot with stripplot
    sns.stripplot(x=var, y='SalePrice', data=data, jitter=0.1, alpha=0.3, color='k')
    plt.show()

# %% Categotical with out Quality meaning.
cat_others = [var for var in cat_vars if var not in qual_vars]
len(cat_others)
# %% Rare labels.
def analyse_rare_labels(df, var, rare_perc):
    df = df.copy()
    tmp = df.groupby(var)['SalePrice'].count() / len(df)
    return tmp[tmp < rare_perc]

for var in cat_others:
    print(analyse_rare_labels(data, var, 0.01))
    print("-----------------------------------")

# --------Visualization.---------------
for var in cat_others:
    sns.catplot(x=var, y='SalePrice', data=data, kind='box', height=4, aspect=1.5)
    sns.stripplot(x=var, y='SalePrice', data=data, jitter=0.1, alpha=0.3, color='k')
    plt.show()
# %%
