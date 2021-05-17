#%%
import pandas as pd
import numpy as np

import matplotlib.pyplot as plt
import seaborn as sns

# to save the trained scaler class.
import joblib

import scipy.stats as stats

from sklearn.model_selection import train_test_split

from sklearn.preprocessing import MinMaxScaler,Binarizer

from sklearn.linear_model import Lasso

from sklearn.metrics import mean_squared_error, r2_score

# from feature-engine.
from feature_engine.imputation import (
    AddMissingIndicator,
    MeanMedianImputer,
    CategoricalImputer
)
from feature_engine.encoding import (
    RareLabelEncoder,
    OrdinalEncoder
)
from feature_engine.transformation import (
    LogTransformer,
    YeoJohnsonTransformer
)

from feature_engine.selection import (
    DropFeatures
)

from feature_engine.wrappers import (
    SklearnTransformerWrapper
)

pd.set_option('display.max_columns', None)

#%% Reproducibility: Setting the seed.
data = pd.read_csv('train.csv', encoding='utf-8')

print(data.shape)
print(data.head())

# %%
X_train, X_test, y_train, y_test = train_test_split(data.drop(['Id', 'SalePrice'], axis=1),
    data['SalePrice'],
    test_size=0.1,
    random_state=0
)
X_train.shape, X_test.shape
# %% Target.
y_train = np.log(y_train)
y_test = np.log(y_test)

# %% Categorical.
X_train['MSSubClass'] = X_train['MSSubClass'].astype('O')
X_test['MSSubClass'] = X_test['MSSubClass'].astype('O')

cat_vars = [var for var in data.columns if data[var].dtype == 'O']
cat_vars_with_na = [var for var in cat_vars if X_train[var].isnull().sum() > 0]

# variables to impute with the string missing. ()
with_string_missing = [var for var in cat_vars_with_na if X_train[var].isnull().mean() > 0.1]
with_frequent_category = [var for var in cat_vars_with_na if X_train[var].isnull().mean() < 0.1]

# variables to impute with the most frequent category
# %% Missing values -- Categorical -- Missing.
cat_imputer_missing = CategoricalImputer(
    imputation_method='missing',
    variables=with_string_missing
)
cat_imputer_missing.fit(X_train)
print(cat_imputer_missing.imputer_dict_)

X_train = cat_imputer_missing.transform(X_train)
X_test = cat_imputer_missing.transform(X_test)

# %% Missing values -- Categorical -- Frequency.
cat_imputer_frequent = CategoricalImputer(
    imputation_method='frequent',
    variables=with_frequent_category
)
cat_imputer_frequent.fit(X_train)
print(cat_imputer_frequent.imputer_dict_)

X_train = cat_imputer_frequent.transform(X_train)
X_test = cat_imputer_frequent.transform(X_test)

# %% Varief whether there are missing value.
X_train[cat_vars_with_na].isnull().sum()
[var for var in cat_vars_with_na if X_test[var].isnull().sum() > 0]
# %% Numerical variables.
num_vars = [var for var in X_train.columns if var not in cat_vars and var != 'SalePrice']
vars_with_na = [var for var in num_vars if X_train[var].isnull().sum() > 0]
print(len(vars_with_na))

X_train[vars_with_na].isnull().mean()

# %% Missing values -- Numerical -- add missing indicator.
missing_ind = AddMissingIndicator(variables=vars_with_na)
missing_ind.fit(X_train)
X_train = missing_ind.transform(X_train)
X_test = missing_ind.transform(X_test)

# check the binary missing indicator variables
X_train[['LotFrontage_na', 'MasVnrArea_na', 'GarageYrBlt_na']].head()
# %% # %% Missing values -- Numerical -- add missing indicator.
mean_imputer = MeanMedianImputer(
    imputer_method='mean',
    variables=vars_with_na
)
mean_imputer.fit(X_train)
print(mean_imputer.imputer_dict_)

X_train = mean_imputer.transform(X_train)
X_test = mean_imputer.transform(X_test)

# %% Varief whether there are missing value.
X_train[cat_vars_with_na].isnull().sum()
[var for var in cat_vars_with_na if X_test[var].isnull().sum() > 0]

#%% Temporal variables.
def elapsed_years(df, var):
    df[var] = df['YrSold'] - df[var]
    return df
for var in ['YearBuilt', 'YearRemodAdd', 'GarageYrBlt']：
    X_train = elapsed_years(X_train, var)
    X_test = elapsed_years(X_test, var)

# now we drop YrSold.
drop_features = DropFeatures(features_to_drop=['YrSold'])
X_train = mean_imputer.fit_transform(X_train)
X_test = mean_imputer.transform(X_test)

# %% Numerical variable -- transformation.
log_transformer = LogTransformer(
    variables=["LotFrontage", "1stFlrSF", "GrLivArea"],
)
X_train = log_transformer.fit_transform(X_train)
X_test = log_transformer.transform(X_test)

# check that test set does not contain null values in the engineered variables
[var for var in ["LotFrontage", "1stFlrSF", "GrLivArea"] if X_test[var].isnull().sum() > 0]