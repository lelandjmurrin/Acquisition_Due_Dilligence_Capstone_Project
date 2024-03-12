import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.model_selection import train_test_split
from sklearn.decomposition import PCA
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import FunctionTransformer, OneHotEncoder, StandardScaler
from sklearn.compose import ColumnTransformer
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
from sklearn.cluster import KMeans

import statsmodels.formula.api as smf

import pickle

class OLSWrapper(object):
       """ A sklearn-style wrapper for statsmodels OLS """
       def __init__(self, cols):
           self.cols_ = cols
           self.__module__ = "ols"

       def fit(self, X, y):
           df_pca = pd.DataFrame(X, columns = self.cols_, index = y.index).join(y)
           self.model_ = smf.ols(formula = y.name + " ~ " + "+".join(df_pca.columns[:-1]), data = df_pca)
           self.results_ = self.model_.fit()
           return self

       def predict(self, X):
           return self.results_.predict(X)
       
       def transform(self, X):
           return X
       
       def summary(self):
           print(self.results_.summary())
           
       
       
class XYHelper():
    def __init__(self, data, target = None):
        self.data = data.copy()
        self.target = target
        self.X = self.data.drop(self.target, axis = 1) if self.target is not None else self.data
        self.y = self.data[self.target] if self.target is not None else None

        
    def create(self, other):
        return XYHelper(other, self.target)


class DataHelper():
    def __init__(self, data, target = None):
        self.data = XYHelper(data, target)
        train_data, test_data = train_test_split(self.df, test_size = 0.3, random_state = 0)
        self.train, self.test = self.data.create(train_data), self.data.create(test_data)
        
    def __str__(self):
        return "DataHelper: " + str(self.df.shape)
    
    def __repr__(self):
        return "DataHelper: " + str(self.df.shape)
    
    def __add__(self, other):
        return DataHelper(pd.concat([self.df, other.df], axis = 0))
    
    def __eq__(self, other):
        return self.df.equals(other.data)
    
    def __sub__(self, other):
        return DataHelper(self.df.loc[self.df.index.difference(other.df.index), :])

    def barplot(self, grp = ""):
        df_melt = self.df.pipe(self.melt)
        df_barplot = df_melt.index.to_series().str.split("_", expand = True).set_axis(["Month", "Day", "Year"], axis = 1).drop_duplicates().join(df_melt).query("Feature.str.endswith(@grp)")

        plt.figure(figsize = (20,10))
        sns.barplot(data = df_barplot,
                        x = "Month",
                        y = "Amount",
                        hue = "Year",
                        errorbar = None,
                        order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                    )
        
    def scale(self):
        return DataHelper(self.df.select_dtypes(include = "number").apply(lambda x: (x-x.mean())/x.std()))
        
    @property
    def df(self):
        return self.data.data
    
    @property
    def X(self):
        return self.data.X
    
    @property
    def y(self):
        return self.data.y
    
    def melt(self):
        return DataHelper(self.df.melt(ignore_index = False, var_name = "Feature", value_name = "Amount"))

    def filter(self, cols):
        if isinstance(cols, int):
            return DataHelper(self.df.T.head(cols).T)
        
        if isinstance(cols, pd.Series):
            if cols.dtype == 'bool':
                return DataHelper(self.df[self.df.columns[cols]])
            
        if callable(cols):
            s_bool = self.df.apply(cols)
            return DataHelper(self.df[self.df.columns[s_bool]])
        
        if isinstance(cols, str):
            return DataHelper(self.df.filter(regex = cols))
        
    @property
    def T(self):
        return DataHelper(self.df.T)
    
    @property
    def index(self):
        return self.df.index
    
    @property
    def columns(self):
        return self.df.columns


    def run_pipeline(self, steps, **kwargs):
        arg_list = []
        for s in steps:
            if "pca" == s:
                if "pca_n_components" in kwargs:
                    arg_list.append((s, PCA(n_components = kwargs["pca_n_components"])))
                else: 
                    arg_list.append((s, PCA(5)))
                
            if "ols" == s:
                arg_list.append(("model", OLSWrapper(self.X.columns)))

            if "ss" == s:
                arg_list.append((s, StandardScaler()))
            
            if "dummy" == s:
                if "dummy_cols" in kwargs:
                    categorical_features = kwargs["dummy_cols"]
                    cols_transformer = ColumnTransformer([
                                                            ("encoder", OneHotEncoder(handle_unknown="ignore"), categorical_features)
                                                        ], 
                                                            remainder = "passthrough"
                                                        )
                    arg_list.append((s, cols_transformer))

            if "rfr" == s:
                if "rfr_kwargs" in kwargs:
                    arg_list.append(("model", RandomForestRegressor(random_state = 0, **kwargs["rfr_kwargs"])))
                else:
                    arg_list.append(("model", RandomForestRegressor(random_state = 0)))
            
            if "gbr" == s:
                if "gbr_kwargs" in kwargs:
                    arg_list.append(("model", GradientBoostingRegressor(random_state = 0, **kwargs["gbr_kwargs"])))
                else:
                    arg_list.append(("model", GradientBoostingRegressor(random_state = 0)))

            if "kmeans" == s:
                if "kmeans_kwargs" in kwargs:
                    arg_list.append(("model", KMeans(n_init = 10, random_state = 0, **kwargs["kmeans_kwargs"])))
                else: 
                    arg_list.append(("model", KMeans(n_init = 10, random_state = 0)))

        pipe = Pipeline(arg_list)
        
        if "model" in pipe.get_params() and "ensemble" in pipe["model"].__module__: #using shortcircuiting if "model" not in pipe then it won't evaluate the 2nd condition
            pipe.fit(self.data.X, self.data.y)
        else:
            pipe.fit_transform(self.data.X, self.data.y)
        return pipe
    
class GLDataHelper(DataHelper):
    def group_desc(self, grp, col):
        return GLDataHelper(self.df.groupby(grp)[col].describe())

    def get_top_vcounts(self, col = "feature_value", top_n = 10):
       return self.df[col].value_counts(normalize = True).head(top_n).index.values
        
    def get_top_diff_amt (self, col = "feature_value", top_n = 10):
        return self.df.groupby(col).idx_diff.sum().sort_values(ascending = False).head(top_n).index.values
    
    def save_pkl (self, file_name):
        pickle.dump(self, open(file_name, "wb"))

    def load_pkl (file_name):
        return pickle.load(open(file_name, "rb"))