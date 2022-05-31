from markupsafe import Markup
from sklearn.linear_model import LogisticRegression
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import PolynomialFeatures
from sklearn.metrics import f1_score,mean_squared_error
from sklearn.pipeline import make_pipeline, _fit_transform_one
import numpy as np
import pandas as pd


if __name__ == "__main__":



    '''x_data = np.array(range(1,13)).reshape(-1, 1)
    y_data = np.array([1,0,1,1,1,1,0,1,1,1,0,0])
    print(x_data,y_data)
    X_train, X_test, y_train, y_test = train_test_split(x_data, y_data, test_size = 0.26)'''


