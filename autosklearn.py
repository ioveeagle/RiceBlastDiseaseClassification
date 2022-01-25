# coding=utf-8
import pandas as pd
import numpy as np
import sklearn.model_selection
import sklearn.datasets
import sklearn.metrics
from sklearn.metrics import confusion_matrix, recall_score, classification_report, precision_score
from sklearn import preprocessing
from sklearn.model_selection import train_test_split

import autosklearn.classification

dataFile = 'newDataTo106.csv'
resultFile = 'newDataResultTo106.csv'
testFile = '107Data.csv'
testResultFile = '107DataResult.csv'

def execute():

  records = pd.read_csv(dataFile)
  results = np.ravel(pd.read_csv(resultFile))
  testRecords = pd.read_csv(testFile)
  testResults = np.ravel(pd.read_csv(testResultFile))

  X_train = records
  y_train = results
  X_test = testRecords
  y_test = testResults
  
  automl = autosklearn.classification.AutoSklearnClassifier(
      time_left_for_this_task=120,
      per_run_time_limit=30,
      delete_tmp_folder_after_terminate=False,
      resampling_strategy='cv',
      resampling_strategy_arguments={'folds': 5},
    )

  print(X_train)
  print(y_train)
  # fit() changes the data in place, but refit needs the original data. We
  # therefore copy the data. In practice, one should reload the data
  automl.fit(X_train.copy(), y_train.copy(), dataset_name='digits')
  # During fit(), models are fit on individual cross-validation folds. To use
  # all available data, we call refit() which trains all models in the
  # final ensemble on the whole dataset.
  automl.refit(X_train.copy(), y_train.copy())

  print(automl.show_models())

  predictions = automl.predict(X_test)
  print(y_test)
  print(predictions)
  print("Accuracy score", sklearn.metrics.accuracy_score(y_test, predictions))
  print(classification_report(y_test, predictions))

if __name__ == "__main__":
  execute()