"""A snippet to auto assign marks to UIU students' home assignments
Works with the eLMS auto-generated Grade-sheet in CSV Format"""

import pandas as pd

from pathlib import Path

print("Please save & close the file before pasting the path. Enter CSV full file path.")
filePath = input("Please add .csv at the end if not already pasted: ")

fileName = Path(filePath)
# print(fileName)

dataframe = pd.read_csv(fileName, index_col=0)

fullMarks_condition = dataframe['Status'].str.contains("Submitted")  # Those who submitted
noMarks_condition = ~dataframe['Status'].str.contains("Submitted")  # Those who didn't submit
# print(grading_condition)

dataframe.loc[fullMarks_condition, 'Grade'] = dataframe['Maximum Grade']  # Assign Full Marks
dataframe.loc[noMarks_condition, 'Grade'] = 0  # Assign 0 Marks

dataframe.to_csv(fileName)
print(dataframe[['Full name', 'Grade']])
