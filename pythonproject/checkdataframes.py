"""This function will be used by python program check file types and use pandas to read different file types."""
# Imports
import pandas as pd



def checkfiletype(file_path):
    # Store file name as a string
    filename = r"{}".format(file_path)

    if ".xlsx" in filename:
        data_frame = pd.read_excel(filename)
    elif ".json" in filename:
        data_frame = pd.read_json(filename)
    elif ".csv" in filename:
        data_frame = pd.read_csv(filename)
    elif ".txt" in filename:
        data_frame = pd.read_csv(filename)
    else:
        print("Error")

    return data_frame


def get_columns(file_path):
    filename = r"{}".format(file_path)
    if ".xlsx" in filename:
        data_frame = pd.read_excel(filename)
    elif ".json" in filename:
        data_frame = pd.read_json(filename)
    elif ".csv" in filename:
        data_frame = pd.read_csv(filename)
    elif ".txt" in filename:
        data_frame = pd.read_csv(filename)
    else:
        print("Error")
    column_list = list(data_frame.columns.values)

    return column_list
