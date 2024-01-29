import pandas as pd
import os
import numpy as np

def merge_excel_files(file_paths):
    all_data = []
    for file_path in file_paths:
        df = pd.read_excel(file_path, engine='openpyxl') 
        all_data.append(df)
    all_data = [df.assign(Student_ID=df.index) for df in all_data]
    return pd.concat(all_data, ignore_index=True)

def calculate_overall_scores(data):
    data['Completion time'] = pd.to_numeric(data['Completion time'], errors='coerce')
    email_column = 'Email'
    name_column = 'Name'
    data[email_column] = data[email_column].astype(str)
    data[name_column] = data[name_column].astype(str)
    
    start_time_column = 'Start time'
    if start_time_column in data.columns:
        data[start_time_column] = pd.to_datetime(data[start_time_column], errors='coerce')
    numeric_columns = data.select_dtypes(include=[np.number]).columns
    data[numeric_columns] = data[numeric_columns].apply(pd.to_numeric, errors='coerce')
    data['Overall_Score'] = data[numeric_columns].sum(axis=1)
    return data

def identify_exceptional_cases(data):
    exceptional_cases = data[data.iloc[:, 2:].count(axis=1) < len(data.columns) - 2]
    return exceptional_cases

def generate_reports(data, output_path, quiz_names):
    id_column_name = 'Student_ID'  # Replace with the correct column name
    overall_summary = data[[id_column_name, 'Overall_Score']].sort_values(by='Overall_Score', ascending=False)
    overall_summary.to_excel(os.path.join(output_path, 'Overall_Summary.xlsx'), index=False)
    exceptional_cases = identify_exceptional_cases(data)
    exceptional_cases.to_excel(os.path.join(output_path, 'Exceptional_Cases.xlsx'), index=False)

excel_files = [
    r"C:\Python\Basic_Python_Quiz_For_Beginners(1-8).xlsx",
    r"C:\Python\Conditions_quiz(1-6).xlsx",
    r"C:\Python\Dictionaries_quiz(1-6).xlsx",
    r"C:\Python\List_Quiz(1-8).xlsx",
    r"C:\Python\Loops_quiz(1-6).xlsx",
    r"C:\Python\Numbers_and_Booleans_quiz(1-5).xlsx",
    r"C:\Python\Strings_Quiz(1-5).xlsx",
    r"C:\Python\Variables_and_Types_quiz(1-6).xlsx"
]
merged_data = merge_excel_files(excel_files)
quiz_names = [f"Quiz {i+1}" for i in range(len(excel_files))]
merged_data['Quiz_Name'] = np.repeat(quiz_names, len(merged_data) // len(quiz_names) + 1)[:len(merged_data)]
merged_data = calculate_overall_scores(merged_data)
output_path = r"C:\Python"
os.makedirs(output_path, exist_ok=True)
generate_reports(merged_data, output_path, quiz_names)
merged_data.to_excel(os.path.join(output_path, 'Merged_Results.xlsx'), index=False)
