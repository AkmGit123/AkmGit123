#author ashit kumar maharana
import pandas as pd
from openpyxl import Workbook
import chardet

def detect_encoding(file_path):
    with open(file_path, 'rb') as csvfile:
        result = chardet.detect(csvfile.read())
        return result['encoding']

def read_test_cases_from_csv(file_path):
    encoding = detect_encoding(file_path)  
    df = pd.read_csv(file_path, encoding=encoding)

    test_cases = []
    for index, row in df.iterrows():
        steps = row["Test Steps"] if pd.notna(row["Test Steps"]) else ""
        expected_results = row["Expected Results"] if pd.notna(row["Expected Results"]) else ""

        test_cases.append({
            "test_case_id": row["Test Case ID"],
            "test_scenario": row["Test Scenario"],
            "steps": steps.split("\n"),
            "expected_results": expected_results.split("\n"),
            "rtm_id": row.get("RTM ID", "")
        })
    return test_cases

def write_test_cases_to_excel(test_cases, output_file_path):
    wb = Workbook()
    work_sheet = wb.active
    work_sheet.title = "Test Cases"

    work_sheet.append(["Test Case ID", "Test Scenario", "Test Steps", "Expected Result", "RTM ID"])

    for case in test_cases:
        for step, result in zip(case["steps"], case["expected_results"]):
            work_sheet.append([
                case["test_case_id"],
                case["test_scenario"],
                step,
                result,
                case["rtm_id"]
            ])

    wb.save(output_file_path)
    print(f"Data successfully written to {output_file_path}")

if __name__ == "__main__":
    input_file_path = input("Enter the path to the CSV file: ")
    output_file_path = input("Enter the desired output Excel file path: ")

    try:
        test_cases = read_test_cases_from_csv(input_file_path)
        write_test_cases_to_excel(test_cases, output_file_path)
    except FileNotFoundError:
        print("Error: The specified file was not found.")
    except KeyError as e:
        print(f"Error: Missing expected column in CSV file: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
