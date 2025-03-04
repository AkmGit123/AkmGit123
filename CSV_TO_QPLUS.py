import pandas as pd
from openpyxl import Workbook

# Function to read the test cases from the CSV file
def read_test_cases_from_csv(file_path):
    # Read the CSV file into a pandas DataFrame
    df = pd.read_csv(file_path)

    test_cases = []
    # Iterate over each row in the DataFrame
    for index, row in df.iterrows():
        # Safely handle NaN values by replacing them with empty strings
        steps = row["Test Steps"] if pd.notna(row["Test Steps"]) else ""
        expected_results = row["Expected Results"] if pd.notna(row["Expected Results"]) else ""

        test_cases.append({
            "test_case_id": row["Test Case ID"],
            "test_scenario": row["Test Scenario"],
            "steps": steps.split("\n"),  # Split steps by newline
            "expected_results": expected_results.split("\n"),  # Split expected results by newline
            "rtm_id": row.get("RTM ID", "")  # Handle RTM ID if it exists
        })
    return test_cases

# Function to write the test cases into an Excel sheet using openpyxl
def write_test_cases_to_excel(test_cases, output_file_path):
    # Create a new workbook and select the active sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Test Cases"

    # Write the header row
    ws.append(["Test Case ID", "Test Scenario", "Test Steps", "Expected Result", "RTM ID"])

    # Iterate over the test cases and add them to the Excel sheet
    for case in test_cases:
        for step, result in zip(case["steps"], case["expected_results"]):
            ws.append([
                case["test_case_id"],  # Test Case ID (repeated for each step)
                case["test_scenario"],  # Test Scenario (repeated for each step)
                step,  # Test Step
                result,  # Expected Result
                case["rtm_id"]  # RTM ID (repeated for each step, or blank if missing)
            ])

    # Save the workbook to the output file
    wb.save(output_file_path)
    print(f"Data successfully written to {output_file_path}")

# Main execution
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
