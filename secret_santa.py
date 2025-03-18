
import pandas as pd
import random

employee_file = "Employee-List.xlsx"
last_year_file = "Secret-Santa-Game-Result-2023.xlsx"
output_file = "Secret-Santa-Assignments-2024.xlsx"

def read_excel(file_path):
    """Reads an Excel file and returns a DataFrame."""
    try:
        return pd.read_excel(file_path)
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found.")
        return pd.DataFrame()
    except Exception as e:
        print(f"Error reading '{file_path}': {e}")
        return pd.DataFrame()

def get_previous_assignments(last_year_data):
    """Creates a dictionary of last year's assignments for quick lookup."""
    return dict(zip(last_year_data["Employee_Name"], last_year_data["Secret_Child_Name"]))

def assign_secret_santa(employees, last_year_assignments):
    """Assigns Secret Santa pairs, avoiding previous year's assignments."""
    available_recipients = employees.copy()
    random.shuffle(available_recipients)

    assignments = []
    for _ in range(100):  # Try 100 times to find a valid assignment
        temp_recipients = available_recipients.copy()
        random.shuffle(temp_recipients)

        temp_assignments = []
        valid = True
        for giver in employees:
            giver_name = giver["Employee_Name"]

            # Filter valid recipients
            possible_recipients = [
                recipient for recipient in temp_recipients
                if recipient["Employee_Name"] != giver_name and 
                   recipient["Employee_Name"] != last_year_assignments.get(giver_name, None)
            ]

            if not possible_recipients:
                valid = False
                break

            chosen_recipient = random.choice(possible_recipients)
            temp_recipients.remove(chosen_recipient)

            temp_assignments.append({
                "Employee_Name": giver_name,
                "Employee_EmailID": giver["Employee_EmailID"],
                "Secret_Child_Name": chosen_recipient["Employee_Name"],
                "Secret_Child_EmailID": chosen_recipient["Employee_EmailID"]
            })

        if valid:
            return temp_assignments

    print("Error: Could not find a valid Secret Santa assignment after multiple attempts.")
    return []

def write_excel(file_path, data):
    """Writes the assignments to an Excel file."""
    try:
        df = pd.DataFrame(data)
        df.to_excel(file_path, index=False)
        print(f"Secret Santa assignments saved to '{file_path}'.")
    except Exception as e:
        print(f"Error writing '{file_path}': {e}")

def main():
    employees_df = read_excel(employee_file)
    last_year_df = read_excel(last_year_file)

    if employees_df.empty:
        print("No employee data available. Exiting.")
        return

    employees = employees_df.to_dict(orient="records")
    last_year_assignments = get_previous_assignments(last_year_df)

    secret_santa_assignments = assign_secret_santa(employees, last_year_assignments)

    if secret_santa_assignments:
        write_excel(output_file, secret_santa_assignments)

if __name__ == "__main__":
    main()
