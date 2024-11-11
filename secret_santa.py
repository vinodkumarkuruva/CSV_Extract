import pandas as pd
import random

# Class to represent each employee
class Employee:
    def __init__(self, name, email):
        self.name = name
        self.email = email
        self.secret_child = None
    
    def assign_secret_child(self, child):
        self.secret_child = child
        
    
    def __repr__(self):
        # Avoid recursion by only including secret_child if it's set and non-recursive
        return f"Employee(Name: {self.name}, Email: {self.email}, SecretChild: {self.secret_child.name if self.secret_child else None})"


# Class to handle file operations
class FileHandler:
    @staticmethod  # because it doesn't need access to any instance variables; it operates only on the given file_path.
    def read_excel_file(file_path):
        """Read the Excel file and return the data as a list of dictionaries."""
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
            data = df.to_dict(orient='records')
            return data
        except Exception as e:
            print(f"Error reading Excel file {file_path}: {e}")
            return None

    @staticmethod
    def save_assignments_to_csv(assignments, output_file):
        """Save the Secret Santa assignments to a CSV file."""
        try:
            output_data = [{
                'Employee_Name': emp.name,
                'Employee_EmailID': emp.email,
                'Secret_Child_Name': emp.secret_child.name,
                'Secret_Child_EmailID': emp.secret_child.email
            } for emp in assignments]

            df = pd.DataFrame(output_data)
            df.to_csv(output_file, index=False)
            print(f"Assignments saved to {output_file}")
        except Exception as e:
            print(f"Error saving assignments to CSV: {e}")


# Class to assign Secret Santa
class SecretSantaAssigner:
    def __init__(self, employees, previous_assignments):
        self.employees = [Employee(emp['Employee_Name'], emp['Employee_EmailID']) for emp in employees]
        self.previous_assignments = {assignment['Employee_Name']: assignment['Secret_Child_Name'] for assignment in previous_assignments}
    
    def assign_secret_children(self):
        """Generate Secret Santa assignments based on employee data and previous assignments."""
        remaining_employees = self.employees[:]
        
        for employee in self.employees:
            possible_children = [
                e for e in remaining_employees
                if e != employee and e.name != self.previous_assignments.get(employee.name)
            ]
            
            if not possible_children:
                print(f"Could not assign a secret child to {employee.name}. Assignment failed.")
                return None
            
            secret_child = random.choice(possible_children)
            employee.assign_secret_child(secret_child)
            remaining_employees.remove(secret_child)
        
        return self.employees


# Main application class to coordinate the Secret Santa process
class SecretSantaApp:
    def __init__(self, employee_file, previous_assignments_file, output_file):
        self.employee_file = employee_file
        self.previous_assignments_file = previous_assignments_file
        self.output_file = output_file

    def run(self):
        # Load data from files
        employees_data = FileHandler.read_excel_file(self.employee_file)
        previous_assignments_data = FileHandler.read_excel_file(self.previous_assignments_file)
        
        if not employees_data or not previous_assignments_data:
            print("Failed to load data. Exiting.")
            return
        
        # Assign Secret Santa pairs
        assigner = SecretSantaAssigner(employees_data, previous_assignments_data)
        assignments = assigner.assign_secret_children()
        
        if assignments is not None:
            # Save assignments to a CSV file
            FileHandler.save_assignments_to_csv(assignments, self.output_file)


# Example usage of the OOP-based Secret Santa program
if __name__ == "__main__":
    employee_file = r'files\Employee-List.xlsx'
    previous_assignments_file = r'files\Secret-Santa-Game-Result-2023.xlsx'
    output_file = r'files\Secret-Santa-Assignments-2024.csv'

    # Initialize and run the Secret Santa app
    app = SecretSantaApp(employee_file, previous_assignments_file, output_file)
    app.run()
