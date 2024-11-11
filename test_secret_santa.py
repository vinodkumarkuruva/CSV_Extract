import pytest
from unittest.mock import patch, MagicMock
import pandas as pd
import random
from io import StringIO

# Import the necessary classes from your application code
from secret_santa import FileHandler, Employee, SecretSantaAssigner, SecretSantaApp  # Update with the actual module name

# Mock employee and previous assignment data
mock_employees_data = [
    {'Employee_Name': 'Alice', 'Employee_EmailID': 'alice@example.com'},
    {'Employee_Name': 'Bob', 'Employee_EmailID': 'bob@example.com'},
    {'Employee_Name': 'Charlie', 'Employee_EmailID': 'charlie@example.com'}
]

mock_previous_assignments_data = [
    {'Employee_Name': 'Alice', 'Secret_Child_Name': 'Bob'},
    {'Employee_Name': 'Bob', 'Secret_Child_Name': 'Charlie'},
    {'Employee_Name': 'Charlie', 'Secret_Child_Name': 'Alice'}
]

# Define a fixture to mock the file handler
@pytest.fixture
def mock_file_handler():
    with patch.object(FileHandler, 'read_excel_file', side_effect=[mock_employees_data, mock_previous_assignments_data]):
        with patch.object(FileHandler, 'save_assignments_to_csv') as mock_save_csv:
            yield mock_save_csv

# Test the SecretSantaAssigner logic
def test_assign_secret_children(mock_file_handler):
    # Create a SecretSantaAssigner instance
    assigner = SecretSantaAssigner(mock_employees_data, mock_previous_assignments_data)
    
    # Run the assignment process
    assignments = assigner.assign_secret_children()
    
    # Check that all employees have been assigned a secret child and the assignments are correct
    assert len(assignments) == 3  # Three employees should be assigned
    for employee in assignments:
        assert employee.secret_child is not None  # Every employee should have a secret child
    
    # Check the correctness of the secret child assignments
    assert assignments[0].secret_child.name != 'Alice'  # Alice shouldn't be her own secret child
    assert assignments[1].secret_child.name != 'Bob'  # Bob shouldn't be his own secret child
    assert assignments[2].secret_child.name != 'Charlie'  # Charlie shouldn't be his own secret child

# Test the FileHandler saving to CSV
def test_save_assignments_to_csv(mock_file_handler):
    # Mock the list of employee objects with their secret children
    employee1 = Employee('Alice', 'alice@example.com')
    employee2 = Employee('Bob', 'bob@example.com')
    employee3 = Employee('Charlie', 'charlie@example.com')

    # Assign secret children to employees
    employee1.assign_secret_child(employee2)
    employee2.assign_secret_child(employee3)
    employee3.assign_secret_child(employee1)

    # Create a list of employee objects
    assignments = [employee1, employee2, employee3]

    # Convert Employee objects to dictionaries for saving
    assignments_dict = [
        {
            'Employee_Name': emp.name,
            'Employee_EmailID': emp.email,
            'Secret_Child_Name': emp.secret_child.name if emp.secret_child else None,
            'Secret_Child_EmailID': emp.secret_child.email if emp.secret_child else None
        }
        for emp in assignments
    ]

    # Call the method to save assignments to CSV
    FileHandler.save_assignments_to_csv(assignments_dict, 'secret_santa_assignments.csv')  # Pass the dict list

    # Verify that the mock_save_csv method was called
    mock_file_handler.assert_called_once()

    # Verify the arguments passed to the mock function
    args, kwargs = mock_file_handler.call_args
    assert args[0] == assignments_dict  # Compare the correct object (list of dictionaries)


def test_secret_santa_app(mock_file_handler):
    # Mock the app initialization and run method
    app = SecretSantaApp('employee_file.xlsx', 'previous_assignments_file.xlsx', 'output_file.csv')

    # Use patch to mock the save_assignments_to_csv method
    with patch.object(FileHandler, 'save_assignments_to_csv', mock_file_handler):
        app.run()

    # Verify that save_assignments_to_csv was called
    mock_file_handler.assert_called_once()
