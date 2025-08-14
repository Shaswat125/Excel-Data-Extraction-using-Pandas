import pandas as pd
import numpy as np
from faker import Faker

def generate_excel_data(filename='test_data.xlsx', rows=1000, cols=20):
    fake = Faker()
    np.random.seed(42)

    # Define sample columns
    columns = [
        'ID', 'Name', 'Email', 'City', 'Country', 'Join Date', 'Last Login',
        'Status', 'Age', 'Score', 'Salary', 'Bonus', 'Department', 'Role',
        'Manager', 'Experience', 'Rating', 'Projects', 'Working Hours', 'Active'
    ]

    if cols < len(columns):
        columns = columns[:cols]

    data = []
    for i in range(rows):
        row = [
            i + 10000,
            fake.name(),
            fake.email(),
            fake.city(),
            fake.country(),
            fake.date_between(start_date='-10y', end_date='-1y'),
            fake.date_between(start_date='-1y', end_date='today'),
            np.random.choice(['Active', 'Inactive', 'Pending']),
            np.random.randint(21, 60),
            np.round(np.random.uniform(50, 100), 2),
            np.round(np.random.uniform(30000, 120000), 2),
            np.round(np.random.uniform(1000, 10000), 2),
            np.random.choice(['Sales', 'Tech', 'Marketing', 'Support', 'HR']),
            np.random.choice(['Analyst', 'Consultant', 'Manager', 'Lead', 'Associate', 'Director']),
            fake.name(),
            np.random.randint(1, 15),
            np.round(np.random.uniform(1.0, 5.0), 1),
            np.random.randint(0, 20),
            np.round(np.random.uniform(20, 60), 1),
            np.random.choice([True, False])
        ]
        data.append(row[:cols])  # trim to column count

    df = pd.DataFrame(data, columns=columns)
    df.to_excel(filename, index=False)
    print(f"Generated Required testing excel file with given file name: '{filename}'")

generate_excel_data(filename='Employees.xlsx', cols=20, rows=2000)