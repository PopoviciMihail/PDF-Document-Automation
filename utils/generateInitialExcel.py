import pandas as pd

# Create sample employee data
data = {
    "Name": ["Alice Smith", "Bob Johnson", "Charlie Lee"],
    "Position": ["Software Developer", "Data Analyst", "Project Manager"],
    "StartDate": ["2025-10-01", "2025-10-05", "2025-10-10"]
}

df = pd.DataFrame(data)
df.index += 1  
df.index.name = 'ID'


df.to_excel("../employees.xlsx", index=True)

print("employees.xlsx has been created successfully!")
