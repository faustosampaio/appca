import pandas as pd

# Example DataFrame
data = {'Column1': ['Banana', 'Apple', 'Berry', 'Cherry'],
        'Column2': ['Bear', 'Dog', 'Cat', 'Elephant']}
df = pd.DataFrame(data)

# Check if column names start with 'B'
start_with_B = df.columns[df.columns.str.startswith('B')]

print(start_with_B)
