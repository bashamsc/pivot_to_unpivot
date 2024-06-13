import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.pivot.fields import Missing

file_path = 'path/to/your/data.xlsx'

workbook = load_workbook(file_path)
worksheet = workbook['Sheet_Name']

# Name of desired pivot table (the same name that appears within Excel)

for p in worksheet._pivots:
    print(p.name)
    pivot_name=p.name
	
	
# Extract the pivot table object from the worksheet
pivot_table = [p for p in worksheet._pivots if p.name == pivot_name][0]

# Printing the field names
for field in pivot_table.cache.cacheFields:
    if field.sharedItems.count > 0:
        print(field.name)
		


# Extract a dict of all cache fields and their respective values
fields_map = {}
for field in pivot_table.cache.cacheFields:
    if field.sharedItems.count > 0:
        # take care of cases where f.v returns an AttributeError because the cell is empty
        # fields_map[field.name] = [f.v for f in field.sharedItems._fields]
        l = []
        for f in field.sharedItems._fields:
            try:
                l += [f.v]
            except AttributeError:
                l += [""]
        fields_map[field.name] = l
		

# Extract all rows from cache records. Each row is initially parsed as a dict
column_names = [field.name for field in pivot_table.cache.cacheFields]
rows = []
for record in pivot_table.cache.records.r:
    # If some field in the record in missing, we replace it by NaN
    record_values = [
        field.v if not isinstance(field, Missing) else np.nan for field in record._fields
    ]

    row_dict = {k: v for k, v in zip(column_names, record_values)}

    # Shared fields are mapped as an Index, so we replace the field index by its value
    for key in fields_map:
        row_dict[key] = fields_map[key][row_dict[key]]

    rows.append(row_dict)
	
#Converting dict to pandas dataframe

df = pd.DataFrame.from_dict(rows)

#Writing to an csv file

df.to_csv("output.csv", index=False)
