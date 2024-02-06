import pandas as pd

# Read the existing Excel file
input_file = 'mitre.xlsx'
df = pd.read_excel(input_file)

# Create a new DataFrame to store the formatted data
formatted_df = pd.DataFrame(columns=['Tactics', 'Technique', 'Description', 'Sub-Technique', 'Detection'])

# Iterate through each row in the original DataFrame
for index, row in df.iterrows():
   
    # tactic_id = row['tactics']
    technique_id = row['ID']
    technique_name = row['name']
    description = row['description']
    is_sub_technique = row['is sub-technique']
    
    # Check if it's a sub-technique
    if is_sub_technique:
        sub_technique_id = f"{technique_id} .{index + 1:03d}"  # Example: T1548.002
        formatted_df = formatted_df.append({'Tactics': row['tactics'],
                                            'Technique': technique_id.split('.')[0],
                                            'Description': description,
                                            'Sub-Technique': technique_id,
                                            'Detection': row['detection']},
                                           ignore_index=True)
    else:
        formatted_df = formatted_df.append({'Tactics': row['tactics'],
                                            'Technique': technique_id.split('.')[0],
                                            'Description': description,
                                            'Sub-Technique': ' ',
                                            'Detection': row['detection']},
                                           ignore_index=True)
        
# Sort the DataFrame by Tactics (you can adjust the sorting criteria as needed)
formatted_df.sort_values(by='Tactics', inplace=True)

# Write the formatted data to a new Excel file
output_file = 'file.xlsx'
formatted_df.to_excel(output_file, index=False)

print(f"Formatted data saved to {output_file}")
