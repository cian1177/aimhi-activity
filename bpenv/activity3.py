import pandas as pd
import json

def create_files():
    # You can set up the file creation process here
    # For example, to create empty Excel files, you can use pandas:
    pd.DataFrame(columns=['id', 'type']).to_excel('Type.xlsx', index=False)
    pd.DataFrame(columns=['id', 'form_group']).to_excel('FormGroup.xlsx', index=False)
    pd.DataFrame(columns=['id', 'Type', 'Report Forms Group', 'Column3', 'Column4']).to_excel('Report.xlsx', index=False)

def import_type():
    # Read source.csv and get the values from the Type column
    df_source = pd.read_csv('source.csv')
    types = df_source['Type'].unique()

    # Read Type.xlsx to check for existing types
    try:
        df_type = pd.read_excel('Type.xlsx')
        existing_types = df_type['type'].unique()
    except FileNotFoundError:
        existing_types = []

    # Find new types and add them to Type.xlsx with auto-incremented ids
    new_types = set(types) - set(existing_types)
    new_type_df = pd.DataFrame({'type': list(new_types)})
    if not new_type_df.empty:
        new_type_df.insert(0, 'id', range(1, 1 + len(new_type_df)))
        df_type = pd.concat([df_type, new_type_df])
        df_type.to_excel('Type.xlsx', index=False)

def import_form_group():
    # Read source.csv and get the values from the 'Report Forms Group' column
    df_source = pd.read_csv('source.csv')
    form_groups = df_source['Report Forms Group'].unique()

    # Read FormGroup.xlsx to check for existing form groups
    try:
        df_form_group = pd.read_excel('FormGroup.xlsx')
        existing_form_groups = df_form_group['form_group'].unique()
    except FileNotFoundError:
        existing_form_groups = []

    # Find new form groups and add them to FormGroup.xlsx with auto-incremented ids
    new_form_groups = set(form_groups) - set(existing_form_groups)
    new_form_group_df = pd.DataFrame({'form_group': list(new_form_groups)})
    if not new_form_group_df.empty:
        new_form_group_df.insert(0, 'id', range(1, 1 + len(new_form_group_df)))
        df_form_group = pd.concat([df_form_group, new_form_group_df])
        df_form_group.to_excel('FormGroup.xlsx', index=False)

def import_report():
    # Read source.csv and get values from all columns
    df_source = pd.read_csv('source.csv')

    # Read Type.xlsx and FormGroup.xlsx
    df_type = pd.read_excel('Type.xlsx')
    df_form_group = pd.read_excel('FormGroup.xlsx')

    # Replace values in the 'Type' column with the 'id' of the 'type' found in Type.xlsx
    df_report = df_source.copy()
    df_report['Type'] = df_report['Type'].map(df_type.set_index('type')['id'])

    # Replace values in the 'Report Forms Group' column with the 'id' of the 'form_group' found in FormGroup.xlsx
    df_report['Report Forms Group'] = df_report['Report Forms Group'].map(df_form_group.set_index('form_group')['id'])

    # Save the updated data to Report.xlsx
    df_report.to_excel('Report.xlsx', index=False)

def display_report():
    # Read FormGroup.xlsx to get original values for 'Type' and 'Report Forms Group' columns
    df_form_group = pd.read_excel('FormGroup.xlsx')
    form_group_dict = df_form_group.set_index('id')['form_group'].to_dict()
    df_type = pd.read_excel('Type.xlsx')
    type_dict = df_type.set_index('id')['type'].to_dict()

    # Convert the data to JSON format with original values for 'Type' and 'Report Forms Group' columns
    df_report = pd.read_excel('Report.xlsx')
    report_data = df_report.to_dict(orient='records')
    json_data = []
    for row in report_data:
        json_row = {}
        for key, value in row.items():
            if key in ['Type']:
                # Check if the value exists in the form_group_dict before accessing it
                json_row[key] = type_dict.get(value, value)
            elif key in ['Report Forms Group']:
                json_row[key] = form_group_dict.get(value, value)
            else:
                json_row[key] = value
        json_data.append(json_row)

    return json.dumps(json_data, indent=4)

# Call the functions to perform the operations
create_files()
import_type()
import_form_group()
import_report()
json_data = display_report()
print(json_data)
