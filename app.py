import json
import pandas as pd
##processing each item in json file 
def process_json(data, parent_key='', sep=':'):
    items = {}
    ##check if dictionary
    if isinstance(data, dict):
        for k, v in data.items():
            new_key = parent_key + sep + k if parent_key else k
            if isinstance(v, (dict, list)):
                items.update(process_json(v, new_key, sep))
            else:
                items[new_key] = v
    ##check for list            
    elif isinstance(data, list):
        for i, v in enumerate(data):
            new_key = parent_key + sep + str(i) if parent_key else str(i)
            if isinstance(v, (dict, list)):
                items.update(process_json(v, new_key, sep))
            else:
                items[new_key] = v
    return items
##function for writing in excel
def convert_to_excel(json_file, excel_file):
    with open(json_file) as f:
        json_data = json.load(f)
    
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as excel_writer:
        for key, value in json_data.items():
            if isinstance(value, (dict, list)):
                df = pd.DataFrame.from_dict(process_json(value), orient='index', columns=['Value'])
                df.to_excel(excel_writer, sheet_name=key)
            else:
                df = pd.DataFrame({key: [value]})
                df.to_excel(excel_writer, sheet_name='Sheet'+ key, index=False)

if __name__ == "__main__":
    json_file = 'sample.json'
    excel_file = 'output.xlsx'
    convert_to_excel(json_file, excel_file)
