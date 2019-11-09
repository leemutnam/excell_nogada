import pandas as pd
import os

project_dir_path = os.path.dirname(os.path.abspath(__file__))
origin_dir_name = '/excell_origin'
origin_dir_path = os.path.abspath(''.join([project_dir_path, origin_dir_name]))
result_dir_name = '/excell_result'
result_dir_path = os.path.abspath(''.join([project_dir_path, result_dir_name]))

origin_file_list = os.listdir(origin_dir_path)
latest_idx = len(origin_file_list) - 1
origin_file_name = origin_file_list[latest_idx]

origin_file_path = ''.join([origin_dir_path, '/', origin_file_name])
xl = pd.ExcelFile(origin_file_path)
sheet_names = xl.sheet_names

addr_keys = ['주소', '지번주소', '도로명주소']
file_type = '.xlsx'

for name in sheet_names:
    data = pd.read_excel(origin_file_path, sheet_name=name)
    data_columns = data.columns
    keys_existence = False

    for key in addr_keys:
        if key in data_columns:
            keys_existence = True

    if keys_existence:
        result_file = ''.join([result_dir_path, '/', name, file_type])
        print(result_file)
        data.to_excel(result_file, sheet_name='Sheet1')
