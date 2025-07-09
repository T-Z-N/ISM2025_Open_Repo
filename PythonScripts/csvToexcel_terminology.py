import pandas as pd

csv_file_path = 'query-result.csv'
excel_file_path = 'class_annotations.xlsx'

try:
    df = pd.read_csv(csv_file_path)

    df.fillna('', inplace=True)

    with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Annotations', index=False)

        worksheet = writer.sheets['Annotations']


        if not df.empty:
            current_class_uri = None
            merge_start_row = 2  

            for idx in range(len(df)):

                excel_current_row = idx + 2

                if current_class_uri is None or df.loc[idx, 'class'] != current_class_uri:
                    if current_class_uri is not None and excel_current_row - 1 > merge_start_row:
                        worksheet.merge_cells(start_row=merge_start_row, start_column=1,
                                              end_row=excel_current_row - 1, end_column=1)

                    current_class_uri = df.loc[idx, 'class']
                    merge_start_row = excel_current_row

            if current_class_uri is not None and (len(df) + 1) > merge_start_row: 
                worksheet.merge_cells(start_row=merge_start_row, start_column=1,
                                      end_row=len(df) + 1, end_column=1)

        print(f"Excel File '{excel_file_path}' succesfuly created.")

except FileNotFoundError:
    print(f"Error: '{csv_file_path}' file is missing.")
except Exception as e:
    print(f"There is an error: {e}")