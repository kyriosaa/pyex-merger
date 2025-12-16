# !!! make sure u close all the excel files before running this !!!

import pandas as pd
import glob
import os
import config
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill

folder_path = config.PATH
output_file = 'MasterReport.xlsx'

def merge_excel(path, output_name):
    all_files = glob.glob(os.path.join(path, "*.xlsx")) # os.path.join to account for Mac OS
    print(f"{len(all_files)} excel files to process")
    all_data = [] # hold tables from each file
    
    for filename in all_files:
        try:
            wb = load_workbook(filename, data_only=True) # data_only=True ignores formulas
            sheet = wb.active
            row_data = {}
            row_data['Source File'] = os.path.basename(filename)
            
            # extract cells
            for col_name, cell_address in config.extraction_cells.items():
                try:
                    val = sheet[cell_address].value
                    row_data[col_name] = val
                except Exception as error:
                    row_data[col_name] = None
                    print(f"Could not read cell {cell_address} in {filename}")
            
            all_data.append(row_data)
            print(f"Read {os.path.basename(filename)}")
            wb.close()
        except Exception as error:
            row_data[col_name] = None
            print(f"Error reading {filename}: {error}")
        
    # combine tables into one
    if all_data:
        combined_df = pd.DataFrame(all_data)
        
        # export to new excel file
        output_path = os.path.join(path, output_name)
        
        # format output
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            combined_df.to_excel(writer, index=False, sheet_name='Master Report')
            worksheet = writer.sheets['Master Report']
            
            # auto calc column width
            for i, column in enumerate(combined_df.columns):
                max_len = len(str(column))
                if not combined_df.empty:
                    data_len = combined_df[column].astype(str).map(len).max()
                    if not pd.isna(data_len):
                        max_len = max(max_len, data_len)
                
                column_letter = get_column_letter(i + 1)
                
                # formatting requirements
                is_source_file = column == "Source File"
                should_center = column in config.centered_columns
                should_wrap = max_len + 2 > 50
                should_green = column in getattr(config, 'green_columns', [])
                should_red = column in getattr(config, 'red_columns', []) # btw this only turns red if there is text present in the cell (some days theres no event)
                
                # column width wrapping
                if should_wrap:
                    worksheet.column_dimensions[column_letter].width = 50
                else:
                    worksheet.column_dimensions[column_letter].width = max_len + 2
                    
                # green highlighting (for income col)
                green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
                # red highlighting (for daily events)
                red_fill = PatternFill(start_color="FFC1C1", end_color="FFC1C1", fill_type="solid")
                    
                # apply
                for cell in worksheet[column_letter]:
                    # default
                    horiz = None
                    vert = None
                    
                    
                    if should_center:
                        horiz = 'center'
                        vert = 'center'
                    elif is_source_file:
                        vert = 'center'
                        
                    if should_wrap:
                        if not vert: vert = 'center'
                        
                    if should_center or should_wrap or is_source_file:
                        cell.alignment = Alignment(horizontal=horiz, vertical=vert, wrap_text=should_wrap)
                
                    if should_green:
                        cell.fill = green_fill
                        
                    if should_red and cell.value: # cell.value to check for text inside the cell
                        cell.fill = red_fill
                
        print(f"[SUCCESS] All files merged into: {output_path}")
    else:
        print(f"[ERROR] No data found to merge")

if __name__ == "__main__":
    merge_excel(folder_path, output_file)