"""
Excelcreator.py

A modern, clean, and well-documented module for generating formatted Excel files from pandas DataFrames.
Designed for use with Tableau Calculation Extractor.


"""

import os
import pandas as pd


def create_output_paths(base_dir, filename_base):
    """
    Ensures output directory exists and returns full file paths for Excel and HTML outputs.
    """
    output_dir = os.path.join(base_dir, 'outputs')
    os.makedirs(output_dir, exist_ok=True)
    excel_path = os.path.join(output_dir, f"{filename_base}_calculations_table.xlsx")
    html_path = os.path.join(output_dir, f"{filename_base}_lineage_diagram.html")
    return excel_path, html_path


def create_new_file_paths(filename_base):
    """
    Legacy function for backward compatibility.
    Creates output paths relative to the current working directory.
    Returns tuple of excel_path.
    """
    output_dir = os.path.join(os.getcwd(), 'outputs')
    os.makedirs(output_dir, exist_ok=True)
    excel_path = os.path.join(output_dir, f"{filename_base}.xlsx")
    return excel_path


def format_excel(df, excel_path, column_widths=None, sheet_name='Calculations'):
    """
    Writes a DataFrame to a professionally formatted Excel file using xlsxwriter.
    """
    if column_widths is None:
        column_widths = [15, 12, 15, 50, 20, 25, 30]  # Default widths for 7 columns

    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        # Header format
        header_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#B7DEE8', 'border': 1,
            'align': 'center', 'valign': 'vcenter', 'font_size': 11
        })
        # Data format
        data_fmt = workbook.add_format({
            'border': 1, 'text_wrap': True, 'valign': 'top', 'font_size': 10
        })

        # Set column widths and formats
        for col_idx, width in enumerate(column_widths):
            worksheet.set_column(col_idx, col_idx, width, data_fmt)
        # Apply header format
        for col_idx, value in enumerate(df.columns):
            worksheet.write(0, col_idx, value, header_fmt)

        # Page setup
        worksheet.set_paper(9)  # A4
        worksheet.set_landscape()
        worksheet.fit_to_pages(1, 0)
        worksheet.set_margins(left=0.5, right=0.5, top=0.75, bottom=0.75)
        worksheet.set_footer('&CPage &P of &N')

        # Freeze header row
        worksheet.freeze_panes(1, 0)


def save_calculations_to_excel(df, base_dir, filename_base):
    """
    High-level function to save calculations DataFrame to a formatted Excel file.
    Returns the path to the created Excel file.
    """
    excel_path, _ = create_output_paths(base_dir, filename_base)
    format_excel(df, excel_path)
    return excel_path


def create_excel_from_dfs(dfs_to_use, excel_path, column_widths=None):
    """
    Creates an Excel file with multiple sheets from a list of DataFrame dictionaries.
    
    Parameters:
    -----------
    dfs_to_use : list of dict
        List of dictionaries containing:
        - 'df_to_use': pandas DataFrame
        - 'sheetName': name for the Excel sheet
        - 'normalColWidth': list of column widths (optional)
        - 'color': background color for headers (optional, default '#B7DEE8')
    excel_path : str
        Full path where the Excel file should be saved
    column_widths : list, optional
        Default column widths if not specified per sheet
    """
    if column_widths is None:
        column_widths = [15, 12, 15, 50, 20, 25, 30]
    
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        for sheet_config in dfs_to_use:
            df = sheet_config.get('df_to_use')
            sheet_name = sheet_config.get('sheetName', 'Sheet')
            widths = sheet_config.get('normalColWidth', column_widths)
            color = sheet_config.get('color', '#B7DEE8')
            
            # Write DataFrame to sheet
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            
            # Header format with custom color
            header_fmt = workbook.add_format({
                'bold': True, 'bg_color': color, 'border': 1,
                'align': 'center', 'valign': 'vcenter', 'font_size': 11
            })
            # Data format
            data_fmt = workbook.add_format({
                'border': 1, 'text_wrap': True, 'valign': 'top', 'font_size': 10
            })
            
            # Set column widths and formats
            for col_idx, width in enumerate(widths):
                if col_idx < len(df.columns):
                    worksheet.set_column(col_idx, col_idx, width, data_fmt)
            
            # Apply header format
            for col_idx, value in enumerate(df.columns):
                worksheet.write(0, col_idx, value, header_fmt)
            
            # Page setup (A4, landscape)
            worksheet.set_paper(9)
            worksheet.set_landscape()
            worksheet.fit_to_pages(1, 0)
            worksheet.set_margins(left=0.5, right=0.5, top=0.75, bottom=0.75)
            
            # Add footer if specified
            footer = sheet_config.get('footer', 'Page &P of &N')
            worksheet.set_footer(f'&C{footer}')
            
            # Freeze header row
            worksheet.freeze_panes(1, 0)


# Example usage (for testing or integration):
if __name__ == "__main__":
    # Example DataFrame
    data = {
        'Field_Name': ['Sales', 'Profit'],
        'DataType': ['float', 'float'],
        'Type': ['Calculated_Field', 'Default_Field'],
        'Calculation': ['[Quantity] * [Unit Price]', ''],
        'Field_ID': ['[Sales]', '[Profit]'],
        'Datasource': ['Orders', 'Orders'],
        'Worksheets': ['Sheet1, Sheet2', 'Sheet1']
    }
    df = pd.DataFrame(data)
    save_calculations_to_excel(df, os.getcwd(), 'TestWorkbook')
    print("Excel file created.")
