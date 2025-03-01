import os
import platform
import time
import pandas as pd
import numpy as np

def clear_screen():
    """
    Clears the console on Windows, Linux, or Mac.
    """
    if platform.system() == "Windows":
        os.system("cls")
    else:
        os.system("clear")

def csv_to_xlsx_dynamic(input_file):
    # Attempt to load CSV using common delimiters, using the fast 'c' engine.
    delimiters = [',', ';', '\t']
    df = None
    for delim in delimiters:
        try:
            df = pd.read_csv(input_file, delimiter=delim, engine='c')
            break
        except Exception:
            continue

    if df is None:
        print("Failed to read the CSV file with known delimiters.")
        return

    # Replace INF and -INF with NaN, so they become valid "missing" values.
    df.replace([np.inf, -np.inf], np.nan, inplace=True)

    # Fill numeric columns' NaNs with 0, and object/text columns' NaNs with ''.
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df.loc[:, col] = df[col].fillna(0)
        else:
            df.loc[:, col] = df[col].fillna('')

    # Define output file location
    output_file = './output/output.xlsx'
    
    # Use the XlsxWriter engine
    writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

    # Write the DataFrame to Excel
    df.to_excel(writer, sheet_name='Full Data', index=False)

    workbook = writer.book
    worksheet = writer.sheets['Full Data']

    # Freeze the top header row
    worksheet.freeze_panes(1, 0)

    # Define header format: bold, white text on blue background, Arial 12pt with borders.
    header_format = workbook.add_format({
        'bold': True,
        'font_color': 'white',
        'bg_color': '#0B57A4',
        'font_name': 'Arial',
        'font_size': 12,
        'border': 1,
        'align': 'left',
        'valign': 'vcenter'
    })

    # Apply header formatting and adjust column widths.
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        col_width = len(str(value)) + 8
        worksheet.set_column(col_num, col_num, col_width)

    # Define data cell format with borders and Arial 12pt font.
    data_format = workbook.add_format({
        'font_name': 'Arial',
        'font_size': 12,
        'border': 1,
        'align': 'left',
        'valign': 'vcenter'
    })

    # Re-write data cells with the custom data format (note: header row is row 0).
    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            worksheet.write(i + 1, j, df.iat[i, j], data_format)

    # Enable autofilter for the entire data range.
    worksheet.autofilter(0, 0, df.shape[0], df.shape[1] - 1)

    # Finalize and save the workbook.
    writer.close()
    print(f"Formatted Excel file saved as: {output_file}")

if __name__ == '__main__':
    # Clear the screen first
    clear_screen()
    
    # Display a welcome banner
    print("┏━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┓")
    print("┃          ⚙️  CSV TO XLSX CONVERTER  ⚙️           ┃")
    print("┗━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┛")
    print()

    print("This program converts your .csv file to .xlsx")
    print("")
    print("The time required can vary based on file size..")    
    print("")
    
    # Start the timer
    start_time = time.time()

    # Here we run the conversion
    input_file = './input/input.csv'
    csv_to_xlsx_dynamic(input_file)

    # Stop the timer
    end_time = time.time()

    print()
    print("✅ Conversion complete! Thank you for using our tool.")
    print("")
    print(f"⏱️  Time elapsed: {end_time - start_time:.2f} seconds")
    print("")

