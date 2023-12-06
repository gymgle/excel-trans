import openpyxl
import time

from easygoogletranslate import EasyGoogleTranslate


def translate_excel_file(excel_file):
    translator = EasyGoogleTranslate(
        source_language='zh-CN',
        target_language='en',
        timeout=10
    )

    # Load the input Excel file
    wb = openpyxl.load_workbook(excel_file)
    active_sheet = wb.active

    # Iterate through each cell in the active sheet
    for row in active_sheet.iter_rows():
        for cell in row:
            # Skip empty cells
            if not cell.value:
                continue

            time.sleep(0.5)
            # Translate the cell value using Google Translate
            translated_value = translator.translate(cell.value)
            print(f'{cell.value} -> {translated_value}')

            # Overwrite the cell value with the translated value
            cell.value = translated_value.capitalize()

    # Save the modified workbook
    wb.save(excel_file)
    print("Translation completed successfully!")


# Provide the input and output file paths
input_file = 'test.xlsx'

# Call the translation function
translate_excel_file(input_file)
