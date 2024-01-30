import time

import flet as ft

import utils
from translator import create_translator
from utils import *

# The Excel file path need to be translated
trans_file_path = ''
trans_file_sheets = []

# Define the constants
lang_options = [
    ('zh', 'en'),
    ('en', 'zh'),
    ('ru', 'en'),
]
cell_override = 'Override'
cell_append = 'Append'
api_google = 'Google Translate'
api_baidu = 'Baidu Translate'


def lang_opt_to_label(opt):
    src, des = opt
    return f'{src.upper()} ➟ {des.upper()}'


def label_to_lang_opt(label):
    return [x.lower() for x in label.split(" ➟ ")]


trans_mem = {}


def main(page: ft.Page):
    page.window_width = 480
    page.window_height = 300
    page.title = 'Excel Translator'
    page.horizontal_alignment = ft.MainAxisAlignment.CENTER
    page.vertical_alignment = ft.CrossAxisAlignment.CENTER
    page.window_resizable = False

    def pick_files_result(e: ft.FilePickerResultEvent):
        global trans_file_path
        global trans_file_sheets
        clean_trans_result(e)
        ft_result_text.value = (', '.join(map(lambda f: f.name, e.files))
                                if e.files else "Cancelled!")
        if e.files:
            trans_file_path = e.files[0].path
            print(trans_file_path)
            trans_file_sheets = get_excel_sheets(trans_file_path)
            # Clean the sheet from dropdown box
            remove_sheets_from_dropdown()
            if trans_file_sheets:
                add_sheets_to_dropdown()
            print(trans_file_sheets)
        ft_result_text.update()

    def add_sheets_to_dropdown():
        for sheet in trans_file_sheets:
            ft_sheet_dropdown.options.append(ft.dropdown.Option(sheet))
        ft_sheet_dropdown.value = trans_file_sheets[0]
        page.update()

    def remove_sheets_from_dropdown():
        ft_sheet_dropdown.options.clear()
        page.update()

    def clean_trans_result(e):
        if ft_result_text.value:
            ft_result_text.value = ""
            ft_result_text.visible = False
            page.update()

    def translate_click(e):
        print(f'trans_file_path: {trans_file_path}')
        print(f'trans_file_sheet: {ft_sheet_dropdown.value}')
        print(f'language_option: {ft_language_dropdown.value}')
        print(f'cell_option: {ft_cell_options.value}')
        print(f'api_option: {ft_api_options.value}')
        if not trans_file_path or not file_exist(trans_file_path):
            page.dialog = err_dlg
            err_dlg.open = True
            err_dlg.title = ft.Text('Please select an Excel file')
            page.update()
            return

        total_cell = get_excel_cell_total(trans_file_path,
                                          ft_sheet_dropdown.value)
        # Copy to new file
        new_file_path = append_datetime(trans_file_path)
        copy_file(trans_file_path, new_file_path)

        # Start to translate the Excel file
        api_client = 'google' if ft_api_options.value == api_google else 'baidu'
        src_lang, dst_lang = label_to_lang_opt(ft_language_dropdown.value)
        print(src_lang, dst_lang)
        translator = create_translator(api_client, src_lang, dst_lang)

        ft_progress.value = 0
        ft_progress.visible = True
        ft_result_text.visible = False
        page.update()

        wb = openpyxl.load_workbook(new_file_path)
        active_sheet = wb[ft_sheet_dropdown.value]
        # Iterate through each cell in the active sheet
        total_in_progress = 0
        for row in active_sheet.iter_rows():
            for cell in row:
                # Skip empty cells
                if not cell.value:
                    continue

                if type(cell.value) != str:
                    continue

                if cell.value.isnumeric():
                    continue

                if cell.value in trans_mem:
                    translated_value = trans_mem[cell.value]
                else:
                    # Translate the cell value using Google Translate
                    translated_value = translator.translate(cell.value)
                    trans_mem[cell.value] = translated_value

                print(f'{cell.value} -> {translated_value}')

                # Overwrite the cell value with the translated value
                if ft_cell_options.value == cell_override:
                    cell.value = translated_value.capitalize()
                else:
                    cell.value = f'{cell.value}\n{translated_value.capitalize()}'

                total_in_progress += 1
                print(f'progress: {total_in_progress} / {total_cell}')
                ft_progress.value = total_in_progress / total_cell
                page.update()
                time.sleep(0.5)

        # Save the modified workbook
        wb.save(new_file_path)
        print('Translation completed successfully!')

        ft_progress.visible = False
        ft_result_text.value = f'Translation completed: {utils.file_name(new_file_path)}'
        ft_result_text.visible = True
        page.update()

    ft_progress = ft.ProgressBar(width=400, visible=False)
    pick_files_dialog = ft.FilePicker(on_result=pick_files_result)
    ft_result_text = ft.Text(width=400)
    ft_sheet_dropdown = ft.Dropdown(
        label='Sheet Translated',
        hint_text='Select the sheet to be translated',
        width=200,
        options=[],
        autofocus=True,
        on_change=clean_trans_result,
    )
    ft_language_dropdown = ft.Dropdown(
        label='Language Option',
        hint_text='Select the language Option',
        width=200,
        options=[
            ft.dropdown.Option(lang_opt_to_label(e)) for e in lang_options
        ],
        value=lang_opt_to_label(lang_options[0]),
        on_change=clean_trans_result,
    )
    ft_cell_options = ft.Dropdown(
        label='Cell Option',
        hint_text='Select the cell options',
        width=200,
        options=[
            ft.dropdown.Option(cell_override),
            ft.dropdown.Option(cell_append),
        ],
        value=cell_override,
        on_change=clean_trans_result,
    )
    ft_api_options = ft.Dropdown(
        label='Translate API',
        hint_text='Select API',
        width=200,
        options=[
            ft.dropdown.Option(api_google),
            ft.dropdown.Option(api_baidu),
        ],
        value=api_google,
        on_change=clean_trans_result,
    )
    err_dlg = ft.AlertDialog(
        title=ft.Text(''),
        on_dismiss=lambda e: print('Dialog dismissed!'),
    )

    page.overlay.append(pick_files_dialog)
    page.add(
        ft.Container(
            padding=20,
            content=ft.Column(controls=[
                ft.Row(
                    [
                        ft_language_dropdown,
                        ft_cell_options,
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
                ft.Row(
                    [
                        ft_sheet_dropdown,
                        ft_api_options,
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
                ft.Row(
                    [
                        ft.ElevatedButton(
                            'Select Excel File',
                            icon=ft.icons.UPLOAD_FILE,
                            width=200,
                            on_click=lambda _: pick_files_dialog.pick_files(
                                allow_multiple=False),
                        ),
                        ft.ElevatedButton(
                            'Translate',
                            icon=ft.icons.TRANSLATE,
                            width=200,
                            on_click=translate_click,
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
                ft.Row(
                    [
                        ft_progress,
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
                ft.Row(
                    [
                        ft_result_text,
                    ],
                    alignment=ft.MainAxisAlignment.CENTER,
                ),
            ], ),
        ), )


ft.app(target=main)
