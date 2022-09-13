#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author:SHK C.

import os

import PySimpleGUI as sg

from pdf2xl import Pdf2Xl

BUTTON_MAPPING = {'-PDF-': '-FOLDER_PDF-', '-EXCEL-': '-FILE_EXCEL-'}
INCON_SIZE = 7
VERSION = ['V1.0.2', 'V1.0.1', 'V.1.0.0']
UPDATE_NOTE = [
    ['Others: Data structure optimize',
     'Fixed: Raise error when excel file is occupied',
     'Added: 1. Tips for each button\n\
             2. Converting multiple PDF files at once\n\
             3. Progress bar\n\
             4. Threading exception handling\n'],
    ['Fixed: Multipage PDF files handling\n'],
    ['Others: Alpha version debut\n'],
]


def main():
    window = win_main()

    while True:
        event, values = window.read(50)
        
        if event in (None, sg.WIN_CLOSED, '-CLOSE-'):
            break
        if event == 'About':
            notes = []
            for i in range(len(VERSION)):
                notes.append(VERSION[i])
                for j in range(len(UPDATE_NOTE[i])):
                    notes.append(UPDATE_NOTE[i][j])
            notes = '\n'.join(notes)
            sg.popup('Update Notes', f'{notes}', )

        if event in ('-PDF-', '-EXCEL-'):
            window[BUTTON_MAPPING[event]].Click()

        if event == '-CONVERT-':
            if values['-PDF_LIST-'] != '' and values['-INPUT_EXCEL-'] != '':
                pdf_files = list()
                for pdf in values['-PDF_LIST-']:
                    pdf_files.append(values['-INPUT_PDF-'] + "/" + pdf)

                convertor = Pdf2Xl(pdf_files, values['-INPUT_EXCEL-'])
                convertor.start()
                while True:
                    # convertor.convert()
                    if convertor.exit_code == 0:
                        sg.popup(
                            'Info', 'Convert PDF data into Excel successfully!')
                        break
                    elif convertor.exit_code is not None:
                        sg.popup('Error', 'Can not access pdf or excel file!')
                        break
                    # 更新进度条
                    window['-PROGRESS_BAR-'].update(
                        current_count=convertor.progress)
            else:
                sg.popup('Warning', 'Please select pdf and excel files correctly!')

        if event == '-INPUT_PDF-':
            folder = values['-INPUT_PDF-']
            try:
                # Get list of files in folder
                file_list = os.listdir(folder)
            except:
                file_list = []
            fnames = []
            for f in file_list:
                if os.path.isfile(os.path.join(folder, f)) and f.lower().endswith(('.pdf')):
                    fnames.append(f)
            window['-PDF_LIST-'].update(fnames)

    window.close()


def win_main():
    sg.theme('BlueMono')

    menu_def = [['&Help', ['&About']]]

    layout = [
        [
            [sg.Menu(menu_def)],
            [sg.Button(image_filename='.//icon//pdf.png', image_subsample=INCON_SIZE, key='-PDF-',
                       tooltip='Select PDF files folder to convert data'),
             sg.Input(disabled=True, size=(80, 2),
                      enable_events=True, key='-INPUT_PDF-'),
             sg.FolderBrowse(visible=False, key='-FOLDER_PDF-')],
            [sg.Listbox(values=[], enable_events=True,
                        size=(90, 20), key='-PDF_LIST-', select_mode='extended')],
            [sg.HSeparator()],
            [sg.Button(image_filename='.//icon//sheets.png',                        image_subsample=INCON_SIZE, key='-EXCEL-',
                       tooltip='Select Excel file to save data'),
             sg.Input(key='-INPUT_EXCEL-', disabled=True, size=(80, 2)),
             sg.FileBrowse(file_types=(('Excel Files *.xlsx', '*.xlsx'), ('Excel Files *.xls', '*.xls')), visible=False, key='-FILE_EXCEL-')],
            [sg.HSeparator()],
            [sg.ProgressBar(max_value=100, orientation='h',
                            size=(40, 30), key='-PROGRESS_BAR-', expand_x=True, border_width=3)],
            [sg.Button(image_filename='.//icon//convert.png',
                       image_subsample=INCON_SIZE, key='-CONVERT-',
                       tooltip='Save and exit the selected Excel file before converting')]
        ],
        [sg.Image(filename='', key='-PAGE-')]
    ]

    return sg.Window(f'PyAutoOffice {VERSION[0]}', layout, element_justification='center')


if __name__ == '__main__':
    main()
