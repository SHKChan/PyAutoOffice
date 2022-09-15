#!/usr/bin/env python
# -*- coding:utf-8 -*-
# author:SHK C.

from faulthandler import disable
import os
from tkinter import Button

import cv2
import numpy as np
import PySimpleGUI as sg
from Pdf2Xl import Pdf2Xl
from PdfViewer import PdfViewer

LWIDTH = 32
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

    pdfvierwer = PdfViewer()
    cur_page = 0
    renew_page = True

    # Create an empty image in bytes
    img_mat = np.ones((792, 512, 3), np.uint8)*255
    img_bytes = mat2bytes(img_mat)

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

        if event == '-PDF_LIST-':
            selected_pdf = os.path.join(folder, values['-PDF_LIST-'][0])
            if os.path.isfile(selected_pdf):
                pdfvierwer.open(selected_pdf)
                cur_page = 0
                renew_page = True
                _, img_bytes = pdfvierwer.get_page(cur_page)
                window['-PREV-'].update(disabled=False)
                window['-NEXT-'].update(disabled=False)

        if event in ('-PREV-', '-NEXT-'):
            renew_page = True
            cur_page += -1 if event == '-PREV-' else 1
            cur_page, img_bytes = pdfvierwer.get_page(cur_page)

        if renew_page == True:
            window['-CUR-'].update(cur_page+1)
            window['-PAGE-'].update(data=img_bytes)
            renew_page = False

    window.close()


def win_main():
    sg.theme('BlueMono')

    menu_def = [['&Help', ['&About']]]

    convertor = [
        [sg.Button(
            image_filename='.//icon//pdf.png',
            image_subsample=INCON_SIZE,
            key='-PDF-',
            tooltip='Select PDF files folder to convert data'),
         sg.Input(disabled=True,
                  size=(LWIDTH, 2),
                  enable_events=True,
                  key='-INPUT_PDF-'),
         sg.FolderBrowse(visible=False, key='-FOLDER_PDF-')],
        [sg.Listbox(values=[],
                    enable_events=True,
                    size=(1, 20),
                    expand_x=True,
                    expand_y=True,
                    select_mode='extended',
                    key='-PDF_LIST-')],
        [sg.HSeparator()],
        [sg.Button(
            image_filename='.//icon//sheets.png',                        image_subsample=INCON_SIZE,
            key='-EXCEL-',
            tooltip='Select Excel file to save data'),
         sg.Input(disabled=True,
                  size=(1, 2),
                  expand_x=True,
                  key='-INPUT_EXCEL-'),
         sg.FileBrowse(
            file_types=(('Excel Files *.xlsx', '*.xlsx'),
                        ('Excel Files *.xls', '*.xls')),
            visible=False,
            key='-FILE_EXCEL-')],
        [sg.HSeparator()],
        [sg.ProgressBar(max_value=100,
                        orientation='h',
                        size=(LWIDTH, 30),
                        expand_x=True,
                        border_width=3,
                        key='-PROGRESS_BAR-')],
        [sg.Button(image_filename='.//icon//convert.png',
                   image_subsample=INCON_SIZE,
                   key='-CONVERT-',
                   tooltip='Save and exit the selected Excel file before converting')]
    ]

    pageViewer = [
        [sg.Button('Prev', disabled=True, key='-PREV-'),
         sg.Text('Page: '),
         sg.Input('', disabled=True, size=(5, 1), key='-CUR-'),
         sg.Button('Next', disabled=True, key='-NEXT-')],
        [sg.Image(filename='', key='-PAGE-')]
    ]

    layout = [
        [sg.Frame('Convertor', convertor),
         sg.Frame('Page Viewer', pageViewer)]
    ]

    return sg.Window(f'PyAutoOffice {VERSION[0]}', layout, element_justification='center')


def mat2bytes(img_mat: np.ndarray) -> None:
    png = cv2.imencode('.png', img_mat)
    bytes = png[1].tobytes()
    return bytes


def bytes2mat(img_bytes: bytes) -> None:
    ndarray = np.frombuffer(img_bytes, np.int8)
    mat = cv2.imdecode(ndarray, cv2.IMREAD_ANYCOLOR)
    return mat


if __name__ == '__main__':
    main()
