import base64
import cv2
import numpy as np
import pdfplumber as plumber
import PySimpleGUI as sg


def main():
    pdf = plumber.open('files/P-50463-79.pdf')

    # Get png image in bytes
    im = pdf.pages[0].to_image(resolution=150)
    im_png = im._repr_png_()

    # Create an 512*512 np array with values 255
    img_mat = np.ones((512, 512, 3), np.uint8)*255

    window = win_main()

    while True:
        event, values = window.read(50)

        if event in (None, sg.WIN_CLOSED, '-CLOSE-'):
            break

        img_bytes = mat2bytes(img_mat)
        window['-IMAGE-'].update(data=img_bytes)

    window.close()


def mat2bytes(img_mat: np.ndarray) -> None:
    png = cv2.imencode('.png', img_mat)
    bytes = png[1].tobytes()
    return bytes


def bytes2mat(img_bytes: bytes) -> None:
    ndarray = np.frombuffer(img_bytes, np.int8)
    mat = cv2.imdecode(ndarray, cv2.IMREAD_ANYCOLOR)
    return mat


def win_main():
    layout = [
        [sg.Image(filename='', k='-IMAGE-')]
    ]

    return sg.Window('To Image Test', layout, element_justification='center')


if __name__ == '__main__':
    main()
