import PySimpleGUI as sg
from converter import convert_pdf_to_excel

def main():
    layout = [
        [sg.Text('Select PDF File'), sg.Input(), sg.FileBrowse(file_types=(("PDF Files", "*.pdf"),))],
        [sg.Button('Convert'), sg.Button('Exit')],
        [sg.Text('', size=(40, 1), key='-OUTPUT-')],
        [sg.ProgressBar(100, orientation='h', size=(20, 20), key='-PROGRESSBAR-')]
    ]

    window = sg.Window('PDF to Excel Converter', layout)

    while True:
        event, values = window.read()
        if event in (sg.WIN_CLOSED, 'Exit'):
            break
        if event == 'Convert':
            pdf_file = values[0]
            try:
                window['-OUTPUT-'].update('Converting...')
                window['-PROGRESSBAR-'].update(0)
                convert_pdf_to_excel(pdf_file, update_progress=window['-PROGRESSBAR-'].update)
                window['-OUTPUT-'].update('Conversion Complete!')
            except Exception as e:
                window['-OUTPUT-'].update(f'Error: {str(e)}')

    window.close()

if __name__ == "__main__":
    main()