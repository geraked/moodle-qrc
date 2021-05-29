import PySimpleGUI as sg
import os


class GUI:
    def __init__(self, callback):
        self.callback = callback
        self.create_layout()
        self.create_window()

    def create_layout(self):
        self.layout = [
            [sg.Text("Choose a .xlsx File: ")],
            [sg.In(key='-FILE-'),
             sg.FileBrowse(file_types=(("Excel Files", "*.xlsx"),))],
            [sg.Text("Columns"), sg.Slider((1, 20), 2, 1,
                                           orientation="h", size=(37, 15), key="-COLS-",)],
            [sg.Text("Direction"), sg.Radio("RTL", "Radio", True,
                                            key='-RTL-'), sg.Radio("LTR", "Radio", False)],
            [sg.Checkbox('Open After', default=True, key='-OPEN-')],
            [sg.Output(size=(51, 6), key='-OUT-')],
            [sg.Button('Convert')]
        ]

    def create_window(self):
        icon = os.path.join(os.path.dirname(__file__), '../docs/favicon.ico')
        self.window = sg.Window('Moodle Quiz Response Converter',
                                self.layout, font=("Helvetica", 12), icon=icon)
        self.loop()
        self.window.close()

    def loop(self):
        while True:
            event, values = self.window.read()
            if event == sg.WINDOW_CLOSED or event == 'Quit':
                break
            self.file_val = values['-FILE-']
            self.cols_val = int(values['-COLS-'])
            self.rtl_val = values['-RTL-']
            self.open_after_val = values['-OPEN-']
            self.callback(self)
