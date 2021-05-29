import sys
import os
import re


class QRCTest:
    def __init__(self, input_dir, output_dir):
        self.input_dir = os.path.abspath(input_dir)
        self.output_dir = os.path.abspath(output_dir)
        self.add_path()
        self.clean_output()
        self.walk_on_input()

    def add_path(self):
        sys.path.insert(0, os.path.abspath(
            os.path.join(os.path.dirname(__file__), '..')))

    def clean_output(self):
        for (dirpath, dirnames, filenames) in os.walk(self.output_dir):
            for filename in filenames:
                file = os.path.join(dirpath, filename)
                os.remove(file)

    def walk_on_input(self):
        for (dirpath, dirnames, filenames) in os.walk(self.input_dir):
            for filename in filenames:
                input_file = os.path.join(dirpath, filename)
                name = os.path.splitext(filename)[0]
                output_file = os.path.join(self.output_dir, name) + '.docx'
                print('Converting:\t' + input_file)
                self.convert(input_file, output_file)

    def convert(self, input_file, output_file):
        from mdlqrc.qrc import QRC
        name = os.path.basename(input_file).lower()
        if not name.endswith('.xlsx'):
            return
        rtl = 'rtl' in name
        reg = re.search(r'cols?(\d+)', name)
        cols = int(reg.group(1)) if reg else 2
        QRC(xlsx=input_file, docx=output_file,
            rtl=rtl, open_after=False, cols=cols)
