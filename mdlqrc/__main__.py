import os
import sys
from datetime import datetime


sys.path.insert(0, os.path.abspath(
    os.path.join(os.path.dirname(__file__), '..')))


from mdlqrc.qrc import QRC
from mdlqrc.gui import GUI


def gui_callback(self):
    if not self.file_val:
        return
    try:
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        output_file = os.path.splitext(self.file_val)[0] + '.docx'
        print(current_time)
        QRC(xlsx=self.file_val,
            cols=self.cols_val, rtl=self.rtl_val, open_after=self.open_after_val)
        print('Output: ' + output_file)
        print('Operation completed successfully!')
    except Exception as e:
        print(e)
    finally:
        print(30 * '-')


if __name__ == '__main__':
    GUI(gui_callback)
