import os
from qrctest import QRCTest

if __name__ == '__main__':
    input_path = os.path.join(os.path.dirname(__file__), 'input')
    output_path = os.path.join(os.path.dirname(__file__), 'output')
    QRCTest(input_path, output_path)
