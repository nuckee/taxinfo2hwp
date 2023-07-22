import sys
import csv
import re
import os
import time
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QProgressDialog, QMessageBox
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject

def find_last_nonempty_row(csv_file, column):
    # Find the index of the last non-empty row in the specified column of the CSV file
    with open(csv_file, 'rt', encoding='UTF-8') as file:
        reader = csv.reader(file)
        header = next(reader)  # Skip the header row
        last_nonempty_row = None

        for i, row in enumerate(reader, start=2):  # Start from the second row (index 2)
            if row[column]:
                last_nonempty_row = i

    return last_nonempty_row - 1

def replace_values_in_xml(csv_file, xml_file):
    # Read CSV file
    with open(csv_file, 'rt', encoding='UTF-8') as file:
        reader = csv.reader(file)
        data = list(reader)

    # Extract values from CSV (A2 to E2 values)
    values = data[1][0:5]

    # Read XML file
    with open(xml_file, 'rt', encoding='UTF-8') as file:
        xml_content = file.read()

    # Perform value substitution in XML for F2 to L2 up to F(num_tax_numbers) to L(num_tax_numbers)
    for i, value in enumerate(values, start=1):
        placeholder = f'%{chr(64 + i)}2%'  # %A2%, %B2%, ...
        xml_content = re.sub(re.escape(placeholder), value, xml_content)

    column_to_check = 6
    num_tax_numbers = find_last_nonempty_row(csv_file, column_to_check)
 
    # Replace F2 to L2 up to F(num_tax_numbers + 1) to L(num_tax_numbers + 1)
    for row_num in range(1, num_tax_numbers + 2):
        if row_num < len(data):
            for i in range(6, 13):
                if i - 1 < len(data[row_num]):
                    placeholder = f'%{chr(64 + i)}{row_num + 1}%'
                    value = data[row_num][i - 1]
                    print(f'{placeholder}, row_num : {row_num}, Value : {value}')
                    xml_content = re.sub(re.escape(placeholder), value, xml_content)

    # Save the updated XML content to the xml_output_result variable
    xml_output_result = xml_content

    # Print the number of tax numbers and the filename of the output XML file
    print(f'The number of tax numbers in {csv_file}: {num_tax_numbers}')

    return xml_output_result

class MessageSignal(QObject):
    message_signal = pyqtSignal(str)

class ConverterThread(QThread):
    progress_signal = pyqtSignal(int)

    def __init__(self, directory):
        super().__init__()
        self.directory = directory
        self.message_signal = MessageSignal()

    def run(self):
        csv_files = [filename for filename in os.listdir(self.directory) if filename.endswith('.csv')]
        if not csv_files:
            self.message_signal.message_signal.emit('There are no .csv files in the selected directory.')
            return

        total_files = len(csv_files)
        for i, filename in enumerate(csv_files):
            file_path = os.path.join(self.directory, filename)
            num_tax_numbers = find_last_nonempty_row(file_path, 6)
            additional_tax_numbers = num_tax_numbers - 3
            if additional_tax_numbers <= 0:
                xml_file = 'sample.xml'
            else:
                xml_file = f'sample-{additional_tax_numbers}.xml'
            xml_output_result = replace_values_in_xml(file_path, xml_file)

            # '%영문자숫자%' 패턴의 문자열을 제거합니다.
            xml_output_result = re.sub(r'%[a-zA-Z0-9]+%', '', xml_output_result)

            # Save the result to the "sample_output.xml" file
            with open(f'{filename[:-4]}_output.xml', 'wt', encoding='UTF-8') as output_file:
                output_file.write(xml_output_result)

            time.sleep(1)  # 컨버팅 시뮬레이션을 위한 딜레이
            progress_percent = int((i + 1) / total_files * 100)
            self.progress_signal.emit(progress_percent)

        self.message_signal.message_signal.emit('Conversion completed.')


class ConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Converter App')
        self.setGeometry(100, 100, 400, 150)

        self.convert_button = QPushButton('Select Directory', self)
        self.convert_button.setGeometry(150, 30, 100, 40)
        self.convert_button.clicked.connect(self.select_directory)

        self.progress_dialog = QProgressDialog('Converting...', 'Cancel', 0, 100, self)
        self.progress_dialog.setWindowModality(Qt.WindowModal)
        self.progress_dialog.canceled.connect(self.cancel_conversion)
        self.converter_thread = None

        self.message_signal = MessageSignal()
        self.message_signal.message_signal.connect(self.show_message)

    def show_message(self, message):
        QMessageBox.warning(self, 'No CSV Files', message)

    def cancel_conversion(self):
        # Todo
        pass

    def select_directory(self):
        options = QFileDialog.Options()
        directory = QFileDialog.getExistingDirectory(self, 'Select Directory', options=options)
        if directory:
            self.converter_thread = ConverterThread(directory)
            self.converter_thread.progress_signal.connect(self.update_progress)
            self.progress_dialog.setValue(0)
            self.progress_dialog.show()
            self.converter_thread.start()

    def update_progress(self, value):
        self.progress_dialog.setValue(value)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    converter_app = ConverterApp()
    converter_app.show()
    sys.exit(app.exec_())
