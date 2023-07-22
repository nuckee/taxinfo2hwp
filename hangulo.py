import sys
import csv
import re
import zipfile
import os
import shutil
import time
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QFileDialog, QProgressDialog, QMessageBox
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject

def get_config_value(key):
    """
    주어진 key에 해당하는 값을 config 파일에서 가져오는 함수입니다.

    Parameters:
        key (str): 찾을 값의 key

    Returns:
        str: key에 해당하는 값 (찾을 수 없으면 None 반환)
    """
    config = configparser.ConfigParser()
    config.read('config.ini')

    value = None
    if 'Section1' in config and key in config['Section1']:
        value = config['Section1'][key]

    return value

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

def number_to_korean_amount(number):
    """
    주어진 숫자를 한글로 변환하여 금액 표현과 함께 리턴하는 함수입니다.

    Parameters:
        number (int): 변환할 숫자

    Returns:
        str: 한글로 변환된 금액 표현 (예: "128,600원(금일십이만팔천육백원정)")
    """
    korean_numbers = ["일", "이", "삼", "사", "오", "육", "칠", "팔", "구"]
    korean_units = ["", "십", "백", "천"]
    korean_big_units = ["", "만", "억", "조", "경", "해", "경", "조", "억", "만"]

    number_str = str(number)
    result = []
    length = len(number_str)

    for i, digit in enumerate(number_str):
        num = int(digit)
        unit_index = (length - 1 - i) % 4
        if num != 0:
            if i > 0 and unit_index == 0 and number_str[i - 1] == '1':
                result.append(korean_numbers[0])
            else:
                result.append(korean_numbers[num - 1])
            result.append(korean_units[unit_index])
        if unit_index == 0 and i < length - 1:
            result.append(korean_big_units[(length - 1 - i) // 4])

    return "금" + "".join(result) + "원정)"

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

        # Need to make configuration file below later because it has a problem to have text overlap)
        # config_value = get_config_value(key_input)

        # TODO

        pattern = r'%I\d+%'
        matches = re.finditer(pattern, xml_content)
        for match in matches:
            i_start, i_end = match.span()
            linesegarray_end = xml_content.find("</hp:linesegarray>", i_end)
            if linesegarray_end != -1:
                insertion_text = '<hp:lineseg textpos="4" vertpos="1600" vertsize="1000" textheight="1000" baseline="850" spacing="600" horzpos="0" horzsize="4348" flags="393216"/>'
                xml_content = xml_content[:linesegarray_end] + insertion_text + xml_content[linesegarray_end:]

    # Perform value substitution in XML for F2 to L2 up to F(num_tax_numbers) to L(num_tax_numbers)
    for i, value in enumerate(values, start=1):
        placeholder = f'%{chr(64 + i)}2%'  # %A2%, %B2%, ...
        xml_content = re.sub(re.escape(placeholder), value, xml_content)

    column_to_check = 6
    num_tax_numbers = find_last_nonempty_row(csv_file, column_to_check)
 

    total_sum = 0

    # Replace F2 to L2 up to F(num_tax_numbers + 1) to L(num_tax_numbers + 1)
    for row_num in range(1, num_tax_numbers + 2):
        if row_num < len(data):
            for i in range(6, 13):
                if i - 1 < len(data[row_num]):
                    placeholder = f'%{chr(64 + i)}{row_num + 1}%'
                    value = data[row_num][i - 1]
                    if row_num > 0 and i == 12:
                        amount = int(value)
                        total_sum += amount
                    
                    # Need to make configuration file below later

                    # config_value = get_config_value(key_input)

                    # TODO
                    if (row_num > 0 and i > 9):
                        value = "{:,.0f}".format(int(value))
                    print(f'{placeholder}, row_num : {row_num}, Value : {value}')
                    xml_content = re.sub(re.escape(placeholder), value, xml_content)

    # Calculate the sum of amounts from L2 to L(num_tax_numbers + 1)
    korean_amount_num = "{:,.0f}원".format(total_sum)
    korean_amount_str = number_to_korean_amount(total_sum)
    tax_total_amount_str = korean_amount_num + '(' + korean_amount_str  + ')'
    xml_content = re.sub(re.escape("%TAX_TOTAL_AMOUNT_STR%"), tax_total_amount_str, xml_content)

    xml_content = re.sub(re.escape("%TAX_TOTAL_AMOUNT%"), korean_amount_num, xml_content)

    # Save the updated XML content to the xml_output_result variable
    xml_output_result = xml_content

    # Print the number of tax numbers and the filename of the output XML file
    print(f'The number of tax numbers in {csv_file}: {num_tax_numbers}')

    return xml_output_result

class MessageSignal(QObject):
    message_signal = pyqtSignal(str, str, str)

class ConverterThread(QThread):
    progress_signal = pyqtSignal(int)

    def __init__(self, directory):
        super().__init__()
        self.directory = directory
        self.message_signal = MessageSignal()

    def run(self):
        csv_files = [filename for filename in os.listdir(self.directory) if filename.endswith('.csv')]
        if not csv_files:
            self.message_signal.message_signal.emit('No CSV Files', 'There are no .csv files in the selected directory.', 'warning')
            return

        total_files = len(csv_files)
        for i, filename in enumerate(csv_files):
            file_path = os.path.join(self.directory, filename)
            num_tax_numbers = find_last_nonempty_row(file_path, 6)
            hwpx_file = f'template-tax-{num_tax_numbers}.hwpx'
            hwpx_file_name = os.path.splitext(hwpx_file)[0]
            zip_file = f'{hwpx_file_name}.zip'
            shutil.copy(hwpx_file, zip_file)

            with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                extract_to = '.hangulo'
                if not os.path.exists(extract_to):
                    os.makedirs(extract_to)
                hwpx_dir = os.path.join(extract_to, hwpx_file_name)

                if not os.path.exists(hwpx_dir):
                    zip_ref.extractall(hwpx_dir)

            contents_dir = os.path.join(hwpx_dir, 'Contents')
            xml_file = 'section0.xml'

            xml_file_path = os.path.join(contents_dir , xml_file)
            xml_output_result = replace_values_in_xml(file_path, xml_file_path)

            # '%영문자숫자%' 패턴의 문자열을 제거합니다.
            # xml_output_result = re.sub(r'%[a-zA-Z0-9]+%', '', xml_output_result)
           
            # 디렉토리를 복사합니다.
            gen_hwpx_file_name = os.path.splitext(filename)[0]
            gen_hwpx_dir = os.path.join(self.directory, gen_hwpx_file_name)

            # 디렉토리가 존재하는지 확인합니다.
            if os.path.exists(gen_hwpx_dir):
                # 디렉토리를 삭제합니다.
                try:
                    shutil.rmtree(gen_hwpx_dir)
                except OSError as e:
                    self.message_signal.message_signal.emit('변환 실패', f"기존에 생성된 디렉토리 '{gen_hwpx_dir}'를 삭제할 수 없습니다. '{gen_hwpx_dir}' 를 삭제한 뒤 다시 시도해 주세요. 오류: {e}", 'warning')
                    sys.exit(1)
            shutil.copytree(hwpx_dir, gen_hwpx_dir)
            gen_contents_dir = os.path.join(gen_hwpx_dir, 'Contents')
            gen_xml_file_path = os.path.join(gen_contents_dir , xml_file)

            with open(gen_xml_file_path, 'wt', encoding='UTF-8') as output_file:
                output_file.write(xml_output_result)
    
            # 디렉토리를 zip 파일로 압축합니다.
            gen_zip_file_name_path = os.path.join(self.directory, gen_hwpx_file_name)

            gen_zip_file_path = f'{gen_zip_file_name_path}.zip'

            gen_hwpx_file = f'{gen_hwpx_file_name}.hwpx'
            gen_hwpx_file_path = os.path.join(self.directory, gen_hwpx_file)

            # zip 파일이 존재하는지 확인합니다.
            if os.path.exists(gen_zip_file_path):
                # 파일을 삭제합니다.
                os.remove(gen_zip_file_path)

            # Hwpx 파일이 존재하는지 확인합니다.
            if os.path.exists(gen_hwpx_file_path):
                # 파일 이름을 삭제합니다.
                try:
                    os.remove(gen_hwpx_file_path)
                except OSError as e:
                    self.message_signal.message_signal.emit('변환 실패', f"기존에 생성된 Hwp 파일 '{gen_hwpx_file_path}'를 삭제할 수 없습니다. 열린 '{gen_hwpx_file_path}' 를 닫은 뒤 다시 시도해 주세요. 오류: {e}", 'warning')
                    sys.exit(1)
                

            shutil.make_archive(gen_zip_file_name_path, 'zip', gen_hwpx_dir)

            # 디렉토리가 존재하는지 확인합니다.
            if os.path.exists(gen_hwpx_dir):
                # 디렉토리를 삭제합니다.
                try:
                    shutil.rmtree(gen_hwpx_dir)
                except OSError as e:
                    self.message_signal.message_signal.emit('변환 실패', f"기존에 생성된 디렉토리 '{gen_hwpx_dir}'를 삭제할 수 없습니다. '{gen_hwpx_dir}' 를 삭제한 뒤 다시 시도해 주세요. 오류: {e}", 'warning')
                    sys.exit(1)

            shutil.copy(gen_zip_file_path, gen_hwpx_file_path)

            # zip 파일이 존재하는지 확인합니다.
            if os.path.exists(gen_zip_file_path):
                # 파일을 삭제합니다.
                os.remove(gen_zip_file_path)

            time.sleep(1)  # 컨버팅 시뮬레이션을 위한 딜레이
            progress_percent = int((i + 1) / total_files * 100)
            self.progress_signal.emit(progress_percent)

        self.message_signal.message_signal.emit('Conversion completed', 'the conversion is completed.', 'info')

class ConverterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.converter_thread = None

        self.message_signal = MessageSignal()
        self.message_signal.message_signal.connect(self.show_message)

    def show_message(self, title, message, message_type):
        if message_type == 'warning':
            QMessageBox.warning(self, title, message)
        else:
            QMessageBox.information(self, title, message)

    def cancel_conversion(self):
        if self.converter_thread:
            self.converter_thread.requestInterruption()


    def select_directory(self):
        options = QFileDialog.Options()
        directory = QFileDialog.getExistingDirectory(self, 'Select Directory')
        if not directory:
             QMessageBox.information(None, 'Application Closed', 'You have canceled the selection. The application will now exit.')
             sys.exit(0)
        if directory:
            print(f'Selected directory: {directory}')
            self.progress_dialog = QProgressDialog('Converting...', 'Cancel', 0, 100, self)
            self.progress_dialog.setWindowModality(Qt.WindowModal)
            self.progress_dialog.canceled.connect(self.cancel_conversion)
            self.progress_dialog.setValue(0)
            self.progress_dialog.show()
            self.converter_thread = ConverterThread(directory)
            self.converter_thread.progress_signal.connect(self.update_progress)
            self.converter_thread.message_signal.message_signal.connect(self.show_message)
            self.converter_thread.finished.connect(self.conversion_completed)

            self.converter_thread.start()

    def update_progress(self, value):
        self.progress_dialog.setValue(value)

    def conversion_completed(self):
        self.progress_dialog.hide()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    converter_app = ConverterApp()
    converter_app.select_directory()  # Moved select_directory() out of __init__()
    sys.exit(app.exec_())
