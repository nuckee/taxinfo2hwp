import csv
import re

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

    # Replace F2 to L2 up to F(num_tax_numbers + 1) to L(num_tax_numbers + 1)
    for row_num in range(2, num_tax_numbers + 2):
        if row_num < len(data):
            for i in range(6, 13):
                if i - 1 < len(data[row_num]):
                    placeholder = f'%{chr(64 + i)}{row_num}%'
                    value = data[row_num][i - 1]
                    xml_content = re.sub(re.escape(placeholder), value, xml_content)

    # Save the updated XML content to the xml_output_result variable
    xml_output_result = xml_content

    # Print the number of tax numbers and the filename of the output XML file
    print(f'The number of tax numbers in {csv_file}: {num_tax_numbers}')

    return xml_output_result

# Specify the CSV and XML file paths
csv_file = 'input.csv'
column_to_check = 6  # Column G
num_tax_numbers = find_last_nonempty_row(csv_file, column_to_check)


additional_tax_numbers = num_tax_numbers - 3
if additional_tax_numbers <= 0:
    xml_file = 'sample.xml'
else:
    xml_file = f'sample-{additional_tax_numbers}.xml'

# Call the function to perform value substitution and get the updated XML content
xml_output_result = replace_values_in_xml(csv_file, xml_file)

# Save the result to the "sample_output.xml" file
with open('sample_output.xml', 'wt', encoding='UTF-8') as file:
    file.write(xml_output_result)

# Call the function to perform value substitution and generate the output XML file
replace_values_in_xml(csv_file, xml_file)
