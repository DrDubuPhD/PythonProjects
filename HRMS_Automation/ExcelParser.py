import re
from datetime import datetime
import openpyxl
import xlrd


# Defining a function to locate the titles in the excel file

def findstring(word, sheet):
    for i in range(20, sheet.nrows):
        row = sheet.row_values(i)
        for j in range(len(row)):
            if row[j] == word:
                return i, j
    return None


# Opening the sheet for processing

jess_loc = 'HRMS_Automation\Spreadsheets\Promoter Leave Files\Promoter - ANNUAL LEAVE Sept 2020 Jess.xlsx'
koay_loc = 'HRMS_Automation\Spreadsheets\Promoter Leave Files\Promoter - ANNUAL LEAVE Sept 2020 Ms Koay.xlsx'
template_loc = 'HRMS_Automation\Spreadsheets\Promoter Leave Files\Leave Application Data Import.xlsx'

try:
    Jess_leave_template = xlrd.open_workbook(jess_loc)
    Koay_leave_template = xlrd.open_workbook(koay_loc)
    iHRMS_leave_template = openpyxl.load_workbook(filename=template_loc)
except (xlrd.XLRDError, FileNotFoundError) as Error:
    print("\nError! Please ensure that the 3 files are:")
    print("1) Inside the Promoter Leave Files folder")
    print("2) Named correctly")
    exit(1)

jess_sel_sheets = Jess_leave_template.sheets()
koay_sel_sheets = Koay_leave_template.sheets()
sel_sheets = jess_sel_sheets + koay_sel_sheets
output_sheet = iHRMS_leave_template['Sheet1']

# Creating a list to store all 2020 leave entries (Preprocess all the variables)

leave_2020 = []

for sheets in sel_sheets:
    if '2020' in sheets.name:
        leave_2020.append(sheets)

print('\nFound', len(leave_2020), 'employee leave details to be processed')
can_proceed = str(input('\nProceed (y/n): ')).lower()

# Mapping leave names to their codes

leave_categories = {'ANNUAL LEAVE': 'ALM', 'MEDICAL / HOSPITALISATION LEAVE': 'MC', 'UNPAID LEAVE': 'UPL',
                    'ABSENT': 'ABS', 'REPLACEMENT': 'RPL'}

# Date Regex to detect interval dates

date_regex = '[0-9]{2}[-]{1}[0-9]{2}|[0-9]{1}[-]{1}[0-9]{1}'

# Main Parsing Logic Structure

if can_proceed == 'y':
    output_row_count = 2  # Start from the 2nd row
    for person in leave_2020:
        for leave_type in leave_categories:

            # Locating the leave name starting position
            start_row, start_col = findstring(leave_type, person)

            # print("\nFound String in", start_row + 1, start_col)
            print("\nProcessing Sheet: {} for {}".format(person.name, leave_type), end=None)
            isEmpty = True

            for row_pos in range(start_row + 2, person.nrows):  # Begin looping at after 2 rows from the title until eof
                try:  # Check for empty string cell value
                    notEmptyCell = person.cell_value(row_pos, start_col) != ''
                except IndexError:
                    print("Error Parsing Col:", start_col + 1, "Row: ", row_pos + 1)

                if notEmptyCell:  # Begin Parsing here
                    isEmpty = False
                    try:
                        date = person.cell_value(row_pos, start_col)  # Read the cell value (date)
                        day = float(person.cell_value(row_pos, start_col + 1))  # Read the cell value (day)
                    except ValueError:
                        print("Error Parsing Col:", start_col + 1, "Row: ", row_pos + 1)

                    # print("Cell Value (Date):", date, "Cell Value (Day):", day)

                    # Parsing Dates
                    if isinstance(date, str):
                        if re.search(date_regex, date):  # Checked Interval Special Case
                            parsed_date_interval = date.split('/')  # 3 Element List ['Start - End', 'Month', 'Year']
                            parsed_day_start, parsed_day_end = parsed_date_interval[0].split('-')
                            parsed_date = parsed_day_start + '/' + parsed_date_interval[1] + '/' + parsed_date_interval[
                                2]
                            # print(parsed_day_start, parsed_day_end)
                        else:  # Parse as usual
                            try:
                                temp = datetime.strptime(date, '%d.%m.%Y')
                                parsed_date = temp.strftime('%d/%m/%Y')
                            except ValueError:
                                print("Error in Sheet {}".format(person.name))
                    elif isinstance(date, float):  # Parse as float
                        temp = xlrd.xldate_as_datetime(date, Jess_leave_template.datemode)
                        parsed_date = temp.strftime('%d/%m/%Y')

                    if day <= 1:
                        end_date = parsed_date
                    else:
                        end_date = parsed_day_end + '/' + parsed_date_interval[1] + '/' + parsed_date_interval[2]

                    temp_row = [person.name.split("(")[0], leave_categories[leave_type], parsed_date, end_date, day,
                                'Full' if day >= 1 else 'First']

                    print('\t'.join(map(str, temp_row)))

                    temp_row_counter = 0  # List counter

                    for output_columns in range(1, 7):
                        output_sheet.cell(row=output_row_count, column=output_columns, value=temp_row[temp_row_counter])
                        temp_row_counter += 1

                    output_row_count += 1

            if isEmpty:
                print("No leave applied for this category")

elif can_proceed == 'n':
    print("\nPlease check the files and restart the program")
    exit(1)
else:
    print("\nPlease key in only y/n!")
    exit(1)

iHRMS_leave_template.save('HRMS_Automation\Spreadsheets\Output\Promoter Leave File 2020.xlsx')
print("\n\nDone! Please check the output folder for your file")
exit(0)
