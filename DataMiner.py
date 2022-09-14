from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import string
import time
import os


class DataMiner:
    def __init__(self, backup_path, input_string):
        self.timer1 = 0
        self.timer2 = 0
        self.timer3 = 0
        self.backup_path = backup_path
        self.input_string = input_string

    translation_table = dict.fromkeys(map(ord, 'ABCJMXYZ'), None)
    titles = {'bases': ["base_data", "base_name", "X", "Y", "Z", "A", "B", "C"],
              'tools': ["tool_data", "tool_name", "X", "Y", "Z", "A", "B", "C"],
              'loads': ["load_data", "load_name", "M", "X", "Y", "Z", "A", "B", "C", "Ix", "Iy", "Iz"],
              }

    async def create_file(self):
        self.timer1 = time.time()

        data = {'bases':
                    {'names': [],
                     'data': []},

                'tools':
                    {'names': [],
                     'data': []},

                'loads':
                    {'names': [],
                     'data': []},

                }
        config_path = os.path.join(self.input_string, self.backup_path) + "\\KRC\\R1\\System\\$config.dat"

        if os.path.exists(config_path):
            with open(config_path, "r") as f:
                for line in f.readlines():

                    if line.startswith("BASE_DATA"):
                        data['bases']['data'].append(line[line.find("{"):].replace("{", "").replace("}", "")
                                                     .translate(self.translation_table).split(","))
                    elif line.startswith("TOOL_DATA"):
                        data['tools']['data'].append(line[line.find("{"):].replace("{", "").replace("}", "")
                                                     .translate(self.translation_table).split(","))
                    elif line.startswith("LOAD_DATA"):
                        data['loads']['data'].append(line[line.find("{"):].replace("{", "").replace("}", "")
                                                     .translate(self.translation_table).split(","))
                    elif line.startswith("BASE_NAME"):
                        data['bases']['names'].append(line[line.find('"') + 1:-2])
                    elif line.startswith("TOOL_NAME"):
                        data['tools']['names'].append(line[line.find('"') + 1:-2])
                    elif line.startswith("LOAD_NAME"):
                        data['loads']['names'].append(line[line.find('"') + 1:-2])

            self.timer3 = time.time()
            wb = Workbook()
            wb.create_sheet("tools")
            wb.create_sheet("loads")
            sheet = wb.active
            data_key_list = list(data.keys())
            sheet_num = 0

            for name in data_key_list:

                # Here is the excel file configured
                for num1 in range(0, len(self.titles[f'{name}'])):
                    sheet.title = f"{name}"
                    ct = sheet[f"{string.ascii_uppercase[num1]}{1}"]
                    ct.value = self.titles[f'{name}'][num1]
                    if num1 == 0:
                        sheet.column_dimensions[get_column_letter(num1 + 1)].width = 10
                    else:
                        sheet.column_dimensions[get_column_letter(num1 + 1)].width = 20

                for idx in range(0, 128):
                    for num in range(0, len(data[f'{name}']['data'][idx])):
                        if num == 0:
                            cd = sheet[f"{string.ascii_uppercase[num]}{idx + 2}"]
                            cd.value = idx + 1

                            cd = sheet[f"{string.ascii_uppercase[num + 2]}{idx + 2}"]
                            cd.value = float(data[f'{name}']['data'][idx][num])
                        elif num == 1:
                            cd = sheet[f"{string.ascii_uppercase[num]}{idx + 2}"]
                            cd.value = data[f'{name}']['names'][idx]
                            cd = sheet[f"{string.ascii_uppercase[num + 2]}{idx + 2}"]
                            cd.value = float(data[f'{name}']['data'][idx][num])
                        else:
                            cd = sheet[f"{string.ascii_uppercase[num + 2]}{idx + 2}"]
                            cd.value = float(data[f'{name}']['data'][idx][num])
                if len(data.keys()) - 1 == sheet_num:
                    continue
                else:
                    sheet = wb[f"{data_key_list[sheet_num + 1]}"]
                    sheet_num += 1

            wb.save(filename=f"{self.input_string}\\config_data_{self.backup_path}.xlsx")
            self.timer2 = time.time()
        else:
            self.timer1 = 0
            # print(config_path + 'this isn\'t a config!')

    def absolute_time(self):
        if self.timer1 != 0 or self.timer2 != 0:
            return self.timer2 - self.timer1

    def excel_time(self):
        return self.timer2 - self.timer3
