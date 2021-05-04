import sys, openpyxl


class Tabel:
    def __init__(self, data):
        try:
            file = open(data[2], 'r').read()
            number_formated = file.replace('\n', ' ').replace(',', '.').split()
            file_name_save = data[4]
            if data[3].upper() != '-F':
                self.banner()
            else:
                if data[1].upper() == '-S':
                    self.Simple(number_formated, file_name_save)
                else:
                    self.banner()
        except:
            self.banner()

    def banner(self):
        break_line = "\n"*2
        print(break_line + """     ███████████                                                                                     █████              █████              ████ 
     ░░███░░░░░░█                                                                                    ░░███              ░░███              ░░███ 
      ░███   █ ░  ████████   ██████   ████████ █████ ████  ██████  ████████    ██████  █████ ████    ███████    ██████   ░███████   ██████  ░███ 
      ░███████   ░░███░░███ ███░░███ ███░░███ ░░███ ░███  ███░░███░░███░░███  ███░░███░░███ ░███    ░░░███░    ░░░░░███  ░███░░███ ███░░███ ░███ 
      ░███░░░█    ░███ ░░░ ░███████ ░███ ░███  ░███ ░███ ░███████  ░███ ░███ ░███ ░░░  ░███ ░███      ░███      ███████  ░███ ░███░███████  ░███ 
      ░███  ░     ░███     ░███░░░  ░███ ░███  ░███ ░███ ░███░░░   ░███ ░███ ░███  ███ ░███ ░███      ░███ ███ ███░░███  ░███ ░███░███░░░   ░███ 
      █████       █████    ░░██████ ░░███████  ░░████████░░██████  ████ █████░░██████  ░░███████      ░░█████ ░░████████ ████████ ░░██████  █████
     ░░░░░       ░░░░░      ░░░░░░   ░░░░░███   ░░░░░░░░  ░░░░░░  ░░░░ ░░░░░  ░░░░░░    ░░░░░███       ░░░░░   ░░░░░░░░ ░░░░░░░░   ░░░░░░  ░░░░░ 
                                         ░███                                           ███ ░███                                                 
                                         █████                                         ░░██████                                                  
                                        ░░░░░                                           ░░░░░░                                                   

Usage: table_generator.py -S [txt_file_with_numbers.txt] -F [name_to_spreadsheet]\n""")


    class Simple:
        def __init__(self, number_formated, file_name_save):
            file_name_save = file_name_save
            generic_number = 0
            rol = sorted(number_formated)
            elements = []
            absolute_frequence_simple = {}
            total_absolute_frequence_simple = 0
            relative_frequence = []
            total_relative_frequence_simple = 0
            relative_frequence_cumulate = []
            absolute_frequence_cumulate = []
            for numbers in rol:
                if numbers not in elements:
                    elements.append(numbers)
                    absolute_frequence_simple[numbers] = rol.count(numbers)
            for numbers in absolute_frequence_simple.values():
                total_absolute_frequence_simple += float(numbers)
            for numbers in absolute_frequence_simple.values():
                relative_frequence.append(numbers/total_absolute_frequence_simple)
            for numbers in relative_frequence:
                total_relative_frequence_simple += numbers
            for numbers in absolute_frequence_simple.values():
                generic_number += numbers
                absolute_frequence_cumulate.append(generic_number)
            generic_number = 0
            for numbers in relative_frequence:
                generic_number += numbers
                relative_frequence_cumulate.append(generic_number)
            absolute_frequence_simple = list(absolute_frequence_simple.values())
            absolute_frequence_simple.append(total_absolute_frequence_simple)
            relative_frequence.append(total_relative_frequence_simple)
            self.spreadsheet_creator(rol, elements, absolute_frequence_simple,
                                     relative_frequence, absolute_frequence_cumulate,
                                     relative_frequence_cumulate, file_name_save)

        def spreadsheet_creator(self, rol: list, elements: list,
                                absolute_frequence_simple: list, relative_frequence: list,
                                absolute_frequnce_cumulate: list, relative_frequence_cumulate: list,
                                file_name_save: str
                                ):
            spreedsheet = openpyxl.Workbook()
            active_spreedsheet = spreedsheet.active
            counter = 1
            active_spreedsheet["A1"] = "Rol"
            active_spreedsheet["B1"] = "Elements (K)"
            active_spreedsheet["C1"] = "FA"
            active_spreedsheet["D1"] = "FR"
            active_spreedsheet["E1"] = "FAAC"
            active_spreedsheet["F1"] = "FRAC"
            for numbers in rol:
                counter += 1
                active_spreedsheet["A"+str(counter)] = float(numbers)
            counter = 1
            for numbers in elements:
                counter += 1
                active_spreedsheet["B" + str(counter)] = float(numbers)
            counter = 1
            for numbers in absolute_frequence_simple:
                counter += 1
                active_spreedsheet["C" + str(counter)] = float(numbers)
            counter = 1
            for numbers in relative_frequence:
                counter += 1
                active_spreedsheet["D" + str(counter)] = round(numbers, 2)
            counter = 1
            for numbers in absolute_frequnce_cumulate:
                counter += 1
                active_spreedsheet["E" + str(counter)] = float(numbers)
            counter = 1
            for numbers in relative_frequence_cumulate:
                counter += 1
                active_spreedsheet["F" + str(counter)] = round(numbers, 2)
            counter = 1
            spreedsheet.save(file_name_save + ".xlsx")


if __name__ == "__main__":
    Tabel(data=sys.argv)
