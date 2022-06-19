import openpyxl
from pathlib import Path
import os


class WriteToExcel:
    def __init__(self, file_name, sheet_name, start_depth, end_depth, column_start_range, column_end_range, student_id_column = "B") -> None:
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.start_depth = start_depth
        self.end_depth = end_depth
        self.column_start_range = column_start_range
        self.column_end_range = column_end_range
        self.iterating_range = [i for i in range(self.start_depth, self.end_depth+1)]
        print(f"self.iterating_range: {self.iterating_range}")
        self.student_id_column = student_id_column
        self.column_array_range = self.create_column_array()
        self.work_book = openpyxl.load_workbook(self.file_name)
        self.sheet = self.work_book[self.sheet_name]
        self.sheet_marks_position = self.create_sheet_marks_position()

        print(f"sheet_marks_position: {self.sheet_marks_position}")

    
    def range_char(self,start, stop):
        return (chr(n) for n in range(ord(start), ord(stop) + 1))

    def create_column_array(self):
        my_array = []
        for char in self.range_char(self.column_start_range, self.column_end_range):
            my_array.append(char)
        return my_array
    
    def create_sheet_marks_position(self):
        my_dict = {}
        
        for letter in self.column_array_range:
            value = self.sheet[f"{letter}{self.start_depth-1}"].value.split("(")[0].strip()
            my_dict[value] = f"{letter}"

        return my_dict
    
    def find_student_row_number(self, student_id):
        row_no = 0
        print(f"student_id: {student_id}")
        for i in self.iterating_range:
            print("Inside for loop")
            if student_id == self.sheet[f"{self.student_id_column}{i}"].value:
                row_no = i
                self.iterating_range.remove(i)
                print(f"changed iterating range: {self.iterating_range}")

                break
        print(f"student row number about to be returned: {row_no}")
        return row_no



class ReadFromExcel:
    def __init__(self, folder_path, marks_dict, sheet_name, excel_object) -> None:
        self.folder_path = folder_path 
        self.marks_dict = marks_dict
        self.sheet_name = sheet_name
        self.excel_object = excel_object

    
    def read_folder(self):
        path = Path(self.folder_path)
        """
        Reading the each folder of student and taking the student id
        """
        for p in path.iterdir():
            student_id = int(os.path.basename(p).split()[0])
            self.read_from_excel_file(student_id, p)
            
    
    def read_from_excel_file(self, student_id,p):
        for file in p.rglob("*.xlsx"):
                print(f"file: {file}")
                wb = openpyxl.load_workbook(file, data_only=True)
                ws = wb[self.sheet_name]
                marks_value_dict = {}
                for key, value in self.marks_dict.items():
                    # has used this try block to convert 0 string into integer
                    try:
                        marks_value_dict[key] = float(ws[value].value)
                    except ValueError:
                        marks_value_dict[key] = 0

                self.write_to_main_docs(student_id, marks_value_dict)

    
    def write_to_main_docs(self,student_id, marks_value_dict):
        print(f"writing to main docs of id {student_id}")
        print(f"marks_value_dict: {marks_value_dict}")
        print(f"sheet marks position: {self.excel_object.sheet_marks_position}")
 
        if len(marks_value_dict.keys()) == len(self.excel_object.sheet_marks_position.keys()):
            student_row_num = self.excel_object.find_student_row_number(student_id)
            for key, value in marks_value_dict.items():
                # student_row_num = self.excel_object.find_student_row_number(student_id)
                print(f"student_row_num: {student_row_num}")
                if key in self.excel_object.sheet_marks_position:
                    self.excel_object.sheet[f"{self.excel_object.sheet_marks_position[key]}{student_row_num}"] = value
            
            self.excel_object.work_book.save(self.excel_object.file_name)
            print(f"sucessfully saved file")
        
        else:
            print(f"Marks dict is not so valid with")
                
                


