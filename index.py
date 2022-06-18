from flask import Flask, render_template, request
import os
from read_files import WriteToExcel, ReadFromExcel
import json
from time import time

app = Flask(__name__)

@app.route("/", methods=['GET', 'POST'])
def home():
    if request.method=="GET":
        return render_template("index.html")
    else:
        excel_file_name = request.form.get("filename")
        final_sheet_name = request.form.get("final_sheet_name")
        start_depth = request.form.get("start_depth")
        end_depth = request.form.get("end_depth")
        col_start_range = request.form.get("col_start_range").upper()
        col_end_range = request.form.get("col_end_range").upper()
        print(f"col_end_range: {col_end_range}")
        student_col = request.form.get("student_col").upper()
        student_folder = request.form.get("student_folder")
        student_marks_dict = request.form.get("student_marks_dict")

        print(f"before converting into json: {student_marks_dict}")
        student_marks_dict = json.loads(student_marks_dict)

        print(f"after converring student_marks_dict: {student_marks_dict}")
        
        print(f"student_marks_dict type: {type(student_marks_dict)}")
        student_sheet_name = request.form.get("student_sheet_name")
        is_file = os.path.isfile(excel_file_name)
        is_dir = os.path.isdir(student_folder)

        if is_file and is_dir:
            try:
                start_depth = int(start_depth)
                end_depth = int(end_depth)
                # inatilzing a class for a write to excel file
                write_excel = WriteToExcel(
                    file_name=excel_file_name,
                    sheet_name=final_sheet_name,
                    start_depth=start_depth,
                    end_depth=end_depth,
                    column_start_range= col_start_range,
                    column_end_range=col_end_range,
                    student_id_column=student_col
                    )

                # inilatizing a class for a write to excel file
                read_student_directory = ReadFromExcel(
                    folder_path=student_folder,
                    marks_dict=student_marks_dict,
                    sheet_name=student_sheet_name,
                    excel_object=write_excel
                )

                init = time()
                read_student_directory.read_folder()
                print(f"Time taken to run write to excel is {time()-init}")

                return "sucessfully completed the writing to the excel file"

            except Exception as e:
                print(f"error is  {e}")
                return f"Start depth and end depth value must be a integer"
            
        else:
            return f"The provided file or folder path is could not be found in your machine"


        




# @app.route("/", methods=["POST"])
# def mydata():
#     return "Inside post request"

if __name__ == "__main__":
    app.run(debug=True)