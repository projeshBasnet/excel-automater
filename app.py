from flask import Flask, flash, render_template, request, redirect
import os
from read_files import WriteToExcel, ReadFromExcel
import json
from time import time
from insert_marking_template import iterate_folder
from insert_student_marks import handle_insert_student_marks


app = Flask(__name__)

app.secret_key = "a secret key"

@app.route("/", methods=['GET', 'POST'])
def home():
    error = None
    if request.method=="POST":
        excel_file_name = request.form.get("filename")
        final_sheet_name = request.form.get("final_sheet_name")
        start_depth = request.form.get("start_depth")
        end_depth = request.form.get("end_depth")
        col_start_range = request.form.get("col_start_range").upper()
        col_end_range = request.form.get("col_end_range").upper()
        student_col = request.form.get("student_col").upper()
        student_folder = request.form.get("student_folder")
        student_marks_dict = request.form.get("student_marks_dict")
        if request.form.get("merge_pdf"):
            merge_excel_sheet = True
        else:
            merge_excel_sheet = False
        if len(student_marks_dict)>0:
            student_marks_dict = json.loads(student_marks_dict)

        print(f"after converring student_marks_dict: {student_marks_dict}")
        
        print(f"student_marks_dict type: {type(student_marks_dict)}")
        student_sheet_name = request.form.get("student_sheet_name")
        is_file = os.path.isfile(excel_file_name)
        is_dir = os.path.isdir(student_folder)

        if is_file and is_dir:
            try:
                print("inside try block")
                start_depth = int(start_depth)
                end_depth = int(end_depth)


                print(f"merge_excel_sheet, {merge_excel_sheet}")
               
                # inatilzing a class for a write to excel file
                print("initilizing class now")
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
                    excel_object=write_excel,
                    merge_excel = merge_excel_sheet
                )

                print("Initilized the two classes")
                init = time()
                read_student_directory.read_folder()

                print(f"Time taken to run write to excel is {time()-init}")

                flash(f"Marks has been sucessfully written in the excel file")
                return redirect("/")

            except Exception as e:
                error = e
                print(f"error is  {e}")
                return f"Start depth and end depth value must be a integer"
            
        else:
            error =  f"The provided file or folder path  could not be found in your machine"

    return render_template("index.html", error=error)
       


# @app.route("/", methods=['GET', 'POST'])

@app.route("/add_excel_sheet",methods = ["GET","POST"])
def add_excel_sheet():
    if request.method == "POST":
        folder_name = request.form.get("student_folder")
        excel_sheet = request.form.get("excel_file")
        sheet_name = request.form.get("sheet_name")
        student_name_cell = request.form.get("name_cell_value")
        student_id_cell = request.form.get("id_cell_value")

        cell_infos = {"name":student_name_cell,"id":student_id_cell}

        print(f"cell_infos",cell_infos)
        print(f"sheet_name",sheet_name)


        iterate_folder(folder_name,excel_sheet,cell_infos,sheet_name)
    return render_template("add_excel_sheet.html")



@app.route("/insert_student_marks",methods = ["GET","POST"])
def insert_student_marks():
    if request.method=="POST":
        folder_name = request.form.get("foldername")
        marks_infos = json.loads(request.form.get("marks_infos"))

        print(f"folder_name: {folder_name}")
        print(f"marks_infos: {marks_infos}")
        handle_insert_student_marks(folder_name,marks_infos)

        return redirect("/insert_student_marks")
        
    return render_template("insert_student_marks.html")


if __name__ == "__main__":
    app.run(debug=True)