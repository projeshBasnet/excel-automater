from pathlib import Path
import shutil

def iterate_folder(folder_path, excel_sheet):
    path = Path(folder_path)

    for each in path.iterdir():
        if not each.is_file():
            copy_excel_sheet(excel_sheet,each)
        else:
            folder_name = create_folder(each)
            copy_excel_sheet(excel_sheet,folder_name)

        

def copy_excel_sheet(excel_sheet,folder):
    excel_file_name = Path(excel_sheet).name
    file_extension = Path(excel_sheet).suffix
    student_folder_info= Path(folder).name
    shutil.copy(excel_sheet,folder)
    shutil.move(f"{folder}/{excel_file_name}",f"{folder}/{student_folder_info}{file_extension}")

def create_folder(file):
    folder_location = Path(file).parent.absolute()
    student_info = Path(file).stem
    print(f"student_info: {student_info}")

    new_folder_name = Path(f"{folder_location}/{student_info}")
    new_folder_name.mkdir(parents=True, exist_ok=True)
    shutil.move(file,new_folder_name)
    return new_folder_name



    


        
       