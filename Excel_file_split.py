import os
import pandas as pd
import xlsxwriter
from shutil import copyfile
from openpyxl import load_workbook

#'
#MOCK_DATA.xlsx
org_file = input("Please enter the file path name:\n ")       # prompt user to enter the file name or path
org_file_extension = os.path.splitext(org_file)[1]          # split the file extension i.e txt,xlxs,csv,json.....
org_file_name = os.path.splitext(org_file)[0]               # grab the file name from the path given by user
path = os.path.dirname(org_file)                            # grab the directory name
new_file = os.path.join(f"{path},{org_file_name}_2{org_file_extension}")    # the new file name for saving the new file
df = pd.read_excel(org_file)                                   # converting the file into a pandas DataFrame
selected_column = input("Select a column please:\n ")
columns = list(set(df[selected_column].values))                 # this is the column on which we shall split the file into sheets or files based on its values / content



def send_to_files(columns):

    for column in columns:
        df[df[selected_column] == column].to_excel(f'{path}/{column}.{org_file_extension}',sheet_name = column,index = False)
    print("\nProcess Completed Successfully!!")
    print("\nThank You for using this application, Good Bye!!")
    return


def send_to_sheets(columns):
    copyfile(org_file,new_file)

    for column in columns:
        writer = pd.ExcelWriter(new_file, engine = "openpyxl")
        for myname in columns:
            my_data_frame = df.loc[df[selected_column] == myname]
            my_data_frame.to_excel(writer,sheet_name=myname,index = False)
        writer.save()
    print("\nProcess Completed Successfully!!")
    print("\n")
    return








print(f'''\nYour data will be separated based on this value {",".join(columns)},
it will create {len(columns)} file(F) or sheets(S) based on your selection.
if you wish to continue press YES and enter else pres NO to quit''')

while True:
    user_confirmation = input(" \n Ready to continue? please enter 'Yes' or 'No': ").lower()

    if user_confirmation == "yes":

        while True:

            sheet_or_file = input(f"""\nDo you wish to split the {org_file_name} into {len(columns)} files (enter--> F)
                                  or into {len(columns)} sheets (enter--> S)??  """).lower()

            if sheet_or_file == "s":
                send_to_sheets(columns)
                break

            elif sheet_or_file == "f":
                send_to_files(columns)
                break

            else: continue
        break

    elif user_confirmation == "no":
        print("\n Thank you for using this application, Good Bye!!")
        break
    else:
        print("\n Invalid Entry, Please enter 'Yes' or 'No' to Continue.")
        continue

# print(org_file_extension)
# print(org_file_name)
# print(path)
# print(new_file)
# print(columns)