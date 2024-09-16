from tkinter import * 
from ttkbootstrap.constants import * 
import ttkbootstrap as tb 
from datetime import date
from ttkbootstrap.dialogs import Querybox, Messagebox
import pandas as pd 
from docxtpl import DocxTemplate
import os

font=("Calibri light", 9)
font2=("Calibri light", 12)
root = tb.Window(themename = 'solar')
root.geometry('550x280')
root.resizable(False, False)
root.title("NPEC")


def close():
	quit()

def generate_temp():
	print("You will generate an excel template")
	print("You will generate a Report Template")


def load_excel_file(): 
	global file_name 
	name = sheetNameEntry.get()
	sheetNameEntry.delete(0, END)

	# Checking if file_name is not empty
	if name == "":
		print("The file name can't be empty")
	else: 
		file_name = name
		print(f"{file_name}.xlsx")
		return file_name


	# global df
	# df = pd.read_excel("grades.xlsx", sheet_name="Form2")

	

def generate_reports():

	# GETTING THE CLASS SELECTED TO BE USED TO READ A SHEET IN THE "grades.xlsx" FILE 
	Class = class_combo.get()
	class_to_read = Class.replace(" ", "")
	

	# changing Form 3 and Form 4 to 'Form3Analysis' and "Form4Analysis" as reflected in the data source excel file 

	senior_classes = {"Form3": "Form3Analysis", "Form4": "Form4Analysis"}

	if class_to_read == "SelectClass":
		Messagebox.show_error("Kindly select a class to procedee")

	elif class_to_read != "SelectClass":
		if class_to_read == "Form4":
			class_to_read = senior_classes[class_to_read]

		if class_to_read == "Form3":
			class_to_read = senior_classes[class_to_read]

		# READING THE SPREADSHEET FILE	
		df = pd.read_excel("C:/Users/Thoko/Desktop/toDeploy/Scores/grades.xlsx", sheet_name=class_to_read)

		POINTS = {'pointss':[]}

		for index, row in df.iterrows():
		    firstSixGrades = []

		    firstSixGrades.append(pd.to_numeric(row['Eng_Grade'], errors="coerce"))
		    allGradesList = [
		    	pd.to_numeric(row["Agr_Grade"], errors="coerce"),
		    	pd.to_numeric(row["Bio_Grade"], errors="coerce"),
		    	pd.to_numeric(row["Chem_Grade"], errors="coerce"),
		    	pd.to_numeric(row["Chi_Grade"], errors="coerce"),
		    	pd.to_numeric(row["Geo_Grade"], errors="coerce"),
		    	pd.to_numeric(row["His_Grade"], errors="coerce"),
		    	pd.to_numeric(row["Maths_Grade"], errors="coerce"),
		    	pd.to_numeric(row["Phy_Grade"], errors="coerce"),
		    	pd.to_numeric(row["Sod_Grade"], errors="coerce"),
		    	pd.to_numeric(row["Bk_Grade"], errors="coerce")


		    ]
		    
		    # Remove 'NaN' values from allGradesList
		    allGradesList = [grade for grade in allGradesList if pd.notna(grade)]
		    allGradesList.sort()
		    
		    firstSixGrades.extend(allGradesList[:5])
		    # print(f"Grades for {row['firstname']} {row['surname']} are {firstSixGrades}")
		    # Getting points for the first 6 best subjects including English
		    points = sum(firstSixGrades)
		    POINTS['pointss'].append(points)
		# print(POINTS)
		df['pointss'] = POINTS['pointss']
		# df['pointss'].astype(int)
		df['Rank'] = df['pointss'].rank(ascending=True)
		rank_count = df['Rank'].count()
		print(df)
		
		# print(df)
		# DIVING CLASSES INTO JUNIOR AND SENIOR SECTION 
		Senior_section = (senior_classes["Form3"], senior_classes["Form4"])
		Junior_section = ("Form1", "Form2")

		# Reading the Template Documents based on the class entered

		if class_to_read in Senior_section:
			doc = DocxTemplate("C:/Users/Thoko/Desktop/toDeploy/Templates/Senior_Report_Temp.docx")

		elif class_to_read in Junior_section:
			doc = DocxTemplate("C:/Users/Thoko/Desktop/toDeploy/Templates/Junior_Report_Temp.docx")

		# IMPUTING DATA
		def impute_data(df_data):
			if pd.isnull(df_data):
				return "-"
			else:
				return int(df_data)

		def impute_grade(data):
				if pd.isnull(data):
					return "-"

				elif data in ['A', 'B', 'C', 'D', 'F']:
					return data

				elif type(data) != int:
					return int(data)

				else:
					return data

		def subject_count(column):
			length = df[column].count()
			return length


		def finalee(data):
		    if pd.isnull(data):
		        return 0  # Return 0 for NaN values
		    else:
		        return int(data)

		# ASSIGNING THE VARIABLES IN THE WORD TEMPLATE TO THE VALUES IN THE EXCEL FILE 

		for index, row in df.iterrows():

			content = {
				"firstname":row['firstname'], "surname":row['surname'],
				"Agr":impute_data((row['Agr'])), "Agr_Grade":impute_grade(row['Agr_Grade']), "Agr_Pos":impute_data(row['Agr_Pos']),
				"Bio":impute_data(row['Bio']), "Bio_Grade":impute_grade(row['Bio_Grade']), "Bio_Pos":impute_data(row['Bio_Pos']), 
				"Chem":impute_data(row['Chem']), "Chem_Grade":impute_grade(row['Chem_Grade']), "Chem_Pos":impute_data(row['Chem_Pos']), 
				"Chi":impute_data(row['Chi']), "Chi_Grade":impute_grade(row['Chi_Grade']), "Chi_Pos":impute_data(row['Chi_Pos']), 				
				"Eng":impute_data(row['Eng']), "Eng_Grade":impute_grade(row['Eng_Grade']), "Eng_Pos":impute_data(row['Eng_Pos']), 
				"Geo":impute_data(row['Geo']), "Geo_Grade":impute_grade(row['Geo_Grade']), "Geo_Pos":impute_data(row['Geo_Pos']),
				"His":impute_data(row['His']), "His_Grade":impute_grade(row['His_Grade']), "His_Pos":impute_data(row['His_Pos']), 
				"Maths":impute_data(row['Maths']), "Maths_Grade":impute_grade(row['Maths_Grade']), "Maths_Pos":impute_data(row['Maths_Pos']), 
				"Phy":impute_data(row['Phy']), "Phy_Grade":impute_grade(row['Phy_Grade']), "Phy_Pos":impute_data(row['Phy_Pos']), 
				"Sod":impute_data(row['Sod']) , "Sod_Grade":impute_grade(row['Sod_Grade']), "Sod_Pos":impute_data(row['Sod_Pos']),
				"Bk":impute_data(row['Bk']) , "Bk_Grade":impute_grade(row['Bk_Grade']), "Bk_Pos":impute_data(row['Bk_Pos']),
				"Class":Class,
				"Points": finalee(row['pointss']) ,
				"Position": finalee(row['Rank']),
				"Positionn":row['CPos'],
				"total_count": df['Eng'].count(),
				"total_countt":df['Eng'].count(),
				"engcount":subject_count('Eng'),
				'phycount':subject_count('Phy'),
				'checount':subject_count('Chem'),
				'biocount':subject_count('Bio'),
				'mathscount':subject_count('Maths'),
				'sodcount':subject_count('Sod'),
				'hiscount':subject_count('His'),
				'geocount':subject_count('Geo'),
				'agrcount':subject_count('Agr'),
				'bkcount':subject_count('Bk'),
				'chicount':subject_count('Chi')

			}

			# COMMING UP WITH A FILE NAME FOR EACH STUDENT IN THE DATA SOURCE FILE
			doc_name = f"{content['firstname'],content['surname']}.docx"
			doc_name = doc_name.replace(',', "")
			doc_name = doc_name.replace('(', "")
			doc_name = doc_name.replace(')', "")
			doc_name = doc_name.replace("'", "")

			doc_path = os.path.join(Class, doc_name)
			doc.render(content)
			doc.save(doc_path)

		Messagebox.show_info(f"{Class} reports were generated successfully.", "NPEC")		

# The "Configure Spreadsheet Frame" which has the 'Generate Template' and 'Load File' buttons
# and File name entry field

logInFrame = tb.Labelframe(root, bootstyle="success",text="Configure Mark Book")
logInFrame.pack(padx=20, pady=10, fill=X)


frame1 = tb.Frame(logInFrame)
frame1.pack(fill=X)

temp_button = tb.Button(frame1, text="Generate Templates", bootstyle="default, outline", command=generate_temp)
temp_button.grid(row=0, column=0, padx=20, pady=10)


sheetNameLabel = tb.Label(frame1, text="Markbook name: ", font=font2, bootstyle="success")
sheetNameLabel.grid(row=1, column=0, padx=20, pady=10)

sheetNameEntry = tb.Entry(frame1, bootstyle="success", font=font)
sheetNameEntry.grid(row=1, column=1)


load_file_btn = tb.Button(frame1,bootstyle="default, outline",text="Load Markbook", command=load_excel_file)
load_file_btn.grid(row=1, column=2, padx=10)


separator1 = tb.Separator(root, bootstyle="info")
separator1.pack(padx=20, fill=X, pady=5)

# The "Generate Report Cards" Frame

register_frame = tb.Labelframe(root, bootstyle='success', text="Generate Report Forms")
register_frame.pack(padx=20, pady=10, fill=X)

classes = ("Select Class", "Form 1", "Form 2", "Form 3", "Form 4")

class_combo = tb.Combobox(register_frame, bootstyle="success", values=classes)
class_combo.grid(column=0, row=0, padx=10)
class_combo.current(0)

register_btn = tb.Button(register_frame, text="Generate Reports", width=21, bootstyle='default, outline', command=generate_reports)
register_btn.grid(column=1, row=0, padx=10)

close_button = tb.Button(register_frame, text="Close", width=21, bootstyle='danger, outline', command=close)
close_button.grid(column=2, row=0, padx=10, pady=10)


separator2 = tb.Separator(root, bootstyle="info")
separator2.pack(padx=20, fill=X, pady=5)

footer_frame = tb.Frame(root)
footer_frame.pack(fill=X)

exams_depart = tb.Label(footer_frame, text="NPEC Examination Department", bootstyle="success")
exams_depart.pack()
npec_label = tb.Label(footer_frame, text="New Palm Education Center @2024", bootstyle="success")
npec_label.pack()

root.mainloop()