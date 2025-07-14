#Rows - Reg No, Name, Aec, SL, MDC, Major, Minor 1, Minor 2,SGPA


import pandas as pd
import pdfplumber as pp
import os


folder_path = r'D:\GIt\result_analyzer\results'


names=[]
roll_no=[]
Sgpa=[]
aec=[]
major=[]
SL=[]
mdc=[]
minor1=[]
minor2=[]


for filename in os.listdir(folder_path):
    if filename.endswith(".pdf"):
        file_path=os.path.join(folder_path,filename)
        #print(f"\n Reading file: {filename}")

        with pp.open(file_path) as pdf:
            page=pdf.pages[0]

            page_text=page.extract_text()
            if page_text:
                lines=page_text.splitlines()
                for line in lines:
                    if "Reg No" in line:
                        reg_no = line.split("Reg No")[-1].replace(":", "").strip(".")
                        roll_no.append(reg_no)

                    elif line.startswith("Name:"):
                        name=line.split("Name:")[1].strip()
                        names.append(name)
                        
                    elif line.startswith("SGPA:"):
                        sgpa = line.split("SGPA:")[-1].strip().split()[0]
                        try:
                            Sgpa.append(float(sgpa))
                        except ValueError:
                            Sgpa.append(0.0)
                    elif line.startswith("KU2AECENG105"):
                        parts=line.split()
                        try:
                            grade = parts[-3]  
                            aec.append(grade)
                        except IndexError:
                            aec.append("-")
                    elif line.startswith("KU2DSCCAP106"):
                        parts=line.split()
                        try:
                            grade = parts[-3]  
                            major.append(grade)
                        except IndexError:
                            major.append("-")
                    elif line.startswith(("KU2AECARB106","KU2AECHIN104","KU2AECMAL104")):
                        parts=line.split()
                        try:
                            grade = parts[-3]  
                            SL.append(grade)
                        except IndexError:
                            SL.append("-")
                    elif line.startswith(("KU2MDCENG105","KU2MDCMAT101","KU2MDCARB104","KU2MDCMAL102","KU2MDCCOM102","KU2MDCHIN102")):
                        parts=line.split()
                        try:
                            grade = parts[-3]  
                            mdc.append(grade)
                        except IndexError:
                            mdc.append("-")     
                    elif line.startswith(("KU2DSCMAT111","KU2DSCPHL104")):
                        parts=line.split()
                        try:
                            grade = parts[-3]  
                            minor1.append(grade)
                        except IndexError:
                            minor1.append("-") 
                    elif line.startswith(("KU2DSCCOM109")):
                        parts=line.split()
                        try:
                            grade = parts[-3]  
                            minor2.append(grade)
                        except IndexError:
                            minor2.append("-") 



target_len = len(names)  # master length

def pad(lst):
    while len(lst) < target_len:
        lst.append("-")
    return lst

roll_no = pad(roll_no)
aec = pad(aec)
major = pad(major)
SL = pad(SL)
mdc = pad(mdc)
minor1 = pad(minor1)
minor2 = pad(minor2)
Sgpa = pad(Sgpa)


df= pd.DataFrame({
    'Name':names,
    'Reg_no':roll_no,
    'AEC':aec,
    'Major':major,
    'SL': SL,
    'MDC':mdc,
    'MINOR1':minor1,
    'MINOR2':minor2,
    'SGPA':Sgpa
})


df.to_excel('semester2.xlsx',index=False)


import openpyxl
from openpyxl.styles import Border,Side

# Load the workbook and worksheet
wb = openpyxl.load_workbook('semester2.xlsx')
ws = wb.active

thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)


# Auto-adjust column widths
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter  # Get the column letter (e.g., A, B, C)
    for cell in col:
        cell.border = thin_border
        try:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass
    adjusted_width = max_length + 2  # Add a little extra padding
    ws.column_dimensions[column].width = adjusted_width

# Save the updated workbook
wb.save('semester2.xlsx')





