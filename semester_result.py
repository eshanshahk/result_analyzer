import pandas as pd
import pdfplumber as pp
import os

folder_path = r'D:\GIt\result_analyzer\results'

names=[]
roll_no=[]
Sgpa=[]

for filename in os.listdir(folder_path):
    if filename.endswith(".pdf"):
        file_path=os.path.join(folder_path,filename)
        print(f"\n Reading file: {filename}")

        with pp.open(file_path) as pdf:
            page=pdf.pages[0]

            page_text=page.extract_text()

            if page_text:
                lines=page_text.splitlines()
                for line in lines:
                    if line.startswith("Name:"):
                        name=line.split("Name:")[1].strip()
                        names.append(name)
                    elif "Reg No" in line:
                        reg_no = line.split("Reg No")[-1].replace(":", "").strip()
                        roll_no.append(reg_no)
                    elif line.startswith("SGPA:"):
                        sgpa = line.split("SGPA:")[-1].strip().split()[0]
                        Sgpa.append(sgpa)

df= pd.DataFrame({
    'Name':names,
    'Reg_no':roll_no,
    'SGPA':Sgpa
})

df.to_excel('semester2.xlsx',index=False)