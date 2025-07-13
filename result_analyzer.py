import pandas as pd #imported pandas to export the input to excell

def get_details(no_of_students): #used for loop to enter student details and store .
      for i in range(no_of_students):
            Name=input("Enter name : ")
            Class=int(input("Enter class : "))
            students.append({'Name': Name, 'Class': Class})


students=[] #created an empty to store the values
no_of_students=int(input("Enter no of students : "))
get_details(no_of_students) #function call

pd.DataFrame(students).to_excel('marks.xlsx',index=False) #used to create excell file