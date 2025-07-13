import pandas as pd


def get_details(no_of_students):
      for i in range(no_of_students):
            Name=input("Enter name : ")
            Class=int(input("Enter class : "))
            students.append(Name)
            students.append(Class)

students=[]
no_of_students=int(input("Enter no of students : "))
get_details(no_of_students)
print(students)