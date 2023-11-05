from functions import *

if __name__ == '__main__':

    print("Bienvenue aux parrainage ENSPD 20123")

    # Lire les ficheirs excels et cree les Students
    level3_list = create_students(read_excel("level3.xlsx"))
    level4_list = create_students(read_excel("level4.xlsx"))

    # # Pour tester
    # print("I have read and created students successfully!")

    # Attribuer les parains
    assignment_dict = assign_seniors(level3_list, level4_list)

    # # Pour tester
    # print("I have assigned seniors successfully!")

    # Generer des fichiers excels
    for speciality, assignment_list in assignment_dict.items():
        write_excel(speciality, assignment_list)
        create_pdf(speciality, assignment_list)

    # # Pour tester
    # print("I have written Excel files successfully!")


