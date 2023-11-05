import pandas as pd
import random
import collections

from reportlab.pdfgen import canvas
import xlsxwriter

from Student import Student


def read_excel(filename: str) -> list:
    """ Cette fonction c'est pour lire les fichiers excel, elle retourne un dictionnaire contenat tous le contenu
    du fichier excel"""
    df = pd.read_excel(filename)
    excel_list = df.to_dict('records')
    return excel_list


def create_students(excel_list: list) -> list:
    """Cette fonction prend chaque objet de la liste des elements du fichier excel et instanci une classe Student
    pour chaque element correspondant. Elle retourne la liste des Students"""
    student_list = []
    for student_dict in excel_list:
        student = Student(student_dict['matricule'], student_dict['name'], student_dict['speciality'],
                          student_dict['level'])
        student_list.append(student)
    return student_list


# def assign_seniors(level3_list: list, level4_list: list) -> dict:
#     """Cette fonction attribu a chaque filleul un parrain"""
#     assignment_dict = collections.defaultdict(list)
#     for junior in level3_list:
#         speciality = junior.speciality
#         senior_list = [senior for senior in level4_list if senior.speciality == speciality]
#         senior = random.choice(senior_list)
#         assignment_dict[speciality].append((junior, senior))
#     return assignment_dict

def assign_seniors(level3_list: list, level4_list: list) -> dict:
    """Cette fonction attribu a chaque filleul un parrain"""
    assignment_dict = collections.defaultdict(list)
    senior_list = level4_list.copy()
    for junior in level3_list:
        speciality = junior.speciality
        senior_list = [senior for senior in senior_list if senior.speciality == speciality]
        if not senior_list:
            senior_list = [senior for senior in level4_list if senior.speciality == speciality]
        index = random.randrange(len(senior_list))
        senior = senior_list.pop(index)
        assignment_dict[speciality].append((junior, senior))
    return assignment_dict


def write_excel(speciality: str, assignment_list: list) -> None:
    """Cette fonction  sert a generer un fichier excel contenant les differents
    filleuls avec leurs parrains correspondant"""
    workbook = xlsxwriter.Workbook(f"{speciality}.xlsx")
    worksheet = workbook.add_worksheet()
    worksheet.write(0, 0, "level 3 matricule")
    worksheet.write(0, 1, "level 3 name")
    worksheet.write(0, 2, "corresponding level 4 matricule")
    worksheet.write(0, 3, "corresponding level 4 name")
    for i, (junior, senior) in enumerate(assignment_list):
        worksheet.write(i + 1, 0, junior.matricule)
        worksheet.write(i + 1, 1, junior.name)
        worksheet.write(i + 1, 2, senior.matricule)
        worksheet.write(i + 1, 3, senior.name)
    workbook.close()


def create_pdf(speciality: str, assignment_list: list) -> None:
    pdf = canvas.Canvas(f"{speciality}.pdf")
    pdf.setTitle(f"{speciality} parrainage")
    pdf.drawString(50, 750, "level 3 matricule")
    pdf.drawString(150, 750, "level 3 name")
    pdf.drawString(250, 750, "level 4 matricule")
    pdf.drawString(350, 750, "level 4 name")
    for i, (junior, senior) in enumerate(assignment_list):
        pdf.drawString(50, 730 - i * 20, junior.matricule)
        pdf.drawString(150, 730 - i * 20, junior.name)
        pdf.drawString(250, 730 - i * 20, senior.matricule)
        pdf.drawString(350, 730 - i * 20, senior.name)
    pdf.save()
