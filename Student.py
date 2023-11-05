class Student:
    def __init__(self, matricule, name, speciality, level):
        self.matricule = matricule
        self.name = name
        self.speciality = speciality
        self.level = level

    def __str__(self):
        return f"{self.matricule} - {self.name} - {self.speciality} - {self.level}"