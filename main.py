from docxtpl import DocxTemplate
from docx2pdf import convert
import datetime
import shutil
import csv
import os

class Lecture:
    def __init__(self, name, duration, date):
        self.name = name
        self.duration = duration
        self.date = date

    def __str__(self):
        return f'Informações da palestra {self.name}:\nNome: {self.name}\nDuração: {self.duration}\nData: {self.date}'
    
class Participant:
    def __init__(self, name, registration, email, category, hoursGranted):
        self.name = name
        self.registration = registration
        self.email = email
        self.category = category
        self.hoursGranted = hoursGranted

    def __str__(self):
        return f'Informações do participante {self.name}:\nEmail: {self.email}\nMatrícula: {self.registration}\nCategoria: {self.category}\nHoras concedidas: {self.hoursGranted}'


document = DocxTemplate("template.docx")
base_folder = "./certificados"

if "certificados" in os.listdir():
    shutil.rmtree(base_folder)
    os.mkdir(base_folder)

with open('attendance.csv', mode='r', encoding="UTF-8") as file:
    next(file)
    csvFile = csv.reader(file)

    for line in csvFile:
        lecture = Lecture(line[0], "2h", "19/05/2023")
        participant = Participant(line[1].title().strip(), line[2], line[3], line[4], line[5])
        time = datetime.datetime.now()

        context = {
            "name": participant.name,
            "registration": participant.registration,
            "lectureName": lecture.name,
            "category": participant.category,
            "hoursGranted": participant.hoursGranted,
            "now": time
        }

        documentName = f'./certificados/certificado-de-participacao_{participant.name.replace(" ", "-")}_{lecture.name}'
        document.render(context)
        document.save(f'{documentName}.docx')

        convert(f'{documentName}.docx', f'{documentName}.pdf')

                

        


