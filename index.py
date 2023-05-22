import csv
import aspose.words as aw
from docxtpl import DocxTemplate

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


document = DocxTemplate("modelo_certificado_dasi.docx")

with open('Presentes - Introdução à Ciência de Dados 19 05 23 - Página1.csv', mode='r', encoding="UTF-8") as file:
    # Ignoring file first line
    next(file)
    csvFile = csv.reader(file)

    for line in csvFile:
        lecture = Lecture(line[0], "2h", "19/05/2023")
        participant = Participant(line[1].title().strip(), line[2], line[3], line[4], line[5])

        context = {
            "name": participant.name,
            "registration": participant.registration,
            "lectureName": lecture.name,
            "category": participant.category,
            "hoursGranted": participant.hoursGranted
        }

        documentName = f'./certificados/certificado-de-participacao-{participant.name.replace(" ", "-")}'
        document.render(context)
        document.save(f'{documentName}.docx')

        


