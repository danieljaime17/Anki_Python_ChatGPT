import openai
from openpyxl import Workbook
from openpyxl import load_workbook
import time


# Tutorial en vídeo: https://youtu.be/1Pl1FWHKHXQ

# Genera una API Key desde https://openai.com/api
openai.api_key = "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"


def ChatGPT(pregunta):

    prompt = pregunta

    completion = openai.Completion.create(engine="text-davinci-003",
                                          prompt=prompt,
                                          max_tokens=2048)

    return completion.choices[0].text




document = 'Sustantivos_Aleman_Completo.xlsx'

Book = load_workbook(document)

Page = Book['Hoja 1']




contadorVertical = 1

while (str(Page.cell(contadorVertical,1).value) != 'None'):

    #fill the column of "Sustantivo aleman"
    if (str(Page.cell(contadorVertical,2).value) == 'None'):

        Page.cell(contadorVertical,2).value = ChatGPT("Traduceme esta palabra del español al aleman : " + str(Page.cell(contadorVertical,1).value))
        time.sleep(20)
        print(str(Page.cell(contadorVertical,1).value) + " - " + str(Page.cell(contadorVertical,2).value))
        
    #fill the column of "Palabra plural en aleman"
    if (str(Page.cell(contadorVertical,3).value) == 'None'):

        #respuesta = ChatGPT("escribe " + str(Page.cell(contadorVertical,2).value) + "en plural en aleman, escribeme solo el articulo dereminado y la palabra")
        respuesta = Page.cell(contadorVertical,2).value = ChatGPT("Traduceme esta palabra del español al aleman en plural : " + str(Page.cell(contadorVertical,1).value))
        print(respuesta)
        #respuesta = respuesta.split(":")
        Page.cell(contadorVertical,3).value = respuesta
        time.sleep(20)
        print(str(Page.cell(contadorVertical,2).value) + " - " + str(Page.cell(contadorVertical,3).value))
        
    
    contadorVertical += 1
    Book.save('Sustantivos_Aleman_Completo.xlsx')


Book.save('Sustantivos_Aleman_Completo.xlsx')
Book.close()