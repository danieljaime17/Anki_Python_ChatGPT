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

        respuesta = ChatGPT("schreibt den Plural von " + str(Page.cell(contadorVertical,2).value) + "mit seinem Artikel")
        
        if len(respuesta.split(" ")) == 2 or len(str(Page.cell(contadorVertical,1).value).split(" ")) != 2:
            print("la respuesta de gpt es correta")
            Page.cell(contadorVertical,3).value = respuesta
            print(str(Page.cell(contadorVertical,2).value) + " - " + str(Page.cell(contadorVertical,3).value))

        else:
            print("la respuesta de gpt no es correta: " + respuesta)
       
        time.sleep(20)
    
    #fill the column of "Frase en Aleman"
    if (str(Page.cell(contadorVertical,4).value) == 'None'):
        FraseAleman = ChatGPT("Schreibe einen Satz auf Deutsch mit diesem Wort: " + str(Page.cell(contadorVertical,2).value))
        Page.cell(contadorVertical,4).value = FraseAleman
        print(FraseAleman)
        time.sleep(20)

    #fill the column of frase en español
    if (str(Page.cell(contadorVertical,5).value) == 'None'):
        FraseEspañol = ChatGPT("Übersetze diesen Satz ins Spanische: " + str(Page.cell(contadorVertical,4).value))
        Page.cell(contadorVertical,5).value = FraseEspañol
        print(FraseEspañol)
        time.sleep(20)

    contadorVertical += 1
    Book.save('Sustantivos_Aleman_Completo.xlsx')


Book.save('Sustantivos_Aleman_Completo.xlsx')
Book.close()