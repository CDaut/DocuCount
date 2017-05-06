from pptx import Presentation
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def count(p):

    #list for all strings
    text = []

    #get all text
    for slide in p.slides: #each slide
        for shape in slide.shapes:  #each shape
            if(not shape.has_text_frame): #continue if the shape has no text frames
                continue
            for paragraph in shape.text_frame.paragraphs: #for each paragraph

                for run in paragraph.runs: #for each run
                    text.append(run.text)  #get text and store

    return text


#####MAIN#####


pres = "null"

while(pres == "null"):

    #get file chooser
    Tk().withdraw()
    filename = askopenfilename()

    if(filename == ""):
        break

    try:

        #get presentation
        pres = Presentation(filename)

        #count file
        text = count(pres)

        #create list to store all words
        allWords = []
        allWordsTemp = []

        #each text line
        for n in range(len(text)):

            tempLine = text[n]

            #each word in tempLine
            for i in range(len(tempLine.split(" "))):

                #add each word
                tempWordsSplit = tempLine.split(" ")
                allWordsTemp.append(tempWordsSplit[i])



        #remove word if its just a space
        for n in range(len(allWordsTemp)):

            #if the word is a space it may not be counted
            if(allWordsTemp[n] == "" or allWordsTemp[n] == " "):
                    pass

            else:
                allWords.append(allWordsTemp[n])

        print("-------------------------------------------")
        print("Total words: " + str(len(allWords)))

    except Exception as exc:

        print("File not supported!")
        pres = "null"
        print(exc)
