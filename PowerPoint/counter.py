from pptx import Presentation
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def count(p):

    #list for all strings
    text = []

    #get all text
    for slide in p.slides: #each slide
        for shape in slides.shape:  #each shape
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

        text = count(pres)
        print(text)

    except Exception as exc:

        print("File not supported!")
        pres = "null"
