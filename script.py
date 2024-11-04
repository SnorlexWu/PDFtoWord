import os
from pdf2docx import Converter

def convert_pdf_to_word(input_path,output_path):
    if not os.path.exists(output_path):
        print("Output path does ont exist")
    if not os.path.exists(input_path):
        print("Input path does not exist")

    #loop thorugh all files in the input folder
    for filename in os.listdir(input_path):
        if filename.endswith(".pdf"): #identity pdf in directory
            pdf_path = os.path.join(input_path,filename) #get the exact file pth for the pdf
            word_path = os.path.join(output_path,f"{os.path.splitext(filename)[0]}.docx") #set the output of the converted pdf

            try:  #convert function
                cv = Converter(pdf_path)
                cv.convert(word_path,start=0,end=None)
                cv.close()
                print(f"Converted: {pdf_path} to {word_path}")
            except Exception as e: #error handling
                print(f"Failed to convert {pdf_path}")

if __name__ == "__main__":
    input_path = "C:\\Users\\Alex Wu\\Documents\\Python Scripts\\PDF to Word Converter\\Input"
    output_path = "C:\\Users\\Alex Wu\\Documents\\Python Scripts\\PDF to Word Converter\\Output"
    convert_pdf_to_word(input_path, output_path)
