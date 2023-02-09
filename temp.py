
import os
import re
import cv2
import docx2txt
import tkinter as tk
import customtkinter
from tkinter import *
from tkinter.filedialog import *
import tkinter.scrolledtext as st
from docx2python import docx2python

# Constants
ROOF_TYPES = ["4-Ply","Single ply", "Shingled","Modified Bitumen"]

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("dark-blue")

root = customtkinter.CTk()
root.geometry("300x300")
root.title("Roof Images Extraction")
root.resizable(width=False, height=False)

# Creates a new folder name if wanted folder name already exists
def nextnonexistent(f):
    fnew = f
    root, ext = os.path.splitext(f)
    i = 0
    while os.path.exists(fnew):
        i += 1
        fnew = '%s_%i%s' % (root, i, ext)
    return fnew
  
# Selects and saves file for the Extracted Images
def imageDataExtraction():
    
    filetypes = [("DOCX Files", "*.docx")]
    selected_files = askopenfilenames(filetypes = filetypes)  

    # Locates the users download path
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads') # The full path to a users download folder
    os.chdir(desktop_path) # changes directory to the users downloads folder
    download_path = desktop_path + "\\" + nextnonexistent("Extracted_Images") # The path we want to create for extracted images
    os.mkdir(download_path)
    os.chdir(download_path) 

    counter  = 0
    roof_types_seen = set()

    # Finds the roof type for a given file
    for file in selected_files:
        file_contents = docx2txt.process(file)
        formatted_filecontents = file_contents.replace("\n"," ")
        survey_rooftype = re.findall("(?<=Roof Type\s\s)(.*)(?=\s\sRoof Access)", formatted_filecontents)
        survey_rooftype = str(survey_rooftype).strip("[]").strip("'").strip()
        counter += 1

    # Checks each survey for its roof type and extracts and organizes images into folders based on roof type
        if survey_rooftype not in ROOF_TYPES:
            survey_rooftype = "Unknown Rooftype"

        file_path = f"{survey_rooftype}\\File {counter} Images"
        os.makedirs(file_path)
        extracted_imgs = docx2txt.process(file, download_path +"\\" + file_path)

        roof_types_seen.add(survey_rooftype)  
           
    # Text displayed whenever "Get Images" button is clicked
    def image_info():
        info_box.insert("end", 
        f"New folder created in {os.getlogin()}'s Downloads folder\n"
        f'Number of selected files: {len(selected_files)}\n'
        f'Roof types: {str(roof_types_seen).strip("{}").strip("")} \n\n')
        info_box.configure(state="disabled")
    info_box.configure(state="normal")
    image_info()
   

frame = customtkinter.CTkFrame(master = root)
frame.pack(pady=20, padx = 20, fill="both", expand =True)

button = customtkinter.CTkButton(master=frame, text="GET IMAGES", command = imageDataExtraction, width= 400)
button.pack(pady=6, padx=10)

info_box = customtkinter.CTkTextbox(master=frame, width=400)
info_box = Text(master=frame,wrap=WORD, height= 200, width=200, font=("Computer Modern", 6)) 
info_box.pack( pady=6, padx=10)
info_box.insert("0.0", 'Instructions:\nPress the button above to select files and extract their images based on roof type.\n\n')

info_box.configure(state="disabled")

info_box_scrollbar = customtkinter.CTkScrollbar(master=frame, command=info_box.yview)
info_box_scrollbar.pack(pady=12, padx=10)

root.mainloop()