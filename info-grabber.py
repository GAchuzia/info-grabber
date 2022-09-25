import os
import re
import docx2txt
import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter.filedialog import *
from docx2python import docx2python

root = tk.Tk() # Manages the components of the tkinter application
root.title("Info-Grabber")
# Sets the width and height of the application
canvas = tk.Canvas(root, width = 540, height = 150)
canvas.grid(columnspan = 2, rowspan = 2, )



# Selects and save files for Building Data
def buildingDataExtraction():
    buildingDataText.set("loading...") # displays after clicking the browse button
    filetypes = [("DOCX Files", "*.docx")] #rename variable
    selected_files = askopenfilenames(filetypes = filetypes)
    
    # docx to table
    first_time = True # Make sure varibale is not declared more than once
    for file in selected_files:
        file_contents = docx2txt.process(file)
        formatted_filecontents = file_contents.replace("\n"," ") #Turns docx file into formatted text

        company_name = re.findall("(?<=Company Name\s\s)(.*)(?=Property Manager)|(?<=Company)(.*)(?=Contact Name)|(?<=Company:)(.*)(?=PREPARED BY:)",formatted_filecontents)
        property_manager = re.findall("(?<=Property Manager\s\s)(.*)(?=Building Address)|(?<=Contact Name)(.*)(?=Building Name)|(?<=Contact:)(.*)(?=Building Name:)|(?<= Property Manager)(.*)(?=Address)",formatted_filecontents)
        building_address = re.findall("(?<=Building Address\s\s)(.*)(?=City)|(?<= Building Address:)(.*)(?=,)|(?<=Building Address)(.*)(?=Roof Type:)|(?<=Building Address)(.*)(?=Job #)",formatted_filecontents)
        file_city = re.findall("(?<=City\s\s)(.*)(?=Province)|(?<= ,)(.*)(?= Type of Roof: )",formatted_filecontents)
        file_province = re.findall("(?<=Province\s\s)(.*)(?=Job #)",formatted_filecontents)
        job_number = re.findall("(?<=Job #\s\s)(.*)(?=P.O.#)", formatted_filecontents)
        po_number = re.findall("(?<=P.O.#\s\s)(.*)(?=GENERAL INFORMATION)",formatted_filecontents)

        new_dict = {"Company Name": str(company_name).strip("[]").strip("()"),
                    "Property Manager": str(property_manager).strip("[]").strip("()"),
                    "Building Address": str(building_address).strip("[]").strip("()"),
                    "City": str(file_city).strip("[]").strip("()"),                   
                    "Province": str(file_province).strip("[]").strip("()"),
                    "Job #": str(job_number).strip("[]").strip("()"),
                    "P.O #": str(po_number).strip("[]").strip("()"),
                    }

        new_dict = {k: [v] for k, v in new_dict.items()} 
        if first_time: # Same as first_time == True
            first_time = False
            df2export = pd.DataFrame.from_dict(new_dict) # Makes dataframe if its the first time
        else: 
            df2export = df2export.append(pd.DataFrame(new_dict), ignore_index = True) # Adds to the existing dataframe ()
            df2export.dropna()

    saving_path = filedialog.asksaveasfile(mode = 'w', defaultextension = ".csv")
    df2export.to_csv(saving_path, index = False, line_terminator='\n')
    saving_path.close()
    print(df2export)
    buildingDataText.set("Extract Building Data")



# Selects and save files for Survey Data
def surveyDataExtraction():
    surveyDataText.set("loading...") 
    filetypes = [("DOCX Files", "*.docx")] 
    selected_files = askopenfilenames(filetypes = filetypes)

    first_time = True 
    for file in selected_files:
        file_contents = docx2txt.process(file)
        formatted_filecontents = file_contents.replace("\n"," ")

        survey_code = re.findall("(?<=Job #\s\s)(.*)(?=P.O.#)",formatted_filecontents)
        job_date =  re.findall("(?<=Date of Job\s\s)(.*)(?<=2022)|(?<=Date of Scan\s\s)(.*)(?<=2022)|(?<=Date of Maintenance)(.*)(?<=2022)",formatted_filecontents)
        specified_work =  re.findall("(?<=d Work\s\s)(.*)(?=\s\sPilot)|(?<=Work\s\s)(.*)(?=\s\sLead)",formatted_filecontents)
        file_pilot =  re.findall("(?<=Lead Technicians\s\s)(.*)(?=\s\sValidated By)|(?<=Pilot)(.*)(?=\s\sValidated By)",formatted_filecontents)
        validated_by =  re.findall("(?<=Validated By\s\s)(.*)(?=\s\sWEATHER CONDITIONS:)",formatted_filecontents)
        temp_external =  re.findall("(?<=Temperature\s\s)(.*)(?=\sCloud Cover)|(?<=Exterior Temperature\s\s)(.*)(?<=\s\sNot Applicable  Interior Temperature)",formatted_filecontents) #Not Working Properly
        cloud_cover =  re.findall("(?<=Cloud Cover\s\s)(.*)(?=\s\sWind Condition)|(?<=Cloud Cover\s\s)(.*)(?=Wind Speed)",formatted_filecontents)
        wind_condition_speed =  re.findall("(?<=Wind Speed\s\s)(.*)(?=Wind Direction)|(?<=Wind Conditions\s\s)(.*)(?=Wind Direction)|(?<=Wind Conditions\s\s)(.*)(?=ROOF CONDITION:\s\s)",formatted_filecontents)
        wind_direction =  re.findall("(?<=Wind Direction\s\s)(.*)(?=ROOF CONDITION:\s\s)|(?<=Wind Direction)(.*)(?=\s\sBuilding Photo)",formatted_filecontents)
        construction_date =  re.findall("(?<=of Construction\s\s)(.*)(?=Roof Type)",formatted_filecontents)
        roof_access =  re.findall("(?<=Roof Access\s\s)(.*)(?<=Roof hatch)",formatted_filecontents) # Not working properly
        
        new_dict = {
                    "surveyCode": str(survey_code).strip("[]").strip("()").strip("''"),
                    "jobDate": str(job_date).strip("[]").strip("()"), 
                    "specifiedWork": str(specified_work).strip("[]").strip("()"),
                    "technician": str(file_pilot).strip("[]").strip("()"), 
                    "validator": str(validated_by).strip("[]").strip("()").strip("''"), 
                   # "typeRoof": str(roof_types).strip("[]").strip("()"),
                    "tempExternal": str(temp_external).strip("[]").strip("()").strip("''"), 
                    "cloud": str(cloud_cover).strip("[]").strip("()").strip("''"),
                    "Wind Condition": str(wind_condition_speed).strip("[]").strip("()").strip("''"),
                    "Wind Direction": str(wind_direction).strip("[]").strip("()").strip("''"),
                    "ageConstruction": str(construction_date).strip("[]").strip("()").strip("''"),
                    "accessRoof": str(roof_access).strip("[]").strip("()").strip("''"),
                    }          

        new_dict = {k: [v] for k, v in new_dict.items()}
        if first_time:
            first_time = False
            df2export = pd.DataFrame.from_dict(new_dict) 
        else: 
            df2export = df2export.append(pd.DataFrame(new_dict), ignore_index = True) 
            df2export.dropna()
        
    saving_path = filedialog.asksaveasfile(mode = 'w', defaultextension = ".csv")
    df2export.to_csv(saving_path, index = False, line_terminator='\n')
    saving_path.close()
    print(df2export)
    surveyDataText.set("Extract Survey Data") 



# Selects and saves file for the Extracted Images
def imageDataExtraction():
    imageDataText.set("loading...")
    filetypes = [("DOCX Files", "*.docx")]
    selected_files = askopenfilenames(filetypes = filetypes)
    
    # Locates the users download path
    desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Downloads')
    # The full path to a users download folder
    download_path = desktop_path + "\Extracted_Images\\"
    os.chdir(desktop_path)
           
    # Error checking to see if "Extracted_Image" folder exists, creates said folder if not       
    if os.path.exists(download_path) == True:
        print("Folder 'Extracted_Images' already exist.")
    else:
        new_imagery_directory = os.mkdir('Extracted_Images')
        print('A folder named "Extracted_Images" has been created.')
        
    count = 1
    # Makes a new image folder in users download folder for each file selected
    for file in selected_files:
        os.chdir(download_path) 
        new_folders = os.mkdir(f'File {count} Images')
        extracted_images = docx2txt.process(file, download_path + "\\" + f'File {count} Images')
        count += 1
    imageDataText.set("Extract Survey Data") 



# Renames selected files
def renameFile():
    pass



# Gives the Building Data button its attributes
buildingDataText = tk.StringVar()
buildingDataBtn = tk.Button(root, textvariable = buildingDataText, command = buildingDataExtraction, font = "Calibri", bg = "#007940", fg = "white", height = 1, width = 12 )
buildingDataText.set("Extract Building Data")
buildingDataBtn.grid(column = 0, row =0, pady=6, padx=6, sticky= "nsew")

# Gives the Survey Reults Data  button its attributes
imageDataText = tk.StringVar()
imageDataBtn = tk.Button(root, textvariable = imageDataText, command = imageDataExtraction,font = "Calibri", bg = "#007940", fg = "white", height = 1, width = 12 )
imageDataText.set("Extract Images")
imageDataBtn.grid(column = 1, row = 0, pady=6, padx=6, sticky= "nsew")

# Gives the Survey Data button its attributes
surveyDataText = tk.StringVar()
surveyDataBtn = tk.Button(root, textvariable = surveyDataText, command = surveyDataExtraction,font = "Calibri", bg = "#007940", fg = "white", height = 1, width = 12 )
surveyDataText.set("Extract Survey Data")
surveyDataBtn.grid(column = 0, row = 1, pady=6, padx=6, sticky= "nsew")

# Gives the Rename File button its attributes
renameFileText = tk.StringVar()
renameFileBtn = tk.Button(root, textvariable = renameFileText, command = renameFile,font = "Calibri", bg = "#007940", fg = "white", height = 1, width = 12 )
renameFileText.set("Rename File(s)")
renameFileBtn.grid(column = 1, row = 1, pady=6, padx=6, sticky= "nsew")

root.mainloop() #Also manages all tkinter components, do NOT put any code below this: it won't work