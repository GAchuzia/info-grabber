import re
import docx2txt
import pandas as pd
import tkinter as tk

from tkinter import filedialog
from tkinter.filedialog import *
from docx2python import docx2python

"""
DISCLAIMER
This is an older version of the application, and has some few bugs in it.
As noted by DJ, the Survey Data button is not working properly (will be fixed!) and
the Survey Data Results button does absolutely nothing (will be fixed in new application). 
"""

root = tk.Tk() # Manages the components of the tkinter application

# Sets the width and height of the application
canvas = tk.Canvas(root, width = 650, height = 150)
canvas.grid(columnspan = 4, rowspan = 10)

# Displays the app's instructions
app_instructs = tk.Label(root, text = "Choose Action:", font = "Calibri")
#app_instructs.place(relx = 0.5, rely = .7, anchor = "center")
app_instructs.grid(column = 2, row = 1)


# Selects and save files for Building Data
def buildingData_file_selection():
    buildingData_text.set("loading...") # displays after clicking the browse button
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
    buildingData_text.set("Extract Building Data")


# Selects and save files for Survey Data
def surveyData_file_selection():
    surveyData_text.set("loading...") 
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
        #roof_types =  re.findall("(?<=Roof Type\s\s)(.*)(?=Roof Access)",formatted_filecontents)
        #roof_condition = re.findall( ,formatted_filecontents)
        #roof_life = re.findall(,formatted_filecontents)
        
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
                   # "ageRoof": str().strip("[]").strip("()"),
                   # "conditionRoof": str().strip("[]").strip("()"),
                   # "remainingLifeRoof": str().strip("[]").strip("()"),
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
    surveyData_text.set("Extract Survey Data") 

# Selects and saves file for the Extracted Images
def resultsData_file_selection():
    surveyResultsData_text.set("loading...")
    filetypes = [("DOCX Files", "*.docx")]
    selected_files = askopenfilenames(filetypes = filetypes)
    save_path = "docx2sql_project\Images"

    for file in selected_files:
        extracted_images = docx2txt.process(file, save_path) #Does NOT extract all images :(
    surveyResultsData_text.set("Extract Images")




# Gives the Building Data button its attributes
buildingData_text = tk.StringVar()
buildingData_btn = tk.Button(root, textvariable = buildingData_text, command = buildingData_file_selection, font = "Calibri", bg = "#007940", fg = "white", height = 2, width = 20 )
buildingData_text.set("Extract Building Data")
buildingData_btn.grid(column = 1, row =2)

# Gives the Survey Reults Data  button its attributes
surveyResultsData_text = tk.StringVar()
surveyResultsData_btn = tk.Button(root, textvariable = surveyResultsData_text, command = resultsData_file_selection,font = "Calibri", bg = "#007940", fg = "white", height = 2, width = 20 )
surveyResultsData_text.set("Extract Images")
surveyResultsData_btn.grid(column = 2, row = 2)

# Gives the Survey Data button its attributes
surveyData_text = tk.StringVar()
surveyData_btn = tk.Button(root, textvariable = surveyData_text, command = surveyData_file_selection,font = "Calibri", bg = "#007940", fg = "white", height = 2, width = 20 )
surveyData_text.set("Extract Survey Data")
surveyData_btn.grid(column = 3, row = 2)

root.mainloop() #Also manages all tkinter components, do NOT put any code below this: it won't work