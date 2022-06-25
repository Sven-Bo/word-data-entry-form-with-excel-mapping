import datetime  # Core Python Module
from pathlib import Path  # Core Python Module

import pandas as pd  # pip install pandas openpyxl
import PySimpleGUI as sg  # pip install PySimpleGUI
from docxtpl import DocxTemplate  # pip install docxtpl


# --- LOCATE FILE PATHS ---
base_dir = Path(__file__).parent
mapping_table_path = base_dir / "mapping_table.xlsx"
document_path = base_dir / "template.docx"


# --- WINDOW LAYOUT ---
layout = [
    [sg.Text("Enter your name here:"), sg.Input(key="NAME", do_not_clear=False)],
    [sg.Text("Enter your number here:"), sg.Input(key="NUMBERINPUT", do_not_clear=False)],
    [sg.Button("Submit"), sg.Exit()],
]

window = sg.Window("Data Entry Form", layout, element_justification="right")


# --- EVENT LOOP ---
while True:
    event, values = window.read()
    if event in (sg.WINDOW_CLOSED, "Exit"):
        break
    if event == "Submit":
        name, number_input = values["NAME"], values["NUMBERINPUT"]

        if name and number_input.isdigit():
            number_input = int(number_input)
            if 40 <= number_input <= 140:
                # Get the value pair from Excel & add to dict
                df = pd.read_excel(mapping_table_path)
                text = df.loc[df["NUMBERINPUT"] == number_input].iloc[0]["TEXT"]
                values["TEXT"] = text

                # Render the template, save new word document & inform user
                doc = DocxTemplate(document_path)
                doc.render(values)
                output_path = base_dir / f"{name}-data.docx"
                doc.save(output_path)
                sg.popup("File saved", f"File has been saved here: {output_path}")
            else:
                sg.popup_error("Please enter a number between 40 & 140")

        else:
            sg.popup_error("Invalid Entry")

window.close()
