import PySimpleGUI as sg
import pandas as pd

# Set PySimpleGUI theme
sg.theme('DarkBlue')

# Set the name of the Excel file to use
EXCEL_FILE = "Excel Projects 1.xlsx"

# Load the Excel file into a Pandas dataframe
df = pd.read_excel(EXCEL_FILE)

# Define the layout of the input form
layout = [
    [sg.Text("EE_Course"), sg.Combo(['EE_LE', 'EE_LAB', 'EE_PR'], key="EE_Course")],
    [sg.Text("EE_Chapter"), sg.InputText(key="EE_Chapter")],
    [sg.Text("Link"), sg.InputText(key="Link")],
    [sg.Text("Comments"), sg.Multiline(key="Comments")],
    [sg.Submit(), sg.Button('Clear'), sg.Exit()]
]

# Create the input form window
window = sg.Window('EE_S2', layout)

# Define function to clear the input fields
def clear_input():
    for key in values:
        window[key]('')
    
    return 

# Main event loop for the input form window
while True:
    event, values = window.read()

    # Exit the program if the user closes the window or clicks the "Exit" button
    if event == sg.WIN_CLOSED or event == 'Exit':
        break

    # If the user clicks the "Submit" button, add the data to the Excel file and clear the input fields
    if event == 'Submit':
        df = df.append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)  # Set index=False to exclude the row numbers from the output file
        sg.popup("Data Entered")
        clear_input()

    # If the user clicks the "Clear" button, clear the input fields
    if event == 'Clear':
        clear_input()

# Close the input form window
window.close()
