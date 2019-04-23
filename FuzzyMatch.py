"""
    FuzzyMatch
    ==========

    [what this program does]

    Written by:  James Garijo-Garde
            on:  4/23/2019
            for: The Tufts University Office of Institutional Research &
                 Evaluation (OIRE)
"""

from openpyxl import load_workbook

# Provides an initial launch message to the user and prompts for the survey mode
def launch():
    print("\n###############################################################################")
    print("### Fuzzy String Matcher for 'Significant Impact' and 'Best Course' Surveys ###")
    print("###############################################################################\n")
    print("Type a two-letter mode code to begin, or visit")
    print("https://github.com/JdoubleG/OIRE-Fuzzy-String for documentation.\n")
    print("'si' = 'Significant Impact'  'bc' = 'Best Course'\n")
    mode = input("Mode code:  ")
    loop = True
    while loop == True:
        if mode == "si":
            print("### Significant Impact ###")
            loop = False
            # Launches the significant impact survey processor
            significantImpact()
        elif mode == "bc":
            print("### Best Course ###")
            loop = False
            # Launches the best course survey processor
            bestCourse()
        else:
            mode = input("Input unrecognized. Please re-enter mode code:  ")

def significantImpact():
    try:
        wb = load_workbook("si_input.xlsx")
    except:
        print("\nERROR: Input file could not be opened! Ensure that it is correcty named")
        print("'si_input.xlsx' and that it is in the same directory as this script before")
        print("attempting to run it again.")
        exit()
    try:
        ws = wb["RAW"]
    except:
        print("\nERROR: Data sheet could not be opened! Ensure that it is correcty named 'RAW'")
        print("before attempting to run the script again.")
        exit()
    # do fuzzy string stuff
    ws.title = "OUTPUT"
    wb.save("si_output.xlsx")
    print("\n\nSurvey fuzzy string processing complete! File saved as 'si_output.xlsx\n")

def bestCourse():
    try:
        wb = load_workbook("bc_input.xlsx")
    except:
        print("\nERROR: Input file could not be opened! Ensure that it is correcty named")
        print("'bc_input.xlsx' and that it is in the same directory as this script before")
        print("attempting to run it again.")
        exit()
    try:
        ws = wb["RAW"]
    except:
        print("\nERROR: Data sheet could not be opened! Ensure that it is correcty named 'RAW'")
        print("before attempting to run the script again.")
        exit()
    # do fuzzy string stuff
    ws.title = "OUTPUT"
    wb.save("bc_output.xlsx")
    print("\n\nSurvey fuzzy string processing complete! File saved as 'bc_output.xlsx\n")


#########  MAIN  #########
if __name__ == "__main__":
    launch()