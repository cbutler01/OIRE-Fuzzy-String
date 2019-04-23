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
    print("Type a two-letter mode code to begin:")
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
    wb = load_workbook("si_input.xlsx")
    rawSheet = wb["RAW"]

def bestCourse():
    pass

if __name__ == "__main__":
    launch()