"""
    FuzzyMatch
    ==========

    This program processes the names and departments of individuals mentioned as
    having a significant impact on students and the course number and title of
    classes students mention as favorites during their time at Tufts.

    Written by:  James Garijo-Garde
            on:  4/25/2019
            for: The Tufts University Office of Institutional Research &
                 Evaluation (OIRE)
"""

from openpyxl import load_workbook
from fuzzywuzzy import fuzz

# for debugging
import os


# Provides an initial launch message to the user and prompts for the survey mode
def launch():
    print("\n###############################################################################")
    print("### Fuzzy String Matcher for 'Significant Impact' and 'Best Course' Surveys ###")
    print("###############################################################################\n")
    print("Type a two-letter mode code to begin, or visit")
    print("https://github.com/JdoubleG/OIRE-Fuzzy-String for documentation.\n")
    print("'si' = 'Significant Impact'  'bc' = 'Best Course'\n")
    mode = input("Mode code:  ")
    # loop to ensure input is valid
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
    print(os.getcwd())
    try:
        wb = load_workbook("si_input.xlsx")
    except:
        print("\nERROR: Input file could not be opened! Ensure that it is correctly named")
        print("'si_input.xlsx' and that it is in the same directory as this script before")
        print("attempting to run it again.")
        exit()
    try:
        ws = wb["RAW"]
    except:
        print("\nERROR: Data sheet could not be opened! Ensure that it is correctly named 'RAW'")
        print("before attempting to run the script again.")
        exit()
    end = ws.max_row
    
    ## fuzzy comparisons of names
    i = 2  # starting index of the new grouping of names
    # 3 = the column with the existing names
    prevName = str(ws.cell(2, 3).value)
    nextName = str(ws.cell(3, 3).value)
    if prevName == None:
        prevName = "-99"
    if nextName == None:
        nextName = "-99"
    # a dictionary of names appearing in the file similar to each other
    names = {prevName: 1}
    row_index = 2
    while row_index <= end:
        # 90 = threshold of minimum allowable similarity after passing it into
        # the FuzzyWuzzy algorithm.
        if fuzz.token_sort_ratio(prevName, nextName) >= 90:  # TEST THIS THRESHOLD!
            if nextName in names:
                # if an instance of the name is already in the dictionary,
                # update the value
                names[nextName] = names[nextName] + 1
            else:
                # otherwise, add the instance of the name to the dictionary
                names.setdefault(nextName, 1)
        else:
            # find the key with the max value (the on appearing most often in
            # the data)
            keys = list(names.keys())
            if len(keys) > 1:
                topScore = 0
                bestName = ""
                for j in keys:
                    # print("Loop 3")
                    if names[j] > topScore:
                        topScore = names[j]
                        bestName = j
            else:
                bestName = prevName
            # for all elements with similar names, give them the same name
            for j in range(i, row_index + 1):
                # 4 = the column with the new names
                ws.cell(j, 4).value = bestName
            # update the starting index of the new grouping of names
            i = row_index + 1
            # reset the names dictionary
            names = {nextName: 1}
        # update the row index
        row_index = row_index + 1
        prevName = str(ws.cell(row_index, 3).value)
        nextName = str(ws.cell(row_index + 1, 3).value)
        if prevName == None:
            prevName = "-99"
        if nextName == None:
            nextName = "-99"
    
    ## fuzzy comparison of department/subject
    print("Subject")
    i = 2  # starting index of the new grouping of subjects
    prevSubject = str(ws.cell(2, 5).value)
    nextSubject = str(ws.cell(3, 5).value)
    if prevSubject == None:
        prevSubject = "-99"
    if nextSubject == None:
        nextSubject = "-99"
    # a dictionary of subjects appearing in the file similar to each other
    subjects = {prevSubject: 1}
    row_index = 2
    while row_index <= end:
        # 90 = threshold of minimum allowable similarity after passing it into
        # the FuzzyWuzzy algorithm. The subject loop adds a check to make sure
        # it is comparing the same person for whom the subject is associated.
        if ws.cell(row_index, 4).value == ws.cell(row_index + 1, 4).value:  # TEST THIS THRESHOLD!
            if nextSubject in subjects:
                # if an instance of the subject is already in the
                # dictionary, update the value
                subjects[nextSubject] = subjects[nextSubject] + 1
            else:
                # otherwise, add the instance of the subject to the
                # dictionary
                subjects.setdefault(nextSubject, 1)
        else:
            # find the key with the max value (the on appearing most often in
            # the data)
            keys = list(subjects.keys())
            if len(keys) > 1:
                topScore = 0
                bestSubject = ""
                for j in keys:
                    if subjects[j] > topScore:
                        topScore = subjects[j]
                        bestSubject = j
            else:
                bestSubject = prevSubject
            # for all elements with similar subjects, give them the same
            # subject
            for j in range(i, row_index + 1):
                # 6 = the column with the new subjects
                ws.cell(j, 6).value = str(bestSubject)
            # update the starting index of the new grouping of subjects
            i = row_index + 1
            # reset the subjects dictionary
            subjects = {nextSubject: 1}
        # update the row index
        row_index = row_index + 1
        prevSubject = str(ws.cell(row_index, 5).value)
        nextSubject = str(ws.cell(row_index + 1, 5).value)
        if prevSubject == None:
            prevSubject = "-99"
        if nextSubject == None:
            nextSubject = "-99"

    ws.title = "OUTPUT"
    try:
        wb.save("si_output.xlsx")
    except:
        print("\nERROR: Data sheet could not be saved!")
        exit()
    print("\n\nSurvey fuzzy string processing complete! File saved as 'si_output.xlsx\n")


def bestCourse():
    try:
        wb = load_workbook("bc_input.xlsx")
    except:
        print("\nERROR: Input file could not be opened! Ensure that it is correctly named")
        print("'bc_input.xlsx' and that it is in the same directory as this script before")
        print("attempting to run it again.")
        exit()
    try:
        ws = wb["RAW"]
    except:
        print("\nERROR: Data sheet could not be opened! Ensure that it is correctly named 'RAW'")
        print("before attempting to run the script again.")
        exit()
    # do fuzzy string stuff
    # end fuzzy string stuff
    ws.title = "OUTPUT"
    try:
        wb.save("bc_output.xlsx")
    except:
        print("\nERROR: Data sheet could not be saved!")
        exit()
    print("\n\nSurvey fuzzy string processing complete! File saved as 'bc_output.xlsx\n")


#########  MAIN  #########
if __name__ == "__main__":
    launch()
    input("Press the enter key to exit.")