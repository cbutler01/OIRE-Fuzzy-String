OIRE-Fuzzy-String
=================

A fuzzy string matching script for the Tufts OIRE's Significant Impact and Best
Courses surveys.

Written by:  James Garijo-Garde | 
        for: The Tufts University Office of Institutional Research and
             Evaluation (OIRE) |
        in:  April and May, 2019

## Purpose of this Program
This program is a command line script that automates a survey processing step
for the
[Tufts University Office of Institutional Research and Evaluation (OIRE)](https://provost.tufts.edu/institutionalresearch).
This program processes the names and departments of individuals mentioned as
having a significant impact on students and the course number and title of
classes students mention as favorites during their time at Tufts.

## Steps to Completion
- [x] Outline the steps that must be taken before the Excel file is read by the
  script
- [x] Provide users with an initial informational launch and the option to
  select the survey mode ("Significant Impact" or "Best Course")
- [x] Implement the algorithm for processing the "Significant Impact" survey
- [x] Implement the algorithm for processing the "Best Course" survey
- [ ] Add matching of professor names in the "Best Course" survey
- [x] Provide output acknowledging the completion of the script and outline the
  expected output
- [ ] Debug and adjust fuzzy string thresholds

## Steps before Running
Before running this program, make sure you have
[Python 3](https://www.python.org/downloads) and
[pip](https://pypi.org/project/pip) installed on your computer, as well as the
[FuzzyWuzzy](https://chairnerd.seatgeek.com/fuzzywuzzy-fuzzy-string-matching-in-python)
and [openpyxl](https://openpyxl.readthedocs.io/en/stable) Python packages.

It is then important to make sure the Excel file is ready for the script.
1. Ensure that the survey's Excel file is in the same directory as the script.
2. Ensure that the file is named appropriately.
   - If you are processing the Significant Impact survey data, the file must be
     named `si_input.xlsx`.
   - If you are processing the Best Course survey data, the file must be named
     `bc_input.xlsx`.
3. Rename the sheet in the Excel file with the data needed by the script "RAW".
4. Run a spell check on the appropriate columns:
   - For the Significant Impact survey, spell check the "Individual:-Department/
     office/place of work" column.
   - For the Best Course survey, spell check the "BC: Course Title" column.
5. Sort the Significant Impact survey data alphabetically by the
   "Individual:-Name" column and the Best Course survey alphabetically by the
   "BC: Department" column.
6. In the Significant Impact survey, try to find and split all the instances of
   grouped faculty: select the entire "Individual:-Name column" (column C) and
   perform a search for "and " (note the trailing space). This is not entirely
   necessary, but it will make things easier down the road.
7. Insert new columns to the right of the columns undergoing processing:
   - This will be the "Individual:-Name" and "Individual:-Department/office/
     place of work" columns in the Significant Impact survey. There should be
     new empty columns in columns D and F.
   - This will be the "BC: Course Number", "BC: Course Title", "BC: Prof First
     Name", and "BC: Prof Last Name" columns in the Best Course survey. There
     should be new empty columns in columns E and G.
8. The Best Course survey is split into two phases. Complete the prompted steps
   after phase 1 in order to proceed with phase 2.

## Running the Script
* Run the script by double-clicking the script file in your file explorer.
* Run the script with Python in a Bash command line using the following syntax:
  `python FuzzyMatch.py`. Note that on some systems you may need to replace
  `python` with `python3`.

## Expected Output
When the script has finished running, you should expect to see a file named
either `si_output.xlsx` or `bc_output.xlsx` depending on which survey you
processed in the same directory as the script and the input files. Regardless of
the file name, the file should contain a sheet entitled "OUTPUT" in which the
previously empty columns inserted by the user are filled with new values.

## Potential Next Steps
This algorithm needs to be debugged, and edge cases must be considered! The
current implementation does not use fuzzy string recognition for the department
category in the Significant Impact survey. It might be a good idea to somehow
include that as a means of including similar subject names that are not exactly
the same as instances of the same subject, which would in theory lead to the
algorithm being better at picking a value for that category. It would also be a
good idea to modularize the code more than it is currently since the steps in
each of the sections of the program are not overtly dissimilar.

Further, there is currently no implementation to match the names of the
professors of the courses in the Best Course survey.