OIRE-Fuzzy-String
=================

A fuzzy string matching script for the Tufts OIRE's Significant Impact and Best
Courses surveys.

Written by:  James Garijo-Garde | 
        for: The Tufts University Office of Institutional Research & Evaluation
             (OIRE) |
        in:  April, 2019

## Purpose of this Script
... ...

## Steps before Running
Before running this program, make sure you have
[Python 3](https://www.python.org/downloads) and
[pip](https://pypi.org/project/pip) installed on your computer, as well as the
[FuzzyWuzzy](https://github.com/seatgeek/fuzzywuzzy) and
[openpyxl](https://openpyxl.readthedocs.io/en/stable) Python packages.

It is then important to make sure the Excel file is ready for the script.
1. Ensure that the survey's Excel file is in the same directory as the script.
2. Ensure that the file is named appropriately.
   - If you are processing the Significant Impact survey data, the file must be
     named "si_input.xlsx".
   - If you are processing the Best Course survey data, the file must be named
     "bc_input.xlsx".
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
   - This will be the "BC: Course Number" and "BC: Course Title" columns in the
     Best Course survey. There should be new empty columns in columns E and G.

## Steps to Completion
- [x] Outline the steps that must be taken before the Excel file is read by the
  script
- [x] Provide users with an initial informational launch and the option to
  select the survey mode ("Significant Impact" or "Best Course")
- [ ] Implement the algorithm for processing the "Significant Impact" survey
- [ ] Implement the algorithm for processing the "Best Course" survey