"""
# get the different column categories (i.e. name, email, address, etc.)
columns = teacher_registration.columns;

# the number of rows of the sheet (i.e. the number of individuals who submitted form data)
numRows = len(teacher_registration)

# pd.isna(thing) --> checks if a thing from a cell is NaN or not

# list(team_registration.iloc[:, [8, 13, 18, 23, 28, 33, 38]]) --> locating column by its index, and converting to a list
"""

# imports
import pandas as pd
import math

# all the lunch options as displayed in the actual Google Form that participants fill out (modify the options as needed)
lunch_options = {
    "2 slices of pepperoni": {"pepperoni": 2},
    "2 slices of cheese": {"cheese": 2},
    "2 slices of vegetarian": {"vegetarian": 2},
    "1 slice of pepperoni and 1 slice of cheese": {"pepperoni": 1, "cheese": 1},
    "1 slice of pepperoni and 1 slice of vegetarian": {"pepperoni": 1, "vegetarian": 1},
    "1 slice of cheese and 1 slice of vegetarian": {"cheese": 1, "vegetarian": 1},
    "1 slice of pepperoni": {"pepperoni": 1},
    "1 slice of cheese": {"cheese": 1},
    "1 slice of vegetarian": {"vegetarian": 1}
}

# master list of all the pizza types and how many SLICES of each pizza are required
lunch_master_list = {
    "pepperoni": 0,
    "cheese": 0,
    "vegetarian": 0
}

############################################################################
# BEGIN PRE-PROCESSING EXCEL FILES
############################################################################

# enter the Excel file name of the teacher registration data
teacher_registration_filename = "Sample Teacher Registration Data.xlsx"

# enter the Excel file name of the team registration data
team_registration_filename = "Sample Team Registration Data.xlsx"

# read in the Excel files
teacher_registration = pd.read_excel(teacher_registration_filename)
team_registration = pd.read_excel(team_registration_filename)

############################################################################
# END PRE-PROCESSING
############################################################################

############################################################################
# BEGIN DATA VALIDATION
############################################################################

# list of all emails aggregated from the teacher registration form
teacher_emails = list(teacher_registration["Email Address | Adresse courriel"])
teacher_emails.extend(list(teacher_registration["Email of Supervisor #2 | Adresse courriel du superviseur #2"]))

# stores the emails from the team registration form that don't match any of the emails from the teacher' form
error_emails = []

# outputs the resulting teacher_emails list
print(teacher_emails)

# iterate through the team registration form's emails and cross-reference them with the list of teacher emails
for email in team_registration["Primary Supervisor Email Address | Adresse courriel du(de la) superviseur(e) primaire"]:
    if email not in teacher_emails:
        error_emails.append(email)

# outputs any error emails, if our list of error emails is non-empty
if error_emails:
    print("\nThe following emails from the Team Registration form did not have a matching teacher email from the Teacher Registration form:\n")
    for error_email in error_emails:
        print(error_email)
    exit(1)  # stop execution of code
else:
    print("\nAll emails from the Team Registration form matched with at least one teacher email from the Teacher Registration form.\n")

############################################################################
# END DATA VALIDATION
############################################################################

############################################################################
# BEGIN PIZZA ORDERS PROCESSING
############################################################################

groups = teacher_registration["School or Community Name | Nom de l'école ou du communauté"]
teacher_lunch_choices = list(teacher_registration["Lunch Choice"])
teacher_lunch_choices.extend(list(teacher_registration["Lunch Choice.1"]))

# 8, 13, 18, 23, 28, 33, 38 representative of the index of the columns (0-based indexing) --> 1 column/participant
team_lunch_choices = list(team_registration["Lunch Choice"])

for i in range(1, 7, 1):
    team_lunch_choices.extend(list(team_registration[f'Lunch Choice.{i}']))

# creating master list of pizza choices (by splices) for the entire competition
for choice in teacher_lunch_choices:
    if choice in lunch_options:
        for pizza_choice in lunch_options[choice]:
            lunch_master_list[pizza_choice] += lunch_options[choice][pizza_choice]

for choice in team_lunch_choices:
    if choice in lunch_options:
        for pizza_choice in lunch_options[choice]:
            lunch_master_list[pizza_choice] += lunch_options[choice][pizza_choice]

# for pizza in lunch_master_list:
#     lunch_master_list[pizza] = math.ceil(lunch_master_list[pizza] / 8)

print(lunch_master_list)

# print(df["Full Name | Nom complet"][1])

#for column in columns:
#    i = 0
#    while i < numRows:
#       print(df[column][i])
#        i += 1

string = "Lunch Choice"

#print(string in columns) # check if column category is in our array of categories

############################################################################
# END PIZZA ORDERS PROCESSING
############################################################################

