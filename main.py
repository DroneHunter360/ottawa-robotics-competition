"""
Author: Steven Hua (Event planning)
Usage: Aggregate Excel file registration data for the ORC into meaningful data used for logistical planning purposes
Last modified: February 21, 2023
"""

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

tshirt_options = {
    "S",
    "M",
    "L",
    "XL",
    "XXL"
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

# create dictionary to store key-value pairs relating supervisors' email addresses to the school/community group
# associated to it
school_email_pairs = {}

# iterate through the team registration form's emails and cross-reference them with the list of teacher emails
for email in team_registration["Primary Supervisor Email Address | Adresse courriel du(de la) superviseur(e) primaire"]:
    if email not in teacher_emails:
        error_emails.append(email)

# outputs any error emails, if our list of error emails is non-empty
if error_emails:
    print("The following emails from the Team Registration form did not have a matching teacher email from the Teacher Registration form:\n")
    for error_email in error_emails:
        print(error_email)
    exit(1)  # stop execution of code
else:
    print("All emails from the Team Registration form matched with at least one teacher email from the Teacher Registration form.\n")

############################################################################
# END DATA VALIDATION
############################################################################

############################################################################
# BEGIN DICTIONARY SETUP
############################################################################

model = {}

# list of the school/community groups
groups = teacher_registration["School or Community Name | Nom de l'école ou du communauté"]

# iterate through each group
for i in range(0, len(groups), 1):
    # emails attribute contains a list of all teacher/supervisor emails for that group
    # members attribute contains a separate object for each member of that group, with each object containing
    # information about their lunch choice, t-shirt size, etc.
    model[groups[i]] = {"emails": [], "members": {}}

    # only add an email if it is a non-empty cell field
    if not pd.isna(teacher_registration.iloc[i]["Email Address | Adresse courriel"]):
        model[groups[i]]["emails"].append(teacher_registration.iloc[i]["Email Address | Adresse courriel"])

    # only add an email if it is a non-empty cell field
    if not pd.isna(teacher_registration.iloc[i]["Email of Supervisor #2 | Adresse courriel du superviseur #2"]):
        model[groups[i]]["emails"].append(teacher_registration.iloc[i]["Email of Supervisor #2 | Adresse courriel du superviseur #2"])

    # add TEACHERS' lunch choices and shirt sizes to their dictionary object
    teacher_name = teacher_registration.iloc[i]["Full Name | Nom complet"]
    model[groups[i]]["members"][teacher_name] = {"lunch_choice": '', "shirt_size": '', "isStudent": False}
    # only add a lunch choice if it is a non-empty cell field
    if not pd.isna(teacher_registration.iloc[i]["Lunch Choice"]):
        model[groups[i]]["members"][teacher_name]["lunch_choice"] = teacher_registration.iloc[i]["Lunch Choice"]
    if not pd.isna(teacher_registration.iloc[i]["T-Shirt Size"]):
        model[groups[i]]["members"][teacher_name]["shirt_size"] = teacher_registration.iloc[i]["T-Shirt Size"]

    teacher_name = teacher_registration.iloc[i]["Full Name of Supervisor #2 | Nom complet du superviseur #2"]
    model[groups[i]]["members"][teacher_name] = {"lunch_choice": '', "shirt_size": '', "isStudent": False}
    if not pd.isna(teacher_registration.iloc[i]["Lunch Choice.1"]):
        model[groups[i]]["members"][teacher_name]["lunch_choice"] = teacher_registration.iloc[i]["Lunch Choice.1"]
    # only add a shirt size if it is a non-empty cell field
    if not pd.isna(teacher_registration.iloc[i]["T-Shirt Size.1"]):
        model[groups[i]]["members"][teacher_name]["shirt_size"] = teacher_registration.iloc[i]["T-Shirt Size.1"]

# aggregate the information for each student of each group
for row in range(0, len(team_registration), 1):
    # determine which school/community grop this member is associated with
    primary_email = team_registration.iloc[row]["Primary Supervisor Email Address | Adresse courriel du(de la) superviseur(e) primaire"]
    group_school = ""

    for group in model:
        if primary_email in model[group]["emails"]:
            group_school = group
            break

    # add STUDENTS' lunch choices to their dictionary object
    student_name = team_registration.iloc[row]['Full Name of Student #1 | Nom complet d\'élève #1']
    model[group_school]["members"][student_name] = {"lunch_choice": '', "shirt_size": '', "team_name": '', "isStudent": True} # specify that each student is an object
    model[group_school]["members"][student_name]["lunch_choice"] = team_registration.iloc[row]['Lunch Choice']
    model[group_school]["members"][student_name]["shirt_size"] = team_registration.iloc[row]['T-Shirt Size']
    model[group_school]["members"][student_name]["team_name"] = team_registration.iloc[row]['Team Name | Nom d\'équipe']
    for j in range(1, 7, 1):
        student_name = team_registration.iloc[row][f'Full Name of Student #{j+1} | Nom complet d\'élève #{j+1}']
        model[group_school]["members"][student_name] = {"lunch_choice": '', "shirt_size": '', "team_name": '', "isStudent": True}  # specify that each student is an object

        if not pd.isna(team_registration.iloc[row][f'Lunch Choice.{j}']):
            model[group_school]["members"][student_name]["lunch_choice"] = team_registration.iloc[row][f'Lunch Choice.{j}']

        if not pd.isna(team_registration.iloc[row][f'T-Shirt Size.{j}']):
            model[group_school]["members"][student_name]["shirt_size"] = team_registration.iloc[row][f'T-Shirt Size.{j}']

        if not pd.isna(team_registration.iloc[row]['Team Name | Nom d\'équipe']):
            model[group_school]["members"][student_name]["team_name"] = team_registration.iloc[row]['Team Name | Nom d\'équipe']

print(model)
print("\n")
############################################################################
# END DICTIONARY SETUP
############################################################################

############################################################################
# FUNCTIONS FOR LIST CREATIONS
############################################################################

def create_general_pizza_list():
    result = {
        "pepperoni": 0,
        "cheese": 0,
        "vegetarian": 0,
        "pepperoni_pizzas": 0,
        "cheese_pizzas": 0,
        "vegetarian_pizzas": 0
    }

    for g in model:
        for member in model[g]["members"]:
            lunch_choice = model[g]["members"][member]["lunch_choice"]
            if lunch_choice in lunch_options:
                for pizza_type in lunch_options[lunch_choice]:
                    result[pizza_type] += lunch_options[lunch_choice][pizza_type]

    result["pepperoni_pizzas"] = math.ceil(result["pepperoni"] / 8)
    result["cheese_pizzas"] = math.ceil(result["cheese"] / 8)
    result["vegetarian_pizzas"] = math.ceil(result["vegetarian"] / 8)
    return result

def create_pizza_list_by_school():
    result = {}

    for g in model:
        result[g] = {"supervisors": {}, "students": {}}
        for member in model[g]["members"]:
            if (model[g]["members"][member]["isStudent"]):
                result[g]["students"][member] = model[g]["members"][member]["lunch_choice"]
            else:
                result[g]["supervisors"][member] = model[g]["members"][member]["lunch_choice"]
    return result

print(create_pizza_list_by_school())
