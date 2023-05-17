"""
Author: Steven Hua
Usage: Aggregate Excel file registration data for the ORC into meaningful data used for logistical planning purposes
Last modified: May 11, 2023
"""

# imports
import pandas as pd
import csv
import math
import os
import requests
import datetime

# global invoice number variable
invoice_number = 2  # David Huynh specified it will start at 002

# Generate the invoice number string with leading zeros
invoice_number_str = "{:03d}".format(invoice_number)

# all the lunch options as displayed in the actual Google Form that participants fill out (modify the options as needed)
lunch_options = {
    "2 pepperoni pizza slices | 2 pointes de pizza au pepperoni": {"pepperoni": 2},
    "2 cheese pizza slices | 2 pointes de pizza au fromage": {"cheese": 2},
    "2 vegetarian pizza slices | 2 pointes de pizza végétarienne": {"vegetarian": 2},
    "1 pepperoni and 1 cheese pizza slices | 1 pointe de pizza au pepperoni et 1 pointe de pizza au fromage": {"pepperoni": 1, "cheese": 1},
    "1 pepperoni and 1 vegetarian pizza slices | 1 pointe de pizza au pepperoni et 1 pointe de pizza végétarienne": {"pepperoni": 1, "vegetarian": 1},
    "1 cheese pizza and 1 vegetarian pizza slices | 1 pointe de pizza au fromage et 1 pointe de pizza végétarienne": {"cheese": 1, "vegetarian": 1},
    "1 pepperoni pizza slice | 1 pointe de pizza au pepperoni": {"pepperoni": 1},
    "1 cheese pizza slice | 1 pointe de pizza au fromage": {"cheese": 1},
    "1 vegetarian pizza slice | 1 pointe de pizza végétarienne": {"vegetarian": 1}
}

tshirt_options = {
    "S | P",
    "M",
    "L | G",
    "XL | TG",
    "XXL | TTG",
    "No T-shirt | Aucun T-shirt"
}

gender_options = {
    "Female | Femelle",
    "Male | Mâle",
    "Other | Autre",
    "Prefer not to disclose | Préfère de ne pas divulger"
}

grade_level_options = {
    "1st grade | 1e année": 1,
    "2nd grade | 2e année": 2,
    "3rd grade | 3e année": 3,
    "4th grade | 4e année": 4,
    "5th grade | 5e année": 5,
    "6th grade | 6e année": 6,
    "7th grade | 7e année": 7,
    "8th grade | 8e année": 8,
    "9th grade | 9e année": 9,
    "10th grade | 10e année": 10,
    "11th grade | 11e année": 11,
    "12th grade | 12e année": 12,
}

############################################################################
# BEGIN PRE-PROCESSING EXCEL FILES
############################################################################

# enter the Excel file name of the teacher registration data
teacher_registration_filename = "ORC 2023 - Supervisor Registration (Responses).xlsx"

# enter the Excel file name of the team registration data
team_registration_filename = "ORC 2023 - Team Registration (Responses).xlsx"

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
temp_emails = list(teacher_registration["Email Address | Adresse courriel"])
temp_emails.extend(list(teacher_registration["Email of Supervisor #2 | Adresse courriel du superviseur #2"]))

teacher_emails = []

for email in temp_emails:
    if not pd.isna(email):
        teacher_emails.append(email.upper())

# stores the emails from the team registration form that don't match any of the emails from the teacher' form
error_emails = []

# iterate through the team registration form's emails and cross-reference them with the list of teacher emails
for email in team_registration["Primary Supervisor Email Address | Adresse courriel du(de la) superviseur(e) primaire"]:
    if email.upper() not in teacher_emails:
        error_emails.append(email)

# outputs any error emails, if our list of error emails is non-empty
if error_emails:
    print("The following emails from the Team Registration form did not have a matching teacher email from the Teacher Registration form:\n")
    for error_email in error_emails:
        print(error_email)
else:
    print("All emails from the Team Registration form matched with at least one teacher email from the Teacher Registration form.\n")

############################################################################
# END DATA VALIDATION
############################################################################

############################################################################
# BEGIN DICTIONARY SETUP
############################################################################

# this dictionary object will represent the root structure of the registration data
model = {}

# list of the school/community groups
groups = teacher_registration["School or Community Name | Nom de l'école ou du communauté"]

# 'temp school' will serve as a temporary store for all students without an associated supervisor/teacher
model['temp_school'] = {
    "emails": [],
    "members": {},
    "address": {},
    "num_students": 0,
    "rates": {
        "challenge_rate": 0,
        "lunch5_rate_quantity": 0,
        "lunch75_rate_quantity": 0,
        "girl_discount_quantity": 0,
        "highschool_discount_quantity": 0
    }
}

# iterate through each school/community group to initialize all related attributes
for i in range(0, len(groups), 1):
    """
    1. 'emails' attribute contains a list of all teacher/supervisor emails for that group
    2. 'members' attribute contains a separate object for each member of that group, with each object containing
    information about their lunch choice, t-shirt size, etc.
    3. 'address' attribute is the school/group's address
    4. 'rates' is a dictionary object that stores the number of students that apply to the different types of rates
    """
    model[groups[i]] = {
        "emails": [],
        "members": {},
        "address": {},
        "num_students": 0,
        "rates": {
            "challenge_rate": 0,
            "lunch5_rate_quantity": 0,
            "lunch75_rate_quantity": 0,
            "girl_discount_quantity": 0,
            "highschool_discount_quantity": 0
        }
    }

    # initialize the attributes to the related data from the registration data pandas DataFrame object
    model[groups[i]]["address"]["supervisor_name"] = teacher_registration.iloc[i]["Full Name | Nom complet"]
    model[groups[i]]["address"]["number"] = teacher_registration.iloc[i]["House Number / Unit Number | Numéro de maison / Numéro d'unité"]
    model[groups[i]]["address"]["street"] = teacher_registration.iloc[i]["Street Name and Type | Nom et type de rue"]
    model[groups[i]]["address"]["city"] = teacher_registration.iloc[i]["City | Ville"]
    model[groups[i]]["address"]["province"] = teacher_registration.iloc[i]["Province"]
    model[groups[i]]["address"]["postal_code"] = teacher_registration.iloc[i]["Postal Code | Code postal"]

    # only add an email if it is a non-empty cell field
    if not pd.isna(teacher_registration.iloc[i]["Email Address | Adresse courriel"]):
        model[groups[i]]["emails"].append(teacher_registration.iloc[i]["Email Address | Adresse courriel"].upper())

    # only add an email if it is a non-empty cell field
    if not pd.isna(teacher_registration.iloc[i]["Email of Supervisor #2 | Adresse courriel du superviseur #2"]):
        model[groups[i]]["emails"].append(teacher_registration.iloc[i]["Email of Supervisor #2 | Adresse courriel du superviseur #2"].upper())

    # add TEACHERS' lunch choices and shirt sizes to their dictionary object
    if not pd.isna(teacher_registration.iloc[i]["Full Name | Nom complet"]): # only create the teacher object if it has a valid name
        teacher_name = teacher_registration.iloc[i]["Full Name | Nom complet"]
        model[groups[i]]["members"][teacher_name] = {"lunch_choice": '', "shirt_size": '', "isStudent": False}

        # only add a lunch choice if it is a non-empty cell field
        if not pd.isna(teacher_registration.iloc[i]["Lunch Option for primary supervisor | Option pour le dîner pour le/la superviseur(e) primaire"]):
            model[groups[i]]["members"][teacher_name]["lunch_choice"] = teacher_registration.iloc[i]["Lunch Option for primary supervisor | Option pour le dîner pour le/la superviseur(e) primaire"]
        if not pd.isna(teacher_registration.iloc[i]["T-shirt Size for the primary supervisor | Grandeur de T-shirt pour le/la superviseur(e) primaire"]):
            model[groups[i]]["members"][teacher_name]["shirt_size"] = teacher_registration.iloc[i]["T-shirt Size for the primary supervisor | Grandeur de T-shirt pour le/la superviseur(e) primaire"]

        # determine which lunch rate to use for the given supervisor
        string = model[groups[i]]["members"][teacher_name]["lunch_choice"]
        if string:
            if string[0] == '2' or string.count('1') == 4: # either 2 slices of one type of pizza, or 1 slice of one type, and another 1 slice of another type (thus 2 total) ** we check== 4 because it's english and french, so the numbers are doubled
                model[groups[i]]["rates"]["lunch75_rate_quantity"] += 1
            elif string.count('1') == 2:
                model[groups[i]]["rates"]["lunch5_rate_quantity"] += 1

    # REFACTORING OPPORTUNITY --> the two blocks of supervisor initialization code started off small, but now they are larger
    if not pd.isna(teacher_registration.iloc[i]["Full Name of Supervisor #2 | Nom complet du superviseur #2"]):
        teacher_name = teacher_registration.iloc[i]["Full Name of Supervisor #2 | Nom complet du superviseur #2"]
        model[groups[i]]["members"][teacher_name] = {"lunch_choice": '', "shirt_size": '', "isStudent": False}

        if not pd.isna(teacher_registration.iloc[i]["Lunch Option for secondary supervisor | Option pour le dîner pour le/la superviseur(e) secondaire"]):
            model[groups[i]]["members"][teacher_name]["lunch_choice"] = teacher_registration.iloc[i]["Lunch Option for secondary supervisor | Option pour le dîner pour le/la superviseur(e) secondaire"]
        # only add a shirt size if it is a non-empty cell field
        if not pd.isna(teacher_registration.iloc[i]["T-shirt Size for the secondary supervisor | Grandeur de T-shirt pour le/la superviseur(e) secondaire"]):
            model[groups[i]]["members"][teacher_name]["shirt_size"] = teacher_registration.iloc[i]["T-shirt Size for the secondary supervisor | Grandeur de T-shirt pour le/la superviseur(e) secondaire"]

        # determine which lunch rate to use for the given student
        string = model[groups[i]]["members"][teacher_name]["lunch_choice"]
        if string:
            if string[0] == '2' or string.count('1') == 4:
                model[groups[i]]["rates"]["lunch75_rate_quantity"] += 1
            elif string.count('1') == 2:
                model[groups[i]]["rates"]["lunch5_rate_quantity"] += 1

# aggregate the information for every student in each group
for row in range(0, len(team_registration), 1):
    # determine which school/community grop this member is associated with
    primary_email = team_registration.iloc[row]["Primary Supervisor Email Address | Adresse courriel du(de la) superviseur(e) primaire"].upper()
    group_school = ""

    # REFACTORING OPPORTUNITY --> refactor the structure of our model so that we associate this information earlier on
    for group in model:
        if primary_email in model[group]["emails"]:
            group_school = group
            break

    # if a matching supervisor email was NOT found, place all the students in an unnamed school group
    if group_school == "":
        group_school = 'temp_school'

    # determine the price we charge dependent on the number of challenges being done
    challenges = team_registration.iloc[row]["Challenges | Concours"]
    if len(challenges.split(",")) > 1:  # if there's more than one challenge that the team is participating in
        model[group]["rates"]["challenge_rate"] = 35
    else:
        model[group]["rates"]["challenge_rate"] = 30

    # add STUDENTS' attributes to their dictionary object
    for j in range(1, 9):

        # REFACTORING OPPORTUNITY --> refactor the Google Form itself so that we don't need to create
        # separate condition for students on row 2 of registration excel sheet
        if j == 2:
            if not pd.isna(team_registration.iloc[row]['Full Name of Student #2 | Nom complet de l\'élève #2']):
                student_name = team_registration.iloc[row]['Full Name of Student #2 | Nom complet de l\'élève #2']
                model[group_school]["members"][student_name] = {
                    "lunch_choice": '',
                    "shirt_size": '',
                    "team_name": '',
                    "gender": '', "grade": 0,
                    "isStudent": True
                }

                # initializing all non-null student registration data
                if not pd.isna(team_registration.iloc[row]['Lunch Option for student #2 | Option pour le dîner d\'élève #2']):
                    model[group_school]["members"][student_name]["lunch_choice"] = team_registration.iloc[row]['Lunch Option for student #2 | Option pour le dîner d\'élève #2']

                if not pd.isna(team_registration.iloc[row]['T-shirt Size for student #2 | Grandeur de T-shirt pour l\'élève #2']):
                    model[group_school]["members"][student_name]["shirt_size"] = team_registration.iloc[row]['T-shirt Size for student #2 | Grandeur de T-shirt pour l\'élève #2']

                if not pd.isna(team_registration.iloc[row]['Team Name | Nom d\'équipe']):
                    model[group_school]["members"][student_name]["team_name"] = team_registration.iloc[row]['Team Name | Nom d\'équipe']

                if not pd.isna(team_registration.iloc[row]['Grade Level of Student #2 | Année scolaire d\'élève #2']):
                    grade = grade_level_options[team_registration.iloc[row]['Grade Level of Student #2 | Année scolaire d\'élève #2']]
                    model[group_school]["members"][student_name]["grade"] = grade

                if not pd.isna(team_registration.iloc[row]['Gender of Student #2 | Sexe de d\'élève #2']):
                    model[group_school]["members"][student_name]["gender"] = team_registration.iloc[row]['Gender of Student #2 | Sexe de d\'élève #2']

                # determine which lunch rate to use for the given student
                # REFACTORING OPPORTUNITY --> Can abstract below functionality into a function due to repeated usage
                string = model[group_school]["members"][student_name]["lunch_choice"]
                if string:
                    if string[0] == '2' or string.count('1') == 4:
                        model[group_school]["rates"]["lunch75_rate_quantity"] += 1
                    elif string.count('1') == 2:
                        model[group_school]["rates"]["lunch5_rate_quantity"] += 1

                # determine if female participant discount applies
                if model[group_school]["members"][student_name]["gender"] == "Female | Femelle":
                    model[group_school]["rates"]["girl_discount_quantity"] += 1

                # determine if highschool discount applies
                if model[group_school]["members"][student_name]["grade"] >= 9 and model[group_school]["members"][student_name]["grade"] <= 12:
                    model[group_school]["rates"]["highschool_discount_quantity"] += 1

                # increment number of students in this group
                model[group_school]["num_students"] += 1

        else:
            if not pd.isna(team_registration.iloc[row][f'Full Name of Student #{j} | Nom complet d\'élève #{j}']):
                student_name = team_registration.iloc[row][f'Full Name of Student #{j} | Nom complet d\'élève #{j}']
                model[group_school]["members"][student_name] = {
                    "lunch_choice": '',
                    "shirt_size": '',
                    "team_name": '',
                    "gender": '',
                    "grade":0,
                    "isStudent": True
                }

                if not pd.isna(team_registration.iloc[row][f'Lunch Option for student #{j} | Option pour le dîner d\'élève #{j}']):
                    model[group_school]["members"][student_name]["lunch_choice"] = team_registration.iloc[row][f'Lunch Option for student #{j} | Option pour le dîner d\'élève #{j}']

                if not pd.isna(team_registration.iloc[row][f'T-shirt Size for student #{j} | Grandeur de T-shirt pour l\'élève #{j}']):
                    model[group_school]["members"][student_name]["shirt_size"] = team_registration.iloc[row][f'T-shirt Size for student #{j} | Grandeur de T-shirt pour l\'élève #{j}']

                if not pd.isna(team_registration.iloc[row]['Team Name | Nom d\'équipe']):
                    model[group_school]["members"][student_name]["team_name"] = team_registration.iloc[row]['Team Name | Nom d\'équipe']

                if not pd.isna(team_registration.iloc[row][f'Grade Level of Student #{j} | Année scolaire d\'élève #{j}']):
                    grade = grade_level_options[team_registration.iloc[row][f'Grade Level of Student #{j} | Année scolaire d\'élève #{j}']]
                    model[group_school]["members"][student_name]["grade"] = grade

                if not pd.isna(team_registration.iloc[row][f'Gender of Student #{j} | Sexe de d\'élève #{j}']):
                    model[group_school]["members"][student_name]["gender"] = team_registration.iloc[row][f'Gender of Student #{j} | Sexe de d\'élève #{j}']

                # determine which lunch rate to use for the given student
                string = model[group_school]["members"][student_name]["lunch_choice"]
                if string:
                    if string[0] == '2' or string.count('1') == 4:
                        model[group_school]["rates"]["lunch75_rate_quantity"] += 1
                    elif string.count('1') == 2:
                        model[group_school]["rates"]["lunch5_rate_quantity"] += 1

                # determine if female participant discount applies
                if model[group_school]["members"][student_name]["gender"] == "Female | Femelle":
                    model[group_school]["rates"]["girl_discount_quantity"] += 1

                # determine if highschool discount applies
                if model[group_school]["members"][student_name]["grade"] >= 9 and model[group_school]["members"][student_name]["grade"] <= 12:
                    model[group_school]["rates"]["highschool_discount_quantity"] += 1

                # increment number of students in this group
                model[group_school]["num_students"] += 1

print('Here is the completed model representation of the registration data:')
print(model)
############################################################################
# END DICTIONARY SETUP
############################################################################

############################################################################
# FUNCTIONS FOR LIST CREATIONS
############################################################################

# total sum of slices by type of pizza, divided by 8, and rounded up by 1 (FORMATTED FOR EASY CONVERSION TO CSV)
def create_general_pizza_list():
    result = [
        ['pepperoni', 'cheese', 'vegetarian', 'pepperoni_pizzas', 'cheese_pizzas', 'vegetarian_pizzas'],
        [0, 0, 0, 0, 0, 0]
    ]

    for g in model:
        for member in model[g]["members"]:
            lunch_choice = model[g]["members"][member]["lunch_choice"]
            if lunch_choice in lunch_options:
                for pizza_type in lunch_options[lunch_choice]:
                    index = result[0].index(pizza_type)
                    result[1][index] += lunch_options[lunch_choice][pizza_type]

    for i in range(3, 6, 1):
        result[1][i] = math.ceil(result[1][i-3] / 8)

    return result

# sum of slices by type of pizza, organized by school/community group (FORMATTED FOR EASY CONVERSION TO CSV)
def create_pizza_list_by_school1():
    result = [['school', 'pepperoni', 'cheese', 'vegetarian']]

    for g in model:
        tempArr = [g, 0, 0, 0]

        for member in model[g]["members"]:
            lunch_choice = model[g]["members"][member]["lunch_choice"]
            if lunch_choice in lunch_options:
                for pizza_type in lunch_options[lunch_choice]:
                    index = result[0].index(pizza_type)
                    tempArr[index] += lunch_options[lunch_choice][pizza_type]

        result.append(tempArr)

    for i in range(1, len(result), 1):
        for j in range(1, len(result[i]), 1):
            if result[i][j] == 8:
                str_result = '1 Pizza'
            elif result[i][j] > 8:
                str_result = f"{result[i][j] // 8} Pizza and " + str(result[i][j] % 8) + ' slice(s)'
            else:
                str_result = str(result[i][j]) + ' Slice(s)'

            result[i][j] = str_result
    return result

# student and supervisors' pizza orders, listed out separately and organized by school/community group (FORMATTED FOR EASY CONVERSION TO CSV)
def create_pizza_list_by_school2():
    result = [['School', 'Person', 'Order']]

    for g in model:
        for member in model[g]["members"]:
            lunch_order = model[g]["members"][member]["lunch_choice"]
            result.append([g, member, lunch_order])

    return result

# total number of t-shirts by size, keeping supervisor and student t-shirts SEPARATE (FORMATTED FOR EASY CONVERSION TO CSV)
def create_general_tshirt_list():
    # FIRST sub-list within result => STUDENTS
    # SECOND sub-list within result => SUPERVISORS
    result = [[['S | P', 'M', 'L | G', 'XL | TG', 'XXL | TTG', 'No T-shirt | Aucun T-shirt'], [0, 0, 0, 0, 0, 0]], [['S | P', 'M', 'L | G', 'XL | TG', 'XXL | TTG', 'No T-shirt | Aucun T-shirt'], [0, 0, 0, 0, 0, 0]]]

    for g in model:
        for member in model[g]["members"]:
            shirt_size = model[g]["members"][member]['shirt_size']
            if not shirt_size == '':
                index = result[0][0].index(shirt_size)
                if model[g]["members"][member]['isStudent']:
                    result[0][1][index] += 1
                else:
                    result[1][1][index] += 1

    return result

# listing of each student's t-shirt size, by TEAM (FORMATTED FOR EASY CONVERSION TO CSV)
def create_tshirt_list_by_team1():
    result = [['Team', 'Student', 'Size']]

    for g in model:
        for member in model[g]["members"]:
            if model[g]["members"][member]['isStudent']:
                team = model[g]["members"][member]['team_name']
                shirt_size = model[g]["members"][member]['shirt_size']
                result.append([team, member, shirt_size])
    return result

# total number of t-shirts by size for each team (FORMATTED FOR EASY CONVERSION TO CSV)
def create_tshirt_list_by_team2():
    teams = team_registration["Team Name | Nom d'équipe"]
    result = [['Team', 'S | P', 'M', 'L | G', 'XL | TG', 'XXL | TTG', 'No T-shirt | Aucun T-shirt']]
    for team in teams:
        result.append([team, 0, 0, 0, 0, 0, 0])

    for g in model:
        for member in model[g]["members"]:
            if model[g]["members"][member]['isStudent']:
                shirt_size = model[g]["members"][member]['shirt_size']
                if not shirt_size == '':
                    team = model[g]["members"][member]['team_name']
                    shirt_index = result[0].index(shirt_size)
                    team_index = find_team_index(team, result)
                    if not team_index == -1:
                        result[team_index][shirt_index] += 1

    return result

# helper function for create_tshirt_list_by_team2()
def find_team_index(team, list):
    for i in range(1, len(list), 1):
        if list[i][0] == team:
            return i
    return -1

# team supervisors, with their name and school/community group (FORMATTED FOR EASY CONVERSION TO CSV)
def create_supervisor_certificates_list():
    result = [['Name', 'School Name']]

    for g in model:
        for member in model[g]["members"]:
            if not model[g]["members"][member]['isStudent']:
                result.append([member, g])
    return result

# students, with their name, school/community group, and team name (FORMATTED FOR EASY CONVERSION TO CSV)
def create_student_certificates_list():
    result = [['Name', 'School Name', 'Team Name']]

    for g in model:
        for member in model[g]["members"]:
            if model[g]["members"][member]['isStudent']:
                team = model[g]["members"][member]["team_name"]
                result.append([member, g, team])
    return result

# helper function that will create a .csv file with the provided data
def create_csv(filename, arg):
    with open(filename, 'w', encoding='UTF8', newline='') as file:
        writer = csv.writer(file)
        writer.writerows(arg)
        file.close()

def generate_invoice(rates, address, n, invoice_num):
    # Make a POST request to the pdf generator API and receive the response
    if not address:
        return

    url = "https://invoice-generator.com"
    # Get current date
    current_date = datetime.date.today()
    # Format current date as DD/MM/YYYY
    formatted_date = current_date.strftime("%m/%d/%Y")

    data = {"from": "Kelly Xu\nChair, IEEE Ottawa Robotics Competition\n74 Renova Private\nOttawa, ON\nK1G 4C6".upper(),
            "to": f"{address['supervisor_name']}\n{address['number']} {address['street']}\n{address['city']}, {address['province']}\n{address['postal_code']}".upper(),
            "logo": "https://i.ibb.co/km9mGnm/Picture1.png", "date": f"{formatted_date}",
            "due_date": "Payable upon receipt", "number": f"ORC2023-23{invoice_num}",
            "items": [{"name": "Students participated", "quantity": n, "unit_cost": rates["challenge_rate"]},
                      {"name": "Lunch - pizza slices", "quantity": (2 * rates["lunch75_rate_quantity"]) + rates["lunch5_rate_quantity"],  # need to double lunch75 rate since this was for two slices of pizza
                       "unit_cost": 2}],
            "notes": "Payment can be made by credit card, cheque, or cash. For credit card, please email orcinfo@ieeeottawa.ca to make arrangements. For cheques, please address it to \"IEEE Ottawa Section\" and send to the address in the top left of invoice. For cash, please bring exact change on competition day, as any amount remitted over the invoiced total above will be considered a gratuity."}

    response = requests.post(url, json=data)

    # Extract the binary content of the response
    pdf_content = response.content

    # Save the binary content as a PDF file, using the given invoice number argument
    with open(f"./billings/ORC2023-23{invoice_num}.pdf", "wb") as f:
        f.write(pdf_content)

def generate_all_invoices():
    global invoice_number # will reference the global invoice number variable
    global invoice_number_str

    for g in model:
        generate_invoice(model[g]["rates"], model[g]["address"], model[g]["num_students"], invoice_number_str)
        invoice_number += 1
        invoice_number_str = "{:03d}".format(invoice_number)

############################################################################
# END FUNCTIONS FOR LIST CREATIONS
############################################################################

############################################################################
# BEGIN EXCEL FILES GENERATION
############################################################################

# path to the directory that will store all the generated lists
path = './TEST_generated_lists/'

# create the above directory if it does not already exist
if not os.path.exists(path):
    os.mkdir(path)

file_names = ['./TEST_generated_lists/pizza_orders_general.csv', './TEST_generated_lists/pizza_orders_by_school.csv',
              './TEST_generated_lists/pizza_orders_by_individual.csv', './TEST_generated_lists/shirt_orders_by_supervisor.csv',
              './TEST_generated_lists/shirt_orders_by_student.csv', './TEST_generated_lists/shirt_orders_by_team_individual.csv',
              './TEST_generated_lists/shirt_orders_by_team.csv', './TEST_generated_lists/certificates_list_supervisors.csv',
              './TEST_generated_lists/certificates_list_students.csv']

lists = [create_general_pizza_list(), create_pizza_list_by_school1(), create_pizza_list_by_school2(),
         create_general_tshirt_list()[1], create_general_tshirt_list()[0], create_tshirt_list_by_team1(),
         create_tshirt_list_by_team2(), create_supervisor_certificates_list(), create_student_certificates_list()]

for i in range(0, len(file_names), 1):
    create_csv(file_names[i], lists[i])

generate_all_invoices()

############################################################################
# END EXCEL FILES GENERATION
############################################################################
