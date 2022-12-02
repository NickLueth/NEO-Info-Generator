# New Hire Info Grabber
# Created by: Nick Lueth, 8/3/2022
# Last updated: 12/2/2022
import string
from random import randint
import pandas as pd


class Hire:
    def __init__(self, fname, lname, ttl, dpmt, mngr):
        self.first_name = fname
        self.last_name = lname
        self.title = ttl
        self.department = dpmt
        self.manager = mngr
        invalid_characters = ["'", " ", "."]
        temp_last_name = ""
        for character in self.last_name.lower():
            if character not in invalid_characters:
                temp_last_name += character
        self.email = fname[0].lower() + temp_last_name + "@walkercares.org"
        self.username = fname[0].lower() + temp_last_name

        self.password = self.generate_password()
        self.evolv = "Elvis1222"

    @staticmethod
    def generate_password():
        characters = list(string.ascii_letters + string.digits + string.punctuation)
        invalid_characters = ["`", "\"", "'", ":", ";", ","]
        valid_characters = []
        password = ""
        for char in characters:
            if char not in invalid_characters:
                valid_characters.append(char)
        for i in range(0, 12):
            password += valid_characters[randint(0, len(valid_characters)-1)]
        return password

    def __str__(self):
        return f"""{"First name:" : <11} {self.first_name}
{"Last name:" : <11} {self.last_name}
{"Title:" : <11} {self.title}
{"Department:" : <11} {self.department}
{"Manager:" : <11} {self.manager}

{"Email:" : <11} {self.email}
{"Username:" : <11} {self.username}
{"Password:" : <11} {self.password}
{"Evolv:" : <11} {self.evolv}
-------------------------------------------------
"""


def add_hires():
    all_hires = []
    ex_data = pd.read_excel('Example.xlsx') # Change this file to the path of the Excel file that you want to read from
    first_names = ex_data['First Name'].values.tolist()
    last_names = ex_data['Last Name'].values.tolist()
    titles = ex_data['Position'].values.tolist()
    departments = ex_data['Program'].values.tolist()
    managers = ex_data['Supervisor'].values.tolist()
    for i in range(len(first_names)):
        all_hires.append(Hire(first_names[i], last_names[i], titles[i], departments[i], managers[i]))
        print(f"{first_names[i]} {last_names[i]} the new {titles[i]} has been added!")
    return all_hires


def print_hires(hires_list):
    file = open("NewHireInfo\Example.txt", "w+") # Change this file to the path where you want to save the data and what you want to name it
    for hire in hires_list:
        file.write(hire.__str__())


if __name__ == '__main__':
    hires = add_hires()
    print_hires(hires)
