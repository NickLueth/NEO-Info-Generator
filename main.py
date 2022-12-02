# New Hire Info Grabber
# Created by: Nick Lueth, 8/3/2022
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

        temp_last_name = self.last_name.lower().replace(" ", "")
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
    ex_data = pd.read_excel('NovemberOrientation.xlsx')
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
    file = open("NewHireInfo\Test.txt", "w+")
    for hire in hires_list:
        file.write(hire.__str__())


if __name__ == '__main__':
    hires = add_hires()
    print_hires(hires)
