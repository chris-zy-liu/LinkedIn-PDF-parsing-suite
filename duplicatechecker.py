import os
import re
import xlwt
import time

start_time = time.time()

directory = "LinkedIn Project/All CFO LI/"

profiles = []

class Position:
    def __init__(self, position_title):
        self.title = position_title
        self.dates = ""
        self.duration = ""
        self.description = ""

class Education:
    def __init__(self, school):
        self.school = school
        self.dates = ""
        self.degree = ""

class Company:
    def __init__(self, company_name):
        self.name = company_name
        self.positions = []
    def is_same_company(self, other):
        return re.sub("\\(.*?\\)", "", self.name).strip().upper() == re.sub("\\(.*?\\)", "", other).strip().upper()

class Profile:
    def __init__(self, first_name, last_name):
        self.first_name = first_name
        self.last_name = last_name
        self.filename = ""
        self.companies = []
        self.educations = []

def get_text_height(line):
    return re.findall('height="\\d+"', line)[0][8:-1]

def get_text_font(line):
    return re.findall('font="\\d+"', line)[0][6:-1]

def get_content(line):
    return re.findall(">.*<", line)[0][1:-1]

def parse_profile_name(profile):
    regex = re.compile('[^a-zA-Z\' ]')
    for linenumber, line in enumerate(profile):
        if get_text_height(line) == "54":
            if get_text_height(profile[linenumber+1]) == "54":
                profile[linenumber] = re.sub(">.*<", f">{get_content(line)} {get_content(profile[linenumber+1])}<", line)
                del profile[linenumber+1]

    for line in profile:
        if 'height="54"' in line:
            fullname = re.findall(">.*<", line)[0][1:-1]
            if "," in fullname and len(fullname.split()) > 2:
                fullname = fullname[:fullname.find(",")]
            elif "," in fullname:
                fullname = fullname.split(",")
                fullname = fullname[1].strip() + " " + fullname[0].strip()
            names = re.sub('(.*)', "", fullname)
            names = names[:names.find("◆")]
            names = re.sub('Jr.', "", fullname)
            names = re.sub('Sr.', "", fullname)
            names = re.sub('_', " ", fullname)
            names = regex.sub('', fullname).split()
            if len(names) == 3:
                return [names[0] + " " + names[1], names[2]]
            if len(names[0]) == 1:
                return[names[1], names[0]]
            return [names[0], names[1]]

def parse_profile_experiences(profile, person):
    begin_line = -1
    end_line = -1
    pending_dates = False
    text_font = -1
    for linenumber, line in enumerate(profile):
        if len(re.findall('height="33".*>Experience<', line)) != 0:
            begin_line = linenumber
        if len(re.findall('height="33".*>Education<', line)) != 0:
            end_line = linenumber
    if begin_line == -1:
        return
    for linenumber, line in enumerate(profile[begin_line+1:end_line]):
        height = int(get_text_height(line))
        font = int(get_text_font(line))

        if height == 25:
            text_font = -1
            new_company_name = get_content(line)
            for company in person.companies:
                if company.is_same_company(new_company_name):
                    new_company = company
                    break
            else:
                new_company = Company(new_company_name)
                person.companies.append(new_company)
        if height == 24:
            text_font = -1
            pending_dates = True
            try:
                new_company.positions.append(Position(get_content(line)))
            except UnboundLocalError:
                new_company = Company(get_content(line))
                new_company.positions.append(Position(get_content(line)))
            continue
        if height == 22 and pending_dates:
            raw = get_content(line)
            # print(raw)
            raw = raw.split(' (')
            if " (" in raw:
                start_date = raw[0].split('\xa0-\xa0')[0]
                end_date = raw[0].split('\xa0-\xa0')[1]
                duration = raw[1][:-1]
            else:
                if "-" in raw:
                    start_date = raw[0].split('\xa0-\xa0')[1]
                    end_date = raw[0].split('\xa0-\xa0')[1]
                    duration = "Unlisted"
                else:
                    start_date = "Unlisted"
                    end_date = "Unlisted"
                    duration = raw
            try:
                new_company.positions[-1].dates = start_date + " - " + end_date
                new_company.positions[-1].duration = duration
            except IndexError:
                new_company.positions.append(Position("Unlisted"))
                new_company.positions[-1].dates = start_date + " - " + end_date
                new_company.positions[-1].duration = duration

            pending_dates = False
            text_font = int(get_text_font(line))
            continue
        if height == 22 and text_font > 0:
            if int(get_text_font(line)) == text_font:
                new_company.positions[-1].description += (get_content(line).strip()+" ")

def parse_profile_education(profile, person):
    begin_line = -1
    end_line = -1
    pending_dates = False
    for linenumber, line in enumerate(profile):
        if len(re.findall('height="33".*>Education<', line)) != 0:
            begin_line = linenumber
            continue
        if len(re.findall('height="33".*>.*<', line)) != 0 and begin_line != -1:
            end_line = linenumber
            break
    else:
        end_line = len(profile)
    if begin_line == -1:
        return
    for linenumber, line in enumerate(profile[begin_line+1:end_line]):
        height = int(get_text_height(line))
        font = int(get_text_font(line))
        if height == 25:
            new_education = Education(get_content(line))
            person.educations.append(new_education)
            continue
        if height == 22:
            if " · " in line:
                degree = get_content(line).split(" · ")[0]
                dates = get_content(line).split(" · ")[1][1:-1]
                person.educations[-1].degree = degree
                person.educations[-1].dates = dates
                continue
            else:
                degree = get_content(line)
                person.educations[-1].degree = degree




for filename in os.scandir(directory+"/xmls"):
    if filename.is_file():
        print("LOADING: ", filename.name)
        with open(f"{directory}/xmls/{filename.name}", 'r') as profile:
            with open(f"{directory}/processedXMLs/{filename.name}", 'w+') as output:
                for linenum, line in enumerate(profile):
                    if 'font="0"' not in line and len(re.findall("Page \\d+ of \\d+<", line)) == 0 and "</text>" in line:
                        output.write(re.sub("&lt;", "<", re.sub("&gt;", ">", re.sub("&amp;", "&", line))))
                output.seek(0)
                name = parse_profile_name(output.readlines())
                person = Profile(name[0], name[1])
                profiles.append(person)
                person.filename = filename.name
                output.seek(0)
                parse_profile_experiences(output.readlines(), person)
                output.seek(0)
                parse_profile_education(output.readlines(), person)

checks = []

print("\n\n\n\tDUPLICATES DETECTED FROM CURRENT XMLS (if any):\n\n\n")

for profile in profiles:
    for other in checks:
        if other.first_name + other.last_name == (profile.first_name + profile.last_name):
            print(other.filename, profile.filename)
            break
    else:
        checks.append(profile)