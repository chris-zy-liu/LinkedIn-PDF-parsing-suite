import os
import re
import xlwt
import time

start_time = time.time()

directory = "CMO source files Equilar 2023/Traceys CFOs/LI"

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
    for line in profile:
        if 'height="54"' in line:
            fullname = re.findall(">.*<", line)[0][1:-1]
            names = re.sub('(.*)', "", fullname)
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
            new_company.positions.append(Position(get_content(line)))
            continue
        if height == 22 and pending_dates:
            raw = get_content(line)
            raw = raw.split(' (')
            start_date = raw[0].split('\xa0-\xa0')[0]
            end_date = raw[0].split('\xa0-\xa0')[1]
            duration = raw[1][:-1]
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
                for line in profile:
                    if 'font="0"' not in line and len(re.findall("Page \\d+ of \\d+<", line)) == 0 and "</text>" in line:
                        output.write(re.sub("&lt;", "<", re.sub("&gt;", ">", re.sub("&amp;", "&", line))))
                output.seek(0)
                name = parse_profile_name(output.readlines())
                person = Profile(name[0], name[1])
                profiles.append(person)
                output.seek(0)
                parse_profile_experiences(output.readlines(), person)
                output.seek(0)
                parse_profile_education(output.readlines(), person)

# for profile in profiles:
#     print("\n\n\n")
#     print(profile.first_name, profile.last_name)
#     for company in profile.companies:
#         print(company.name)
#         for position in company.positions:
#             print("\t", position.title, position.dates, position.duration)
#             print('\t', position.description)
#     for education in profile.educations:
#         print(education.school, education.dates, education.degree)


# format to an excel document
textstyle = xlwt.easyxf('font: name Calibri')
headerstyle  = xlwt.easyxf('font: name Calibri, bold 1')

wb = xlwt.Workbook()
experience_sheet = wb.add_sheet('Experience Data')
education_sheet = wb.add_sheet('Education Data')


experience_sheet.write(0, 0, "First Name & MI", headerstyle)
experience_sheet.write(0, 1, "Last Name", headerstyle)

education_sheet.write(0, 0, "First Name & MI", headerstyle)
education_sheet.write(0, 1, "Last Name", headerstyle)

for i in range(1, 51):
    column = 2 + (i-1) * 5
    for header in ["Company ", "Title ", "Dates ", "Duration ", "Description "]:
        experience_sheet.write(0, column, header+str(i), headerstyle)
        column += 1

for i in range(1, 51):
    column = 2 + (i-1) * 3
    for header in ["School ", "Degree ", "Dates "]:
        education_sheet.write(0, column, header+str(i), headerstyle)
        column += 1

for index, profile in enumerate(profiles):
    
    experience_sheet.write(index+1, 0, profile.first_name, textstyle)
    experience_sheet.write(index+1, 1, profile.last_name, textstyle)

    column = 2
    
    for company in profile.companies:
        for position in company.positions:
            experience_sheet.write(index+1, column, company.name, textstyle)
            column += 1
            experience_sheet.write(index+1, column, position.title, textstyle)
            column += 1
            experience_sheet.write(index+1, column, position.dates, textstyle)
            column += 1
            experience_sheet.write(index+1, column, position.duration, textstyle)
            column += 1
            experience_sheet.write(index+1, column, position.description, textstyle)
            column += 1

for index, profile in enumerate(profiles):
    
    education_sheet.write(index+1, 0, profile.first_name, textstyle)
    education_sheet.write(index+1, 1, profile.last_name, textstyle)

    column = 2
    
    for education in profile.educations:
        education_sheet.write(index+1, column, education.school, textstyle)
        column += 1
        education_sheet.write(index+1, column, education.degree, textstyle)
        column += 1
        education_sheet.write(index+1, column, education.dates, textstyle)
        column += 1

wb.save('CMO Complete Data.xls')

print("--- %s seconds ---" % (time.time() - start_time))