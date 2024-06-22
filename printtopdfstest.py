import os
import re
import xlwt

directory = "printToPDFsTest"

profiles = []

class Position:
    def __init__(self, position_title, company):
        self.title = position_title
        self.company = company
        self.dates = ""
        self.duration = ""
        self.description = ""

class Education:
    def __init__(self, school):
        self.school = school
        self.dates = ""
        self.degree = ""

class Profile:
    def __init__(self, first_name, last_name):
        self.first_name = first_name
        self.last_name = last_name
        self.positions = []
        self.educations = []

def get_text_height(line):
    return re.findall('height="\\d+"', line)[0][8:-1]

def get_text_font(line):
    return re.findall('font="\\d+"', line)[0][6:-1]

def get_content(line):
    return re.findall(">.+<", line)[0][1:-1]

def is_date(line):
    if re.findall("\\d+ (yrs?|mos?)", line):
        if " - " in line:
            return "position"
        return "company"
    return False

def parse_experience(lines, begin, person):
    lines = lines[begin+1:]
    nextlineisdate = False
    nextlineiscompany = False
    for linenum, line in enumerate(lines):
        ogline = line
        line = re.sub("<a.*\">", "", line)
        line = re.sub("</a>", "", line)

        height = get_text_height(line)
        content = get_content(line)

        try:
            nextline = lines[linenum+1]
            nextcontent = get_content(nextline)
            nextheight = get_text_height(nextline)
        except:
            pass

        if height == "23" and is_date(nextline) == "company":
            nextlineiscompany = False
            company = content
            continue
        if height == "23" and is_date(nextline) == "position":
            nextlineisdate = True
            newjob = Position(content, company)
            person.positions.append(newjob)
            continue
        if height == "23" and not is_date(nextcontent):
            # print(is_date("2006 - Oct 2008 · 2 yrs 10 mos"), content, nextcontent, linenum)
            nextlineisdate = False
            nextlineiscompany = True
            position = content
            continue
        if nextlineiscompany:
            nextlineiscompany = False
            nextlineisdate = True
            newjob = Position(position, content.split(" · ")[0])
            person.positions.append(newjob)
            continue
        if nextlineisdate:
            nextlineisdate = False
            # print(content)
            person.positions[-1].dates = re.sub("\xa0", " ", content.split(" · ")[0])
            person.positions[-1].duration = content.split(" · ")[1]
            continue
        if person.positions and not is_date(content):
            person.positions[-1].description += content+" "

def parse_education(lines, begin, person):
    lines = lines[begin+1:]
    pending_dates = False
    pending_major = False
    for linenum, line in enumerate(lines):
        ogline = line
        line = re.sub("<a.*\">", "", line)
        line = re.sub("</a>", "", line)
        height = get_text_height(line)
        content = get_content(line)
        if height == "23":
            newschool = Education(content)
            person.educations.append(newschool)
            pending_dates = True
            pending_major = True
            continue
        if pending_major and "href=" in ogline:
            if re.sub("([Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec])?.*\\d\\d\\d\\d", "", content) not in ["", " - "]:
                pending_major == False
                person.educations[-1].degree = content
            else:
                pending_major = False
                pending_dates = False
                person.educations[-1].dates = content
            continue
        if pending_dates and "href=" in ogline:
            
            if re.sub("([Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec])?.*\\d\\d\\d\\d", "", content) not in ["", " - "]:
                pending_dates == False
                person.educations[-1].dates = content
            continue
        



def init_reading(lines):
    regex = re.compile('[^a-zA-Z\' ]')
    line = lines[0]
    fullname = re.findall(">.*<", line)[0][1:-1]
    names = re.sub('(.*)', "", fullname)
    names = regex.sub('', fullname).strip().split()
    if len(names) == 3:
        fullname = [names[0] + " " + names[1], names[2]]
    elif len(names[0]) == 1:
        fullname = [names[1], names[0]]
    else:
        fullname = [names[0], names[1]]
    for linenum, line in enumerate(lines):
        if get_text_height(line) == "47":
            doctype = get_content(line)
            startline = linenum
            print(doctype, startline)
            break
    for profile in profiles:
        if profile.first_name == fullname[0] and profile.first_name == fullname[0]:
            if doctype == "Experience":
                parse_experience(lines, startline, profile)
                return
            if doctype == "Education":
                parse_education(lines, startline, profile)
                return
    profile = Profile(fullname[0], fullname[1])
    if doctype == "Experience":
        parse_experience(lines, startline, profile)
    if doctype == "Education":
        parse_education(lines, startline, profile)
    profiles.append(profile)

for filename in os.scandir(directory+"/xmls"):
    if filename.is_file():
        print("LOADING: ", filename.name)
        with open(f"{directory}/xmls/{filename.name}", 'r') as profile:
            with open(f"{directory}/processedXMLs/{filename.name}", 'w+') as output:

                lines = []
                page_count = 0

                for line in profile:
                    if "<page number=\"" in line:
                        page_count += 1
                        page_height = int(re.sub("\\D", "", line.split()[5]))
                    if "</a>" in line:
                        if "www.linkedin.com" not in line:
                            continue
                        if "media-viewer?" in line:
                            continue
                        if re.findall("Show all \\d+ media", line):
                            continue
                    
                    if len(re.findall("Page \\d+ of \\d+<", line)) == 0 and "</text>" in line:
                        new_line = re.sub("&lt;", "<", re.sub("&gt;", ">", re.sub("&amp;", "&", line)))
                        new_line2 = re.sub("<a.*\">", "", new_line)
                        new_line2 = re.sub("</a>", "", new_line2)
                        if get_content(new_line2) in ["Full-time", "Part-time", "Self-employed", "Freelance", "Contract", "Internship", "Apprenticeship", "Seasonal"]:
                            print(line)
                            continue

                        new_line = re.sub("top=\"\\d+\"", f"top=\"{str(int(re.findall('top="\\d+"', new_line)[0][5:-1])+page_height*(page_count-1))}\"", new_line)
                        lines.append(new_line)

                lines.sort(key = lambda x: int(re.sub("\\D", "", x.split()[1])))

                for linenum, line in enumerate(lines):
                    if get_text_height(line) == "23":
                        if get_text_height(lines[linenum+1]) == "23":
                            lines[linenum] = re.sub(">.*<", f">{get_content(line)} {get_content(lines[linenum+1])}<", line)
                            del lines[linenum+1]

                output.writelines(lines)
                init_reading(lines)

# print(profiles)
# for person in profiles:
#     print("*************", person.first_name, person.last_name)
#     for position in person.positions:
#         print(position.title, "at", position.company)
#         print(position.dates, "duration is", position.duration)
#         print(position.description)
#         print("\n")
#     print("\n")
#     for education in person.educations:
#         print(education.degree, "at", education.school)
#         print(education.dates)
#         print("\n")
#     print("\n\n\n\n\n\n\n")




















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
    
    for position in profile.positions:
        experience_sheet.write(index+1, column, position.company, textstyle)
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

wb.save('sample formatting.xls')
