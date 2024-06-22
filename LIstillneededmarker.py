import os
import re

directory = "LinkedIn Project/All CFO LI/"

names = open(f'{directory}/docs/CFO Lists.txt', 'r')
outputnames = open(f'{directory}/docs/stillneeded.txt', 'w+')
alreadyhave = open(f'{directory}/docs/alreadyhave.txt', 'w+')
notclaimed = open(f'{directory}/docs/notclaimed.txt', 'w+')
nocomma = open(f'{directory}/docs/nocomma.txt', 'w+')

still_needed = []

regex = re.compile('[^a-zA-Z,\' ]')

for name in names:
    name = name.strip().title()
    cont = False
    if not name:
        continue
    for filename in os.scandir(directory):
        if filename.is_file():
            processed_filename = filename.name.title()
            processed_name = regex.sub("", name)

            # os.rename(directory+filename.name, directory+filename.name[:filename.name.find(".pdf")+4])
            # os.rename(directory+filename.name, directory+re.sub(" CLAIMED.pdf", ".pdf", filename.name))
            # print(re.sub(" CLAIMED.pdf", ".pdf", filename.name))
            if name.split(", ")[0].strip() in processed_filename.strip() and max(name.split(", ")[1].split(), key=len).strip() in processed_filename.strip():
                # print(max(name.split(", ")[1].split(), key=len).strip())
                if "Millner" in filename.name:
                    print(name, filename.name)
                print(name.strip(), file=alreadyhave)
                cont = True
                if "CLAIMED.pdf" not in filename.name:
                    os.rename(directory+filename.name, directory+filename.name[:-4]+" CLAIMED.pdf")
                break
                # if "CLAIMED"  in processed_filename:
                #     os.rename(directory+filename.name, directory+re.sub("CLAIMED", "", filename.name))

            # if "," not in processed_filename:
            #     print(processed_filename)
    if cont:
        continue
    still_needed.append(name)

for name in still_needed:
    print(name.strip(), file=outputnames)

for filename in os.scandir(directory):
    if filename.is_file():
        if "CLAIMED.pdf" not in filename.name:
            print(filename.name, file=notclaimed)

        if "," not in filename.name:
            print(filename.name, file=nocomma)
            newfile = filename.name
            if " LI CLAIMED.pdf" in newfile:
                name = newfile[:newfile.find(" LI CLAIMED.pdf")]
                name = name.split()
                if len(name) == 2:
                    newname = f"{name[1]}, {name[0]} LI CLAIMED.pdf"
                if len(name) == 3:
                    newname = f"{name[2]}, {name[0]} {name[1]} LI CLAIMED.pdf"
                os.rename(directory+filename.name, directory+newname)
            
            elif " LI.pdf" in newfile:
                name = newfile[:newfile.find(" LI.pdf")]
                name = name.split()
                print(name)
                if len(name) == 2:
                    newname = f"{name[1]}, {name[0]} LI.pdf"
                if len(name) == 3:
                    newname = f"{name[2]}, {name[0]} {name[1]} LI.pdf"
                os.rename(directory+filename.name, directory+newname)

            elif " CLAIMED.pdf" in newfile:
                name = newfile[:newfile.find(" CLAIMED.pdf")]
                name = name.split()
                print(name)
                if len(name) == 2:
                    newname = f"{name[1]}, {name[0]} CLAIMED.pdf"
                if len(name) == 3:
                    newname = f"{name[2]}, {name[0]} {name[1]} CLAIMED.pdf"
                print(name)
                os.rename(directory+filename.name, directory+newname)

            elif ".pdf" in newfile:
                name = newfile[:newfile.find(".pdf")]
                name = name.split()
                if len(name) == 2:
                    newname = f"{name[1]}, {name[0]}.pdf"
                if len(name) == 3:
                    newname = f"{name[2]}, {name[0]} {name[1]}.pdf"
                os.rename(directory+filename.name, directory+newname)

names.close()
outputnames.close()
alreadyhave.close()