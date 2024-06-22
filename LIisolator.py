import os
import shutil
from pdf2image import convert_from_path
import numpy as np

directory = "LinkedIn Project/CEOs/Source Files CEO/"

num = 1

for filename in os.scandir(directory):
    if filename.is_file():
        print("LOADING: ", filename.name, num)
        if ".pdf" not in filename.name:
            num += 1
            continue
        images = convert_from_path(directory+filename.name)
        image = np.array(images[-1])
        pixel = image[10,10]
        print(pixel)
        if list(pixel) == [41, 62, 73]:
            shutil.move(directory+filename.name, directory+"LI/"+filename.name)

        num += 1

        # with open(directory+filename.name, 'r') as profile:
        #     for line in profile:
        #         if 'LinkedIn Corporation' in line:
        #             print("Detected")