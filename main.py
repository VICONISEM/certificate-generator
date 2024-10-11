import cv2
from PIL import ImageFont, ImageDraw, Image
import numpy as np
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog

def select_file(name_of_window):
    filename = filedialog.askopenfilename(title=f"{name_of_window}")
    return filename

def select_folder():
    foldername = filedialog.askdirectory(title="Save location")
    return foldername

def generate_certificate(names, teams, font_of_name, image):
    save_location = select_folder()   
    os.chdir(rf"{save_location}")
    
    for n, t in zip(names, teams):
        team_folder = os.path.join(save_location, t)
        if not os.path.exists(team_folder):
            os.makedirs(team_folder)

        img_pil = Image.fromarray(image)
        draw = ImageDraw.Draw(img_pil)
        
        bbox = draw.textbbox((0, 0), n, font=font_of_name)
        w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
        
        draw.text(((img_pil.width - w) // 2, 525), n, font=font_of_name, fill=(255, 255, 255))
        
        image_new = np.array(img_pil)
        cv2.imwrite(os.path.join(team_folder, f"certificate_for_{n}.png"), image_new)

certificate_template_path = select_file("Select certificate template")
image = cv2.imread(rf"{certificate_template_path}")

Excel_sheet_path = select_file("Select Excel sheet")
sheet = pd.read_excel(rf"{Excel_sheet_path}")
names = list(sheet["NameOfMember"])
teams = list(sheet["NameOfTeam"])

font_path = select_file("Select Font")
font_of_name = ImageFont.truetype(rf"{font_path}", 80)

generate_certificate(names, teams, font_of_name, image)
