import pydicom
import cv2
import os
import imageio
import glob
import pptx
import tkinter
from tkinter import filedialog, messagebox

# Create a tkinter root window
root = tkinter.Tk()
root.withdraw()

# Initial message
messagebox.showinfo("DICOM to PPT", "DICOM to ppt. 2024. Agustín Fernández Cisnal. Valencia (Spain) fecia82@gmail.com")
messagebox.showinfo("Select a folder", "Select a folder containing the DICOM files. All angio files will be converted to mp4 and 9 key frames will be selected. A slide will be created for each angio")

# Open the "Open" dialog box and allow the user to select a folder
folder_path = filedialog.askdirectory(parent=root)

# Search for all files in the specified folder and its subfolders
dcm_files = [f for f in glob.glob(folder_path + "/**/*", recursive=True) if os.path.isfile(f)]

sorted_dcm_files = []   
for dcm_file in enumerate(dcm_files):
    try:
        dcm_data = pydicom.dcmread(dcm_file[1], force=True)
        sorted_dcm_files.append((dcm_data.SeriesNumber, dcm_file[1]))
    except AttributeError:
        #messagebox.showinfo("Error", "Cerrar archivo")
        continue

sorted_dcm_files.sort()
sorted_dcm_files = [x[1] for x in sorted_dcm_files]

num_diapo = 0
prs = pptx.Presentation()
dcm_files = sorted_dcm_files

# Loop through the found files
for i, dcm_file in enumerate(dcm_files):
    try:
        dcm_data = pydicom.dcmread(dcm_file, force=True)
    except PermissionError:
        #messagebox.showinfo("Error", "Cerrar archivo")
        continue
    if dcm_data.Modality != "XA":
        messagebox.showinfo("Info", f"{dcm_file} not XA (not angio)")
        continue
    if "PixelData" not in dcm_data:
        messagebox.showinfo("Info", f"{dcm_file} does not contain pixel datas")
        continue

    video_data = dcm_data.pixel_array
    file_name, file_ext = os.path.splitext(dcm_file)
    new_file_name = f"video{i+1}"
    new_file_path = f"{folder_path}/{new_file_name}.mp4"
    poster_path = f"{folder_path}/{new_file_name}.png"
    img_path = f"{folder_path}/{new_file_name}_m.png"

    imageio.mimwrite(new_file_path, video_data)
    cv2.imwrite(poster_path, video_data[len(video_data)//2])

    for j in range(9):
        image = video_data[j * len(video_data) // 9]
        cv2.imwrite(f"{folder_path}/image_{j+1}.png", image)

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_movie(new_file_path, pptx.util.Cm(13), pptx.util.Cm(3), width=pptx.util.Cm(11), height=pptx.util.Cm(11), mime_type='video/mp4', poster_frame_image=poster_path)

    for j in range(3):
        for k in range(3):
            x = pptx.util.Cm(0.5 + k * 4)
            y = pptx.util.Cm(3 + j * 4)
            picture = slide.shapes.add_picture(f"{folder_path}/image_{j*3+k+1}.png", x, y, width=pptx.util.Cm(4), height=pptx.util.Cm(4))

    num_diapo += 1

# Notify user of all slides added at once
messagebox.showinfo("Slides Added", f" A total of {num_diapo} sliees were successfully added.")

# Prompt for presentation save location
messagebox.showinfo("Save", "Select location to save the presentation")
pres_name = filedialog.asksaveasfilename(defaultextension='.pptx')
prs.save(pres_name)
messagebox.showinfo("Presentation saved", f"Presentation saved in {pres_name}")
