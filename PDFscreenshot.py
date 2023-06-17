#THIS SCRIPT IS TO DEMONSTRATE A WAY TO REMOVE DRM PROTECTION FROM PDF USING LOW TECHNOLOGY
#CODE IS FOR DEMO PURPOSE ONLY
#USE ANY YOUR OWN RISK, NO LIABILITY FOR DAMAGES RESULTING FROM USING
#YOU MUST SET THE VIEW TO PAGE BY PAGE
import pyautogui
import time, io, os, ast
from PIL import Image, ImageDraw
from reportlab.pdfgen import canvas
import tkinter as tk
from tkinter import simpledialog, filedialog

def save_img2pdf(img_list,filename):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pageCompression=1)    
    for img in img_list:
        can.setPageSize(img.size)
        can.drawInlineImage(img, 0, 0, preserveAspectRatio=True)
        can.showPage()
    can.save()
    packet.seek(0)
    with open(filename, "wb") as file:
        file.write(packet.getvalue())

def page_screenshot(total_pages=1, bbox=(2204, 103, 895, 1271)):
    pyautogui.FAILSAFE = False # set pyautogui.FAILSAFE to False. DISABLING FAIL-SAFE IS NOT RECOMMENDED.
    # Delay between capturing screenshots and pressing Page Down
    delay = 0.1  # Adjust the delay as needed
    print('starting in 5 seconds...')
    #time.sleep(5)
    img_list=[]
    # Loop through the pages
    for page in range(total_pages):
        # Capture the screenshot within the bounding box
        screenshot = pyautogui.screenshot(region=bbox)
        padding=(50, 5,5,5)#Define Mask Padding (Left, Top, Right, Bottom)
        mask_regions=[(0, 0, padding[0], bbox[3]),(0,0,bbox[2],padding[1]), 
                    (bbox[2],bbox[3],-padding[2], bbox[3]), (bbox[2],bbox[3],bbox[2], bbox[3]-padding[3])]
        draw=ImageDraw.Draw(screenshot)
        for region in mask_regions:
            draw.rectangle(region,fill=(255,255,255))
        img_list.append(screenshot)
        pyautogui.press('pagedown')
        time.sleep(delay)
    return img_list

# Function to handle Save button click
def save_button_click():
    #Define output PDF filename
    filename= filedialog.asksaveasfilename(title='Choose a file')
    total_pages = int(pageno_entry.get())
    # Specify the bounding box to capture
    screen_region = ast.literal_eval(screen_entry.get())#(top_left_x, top_left_y, width, height)
    print(screen_region)
    #screen_region = (2204, 103, 895, 1271)  # (left, top, width, height) #FOR PDF-XChange Editor
    #screen_region = (2062, 44, 985, 1388) #FOR www.saiglobal.com online view
    #snip a region of screen and save as image, then use pyautogui.locateOnScreen()to get bbox reading
    #location = pyautogui.locateOnScreen('image.png')
    if filename:
        time.sleep(5)
        img_list=page_screenshot(total_pages,screen_region)
        save_img2pdf(img_list,filename)
        print('PDF successfully generate!')
        os.startfile(os.path.realpath(filename))

# Function to update the count label
def update_count(count):
    if count > 0:
        countdown_label.config(text=str(count))
        root.after(1000, update_count, count - 1)
    else:
        countdown_label.config(text="Countdown complete!")

# Create the main Tkinter window
root = tk.Tk()
root.title("PDF SCREENSHOT")

# Create a big labeled frame for PDF-screenshot
pdf_frame = tk.LabelFrame(root, text="PDF-screenshot")
pdf_frame.grid(row=0, column=0, padx=10, pady=10)

# Create the label for the countdown text
countdown_label = tk.Label(pdf_frame, text="Switch App focus before 5s countdown")
countdown_label.grid(row=1, column=0, padx=10, pady=10)

# Create a labeled frame for the screenshot region
frame = tk.LabelFrame(pdf_frame, text="Screenshot Region")
frame.grid(row=2, column=0, padx=10, pady=10)

# Create labels and entry fields for TopLeft-X, TopLeft-Y, Width, Height, and Total Page
labels = ["TopLeft-X", "TopLeft-Y", "Width", "Height", "Total Page"]
entries = []

default_values = [2204, 103, 895, 1271, 1]  # Default values

screen_label = tk.Label(frame, text='topleft-X, topleft-Y, width, height:')
screen_label.grid(row=1, column=0)
screen_entry=tk.Entry(frame,width=25)
screen_entry.insert(tk.END, '(2204, 103, 895, 1271)')
screen_entry.grid(row=1, column=1)
pageno_label=tk.Label(frame,text='Total Page:')
pageno_label.grid(row=2,column=0)
pageno_entry=tk.Entry(frame,width=25)
pageno_entry.insert(tk.END,1)
pageno_entry.grid(row=2, column=1)

# Create a smaller frame for Close button and Save to File button
button_frame = tk.Frame(pdf_frame)
button_frame.grid(row=3, column=0, padx=10, pady=10)

# Create Save to File button
save_button = tk.Button(button_frame, text="Save to File", command=save_button_click)
save_button.grid(row=0, column=0, padx=5)

# Create Close button
close_button = tk.Button(button_frame, text="Close", command=root.destroy)
close_button.grid(row=0, column=1, padx=5)

# Start the Tkinter event loop
root.mainloop()
