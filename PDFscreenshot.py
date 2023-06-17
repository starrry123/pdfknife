#THIS SCRIPT IS TO DEMONSTRATE A WAY TO REMOVE DRM PROTECTION FROM PDF USING LOW TECHNOLOGY
#CODE IS FOR DEMO PURPOSE ONLY
#USE ANY YOUR OWN RISK, NO LIABILITY FOR DAMAGES RESULTING FROM USING
import pyautogui
import time, io
from PIL import Image, ImageDraw
import io
from reportlab.pdfgen import canvas

#Define output PDF filename
filename='Test.pdf'
#Number of pages to capture
total_pages = 1
# Specify the bounding box to capture
bbox = (2204, 103, 895, 1271)  # (left, top, width, height)
#snip a region of screen and save as image, then use pyautogui.locateOnScreen()to get bbox reading
# Specify the image to search for
#image_path = 'image.png'
# Locate the image on the screen
#location = pyautogui.locateOnScreen(image_path)

def save_img2pdf(img_list,filename):
    packet = io.BytesIO()
    can = canvas.Canvas(packet)    
    for img in img_list:
        can.setPageSize(img.size)
        can.drawInlineImage(img, 0, 0, preserveAspectRatio=True)
        can.showPage()
    can.save()
    packet.seek(0)
    outputStream = open(filename, "wb")
    outputStream.write(packet.getvalue())
    outputStream.close()

def page_screenshot(total_pages=1):
    # Delay between capturing screenshots and pressing Page Down
    delay = 0.1  # Adjust the delay as needed
    print('starting in 5 seconds...')
    time.sleep(5)
    img_list=[]
    # Loop through the pages
    for page in range(total_pages):
        # Capture the screenshot within the bounding box
        screenshot = pyautogui.screenshot(region=bbox)
        mask_regions=[(0, 0, 50, bbox[3]),(0,0,bbox[2],5), (bbox[2],bbox[3],-5, bbox[3]), (bbox[2],bbox[3],bbox[2], bbox[3]-3)]
        draw=ImageDraw.Draw(screenshot)
        for region in mask_regions:
            draw.rectangle(region,fill=(255,255,255))
        img_list.append(screenshot)
        pyautogui.press('pagedown')
        time.sleep(delay)
    return img_list

img_list=page_screenshot(total_pages,bbox)
save_img2pdf(img_list,filename)
