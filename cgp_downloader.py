import pyautogui, os
from time import sleep

TOTAL_PAGE=70

WORKDIR=os.path.dirname(os.path.realpath(__file__))
BOOKDIR=os.path.join(WORKDIR,'cgp1')
if not os.path.exists(BOOKDIR):
    os.makedirs(BOOKDIR)

pyautogui.click(x=1000,y=1023)
sleep(5)
for i in range(1,TOTAL_PAGE):
    filename=os.path.join(BOOKDIR, str(i)+'.png')
    pyautogui.screenshot(region=(230,75,1458,928)).save(filename)
    sleep(3)
    pyautogui.click(x=1762,y=538)
    sleep(3)    
    

