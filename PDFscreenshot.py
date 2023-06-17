from tkinter import *
from tkinter import ttk
from tkinter import messagebox #for messagebox.
from tkinter import filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import pdf2image as ph
import os,io,datetime, tabula,copy, re
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4,A3, landscape
from reportlab.lib.colors import black, blue, red,white,green
from PIL import Image, ImageDraw
import io
import pyautogui
import time


def rotate_page(FileName):
    pdf_writer = PdfFileWriter()

    with open(FileName, 'rb') as pdf_in:
        pdf_reader = PdfFileReader(pdf_in)

        for pageNum in range(pdf_reader.numPages):
            page = pdf_reader.getPage(pageNum)
            ODegree = pdf_reader.getPage(pageNum).get('/Rotate')
            print ("Page " + str(pageNum+1) + " orientation degrees is " + str(ODegree))
            if ODegree == 0:
                page.rotateCounterClockwise(90)
            pdf_writer.addPage(page)

    with open('rotated.pdf', 'wb') as pdf_out:
        pdf_writer.write(pdf_out)

def page_orientation (FileName):
    with open(FileName, 'rb') as pdf_in:
        pdf_reader = PdfFileReader(pdf_in)

        print (str(FileName))
        for pageNum in range(pdf_reader.numPages):
            ODegree = pdf_reader.getPage(pageNum).get('/Rotate')
            print ("Page " + str(pageNum+1) + " orientation degrees is " + str(ODegree))

def grid_lines(pdf_name=None):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=eval(pg.get()))
    c.setFont('Helvetica-Bold',6)
    c.setFillColor(red)
    c.setDash(2,1)
    c.setStrokeColor(black); c.setLineWidth(0)
    w,h=eval(pg.get()) #get page size
    x_step,y_step=int(x_spin.get()),int(y_spin.get())
    deg=int(deg_spin.get())   
    x,y = -w, -h
    c.rotate(deg)    

    while x < w or y< h: 
        c.line(x, -h, x, h)
        c.setFillColor(red)
        if CheckV1.get():
            c.drawString(x,10,str(int(x/10)))
            c.drawString(x,830,str(int(x/10)))
            c.drawString(20,y,str(int(y/10)))
            c.drawString(570,y,str(int(y/10)))
        c.line(0, y, 2*w, y)
        c.setFillColor(red)        
        x += x_step
        y += y_step
    c.circle(0, 0, x_step/2.0,fill=1) # draw origin circle
    c.save()

    packet.seek(0)
    grid_stream = PdfFileReader(packet)
    output = PdfFileWriter()
    if pdf_name:
        existing_pdf=PdfFileReader(open(pdf_name, "rb"))
        new_PDF=os.path.join(os.path.dirname(__file__), os.path.splitext(pdf_name)[0]+".GRID"+".pdf")
        for pageNum in range(existing_pdf.numPages):
            page = existing_pdf.getPage(pageNum)
            page.mergePage(grid_stream.getPage(0))
            output.addPage(page)
        with open(new_PDF, "wb") as outputStream:
            output.write(outputStream)
    else:
        new_PDF=os.path.join(os.path.dirname(__file__), 'GRID.'+str(datetime.datetime.timestamp(datetime.datetime.now()))+'.pdf')
        with open(new_PDF,'wb') as pdf:
            pdf.write(packet.getvalue())
            
    os.startfile(new_PDF,'open')

def pdf_grid ():
    if len(listbox2.get(0,END))>0:
        for pdf in listbox2.get(0,END):
            listbox2.delete(0)
            grid_lines(pdf)
    else:
        print('create empty grid file')
        grid_lines()

def convert_image():
    folder_selected = filedialog.askdirectory()

    for PDF in listbox.get(0,END):
        listbox.delete(0)
        images = ph.convert_from_path(PDF)
        for i,img in enumerate(images):
            img.save(os.path.join(folder_selected, os.path.splitext(os.path.basename(PDF))[0]+'_'+str(i)+'.jpg'))

def drop(event):
    for item in listbox.tk.splitlist(event.data):
        listbox.insert(END,item)

def drop2(event):
    for item in listbox2.tk.splitlist(event.data):
        listbox2.insert(END,item)

def drop3(event):
    for item in listbox3.tk.splitlist(event.data):
        listbox3.insert(END,item)

def drop4(event):
    for item in listbox4.tk.splitlist(event.data):
        listbox4.insert(END,item)

def drop7(event):
    for item in listbox7.tk.splitlist(event.data):
        listbox7.insert(END,item)

def add_files_listbox():
    filez = filedialog.askopenfilenames(parent=root,title='Choose a file')
    for item in root.tk.splitlist(filez):
        listbox.insert(END, item)

def move_up3():
    idxs = listbox3.curselection()
    if not idxs:
        return
    for pos in idxs:
        if pos==0:
            continue
        text=listbox3.get(pos)
        listbox3.delete(pos)
        listbox3.insert(pos-1,text)
        listbox3.selection_set(pos-1)

def move_down3():
    idxs = listbox3.curselection()
    if not idxs:
        return
    for pos in idxs:
        if pos==0:
            continue
        text=listbox3.get(pos)
        listbox3.delete(pos)
        listbox3.insert(pos+1,text)
        listbox3.selection_set(pos+1)

def add_files_listbox3():
    filez = filedialog.askopenfilenames(parent=root,title='Choose a file')
    for item in root.tk.splitlist(filez):
        listbox3.insert(END, item)

def pdf_merge_save():
    out_pdf = PdfFileWriter()
    save_path=filedialog.asksaveasfilename(parent=root, initialdir=os.path.dirname(os.path.abspath(listbox3.get(0))),title='Save New PDF to …', filetypes=[('PDF files','*.pdf')],defaultextension='.pdf')
    print(save_path)
    for item in listbox3.get(0,END):
        source_pdf = PdfFileReader(open(item,'rb'),strict=False)
        for p in range(source_pdf.getNumPages()):         
            out_pdf.addPage(copy.copy(source_pdf.getPage(p)))
    with open(save_path, 'wb') as save_path_stream:
        out_pdf.write(save_path_stream)
    listbox3.delete(0,END) 

def add_files_listbox_r():
    filez = filedialog.askopenfilenames(parent=root,title='Choose a file')
    for item in root.tk.splitlist(filez):
        listbox4.insert(END, item)

def save_as_rot():
    save_path=filedialog.asksaveasfilename(parent=root, initialdir=os.path.dirname(os.path.abspath(listbox4.get(0))),title='Save New PDF to …', filetypes=[('PDF files','*.pdf')],defaultextension='.pdf')
    deg=int(re.sub('(\d+)°','\g<1>',degs.get()))
    print (deg)
    page_str= re.split(',| ',pages_r.get().strip())
    page_list=[]
    if page_str[0]:
        for i in page_str :
            if '-' in i:
                pg_range=re.split('-', i)
                if len(pg_range)==2:
                    page_list.extend(list(range(int(pg_range[0])-1,int(pg_range[1]))))
            else:
                page_list.append(int(i)-1)

    pdf_out = PdfFileWriter()
    for pdf in listbox4.get(0,END):
        with open(pdf, 'rb') as pdf_instream:
            pdf_reader = PdfFileReader(pdf_instream)
            for pageNum in range(pdf_reader.numPages):
                page = pdf_reader.getPage(pageNum)
                if pageNum  in page_list or not page_list:
                    page.rotateClockwise(deg)
                pdf_out.addPage(page)
            #save_path=os.path.splitext(os.path.abspath(pdf))[0]+'_r.pdf'
            with open(save_path, 'wb') as pdf_outstream:
                pdf_out.write(pdf_outstream)
        listbox4.delete(0,END)

def add_files_listbox_pdf_excel():
    filez = filedialog.askopenfilenames(parent=root,title='Choose a file')
    for item in root.tk.splitlist(filez):
        listbox5.insert(END, item)

def drop5(event):
    for item in listbox5.tk.splitlist(event.data):
        listbox5.insert(END,item)

def drop6(event):
    for item in listbox6.tk.splitlist(event.data):
        listbox6.insert(END,item)
        
def save_as_pdf_excel():
    import csv,openpyxl
    for pdf in listbox5.get(0,END):
        filename=os.path.basename(pdf)
        filename_base=os.path.splitext(filename)[0]
        output_csv=os.path.join(os.environ['USERPROFILE'], 'Desktop',filename_base+'.csv')
        output_xls=os.path.join(os.environ['USERPROFILE'], 'Desktop',filename_base+'.xlsx')
        print(output_csv,output_xls)
        tabula.convert_into(pdf,output_csv, output_format='csv',pages = "all")
        wb = openpyxl.Workbook()
        ws = wb.active
        with open(output_csv) as f:
            reader = csv.reader(f)
            for row in reader:
                ws.append(row)
        wb.save(output_xls)
        #tabula.convert_into(item,output,pages = "all", all=True)

def add_files_listbox_img2pdf():
    filez = filedialog.askopenfilenames(parent=root,title='Choose a file')
    for item in root.tk.splitlist(filez):
        listbox6.insert(END, item)


def save_img2pdf(img_list,filename):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pageCompression=1)    
    for img in img_list:
        can.setPageSize(img.size)
        can.drawInlineImage(img, 0, 0, preserveAspectRatio=True)
        can.showPage()
    can.save()
    packet.seek(0)
    outputStream = open(filename, "wb")
    outputStream.write(packet.getvalue())
    outputStream.close()

def add_pdf_fmg():
    filez = filedialog.askopenfilenames(parent=root,title='Choose a file')
    for item in root.tk.splitlist(filez):
        listbox7.insert(END, item)

def pdf_compress(filename):
    output_filename=os.path.join(os.path.dirname(filename),os.path.splitext(os.path.basename(filename))[0]+'_compressed.pdf')
    with open(filename, 'rb') as file:
        # Create a PDF object
        pdf = PdfFileReader(file)
        
        # Create a new PDF file
        output = PdfFileWriter()
        
        # Iterate through each page of the original PDF
        for page in range(pdf.getNumPages()):
            # Get the current page
            current_page = pdf.getPage(page)
            
            # Compress the page
            current_page.compressContentStreams()
            
            # Add the compressed page to the new PDF
            output.addPage(current_page)
        
        # Save the compressed PDF to a file
        with open(output_filename, 'wb') as output_file:
            output.write(output_file)
    os.rename(output_filename, filename,output_filename)



def pdf_fmg_rename():
    from PyPDF2 import PdfFileReader
    for filename in listbox7.get(0,END):
        listbox7.delete(0)
        asset_id=os.path.basename(os.path.dirname(filename))  
        with open(filename,'rb') as fp:
            PDFIN=PdfFileReader(fp)
            print(PDFIN.getDocumentInfo())
            createdate=PDFIN.getDocumentInfo()['/CreationDate']
        m=re.search(r'^D:(\d{4})(\d{2})(\d{2})(\d{6}\+)\d',createdate)
        if m is not None:
            date_stamp=m.group(1)+m.group(2)+m.group(3)
            newfile_name=asset_id+'_'+date_stamp+' Report.pdf'
            newfile_path=os.path.dirname(filename)+'/'+newfile_name
            os.rename(filename, newfile_path)
            #if os.path.getsize(newfile_path) > 2 * 1024 * 1024:  # 2MB
            #    pdf_compress(newfile_path)

def save_img2pdf(img_list,save_pdf):
    packet = io.BytesIO()
    can = canvas.Canvas(packet)    
    for img in img_list:
        can.setPageSize(img.size)
        can.drawInlineImage(img, 0, 0, preserveAspectRatio=True)
        can.showPage()
    can.save()
    packet.seek(0)
    outputStream = open(save_pdf, "wb")
    outputStream.write(packet.getvalue())
    outputStream.close()


def page_screenshot(total_pages=1, bbox=(2204, 103, 895, 1271)):
    pyautogui.FAILSAFE = False # set pyautogui.FAILSAFE to False. DISABLING FAIL-SAFE IS NOT RECOMMENDED.
    # Delay between capturing screenshots and pressing Page Down
    delay = 0.1  # Adjust the delay as needed
    #print('starting in 5 seconds...')
    #time.sleep(5)
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

def pdf_screenshot():
    save_path=filedialog.asksaveasfilename(parent=root, initialdir=os.path.dirname(os.path.abspath(listbox3.get(0))),title='Save New PDF to …', filetypes=[('PDF files','*.pdf')],defaultextension='.pdf')
    img_list=page_screenshot()
    save_img2pdf(img_list,save_path)



# Function to handle Save button click
def save_button_click():
    #Define output PDF filename
    filename= filedialog.asksaveasfilename(title='Choose a file')
    #Number of pages to capture
    #total_pages = get_page_number()
    total_pages = int(entries[4].get())
    # Specify the bounding box to capture
    top_left_x = int(entries[0].get())  # Assuming "TopLeft-X" entry is at index 0 in the 'entries' list
    top_left_y = int(entries[1].get())  # Assuming "TopLeft-Y" entry is at index 1 in the 'entries' list
    width = int(entries[2].get())  # Assuming "Width" entry is at index 2 in the 'entries' list
    height = int(entries[3].get())  # Assuming "Height" entry is at index 3 in the 'entries' list
    screen_region = (top_left_x, top_left_y, width, height)
    #screen_region = (2204, 103, 895, 1271)  # (left, top, width, height) #THIS IS FOR PDF-XChange Editor
    #screen_region = (2062, 44, 985, 1388) #THIS IS FOR www.saiglobal.com online view
    #snip a region of screen and save as image, then use pyautogui.locateOnScreen()to get bbox reading
    #import pyautogui
    #location = pyautogui.locateOnScreen('image.png')
    if filename != '':
        time.sleep(5)
        img_list=page_screenshot(total_pages,screen_region)
        save_img2pdf(img_list,filename)
        print('PDF successfully generate!')
        os.startfile(filename)        



img_list=page_screenshot()
save_img2pdf(img_list,'TEST.pdf')
        
root = TkinterDnD.Tk()
root.title("PDF Knife -- A collection of PDF toolkits")
tab=ttk.Notebook(root)
tab1=ttk.Frame(tab)
tab.add(tab1,text='PDF2Photo')
tab2=ttk.Frame(tab)
tab.add(tab2,text='PDF grid')
tab.pack(expand=1,fill='both')
tab3=ttk.Frame(tab)
tab.add(tab3,text='PDF merge')
tab.pack(expand=1,fill='both')
tab4=ttk.Frame(tab)
tab.add(tab4,text='PDF rotate')
tab.pack(expand=1,fill='both')
tab5=ttk.Frame(tab)
tab.add(tab5,text='PDF2Excel')
tab.pack(expand=1,fill='both')
tab6=ttk.Frame(tab)
tab.add(tab6,text='Image2PDF')
tab.pack(expand=1,fill='both')
tab7=ttk.Frame(tab)
tab.add(tab7,text='PDF Screenshot')
tab.pack(expand=1,fill='both')

frame1=LabelFrame(tab1,text="Drop PDF files here")
frame1.grid(row=0,column=0,padx=5, pady=5,sticky=N+E+S+W)
frame2=LabelFrame(tab1,text="Create grid lines in PDF")
frame2.grid(row=1,column=0,padx=5, pady=5,sticky=N+E+S+W)
frame3=LabelFrame(tab2,text="Create PDF grids")
frame3.grid(row=0,column=0,padx=5, pady=5,sticky=N+E+S+W)
frame4=LabelFrame(tab2,text="Create PDF grids")
frame4.grid(row=1,column=0,padx=5, pady=5,sticky=N+E+S+W)
frame5=LabelFrame(tab3,text="Merge PDF")
frame5.grid(row=0,column=0,padx=5, pady=5,sticky=N+E+S+W)
frame6=LabelFrame(tab3,text="List Operation")
frame6.grid(row=1,column=0,padx=5, pady=5,sticky=N+E+S+W)
frame7=LabelFrame(tab4,text="Rotate PDF")
frame7.grid(row=0,column=0,padx=5, pady=5,sticky=N+E+S+W)
frame8=LabelFrame(tab4,text="List Operation")
frame8.grid(row=1,column=0,padx=5, pady=5,sticky=N+E+S+W)

frame9=LabelFrame(tab5,text="PDF table extraction")
frame9.grid(row=0,column=0,padx=5, pady=5,sticky=N+E+S+W)
frame10=LabelFrame(tab5,text="List Operation")
frame10.grid(row=1,column=0,padx=5, pady=5,sticky=N+E+S+W)

frame11=LabelFrame(tab6,text="Convert Images to PDF")
frame11.grid(row=0,column=0,padx=5, pady=5,sticky=N+E+S+W)
frame12=LabelFrame(tab6,text="Image List")
frame12.grid(row=1,column=0,padx=5, pady=5,sticky=N+E+S+W)

frame13=LabelFrame(tab7,text="PDF Screenshot")
frame13.grid(row=0,column=0,padx=5, pady=5,sticky=N+E+S+W)

  

###############PDF to Photo GUI#########################
listbox = Listbox(frame1, width=60)
listbox.grid(row=0,column=0,padx=5, pady=5, ipadx=5, ipady=5, sticky=N+S+E+W)
#for item in sys.argv[1:]:
#    listbox.insert(END, item)
xscrollbar2 = Scrollbar(frame1,orient=HORIZONTAL)
xscrollbar2.grid(row=1, column=0,sticky=E+W)
listbox.config(xscrollcommand=xscrollbar2.set)
xscrollbar2.config(command=listbox.xview)
yscrollbar2 = Scrollbar(frame1,orient=VERTICAL )
yscrollbar2.grid(row=0, column=1,sticky=N+S+W)
listbox.config(yscrollcommand=yscrollbar2.set)
yscrollbar2.config(command=listbox.yview)
listbox.drop_target_register(DND_FILES)
listbox.dnd_bind('<<Drop>>', drop)
button_addfile = Button(frame2, command=add_files_listbox,text='➕ Add Files ')
button_addfile.grid(row=0,column=0,sticky=N+S+E+W)
button_clearfile = Button(frame2, text='➖ Clear List')
button_clearfile.grid(row=0,column=1,sticky=N+S+E+W)
button_clearfile.bind("<Button-1>", lambda e: listbox.delete(0,END))
button_ok = Button(frame2,command=convert_image, text="[Convert]",bg='gold')
button_ok.grid(row=0,column=2,sticky=E)

################PDF Grid GUI##########################
listbox2 = Listbox(frame4, width=60)
listbox2.grid(row=0,column=0,padx=5, pady=5, ipadx=5, ipady=5, sticky=N+S+E+W)
#for item in sys.argv[1:]:
#    listbox2.insert(END, item)
listbox2.drop_target_register(DND_FILES)
listbox2.dnd_bind('<<Drop>>', drop2)

CheckV1 = BooleanVar()
CheckV1.set(True)
#CheckV2 = BooleanVar()
#CheckV2.set(True)
check1=Checkbutton(frame3,text="Add scale numbers?",var=CheckV1)
check1.grid(row=1,column=0,sticky=N+S+W)
#check2=Checkbutton(frame3,text="Add scale numbers?",var=CheckV2)
#check2.grid(row=1,column=1,sticky=N+S+W)
page_sizes=['A4','A3']
pg=StringVar()
pg.set('A4')
pg_menu=OptionMenu(frame3, pg,*page_sizes)
pg_menu.grid(row=1,column=1,sticky=N+S+W)
orits=['Portrait','Landscape']
orit=StringVar()
orit.set('portrait')
orit_menu=OptionMenu(frame3, orit,*orits)
orit_menu.grid(row=1,column=2,sticky=N+S+W)
lab_xy=Label(frame3,text='Grid Size: (pixel)')
lab_xy.grid(row=2,column=0,sticky=N+S+W)
x_spin=Spinbox(frame3,from_=10,to=595)
x_spin.grid(row=2,column=1,sticky=N+S+W)
y_spin=Spinbox(frame3,from_=10,to=842)
y_spin.grid(row=2,column=2,sticky=N+S+W)

lab_deg1=Label(frame3,text='Rotate Degree: ')
lab_deg1.grid(row=3,column=0,sticky=N+S+W)
deg_spin=Spinbox(frame3,from_=0,to=90)
deg_spin.grid(row=3,column=1,sticky=N+S+W)
lab_deg2=Label(frame3,text='°')
lab_deg2.grid(row=3,column=2,sticky=N+S+W)

button_clearfile2 = Button(frame3, text='➖ Clear List')
button_clearfile2.grid(row=4,column=0,sticky=N+S+E+W)
button_clearfile2.bind("<Button-1>", lambda e: listbox2.delete(0,END))
button_grid = Button(frame3,command=pdf_grid, text="Create PDF grid",bg='gold')
button_grid.grid(row=4,column=1,sticky=N+S+W)

##################PDF Merge GUI#################
listbox3 = Listbox(frame5, width=60)
listbox3.grid(row=0,column=0,padx=5, pady=5, ipadx=5, ipady=5, sticky=N+S+E+W)
#for item in sys.argv[1:]:
#    listbox3.insert(END, item)
xscrollbar2 = Scrollbar(frame5,orient=HORIZONTAL)
xscrollbar2.grid(row=1, column=0,sticky=E+W)
listbox3.config(xscrollcommand=xscrollbar2.set)
xscrollbar2.config(command=listbox3.xview)
yscrollbar2 = Scrollbar(frame5,orient=VERTICAL )
yscrollbar2.grid(row=0, column=1,sticky=N+S+W)
listbox3.config(yscrollcommand=yscrollbar2.set)
yscrollbar2.config(command=listbox3.yview)
listbox3.drop_target_register(DND_FILES)
listbox3.dnd_bind('<<Drop>>', drop3)
button_moveup = Button(frame6, command=move_up3,text='↑ Move Up ')
button_moveup.grid(row=0,column=0,sticky=N+S+E+W)
button_movedown = Button(frame6, command=move_down3,text='↓ Move Down ')
button_movedown.grid(row=0,column=1,sticky=N+S+E+W)

button_addfile = Button(frame6, command=add_files_listbox3,text='➕ Add Files ')
button_addfile.grid(row=0,column=2,sticky=N+S+E+W)
button_clearfile3 = Button(frame6, text='➖ Clear List')
button_clearfile3.grid(row=0,column=3,sticky=N+S+E+W)
button_clearfile3.bind("<Button-1>", lambda e: listbox3.delete(0,END))

button_ok = Button(frame6,command=pdf_merge_save, text="Save PDF",bg='gold')
button_ok.grid(row=0,column=4,sticky=N+E+S+W)

##################PDF Rotation GUI#################
listbox4 = Listbox(frame7, width=60)
listbox4.grid(row=0,column=0,padx=5, pady=5, ipadx=5, ipady=5, sticky=N+S+E+W)
listbox4.drop_target_register(DND_FILES)
listbox4.dnd_bind('<<Drop>>', drop4)

lab_r=Label(frame8, text='Pages to be rotated:')
lab_r.grid(row=0,column=0)
pages_r=StringVar()
entry_r=Entry(frame8,width=15,textvariable=pages_r)
entry_r.grid(row=0,column=1)
deg_list=['90°','-90°', '180°']
degs=StringVar()
degs.set('90°')
deg_menu=OptionMenu(frame8, degs,*deg_list)
deg_menu.grid(row=0,column=2,sticky=N+E+S+W)
button_addfile_r = Button(frame8, command=add_files_listbox_r,text='➕ Add Files ')
button_addfile_r.grid(row=1,column=0,sticky=N+S+E+W)
button_clearfile_r = Button(frame8, text='➖ Clear List')
button_clearfile_r.grid(row=1,column=1,sticky=N+S+E+W)
button_clearfile_r.bind("<Button-1>", lambda e: listbox4.delete(0,END))
button_ok_r = Button(frame8,command=save_as_rot, text="Rotate PDF",bg='gold')
button_ok_r.grid(row=1,column=2,sticky=N+E+S+W)

##################PDF Table Extraction GUI#################
listbox5 = Listbox(frame9, width=60)
listbox5.grid(row=0,column=0,padx=5, pady=5, ipadx=5, ipady=5, sticky=N+S+E+W)
listbox5.drop_target_register(DND_FILES)
listbox5.dnd_bind('<<Drop>>', drop5)

button_addfile_pdf_excel = Button(frame10, command=add_files_listbox_pdf_excel,text='➕ Add Files ')
button_addfile_pdf_excel.grid(row=1,column=0,sticky=N+S+E+W)
button_clearfile_pdf_excel = Button(frame10, text='➖ Clear List')
button_clearfile_pdf_excel.grid(row=1,column=1,sticky=N+S+E+W)
button_clearfile_pdf_excel.bind("<Button-1>", lambda e: listbox5.delete(0,END))
button_ok_pdf_excel = Button(frame10,command=save_as_pdf_excel, text="Save Excel",bg='gold')
button_ok_pdf_excel.grid(row=1,column=2,sticky=N+E+S+W)

##################CONVERT IMAGE TO PDF GUI#################
listbox6 = Listbox(frame11, width=60)
listbox6.grid(row=0,column=0,padx=5, pady=5, ipadx=5, ipady=5, sticky=N+S+E+W)
listbox6.drop_target_register(DND_FILES)
listbox6.dnd_bind('<<Drop>>', drop6)

button_addfile_img2pdf = Button(frame12, command=add_files_listbox_img2pdf,text='➕ Add Files ')
button_addfile_img2pdf.grid(row=1,column=0,sticky=N+S+E+W)
button_clearfile_img2pdf = Button(frame12, text='➖ Clear List')
button_clearfile_img2pdf.grid(row=1,column=1,sticky=N+S+E+W)
button_clearfile_img2pdf.bind("<Button-1>", lambda e: listbox6.delete(0,END))
button_ok_img2pdf = Button(frame12,command=save_img2pdf, text="Save PDF",bg='gold')
button_ok_img2pdf.grid(row=1,column=2,sticky=N+E+S+W)

##################CHANGE PDF NAME#################
# listbox7 = Listbox(frame13, width=60)
# listbox7.grid(row=0,column=0,padx=5, pady=5, ipadx=5, ipady=5, sticky=N+S+E+W)
# listbox7.drop_target_register(DND_FILES)
# listbox7.dnd_bind('<<Drop>>', drop7)

# button_addfile_fmg2pdf = Button(frame14, command=add_pdf_fmg,text='➕ Add Files ')
# button_addfile_fmg2pdf.grid(row=1,column=0,sticky=N+S+E+W)
# button_clearfile_fmg2pdf = Button(frame14, text='➖ Clear List')
# button_clearfile_fmg2pdf.grid(row=1,column=1,sticky=N+S+E+W)
# button_clearfile_fmg2pdf.bind("<Button-1>", lambda e: listbox7.delete(0,END))
# button_ok_fmg2pdf = Button(frame14,command=pdf_fmg_rename, text="Save PDF",bg='gold')
# button_ok_fmg2pdf.grid(row=1,column=2,sticky=N+E+S+W)


################PDF SCREENSHOT GUI##########################


 
# Create the label for the countdown text
countdown_label = Label(frame13, text="Switch App focus before 5s countdown")
countdown_label.grid(row=1, column=0, padx=10, pady=10)

# Create a labeled frame for the screenshot region
frame = LabelFrame(frame13, text="Screenshot Region")
frame.grid(row=2, column=0, padx=10, pady=10)

# Create labels and entry fields for TopLeft-X, TopLeft-Y, Width, Height, and Total Page
labels = ["TopLeft-X", "TopLeft-Y", "Width", "Height", "Total Page"]
entries = []

default_values = [2204, 103, 895, 1271, 1]  # Default values

for i, (label_text, default_value) in enumerate(zip(labels, default_values)):
    # Create label
    label = Label(frame, text=label_text + ":")
    label.grid(row=i, column=0, sticky=W, padx=5, pady=5)

    # Create entry field
    entry = Entry(frame, width=10)
    entry.insert(END, default_value)  # Set default value
    entry.grid(row=i, column=1, padx=5, pady=5, sticky=E)
    entries.append(entry)

# Create a smaller frame for Close button and Save to File button
button_frame = Frame(frame13)
button_frame.grid(row=3, column=0, padx=10, pady=10)

# Create Save to File button
save_button = Button(button_frame, text="Save to File", command=save_button_click,bg='gold')
save_button.grid(row=0, column=0, padx=5)
# Create Close button
close_button = Button(button_frame, text="Close", command=root.destroy)
close_button.grid(row=0, column=1, padx=5)


root.mainloop()