from tkinter import *
from tkinter import ttk
from tkinter import messagebox #for messagebox.
from tkinter import filedialog
from TkinterDnD2 import *
import pdf2image as ph
import os,io,datetime
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import Color, black, blue, red,white,green

def grid_lines(pdf_name=None):     
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=A4)
    c.setFont('Helvetica-Bold',6)
    c.setFillColor(red)
    c.setDash(2,1)
    origin = (0, 0)

    x0, y0 = origin
    w, h = (595,842)

    c.setStrokeColor(blue)
    c.setLineWidth(0)
    x_step=int(x_spin.get())
    y_step=int(y_spin.get())
    # draw grid (right now only regular rectangular)
    # draw vertical lines
    x = x0
    while x < w:
        c.line(x, 0, x, h)
        c.setFillColor(black)
        if chkValue2.get()==True:
            c.drawString(x,10,str(int(x/10)))
            c.drawString(x,830,str(int(x/10)))
        x += x_step       
    # draw horizontal lines

    y = y0
    while y < h:
        c.line(0, y, w, y)
        c.setFillColor(red)
        if chkValue2.get()==True:
            c.drawString(20,y,str(int(y/10)))
            c.drawString(570,y,str(int(y/10)))
        y += y_step

    # draw origin circle
    c.circle(x0, y0, x_step/2.0,fill=1)
    c.save()

    #buffer start from 0
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    output = PdfFileWriter()
    page=None
    new_pdf_file_name=None
    if pdf_name is not None:
        existing_pdf=PdfFileReader(open(pdf_name, "rb"))
        page = existing_pdf.getPage(0)
        page.mergePage(new_pdf.getPage(0))
        output.addPage(page)
        new_pdf_file_name=os.path.join(os.path.dirname(__file__), os.path.splitext(pdf_name)[0]+".GRID"+".pdf")
        outputStream = open(new_pdf_file_name, "wb")
        output.write(outputStream)
        outputStream.close()
    else:
        new_pdf_file_name=os.path.join(os.path.dirname(__file__), 'GRID.'+str(datetime.datetime.timestamp(datetime.datetime.now()))+'.pdf')
        pdf=open(new_pdf_file_name,'wb')
        pdf.write(packet.getvalue())
    # Finally output new pdf
    

    os.startfile(new_pdf_file_name,'open')

def pdf_grid ():
    if chkValue1.get()==False:
        grid_lines()
    else:
        for (i,item) in enumerate(listbox2.get(0,END)):
            listbox2.delete(0)
            grid_lines(item)

 

def convert_image():
    folder_selected = filedialog.askdirectory()

    for (i,item) in enumerate(listbox.get(0,END)):
        listbox.delete(0)
        PDF=item
        images = ph.convert_from_path(PDF)
        for i,img in enumerate(images):
            img.save(os.path.join(folder_selected, os.path.splitext(os.path.basename(PDF))[0]+'_'+str(i)+'.jpg'))



def drop(event):
    for item in listbox.tk.splitlist(event.data):
        listbox.insert(END,item)

def drop2(event):
    for item in listbox2.tk.splitlist(event.data):
        listbox2.insert(END,item)

def add_files_listbox():
    filez = filedialog.askopenfilenames(parent=app,title='Choose a file')
    for item in app.tk.splitlist(filez):
        listbox3.insert(END, item)       




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

def drop3(event):
    for item in listbox3.tk.splitlist(event.data):
        listbox3.insert(END,item)

def add_files_listbox3():
    filez = filedialog.askopenfilenames(parent=root,title='Choose a file')
    for item in root.tk.splitlist(filez):
        listbox3.insert(END, item)       


def save_as3():
    pdf_merger = PdfFileMerger(strict=False)
    save_path=filedialog.asksaveasfilename(parent=root, title='Save New PDF to …', filetypes=[('PDF files','*.pdf')],defaultextension='.pdf')
    
    for (i,item) in enumerate(listbox.get(0,END)):
        pdf_merger.append(item)
    with open(save_path, 'wb') as save_path:
        pdf_merger.write(save_path)   
    listbox.delete(0,END)
    
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
###############PDF to Photo GUI#########################
listbox = Listbox(frame1, width=60)
listbox.grid(row=0,column=0,padx=5, pady=5, ipadx=5, ipady=5, sticky=N+S+E+W)
for item in sys.argv[1:]:
    listbox.insert(END, item)
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
for item in sys.argv[1:]:
    listbox2.insert(END, item)
listbox2.drop_target_register(DND_FILES)
listbox2.dnd_bind('<<Drop>>', drop2)

chkValue1 = BooleanVar() 
chkValue1.set(False)
chkValue2 = BooleanVar() 
chkValue2.set(True)
check1=Checkbutton(frame3,text="Add grid to below?",var=chkValue1)
check1.grid(row=1,column=0,sticky=N+S+W)
check2=Checkbutton(frame3,text="Add scale numbers?",var=chkValue2)
check2.grid(row=1,column=1,sticky=N+S+W)
lab_xy=Label(frame3,text='Grid Size: (pixel)')
lab_xy.grid(row=2,column=0,sticky=N+S+W)
x_spin=Spinbox(frame3,from_=10,to=595)
x_spin.grid(row=2,column=1,sticky=N+S+W)
y_spin=Spinbox(frame3,from_=10,to=842)
y_spin.grid(row=2,column=2,sticky=N+S+W)

button_clearfile2 = Button(frame3, text='➖ Clear List')
button_clearfile2.grid(row=3,column=0,sticky=N+S+E+W)
button_clearfile2.bind("<Button-1>", lambda e: listbox2.delete(0,END))
button_grid = Button(frame3,command=pdf_grid, text="Create PDF grid",bg='gold')
button_grid.grid(row=3,column=1,sticky=N+S+W)


##################PDF Merge GUI#################
listbox3 = Listbox(frame5, width=60)
listbox3.grid(row=0,column=0,padx=5, pady=5, ipadx=5, ipady=5, sticky=N+S+E+W)
for item in sys.argv[1:]:
    listbox3.insert(END, item)
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

button_ok = Button(frame6,command=save_as3, text="Save PDF",bg='gold')
button_ok.grid(row=0,column=4,sticky=N+E+S+W)


root.mainloop()

