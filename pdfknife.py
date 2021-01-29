from tkinter import *
from tkinter import ttk
from tkinter import messagebox #for messagebox.
from tkinter import filedialog
from TkinterDnD2 import *
import pdf2image as ph
import os,io,datetime, tabula

from PyPDF2 import PdfFileWriter, PdfFileReader,PdfFileMerger
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A4,A3, landscape
from reportlab.lib.colors import Color, black, blue, red,white,green


def rotate_page(FileName):
    pdf_in = open(FileName, 'rb')
    pdf_reader = PdfFileReader(pdf_in)
    pdf_writer = PdfFileWriter()

    for pageNum in range(pdf_reader.numPages):
        page = pdf_reader.getPage(pageNum)
        ODegree = pdf_reader.getPage(pageNum).get('/Rotate')
        print ("Page " + str(pageNum+1) + " orientation degrees is " + str(ODegree))
        if ODegree == 0: 
            page.rotateCounterClockwise(90)    
        pdf_writer.addPage(page)

    pdf_out = open('rotated.pdf', 'wb')
    pdf_writer.write(pdf_out)
    pdf_out.close()
    pdf_in.close()

def page_orientation (FileName):
    pdf_in = open(FileName, 'rb')
    pdf_reader = PyPDF2.PdfFileReader(pdf_in)


    print (str(FileName))
    for pageNum in range(pdf_reader.numPages):
        page = pdf_reader.getPage(pageNum)
        ODegree = pdf_reader.getPage(pageNum).get('/Rotate')
        print ("Page " + str(pageNum+1) + " orientation degrees is " + str(ODegree))

    pdf_in.close()
    
def grid_lines(pdf_name):     
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=pg.get())
    c.setFont('Helvetica-Bold',6)
    c.setFillColor(red)
    c.setDash(2,1)
    origin = (0, 0)
    
    x0, y0 = (0, 0)
    #(w,h)=tuple(tuple(map(int, tup)) for tup in pg.get())     

    if pg.get()=='A4':
        w,h=int(A4[0]), int(A4[1])
    elif pg.get()=='A3':
        w,h=int(A3[0]), int(A3[1])    
    #w, h = (595,842)
    print (w,h)
    c.setStrokeColor(blue)
    c.setLineWidth(0)
    x_step=int(x_spin.get())
    y_step=int(y_spin.get())

    ##
    
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
        print (pdf_name)
        existing_pdf=PdfFileReader(open(pdf_name, "rb"),strict=False)
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
            print (item)
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


def drop4(event):
    for item in listbox4.tk.splitlist(event.data):
        listbox4.insert(END,item)
        
def add_files_listbox_r():
    filez = filedialog.askopenfilenames(parent=root,title='Choose a file')
    for item in root.tk.splitlist(filez):
        listbox4.insert(END, item)

def save_as_r():
    save_path=filedialog.asksaveasfilename(parent=root, title='Save New PDF to …', filetypes=[('PDF files','*.pdf')],defaultextension='.pdf')
    deg=int(re.sub('(\d+)°','\g<1>',degs.get()))
    page_str=re.split(',| ',pages_r.get().strip())
    page_list=[]
    if page_str[0]:
        for i in page_str :
            if '-' in i:
                    pg_range=re.split('-', i)
                    if len(pg_range)==2: 
                            page_list.extend(list(range(int(pg_range[0])-1,int(pg_range[1]))))
            else:
                page_list.append(int(i)-1)
    for (i,item) in enumerate(listbox4.get(0,END)):
        pdf_in = open(item, 'rb')
        pdf_reader = PdfFileReader(pdf_in)
        pdf_writer = PdfFileWriter()

        for pageNum in range(pdf_reader.numPages):
            page = pdf_reader.getPage(pageNum)
            if pageNum  in page_list:
                    page.rotateClockwise(deg)
            elif len(page_list)==0:
                pass
            #page.compressContentStreams()
            pdf_writer.addPage(page)

        pdf_out = open(save_path, 'wb')
        pdf_writer.write(pdf_out)
        pdf_out.close()
        pdf_in.close()
        listbox4.delete(0,END)

def add_files_listbox_pdf_excel():
    filez = filedialog.askopenfilenames(parent=root,title='Choose a file')
    for item in root.tk.splitlist(filez):
        listbox5.insert(END, item)       

def drop5(event):
    for item in listbox5.tk.splitlist(event.data):
        listbox5.insert(END,item)
        
def save_as_pdf_excel():
    import csv,openpyxl
    for item in listbox5.get(0,END):
        filename=os.path.basename(item)
        filename_base=os.path.splitext(filename)[0]
        output_csv=os.path.join(os.environ['USERPROFILE'] + '\Desktop',filename_base+'.csv')
        output_xls=os.path.join(os.environ['USERPROFILE'] + '\Desktop',filename_base+'.xlsx')
        tabula.convert_into(item,output_csv,pages = "all", all=True)
        wb = openpyxl.Workbook()
        ws = wb.active
        with open(output_csv) as f:
            reader = csv.reader(f)
            for row in reader:
                ws.append(row)
        wb.save(output_xls)
                   
                            
        #tabula.convert_into(item,output,pages = "all", all=True)
        
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
tab.add(tab5,text='PDF Table Extraction')
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
frame7=LabelFrame(tab4,text="Merge PDF")
frame7.grid(row=0,column=0,padx=5, pady=5,sticky=N+E+S+W)
frame8=LabelFrame(tab4,text="List Operation")
frame8.grid(row=1,column=0,padx=5, pady=5,sticky=N+E+S+W)

frame9=LabelFrame(tab5,text="PDF table extraction")
frame9.grid(row=0,column=0,padx=5, pady=5,sticky=N+E+S+W)
frame10=LabelFrame(tab5,text="List Operation")
frame10.grid(row=1,column=0,padx=5, pady=5,sticky=N+E+S+W)

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
page_sizes=['A4','A3']
pg=StringVar()
pg.set('A4')
pg_menu=OptionMenu(frame3, pg,*page_sizes)
pg_menu.grid(row=1,column=2,sticky=N+S+W)
orits=['Portrait','Landscape']
orit=StringVar()
orit.set('portrait')
orit_menu=OptionMenu(frame3, orit,*orits)
orit_menu.grid(row=1,column=3,sticky=N+S+W)
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
button_ok_r = Button(frame8,command=save_as_r, text="Rotate PDF",bg='gold')
button_ok_r.grid(row=1,column=2,sticky=N+E+S+W)


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

root.mainloop()

