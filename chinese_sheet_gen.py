import io,os
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.colors import Color,black,blue,red,white
from reportlab.lib.pagesizes import A4
from datetime import date, time, datetime,timedelta
import reportlab.pdfbase.ttfonts #Import the registered font module of reportlab
hei=reportlab.pdfbase.ttfonts.TTFont('hei','simhei.ttf') #Import font
song=reportlab.pdfbase.ttfonts.TTFont('song','simsun.ttc') #Import font
reportlab.pdfbase.pdfmetrics.registerFont(hei) #Register the font in the current directory
reportlab.pdfbase.pdfmetrics.registerFont(song) #Register the font in the current directory


def grid_lines(pdf_name=None):

    packet = io.BytesIO()
    c = canvas.Canvas(packet,pagesize=A4)#, pagesize=landscape(A3))
    c.setFont('hei',32)
    chinese_chars='在线中文输入'
    string_list=list(chinese_chars)
    margin=40
    grid_size=20
    w,h=A4
    w,h=w-margin,h-margin
    w1,h1=w//grid_size,h//grid_size
    x,y = margin,margin
    while x<w:
        c.setLineWidth(1)
        c.setDash(1,0)           
        c.line(x, margin, x, h)
        x += 2*grid_size      

    while y<h:
        c.setDash(1,0)            
        c.setLineWidth(1)
        c.line(margin, y, w, y)
        y += 2*grid_size

    x,y = margin,margin
    while x<w:
        c.setLineWidth(0.2)
        c.setDash(6,3)
        c.line(x+grid_size,margin,x+grid_size,h)
        x += grid_size      

    while y<h-margin:
        c.setLineWidth(0.2)
        c.setDash(6,3)
        c.line(margin, y+grid_size, w, y+grid_size)
        y += grid_size

    x,y = margin+grid_size,h-1.7*grid_size
    while y>margin and len(string_list)>0:
        c.setFont('hei',32)            
        c.drawCentredString(x, y,string_list.pop(0))
        y -= 2*grid_size    

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
        for page_num in range(existing_pdf.numPages):
            page = existing_pdf.getPage(page_num)
            page.mergePage(new_pdf.getPage(0))
            output.addPage(page)
        new_pdf_file_name=os.path.join(os.path.dirname(__file__), os.path.splitext(pdf_name)[0]+".GRID"+".pdf")
        outputStream = open(new_pdf_file_name, "wb")
        output.write(outputStream)
        outputStream.close()
    else:
        new_pdf_file_name=os.path.join(os.path.dirname(__file__), 'GRID.'+str(datetime.timestamp(datetime.now()))+'.pdf')
        pdf=open(new_pdf_file_name,'wb')
        pdf.write(packet.getvalue())
    # Finally output new pdf


    os.startfile(new_pdf_file_name,'open')



if __name__ == '__main__':
        grid_lines()
	#os.system('pause')
