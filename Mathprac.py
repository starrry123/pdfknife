from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import Color, black, blue, red,white,green
from datetime import date, time, datetime,timedelta
import random, io,os

def write_pdf():
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=A4)
    c.setFont('Helvetica-Bold',20)
    c.#setFillColor(black)
#    c.setDash(2,1)
#    origin = (0, 0)
    styles = [(10, 0, blue)]
    
    x0, y0 = origin
    w, h = (595,842)


    for style in styles:
        # apply styling
        step, lw, col = style
        c.setStrokeColor(col)
        c.setLineWidth(lw)
    
    # draw grid (right now only regular rectangular)
    # draw vertical lines
    x = x0
    while x < w:
        c.line(x, 0, x, h)
        c.setFillColor(black)
        #c.drawString(x,170,str(int(x/10)))
        #c.drawString(x,450,str(int(x/10)))
        x += 150       
    # draw horizontal lines

    y = y0
    while y < h:
        c.line(0, y, w, y)
        c.setFillColor(black)
        c.drawString(10,y+10,str(random.randint(100,999))+random.choice([' + ',' - '])+str(random.randint(100,999))+' = ')
        c.drawString(310,y+10,str(random.randint(100,999))+random.choice([' + ',' - '])+str(random.randint(100,999))+' = ')
        #c.drawString(570,y,str(int(y/10)))
        y += 40

    # draw origin circle
    c.circle(x0, y0, step/2.0)

 
    c.save()

    #buffer start from 0
    packet.seek(0)
    new_pdf = PdfFileReader(packet)
    new_pdf_file_name=os.path.join(os.path.dirname(__file__), 'test.'+str(datetime.timestamp(datetime.now()))+'.pdf')
    pdf=open(new_pdf_file_name,'wb')
    pdf.write(packet.getvalue())

    #os.startfile(new_pdf_file_name,'open')


write_pdf()
