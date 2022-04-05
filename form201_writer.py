# This script is used to bulk generate Western Australia Plant Registration form 201
import openpyxl, io, os
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.colors import black,blue,white,red,pink
from reportlab.lib.pagesizes import A4

xls_text_coords=[[130,450],[330,450],[130,423],[460,423],[130,400],[460,400],[130,333],
    [130,285],[300,285],[460,285],
    [50,242],[300,242],[50,218],[300,218],[50,195],[130,135]]
sec2=[[130,380,'Burrup Road'],[130,360,'Burrup'],[330,360,'WA'],[460,360,'6714']]

def GeneratePDF():
    xls=r'form_list.xlsx'
    wb=openpyxl.load_workbook(filename=xls,read_only=True, keep_vba=True)
    ws=wb.worksheets[0]
    lastRow=4 # Default row 4 or ws.max_row
    while ws.cell(column=1, row=lastRow).value : lastRow+=1
    
    entry_no=lastRow-3 #
    col_num=17
    text=[['']]*col_num
 
    #read in coordinate to a list
    for i in range(entry_no):
        for j in range(col_num):
            cellvalue=ws.cell(row=3+i, column=j+1).value
            if cellvalue is not None:
                text[j]=cellvalue
            else:
                text[j]='UNKNOWN'
        asset_id=text[0]
    
        pdf_name='form201.pdf' # place the file under same path with this script  
        pdf_output=str(asset_id)+'_FORM201(DRAFT).pdf'
        print("Generating File: "+pdf_output)
        write_pdf(pdf_name,pdf_output, text,xls_text_coords)
        
    wb.close()

def write_pdf(pdf_name,pdf_output,text,xls_text_coords):

    def fillout_static_text(canv,alist):
        for text_list in alist:
            x,y,text=text_list
            canv.drawString(x,y,text)

    def add_watermark(canv):
        canv.setFont("Helvetica-Bold", 16); canv.setFillColor(red)
        canv.setFillColor(pink)
        canv.setDash(3,2)
        canv.rect(395,670,190,100,stroke=1,fill=1)
        canv.setFillColor(red)
        canv.drawString(400,750,'DRAFT')
        canv.drawString(400,730,'DO NOT USE!')
        canv.drawString(400,690,str(text[0]))

    def mergepage(page_i,packet):
        packet.seek(0)
        pdf_buffer = PdfFileReader(packet)
        page=existing_pdf.getPage(page_i)
        page.mergePage(pdf_buffer.getPage(page_i))
        output.addPage(page)
        output.write(outputStream)

    #Add 1st page text
    pre_text=['Design Pressure: ', 'Design Temperature: ', 'Volume: ', 'Content Class: ', 'Hazard Level: ']
    packet = io.BytesIO()
    canv = canvas.Canvas(packet, pagesize = A4)
    canv.setAuthor('Haitao Han')
    canv.setFont("Helvetica-Bold", 9); canv.setStrokeColor(blue)
    coords=[[0]*2]*14
    for i in range(0,len(xls_text_coords)):
        coords=tuple(xls_text_coords[i])
        if i <10:
            if i==2: #special case: join Asset ID and plant description
                canv.drawString(coords[0], coords[1]," ".join([text[0],text[3]]))
            elif i==6: # special case: append 'Room' to plant location
                canv.drawString(coords[0], coords[1], 'Room '+text[7])
            else:
                if text[i+1]=='UNKNOWN':
                    canv.setFillColor(red)
                else:
                    canv.setFillColor(blue)
                canv.drawString(coords[0], coords[1],str(text[i+1]))
        elif i==15:
            canv.drawString(coords[0], coords[1],text[i+1])
        else:
            canv.drawString(coords[0], coords[1],pre_text[i-10]+text[i+1])
#    fillout_static_text(canv,sec1)
    canv.setFont("Helvetica-Bold", 14);canv.drawString(30,660,'X')
    canv.setFont("Helvetica-Bold", 9);canv.drawString(72,580,'X')
    fillout_static_text(canv,sec2)
    add_watermark(canv)
    #Fill out section 1 application for registration
    canv.showPage()

    #Add 2nd page text
    canv.setFont("Helvetica-Bold", 9); canv.setFillColor(blue);canv.setStrokeColor(blue)
    add_watermark(canv)
    canv.showPage()
    
    #Add 3rd page text
    canv.showPage()
    canv.save()

    outputStream = open(pdf_output, "wb")
    output = PdfFileWriter()
    output.addMetadata({
        '/Author': 'Haitao Han, haitao.han@applus.com',
        '/Title': 'WorkSafe Plant Registration Form 201',
        '/Subject':'Woodside KGP plant registration application form',
        '/Keywords': text[0],
        '/Producer': 'Hans PDF Generator'
    })
    if not pdf_name:
        outputStream.write(packet.getvalue())
    else:  
        existing_pdf = PdfFileReader(open(pdf_name, "rb"))
        for page_i in range(existing_pdf.numPages):
            mergepage(page_i,packet) 

    outputStream.close()
    #os.startfile(pdf_output,'open')

if __name__ == '__main__':
    GeneratePDF()
