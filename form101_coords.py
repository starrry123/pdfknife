import io, os
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib.colors import black,blue,white,red,pink
from reportlab.lib.pagesizes import A4
#from reportlab.graphics import renderPDF

page1_coords=[[110, 525], [370, 525], [110, 490], [450, 490], [110, 470], [470, 470], [190, 380], [150, 160], 
        [330, 160],[120,435],[120,415],[120,398],[370,398],[530,398],
        [120,345],[120,328],[370,328],[370,313],[120,278], [120,243],[120,225],[370, 225],[530,225],[500,160]]
page2_coords=[[110,655],[110,638],[110,620],[110,603],[370,620],[370,603], [530,620],[110,568],[110,550],
        [370,550],[530,550],[110,420],[500, 160], [50, 373], [320, 373], [50, 350], [320, 350], [50, 325]]

def GeneratePDF():
    
    pdf_name='form101.pdf' # place the file under same path with this script  
    pdf_output='form101_coord.pdf'
    write_pdf(pdf_name,pdf_output)
        

def write_pdf(pdf_name,pdf_output):

    def mark_coords(canv,coords_list):
        canv.setFont("Helvetica-Bold", 9); canv.setFillColor(red);canv.setStrokeColor(red)
        canv.setLineWidth(0.7)
        for coords in coords_list:
            x,y=coords
            l_len=20
            canv.drawString(x+2,y+2,str(x)+','+str(y))
            canv.line(x,y, x+l_len,y)
            canv.line(x,y,x,y+l_len)
            canv.circle(x,y,1,stroke=1,fill=1)



    def mergepage(page_i,packet):
        packet.seek(0)
        pdf_buffer = PdfFileReader(packet)
        page=existing_pdf.getPage(page_i)
        page.mergePage(pdf_buffer.getPage(page_i))
        output.addPage(page)
        output.write(outputStream)


    #Page one addon text
    packet = io.BytesIO()
    canv = canvas.Canvas(packet, pagesize = A4)
    mark_coords(canv,page1_coords)
    canv.showPage()
    mark_coords(canv,page2_coords)
    canv.save()
    
    outputStream = open(pdf_output, "wb")
    output = PdfFileWriter()
    output.addMetadata({
        '/Author': 'Haitao Han, haitao.han@applus.com',
        '/Title': 'WorkSafe Plant Registration Form 101',
        '/Producer': 'Hans PDF Generator'
    })
    if not pdf_name:
        outputStream.write(packet.getvalue())
    else:  
        existing_pdf = PdfFileReader(open(pdf_name, "rb"))
        for page_i in range(existing_pdf.numPages):
            mergepage(page_i,packet) 

    outputStream.close()
    os.startfile(pdf_output,'open')

if __name__ == '__main__':
    GeneratePDF()