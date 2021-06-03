#Batch pdf_merger2.py
import glob
from PyPDF2 import PdfFileMerger
 
def merger(output_path, input_paths):
    pdf_merger = PdfFileMerger(strict=False)
    file_handles = []
 
    for path in input_paths:
        pdf_merger.append(path)
 
    with open(output_path, 'wb') as fileobj:
        pdf_merger.write(fileobj)
 
if __name__ == '__main__':
    paths = glob.glob(r'*.pdf')
    paths.sort()
    merger('combined.pdf', paths)

#ANNOT CLEANING annot-clean2.py
import os, pdfrw, glob

directory= r"."
for filename in os.listdir(directory):
    if filename.startswith("KPP-"): 
        print(os.path.join(directory, filename))
        reader = pdfrw.PdfReader(filename)
        for p in reader.pages:
            if p.Annots:
            # See PDF reference (Sec. 12.5.6) for all annotation types
                p.Annots = [a for a in p.Annots if a.Subtype == "/Link"]

                pdfrw.PdfWriter(filename, trailer=reader).write()
        continue


''' convert AutoCAD SHX PDF annotation text to searchable text'''
import io,os,glob,json,re,shutil
from PyPDF2 import PdfFileWriter, PdfFileReader
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape, A3
from datetime import date, time, datetime,timedelta




def convert_annot(pdf_name,page_num=0):
    v_id=json.load(open(r'C:\Users\H_Han\Desktop\Python_scripts\PID\PID_LATEST\Index_latest.json','r'))
    PSVRD_TEXT=''
    PSVRD_LIST=[]
    print('processing...:', pdf_name)
    f=open(pdf_name, "rb")
    pdf = PdfFileReader(f,strict=False)
    page = pdf.getPage(page_num) 
    objs=[]
    try:
        for annot in page['/Annots'] :
            obj=annot.getObject()
            if '/Contents' in obj:
                com_dict={'text':obj['/Contents'], 'pos':obj['/Rect']}
                objs.append(com_dict)
                #objs.append({'name'=obj['/Contents'], 'pos'=obj['/Rect']})
                #objs.append(obj)
    except:
        print(pdf_name)
        return 0
        pass
##    for i in objs:
##        print (i)
    page_size=pdf.getPage(0).mediaBox.upperRight
    w, h =page_size
    packet = io.BytesIO()
    c = canvas.Canvas(packet,pagesize=landscape(A3))
    c.setFont('Helvetica-Bold',8)
    c.setFillColor(colors.blue)
    if len(objs)>0:
      
        for i, item in enumerate(objs):
            text=item['text']
            llx,lly,urx,ury=item['pos']   #LowerLeftX,LowerLeftY,UpperRightX, UpperRightY from /Rect list
            x1,y1=int(urx), int(lly)
            #print (i, text)
            text_re=re.search(r'^\d{1}(PSV|RD)',text)
            if text_re is not None and i+1<len(objs):

                v1=objs[i-1]['text'] # previous line text
                v1_re=re.search(r'^[29]\d+',v1)
                llx1,lly1,urx1,ury1=objs[i-1]['pos']   # previous line position
                v2=objs[i+1]['text'] # next line text
                v2_re=re.search(r'^[29]\d+',v2)
                llx2,lly2,urx2,ury2=objs[i+1]['pos']    # next line position
                print (text,v1,v2)
                print (v1_re,v2_re)
                if v1_re is not None and v2_re is None:
                    PSVRD_TEXT=text+v1
                elif v1_re is None and v2_re is not None:
                    PSVRD_TEXT=text+v2
                elif v1_re is not None and v2_re is not None:
                    if (abs(llx-llx2)+abs(lly-lly2))> (abs(llx-llx1)+abs(lly-lly1)):
                        PSVRD_TEXT=text+v1
                    else:
                        PSVRD_TEXT=text+v2
                else:
                    pass

                if len(PSVRD_TEXT)>3:
                    PSVRD_LIST.append(PSVRD_TEXT) 
                    print (PSVRD_TEXT) 
                    c.saveState()
                    c.translate(x1,y1)
                    c.rotate(-90)
                    #c.setFont('Helvetica-Bold',6)
                    c.drawString(-25,-10,PSVRD_TEXT)
                    c.restoreState()
                    PSVRD_TEXT=''
            text=text.replace('%%U','')
            text=text.replace('-','')
            text=text.replace('%%u','')
            m=re.search(r'(?P<v_id>^[BDEFRSPX]{1,3})\d{3,5}\w{0,4}\d{0,3}',text)
#            if m is not None  and len(text)<9 and text[:4] in [x[:4] for x in list(v_id.keys())] :
            if m is not None  and len(text)<9 and (90<lly<1110) and (125<urx<790):
                
                c.saveState()
                c.translate(x1,y1)
                c.rotate(-90)
                if m.group('v_id')=='X':
                    c.setFont('Helvetica-Bold',6)
                else:
                    c.setFont('Helvetica-Bold',8)
                c.drawString(-25,-10,text)
                c.restoreState()               
##            if text.startswith('KPP') and text.endswith(os.path.splitext(pdf_name)[0].split('-')[-1]): # Only add searchable keywords to PDF
##                c.saveState()
##                c.translate(x1,y1)
##                c.rotate(-90)
##                c.setFont('Helvetica-Bold',8)
##                c.drawString(-25,-10,text)
##                c.restoreState()
        c.saveState()
        c.translate(825,1160)
        c.rotate(-90)
        c.setFont('Helvetica-Bold',8)
        c.setFillColor(colors.red)
        c.drawString(0,0,'FILE NAME: '+pdf_name+"  !!!WARNING: THIS P&ID DRAWING MAY NOT BE THE LATEST ONE, USE THIS FOR SEARCH PURPOSE ONLY!!!")
        c.restoreState()
    else:
        return 0 #  Quit if no AutoCAD SHX Text found in this PDF
    c.save()
    packet.seek(0)     #buffer start from 0
    output = PdfFileWriter() 
    page.mergePage(PdfFileReader(packet).getPage(0))
    output.addPage(page)
    saved_dir=os.path.join(os.path.dirname(__file__),'converted')
    if not os.path.exists(saved_dir):
        os.mkdir(saved_dir)
    new_pdf_file_name=os.path.join(saved_dir,pdf_name)
    #os.path.splitext(pdf_name)[0]+".annot.pdf"     
    outputStream = open(new_pdf_file_name, "wb")     # Finally output new pdf
    output.write(outputStream)
    outputStream.close()
    f.close()
    #os.startfile(new_pdf_file_name,'open')
    if len(PSVRD_LIST)>0:
        print (pdf_name, PSVRD_LIST)
    return 1

if __name__ == '__main__':

  
    pdfs=glob.glob('*.pdf')
    print ('Total PDF numbers:', len(pdfs))
    pdf_unsuccessful=[]
    for pdf in pdfs:
        if convert_annot(pdf)==0:
            pdf_unsuccessful.append(pdf)

    print (pdf_unsuccessful)
