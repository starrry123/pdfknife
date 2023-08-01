import PyPDF2,io,glob,os
from reportlab.lib.pagesizes import letter,A4
from reportlab.pdfgen import canvas
from reportlab.lib.colors import blue, red, Color,green


def get_page_origin(pdf_reader, page_number):
    # Get the page object from PyPDF2
    page = pdf_reader.getPage(page_number)

    # Get the media box (page size) and extract the bottom-left and top-left coordinates
    media_box = page.mediaBox
    bottom_left = media_box.lowerLeft
    top_left = media_box.upperLeft

    # Return True if the origin is at the bottom left, False if it's at the top left
    return bottom_left == (0, 0)

def add_text_to_pdf(input_pdf, output_pdf, text_to_insert):

    def my_merge_page(inputStream, packet):
        packet.seek(0)
        for page_i in range(inputStream.numPages):
            pdf_buffer = PyPDF2.PdfFileReader(packet)
            page=inputStream.getPage(page_i)
            page.mergePage(pdf_buffer.getPage(page_i))
            output.addPage(page)
        output.write(outputStream)

    packet = io.BytesIO()
    # Create a new PDF canvas using ReportLab
    canv = canvas.Canvas(packet, pagesize=A4)

    # Open the existing PDF using PyPDF2
    pdf_reader = PyPDF2.PdfFileReader(open(input_pdf, 'rb'))
    pdf_writer = PyPDF2.PdfFileWriter()

    # Loop through all pages of the input PDF
    for page_number in range(pdf_reader.getNumPages()):
        # Get the origin of the current page
        is_bottom_left_origin = get_page_origin(pdf_reader, page_number)

        # Get the existing PDF page
        pdf_page = pdf_reader.getPage(page_number)

        # Get the page size
        page_width = float(pdf_page.mediaBox.getWidth())
        page_height = float(pdf_page.mediaBox.getHeight())

        # Draw the existing PDF content onto the new ReportLab canvas
        canv.setPageSize((page_width, page_height))
 
        # Set the font and font size for the text to be inserted
        canv.setFont("Helvetica", 12)

        # Calculate the top-left position for the text with a 10px offset
        if is_bottom_left_origin:
            text_x = 10
            text_y = page_height - 10
        else:
            text_x = 10
            text_y = 10

        # Draw the text at the calculated position
        canv.setFillColor(red)
        canv.drawString(text_x, text_y, text_to_insert)
        # Save the modified PDF page
        canv.showPage()
    # Save the entire PDF with the inserted text
    canv.save()

    outputStream = open(output_pdf, "wb")
    output = PyPDF2.PdfFileWriter()

    if not input_pdf:
        outputStream.write(packet.getvalue())
    else:  
         my_merge_page(pdf_reader, packet)
    outputStream.close()


if __name__ == "__main__":
    for input_pdf_file in glob.glob(r'C:\Users\Haitao\Downloads\*.pdf'):
        output_pdf_file = os.path.join(r'C:\Users\Haitao\Desktop\test',os.path.basename(input_pdf_file))
        text_to_insert = "This text will appear at the top-left offset by 10px."
        print(input_pdf_file)
        add_text_to_pdf(input_pdf_file, output_pdf_file, text_to_insert)
