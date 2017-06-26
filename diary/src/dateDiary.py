from PyPDF2 import PdfFileWriter, PdfFileReader
from io import BytesIO 
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A5
from reportlab.lib.utils import ImageReader

# Barcode
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing 
from reportlab.graphics import renderPDF
import datetime, sys, os, glob


import math

participants = 0
pages = 1


# Algorithm to put A5 pages in an A4 booklet
# taken from https://bitbucket.org/spookylukey/booklet-maker/src
    

class Sheet(object):
    def __init__(self):
        self.front = PrintPage()
        self.back = PrintPage()


class PrintPage(object):
    def __init__(self):
        self.left = PageContainer()
        self.right = PageContainer()


class PageContainer(object):
    def __init__(self):
        self.page = None


def build_booklet(pages):
    # Double sized page, with double-sided printing, fits 4 of the original.
    sheet_count = int(math.ceil(len(pages) / 4.0))

    booklet = [Sheet() for i in range(0, sheet_count)]

    # Assign input pages to sheets

    # This is the core algo. To understand it:
    # * pick up 3 A4 sheets, landscape
    # * number the sheets from 1 to 3, starting with bottom one
    # * fold the stack in the middle to form an A5 booklet
    # * work out what order you need to use the front left,
    #   front right, back left and back right sides.

    def containers():
        # Yields parts of the booklet in the order they should be used.
        for sheet in booklet:
            yield sheet.back.right
            yield sheet.front.left

        for sheet in reversed(booklet):
            yield sheet.front.right
            yield sheet.back.left

    for c, p in zip(containers(), pages):
        c.page = p

    return booklet


def add_double_page(writer, page_size, print_page):
    width, height = page_size
    p = writer.insertBlankPage(width=width, height=height, index=writer.getNumPages())

    # Merge the left page
    l_page = print_page.left.page
    if l_page is not None:
        p.mergePage(l_page)

    # Merge the right page with translation
    r_page = print_page.right.page
    if r_page is not None:
        p.mergeTranslatedPage(r_page, width / 2, 0)

def make_booklet(input_name, output_name, blanks=0):
    reader = PdfFileReader(open(input_name, "rb"))
    pages = [reader.getPage(p) for p in range(0, reader.getNumPages())]
    for i in range(0, blanks):
        pages.insert(0, None)

    sheets = build_booklet(pages)

    writer = PdfFileWriter()
    p0 = reader.getPage(0)
    input_width = p0.mediaBox.getWidth()
    output_width = input_width * 2
    input_height = p0.mediaBox.getHeight()
    output_height = input_height

    page_size = (output_width, output_height)
    # We want to group fronts and backs together.
    for sheet in sheets:
        add_double_page(writer, page_size, sheet.back)
        add_double_page(writer, page_size, sheet.front)

    #    for sheet in sheets:

    writer.write(open(output_name, "wb"))


### Algorithm to take a template and mark it with date, participant ID and a QR code.

def main(argv):
    
    full_path = argv[1]
    pages = int(argv[2])
    email_address = argv[3]
    #pdfmetrics.registerFont(TTFont('Calibri', 'Calibri.ttf'))
    
    for subdir, dirs, files in os.walk(full_path):
        for directory in dirs:
            for template in glob.glob(os.path.join(full_path,directory,"diaryTemplate.pdf")):
                
                date = datetime.datetime(2017,5,29,0,0,0)
                output = PdfFileWriter()
                input_path = template
                input1 = PdfFileReader(open(input_path, "rb"))
                
                
                print("Participant " + str(directory) )
                
                #Cover
                packet = BytesIO()
                can = canvas.Canvas(packet, pagesize=A5)
                w, h = A5
                #can.setFont("Calibri", 20)
                # You can add a text legend
                #can.drawCentredString(w/2, 100, "The University of Manchester")
                
                # Centering the logo
                logo = ImageReader(os.path.join(full_path, "logo.png"))
                iw, ih = logo.getSize()
                image_scale_width = 250
                aspect = ih / float(iw)
                image_scale_height=(image_scale_width * aspect)
                can.drawImage(logo, x=(w/2) - (image_scale_width/2), y=(h/3)-(image_scale_height/2), width=image_scale_width,preserveAspectRatio=True,mask='auto')
                
                can.save()
                packet.seek(0)
                new_pdf = PdfFileReader(packet)
                output.addPage(new_pdf.getPage(0))
                output.addBlankPage()
                    
                for page in range(1, pages+1):
                    input1 = PdfFileReader(open(input_path, "rb"))
                    # create the date and page marks with Reportlab
                    packet = BytesIO()
                    can = canvas.Canvas(packet, pagesize=A5)
                    
                    date += datetime.timedelta(days=1)
                    
                    # Barcode
                    barcode_value = date.strftime('%y%m%d') + directory 

                    qr_code = qr.QrCodeWidget(barcode_value)
                    bounds = qr_code.getBounds()
                    width = bounds[2] - bounds[0]
                    height = bounds[3] - bounds[1]
                    d = Drawing(30, 30, transform=[30./width,0,0,30./height,0,0])
                    d.add(qr_code)
                    renderPDF.draw(d, can, 355, 554)
                    
                    # Header
                    #can.setFont("Calibri", 11)
                    #can.drawString(340, 565, "P" + str(directory).zfill(2) )
                    can.drawString(320, 565, directory )
                    #can.setFont("Calibri", 11)
                    can.drawString(36.5, 565, date.strftime('%A, %d %b %Y') )
                    
                    
                    # Footer
                    #can.setFont("Calibri", 8)
                    can.drawString(36.5, 24, str(page))
                    can.save()
                    
                    
                    #move to the beginning of the StringIO buffer
                    packet.seek(0)
                    new_pdf = PdfFileReader(packet)
                    
                    newPage = input1.getPage(0)
                    newPage.mergePage(new_pdf.getPage(0))
                    output.addPage(newPage)
                    
                # Backcover
                output.addBlankPage()
                packet = BytesIO()
                can = canvas.Canvas(packet, pagesize=A5)
                #can.setFont("Calibri", 10)
                can.drawString(75, 50, "If you find this diary, please email " + email_address)
                can.save()
                packet.seek(0)
                new_pdf = PdfFileReader(packet)
                output.addPage(new_pdf.getPage(0))
                
                # finally, write "output" to document-output.pdf
                #outputPath = "C://Users//julio//Documents//phd//study//collection//visits//diary//sk0" + str(participant) + "//diary" + str(participant) + ".pdf"
                outputPath = os.path.join(full_path, directory, directory + "-a5-single-sided.pdf")
                #outputBookletPath = "C://Users//julio//Documents//phd//study//collection//visits//diary//diariesBooklets//diaryBooklet" + str(participant) + ".pdf"
                outputBookletPath = os.path.join(full_path, directory, directory + "-a4-double-sided.pdf")
                outputStream = open(outputPath, "wb")
                output.write(outputStream)
                
                outputStream.close()
                
                make_booklet(outputPath, outputBookletPath, 0)


if __name__ == "__main__":
    main(sys.argv)