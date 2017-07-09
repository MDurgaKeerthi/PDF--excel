#reads from excel and writes to the given pdf and outputs a newpdf

from xlrd import open_workbook
from pyPdf import PdfFileWriter, PdfFileReader
import StringIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4

packet = StringIO.StringIO()
# create a new PDF with Reportlab
can = canvas.Canvas(packet,pagesize=A4)
can.setFont("Helvetica", 10) 
can.drawString(235, 535, "Head")
can.setFont("Helvetica", 8) 

excelfile = raw_input("Enter the excel file name to read from:")
wb = open_workbook(excelfile)
for s in wb.sheets():
    j = 535
    for row in range(2,s.nrows):
        for col in range(s.ncols):
            if col == s.ncols-1:
               value  = (s.cell(row,col).value)
               #print value
        j = j - 11.9
        can.drawString(245, j, str(value))    
        
can.save()

#move to the beginning of the StringIO buffer
packet.seek(0)
new_pdf = PdfFileReader(packet)

# read your existing PDF
pdfname = raw_input("Enter the pdf name to write to:")
existing_pdf = PdfFileReader(file(pdfname, "rb"))
output = PdfFileWriter()

# add the "watermark" (which is the new pdf) on the existing page
page = existing_pdf.getPage(0)
page.mergePage(new_pdf.getPage(0))
output.addPage(page)

page = existing_pdf.getPage(1)
page.mergePage(new_pdf.getPage(0))
output.addPage(page)

# finally, write "output" to a real file
outpdfname = raw_input("Enter name of pdf to be created:")
outputStream = file(outpdfname, "wb")
output.write(outputStream)
outputStream.close()

