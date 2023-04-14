from PyPDF2 import PdfReader, PdfWriter
import os

def LabelGenerator(date, pages_list):
	if pages_list == []:
		return 'EmptyList'
	else:
		pdf_file_path = os.getcwd() + r'\HarvestLabels\Harvest labels - %s - IGC Copenhagen.pdf' % date
		file_base_name = f'{date} - Custom Farmboard'

		pdf = PdfReader(pdf_file_path)

		pdfWriter = PdfWriter()

		for page_num in pages_list:
		    pdfWriter.add_page(pdf.pages[page_num])

		file_name = os.getcwd() + r'\HarvestLabels\{0}_subset.pdf'.format(file_base_name)

		with open(file_name, 'wb') as f:
		    pdfWriter.write(f)

		os.startfile(file_name)