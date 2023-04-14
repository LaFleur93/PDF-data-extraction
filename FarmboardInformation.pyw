import os
from io import StringIO
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from xlsxwriter import *

def FarmboardInformation(date):
    path = os.getcwd() + r'\HarvestLabels\Harvest labels - %s - IGC Copenhagen.pdf' % date

    output_string = StringIO()
    with open(path, 'rb') as in_file:
        parser = PDFParser(in_file)
        doc = PDFDocument(parser)
        rsrcmgr = PDFResourceManager()
        device = TextConverter(rsrcmgr, output_string, laparams=LAParams())
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        for page in PDFPage.create_pages(doc):
            interpreter.process_page(page)

    output = output_string.getvalue()
    l = output.splitlines()

    i = 0
    count = 0
    data = []

    for i in range(len(l)-1):
        if l[i] == l[i+1] == '':
            count += 1
            if count == 1:
                data.append(l[0:i])
                j = i
            else:
                data.append(l[j+2:i])
            j = i

    #print(data)

    varieties = ['Bs49','Bs105','Che9','Cr26','Dl6','Mls5','Mn24','Mn25','Or17','Pea3','Pr21','Pr26','Rs4','Rs6','Tm5','Sg2','WtC2']
    names = {'Bs49':'Italian Basil','Bs105':'Italian Basil','Che9':'Chervil','Cr26':'Flat Coriander','Dl6':'Dill','Mls5':'Melissa','Mn24':'Peppermint','Mn25':'Peppermint','Or17':'Oregano','Pea3':'Pea Tops','Pr21':'Curly Parsley','Pr26':'Flat Parsley','Rs4':'Rosemary','Rs6':'Rosemary','Tm5':'Thyme','Sg2':'Sage','WtC2':'Watercress'}
    special = ['X','E','E2','E3','NiA1','nursery']
    special_names = {'X':'Crop X','E':'Empty tray default','E2':'Empty bench','E3':'Empty pots tray','NiA1':'Nursery 1','nursery':'NiA1'}
    days = ['Mon','Tue','Wed','Thu','Fri','Sat','Sun']
    month = {'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06','Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12'}
    trays_amounts = [str(x) for x in range(1,27)]

    ids = [[] for x in range(len(data))]
    acres = [[] for x in range(len(data))] 
    benches = [[] for x in range(len(data))]
    dates = [[] for x in range(len(data))]
    codes = [[] for x in range(len(data))]
    varieties_names = [[] for x in range(len(data))]
    plants = [[] for x in range(len(data))] 
    trays = [[] for x in range(len(data))]
    groups = [[] for x in range(len(data))]
    weeks = [[] for x in range(len(data))]
    pages = [[] for x in range(len(data))]
    main_data = []
    i = 0

    largos = []

    for page in data:
        
        for x in page:
            
            # Batch ID
            if len(x) == 12 and x[6] == '.':
                ids[i].append(x)

            pages[i].append(i)

            for variety in varieties:
                if variety in x:
                    if variety == 'Mn25':
                        
                        codes[i].append('Mn24')
                        varieties_names[i].append(names[variety])
                    elif variety == 'Bs105':
                        codes[i].append('Bs49')
                        varieties_names[i].append(names[variety])
                    elif variety == 'Rs6':
                        codes[i].append('Rs4')
                        varieties_names[i].append(names[variety])
                    else:
                        codes[i].append(variety)
                        varieties_names[i].append(names[variety])

            for special_char in special:
                if special_char == x:
                    codes[i].append(special_char)
                    varieties_names[i].append(special_names[special_char])

            if x in trays_amounts:
                trays[i].append(int(x))

                if any('T30' in mystring for mystring in page):
                    groups[i].append("default")
                    plants[i].append(int(x)*30)
                if any('T15' in mystring for mystring in page):
                    groups[i].append("pots")
                    plants[i].append(int(x)*15)
                if '1 Trays' in page:
                    groups[i].append("empty")
                    plants[i].append(int(x)*600)

            # Acre and Bench Nº

            if 'Acre ' in x and x[0] == 'A':
                acres[i].append(int(x[5:7]))
                
            if 'Bench ' in x and x[0] == 'B':
                benches[i].append(int(x[6:8]))


        #codes[i] = list(dict.fromkeys(codes[i]))
        i += 1
    
    for i in range(len(data)):
        out = list(zip(ids[i]*27, acres[i]*27, benches[i]*27, codes[i]*27, varieties_names[i]*27, groups[i], plants[i], trays[i], pages[i]*27))
        main_data += out

    workbook = Workbook('HarvestData.xlsx')
    worksheet = workbook.add_worksheet("HarvestData")

    headers = ['Variety', 'Batch ID', 'Date', 'Acre']

    header_format = workbook.add_format({'bold': True,
                                         'bottom': 2,
                                         'bg_color': '#D9D9D9'})

    for col_num, col_text in enumerate(headers):
        worksheet.write(0, col_num, col_text, header_format)

    widths = [11,17,11,11]

    i = 0

    for w in widths:
        worksheet.set_column(i, i, w)
        i += 1

    i = 1

    google_sheet_names = {'Crop X':'Crop_X','Rosemary':'Rosemary','Thyme':'Thyme','Italian Basil':'Italian_Basil','Flat Coriander':'Flat_Coriander','Sage':'Sage','Curly Parsley':'Curly_Parsley','Flat Parsley':'Flat_Parsley','Green Mint':'Green_Mint','Dill':'Dill','Watercress':'Watercress','Pea Tops':'Pea Tops','Oregano':'Oregano','Melissa':'Melissa','Peppermint':'Green_Mint','Chervil':'Chervil','Empty tray default':'Empty tray default','NiA1':'Nursery'}

    for element in main_data:
        if element[5] == 'default':
            worksheet.write(i,0,google_sheet_names[element[4]])
            worksheet.write(i,1,element[0])
            worksheet.write(i,2,f'{int(date[5:7])}/{int(date[8:10])}/{int(date[0:4])}')
            worksheet.write(i,3,element[1])
            i += 1

    workbook.close()

    return main_data