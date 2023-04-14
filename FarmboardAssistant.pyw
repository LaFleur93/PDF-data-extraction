import os
from tkinter import *
from tkinter import ttk
from tkcalendar import Calendar
from datetime import datetime
from FarmboardInformation import FarmboardInformation
from missing_and_extra_trays import missed_and_extra_planting
from PyPDF2 import PdfReader, PdfWriter
from SubLabelGenerator import LabelGenerator
import babel.numbers
import webbrowser

class FarmboardAssistant:
    def __init__(self, window):
        # Initializations 
        self.wind = window
        self.wind.title('Farmboard Assistant')
        self.wind.geometry('865x615')
        self.wind.resizable(False, False)

        self.date_selected = False

        #-------------------------------------
        # Left column (messages, calendar and general overview)
        #-------------------------------------
        frame_column_1 = LabelFrame(self.wind, borderwidth = '0')
        frame_column_1.place(x=10, y=0, width=280, relheight=1)

        #-------------------------------------
        # Message frame
        #-------------------------------------
        frame_msg = LabelFrame(frame_column_1, text = 'Message')
        frame_msg.place(x=0, y=0, relwidth=.99, height=55)

        self.message = Label(frame_msg, text = 'No message')
        self.message.place(x = 10, y = 4)

        #-------------------------------------
        # Calendar
        #-------------------------------------
        frame_cal = LabelFrame(frame_column_1, text = 'Select date', borderwidth=2)
        frame_cal.place(x=0, y=55, relwidth=.99, height=260)

        # Current date for calendar default
        today = datetime.now()
        day = today.day
        month = today.month
        year = today.year

        self.calendar = Calendar(frame_cal, selectmode = 'day', year = year, month = month, day = day, date_pattern="y-mm-dd")
        self.calendar.place(x = 10, y = 5)

        Button(frame_cal, text = 'Load Harvest Labels', relief="solid", fg='black', bg = '#D3D3D3', 
            font=("Arial",8,"bold"), bd = 1, width = 34, height = 2, command = self.get_date).place(x = 12, y = 195)

        #-------------------------------------
        # Information
        #-------------------------------------
        frame_info = LabelFrame(frame_column_1, text = 'Information', borderwidth=2)
        frame_info.place(x=0, y=315, relwidth=.99, height=74)

        Label(frame_info, text = '1. Select date and load Harvest Labels data.').place(x=5, y=4)
        Label(frame_info, text = '2. Handle information.').place(x=5, y=24) 

        #-------------------------------------
        # Useful links
        #-------------------------------------
        frame_links = LabelFrame(frame_column_1, text = 'Useful links', borderwidth=2)
        frame_links.place(x=0, y=390, relwidth=.99, height=218)

        Label(frame_links, text = 'Farmboard IGC Copenhagen').place(x=5, y=4)
        Label(frame_links, text = '•').place(x=5, y=24)
        link1 = Label(frame_links, text = 'https://farmboard.infarm.com/', fg = 'blue', cursor="hand2")
        link1.place(x=15, y=24)
        link1.bind("<Button-1>", lambda e: self.callback('https://farmboard.infarm.com/locations/f1903b06-fe2c-4131-bfe9-6de379508289'))

        Label(frame_links, text = 'Packing List IGC Copenhagen').place(x=5, y=44)
        Label(frame_links, text = '•').place(x=5, y=64)
        link2 = Label(frame_links, text = 'https://docs.google.com/spreadsheets/', fg = 'blue', cursor="hand2")
        link2.place(x=15, y=64)
        link2.bind("<Button-1>", lambda e: self.callback('https://docs.google.com/spreadsheets/d/1NZjPF6jYjNCZd6MzCOZDXrCCOxwqRdMFC2URcMgwk68/edit#gid=1315992959'))

        Label(frame_links, text = 'IGC Copenhagen missed and extra planting').place(x=5, y=84)
        Label(frame_links, text = '•').place(x=5, y=104)
        link3 = Label(frame_links, text = 'https://docs.google.com/spreadsheets/', fg = 'blue', cursor="hand2")
        link3.place(x=15, y=104)
        link3.bind("<Button-1>", lambda e: self.callback('https://docs.google.com/spreadsheets/d/1azyJv6WsCrQ5mQ2Z0V9R_JQOWz8wtzWi-cro2nlcYuw/edit#gid=0'))

        #-------------------------------------
        # Extracting information
        #-------------------------------------
        center_col_width = 250

        frame_extract = LabelFrame(self.wind, text = 'Extract only selected varieties')
        frame_extract.place(x=294, y=0, width=center_col_width, height=462)

        all_varieties = ['Italian Basil', 'Flat Coriander', 'Thyme', 'Peppermint', 'Rosemary', 'Flat Parsley', 'Curly Parsley', 'Sage', 'Dill', 'Chervil', 'Watercress', 'Melissa', 'Oregano', 'Pea Tops']

        for i in range(len(all_varieties)):
            Label(frame_extract, text = all_varieties[i]+':').place(x = 20, y = 1+25*(i))

        Label(frame_extract, text = '.............................................................................', fg = '#CCCCCC').place(x = 5, y = 343)

        Label(frame_extract, text = 'Special:').place(x = 20, y = 365)

        # both
        self.h_bs49 = StringVar(value=0)
        self.p_bs49 = StringVar(value=0)
        self.h_cr26 = StringVar(value=0)
        self.p_cr26 = StringVar(value=0)
        self.h_tm5 = StringVar(value=0)
        self.p_tm5 = StringVar(value=0)
        self.h_mn24 = StringVar(value=0)
        self.p_mn24 = StringVar(value=0)
        self.h_rs4 = StringVar(value=0)
        self.p_rs4 = StringVar(value=0)
        self.h_pr26 = StringVar(value=0)
        self.p_pr26 = StringVar(value=0)
        self.h_pr21 = StringVar(value=0)
        # herbs only
        self.h_sg2 = StringVar(value=0)
        self.h_dl6 = StringVar(value=0)
        # pots only
        self.p_che9 = StringVar(value=0)
        self.p_wtc2 = StringVar(value=0)
        self.p_mls5 = StringVar(value=0)
        self.p_or17 = StringVar(value=0)
        self.p_pea3 = StringVar(value=0)
        # special
        self.empty_special = StringVar(value=0)
        self.nia1_special = StringVar(value=0)

        h1 = Checkbutton(frame_extract, text = 'Herbs', variable = self.h_bs49, onvalue = 'H Bs49', offvalue = '0').place(x = 110, y = 1+25*(0))
        h2 = Checkbutton(frame_extract, text = 'Herbs', variable = self.h_cr26, onvalue = 'H Cr26', offvalue = '0').place(x = 110, y = 1+25*(1))
        h3 = Checkbutton(frame_extract, text = 'Herbs', variable = self.h_tm5, onvalue = 'H Tm5', offvalue = '0').place(x = 110, y = 1+25*(2))
        h4 = Checkbutton(frame_extract, text = 'Herbs', variable = self.h_mn24, onvalue = 'H Mn24', offvalue = '0').place(x = 110, y = 1+25*(3))
        h5 = Checkbutton(frame_extract, text = 'Herbs', variable = self.h_rs4, onvalue = 'H Rs4', offvalue = '0').place(x = 110, y = 1+25*(4))
        h6 = Checkbutton(frame_extract, text = 'Herbs', variable = self.h_pr26, onvalue = 'H Pr26', offvalue = '0').place(x = 110, y = 1+25*(5))
        h7 = Checkbutton(frame_extract, text = 'Herbs', variable = self.h_pr21, onvalue = 'H Pr21', offvalue = '0').place(x = 110, y = 1+25*(6))
        h8 = Checkbutton(frame_extract, text = 'Herbs', variable = self.h_sg2, onvalue = 'H Sg2', offvalue = '0').place(x = 110, y = 1+25*(7))
        h9 = Checkbutton(frame_extract, text = 'Herbs', variable = self.h_dl6, onvalue = 'H Dl6', offvalue = '0').place(x = 110, y = 1+25*(8))

        p1 = Checkbutton(frame_extract, text = 'Pots', variable = self.p_bs49, onvalue = 'P Bs49', offvalue = '0').place(x = 170, y = 1+25*(0))
        p2 = Checkbutton(frame_extract, text = 'Pots', variable = self.p_cr26, onvalue = 'P Cr26', offvalue = '0').place(x = 170, y = 1+25*(1))
        p3 = Checkbutton(frame_extract, text = 'Pots', variable = self.p_tm5, onvalue = 'P Tm5', offvalue = '0').place(x = 170, y = 1+25*(2))
        p4 = Checkbutton(frame_extract, text = 'Pots', variable = self.p_mn24, onvalue = 'P Mn24', offvalue = '0').place(x = 170, y = 1+25*(3))
        p5 = Checkbutton(frame_extract, text = 'Pots', variable = self.p_rs4, onvalue = 'P Rs4', offvalue = '0').place(x = 170, y = 1+25*(4))
        p6 = Checkbutton(frame_extract, text = 'Pots', variable = self.p_pr26, onvalue = 'P Pr26', offvalue = '0').place(x = 170, y = 1+25*(5))
        p7 = Checkbutton(frame_extract, text = 'Pots', variable = self.p_che9, onvalue = 'P Che9', offvalue = '0').place(x = 170, y = 1+25*(9))
        p8 = Checkbutton(frame_extract, text = 'Pots', variable = self.p_wtc2, onvalue = 'P WtC2', offvalue = '0').place(x = 170, y = 1+25*(10))
        p9 = Checkbutton(frame_extract, text = 'Pots', variable = self.p_mls5, onvalue = 'P Mls5', offvalue = '0').place(x = 170, y = 1+25*(11))
        p10 = Checkbutton(frame_extract, text = 'Pots', variable = self.p_or17, onvalue = 'P Or17', offvalue = '0').place(x = 170, y = 1+25*(12))
        p11 = Checkbutton(frame_extract, text = 'Pots', variable = self.p_pea3, onvalue = 'P Pea3', offvalue = '0').place(x = 170, y = 1+25*(13))

        s1 = Checkbutton(frame_extract, text = 'Empty benches', variable = self.empty_special, onvalue = 'E E2', offvalue = '0').place(x = 80, y = 365)
        s2 = Checkbutton(frame_extract, text = 'Nia1', variable = self.nia1_special, onvalue = 'E nursery', offvalue = '0').place(x = 190, y = 365)

        Button(frame_extract, text = 'Extract selected varieties', relief="solid", fg='black', bg = '#D3D3D3', 
            font=("Arial",8,"bold"), bd = 1, width = 31, height = 2, command = self.custom_varieties).place(x = 10, y = 397)

        #-------------------------------------
        # Open Farmboard file
        #-------------------------------------
        frame_farmboard_file = LabelFrame(self.wind, text = 'Open all benches (Farmboard file)')
        frame_farmboard_file.place(x=294, y=465, width=center_col_width, height=70)

        Button(frame_farmboard_file, text = 'Get PDF', relief="solid", fg='black', bg = '#D3D3D3', 
            font=("Arial",8,"bold"), bd = 1, width = 31, height = 2, command = self.open_labels).place(x = 10, y = 5)

        #-------------------------------------
        # Open Order by Acre number
        #-------------------------------------
        frame_acre_order = LabelFrame(self.wind, text = 'Order all benches per Acre number')
        frame_acre_order.place(x=294, y=538, width=center_col_width, height=70)

        Button(frame_acre_order, text = 'Get PDF', relief="solid", fg='black', bg = '#D3D3D3', 
            font=("Arial",8,"bold"), bd = 1, width = 31, height = 2, command = self.order_per_acre).place(x = 10, y = 5)

        #-------------------------------------
        # General overview frame
        #-------------------------------------
        self.frame_general = LabelFrame(self.wind, border = 0,)
        self.frame_general.place(x=550, y=0, relwidth=.99, relheight=1)

        self.col_width = 5
        self.col_height = 10
        self.herbs_col_start = 5
        self.pots_col_start = 5
        self.pots_height_start = 240
        self.special_height_start = 530

        # Column headers
        
        Label(self.frame_general, text = 'Herbs / Trays', width=17).place(x = self.herbs_col_start, y = self.col_height)
        Label(self.frame_general, text = 'Pots / Trays', width=17).place(x = self.pots_col_start, y = self.pots_height_start + self.col_height)
        Label(self.frame_general, text = 'Special / Benches', width=17).place(x = self.herbs_col_start, y = self.special_height_start + self.col_height)
        
        headers = ['FB','Miss','Extra','Total']
        for i in range(len(headers)):
            Label(self.frame_general, text = headers[i], bg='#CCCCCC', width= self.col_width).place(x = (127+self.herbs_col_start)+43*i, y = self.col_height)
            Label(self.frame_general, text = headers[i], bg='#CCCCCC', width= self.col_width).place(x = (127+self.pots_col_start)+43*i, y = self.pots_height_start + self.col_height)

        Label(self.frame_general, text = 'Bench(es)', bg='#CCCCCC', width= 10).place(x = 132, y = self.special_height_start + self.col_height)
            
        self.herb_varieties = ['Herbs - Flat Parsley', 'Herbs - Curled Parsley', 'Herbs - Thyme', 'Herbs - Flat Coriander', 'Herbs - Rosemary',
        'Herbs - Green Mint', 'Herbs - Sage', 'Herbs - Italian Basil', 'Herbs - Dill']

        self.pot_varieties = ['Pots - Italian Basil', 'Pots - Flat Coriander', 'Pots - Thyme', 'Pots - Peppermint', 'Pots - Rosemary', 'Pots - Chervil',
        'Pots - Watercress', 'Pots - Melissa', 'Pots - Oregano', 'Pots - Flat Parsley', 'Pots - Pea Tops']

        special_headers = ['Empty benches', 'Nursery benches']

        for i in range(2):
            Label(self.frame_general, text = special_headers[i], bg='#CCCCCC', width=17).place(x = 5, y = self.special_height_start + self.col_height+(1+i)*23)

        self.clear_columns()

    def get_date(self):
        try:
            self.clear_columns()
            self.date = self.calendar.get_date()

            self.main_data = FarmboardInformation(self.date)
            missing_and_extra = missed_and_extra_planting(self.date)

            missing_herbs = missing_and_extra[0]
            missing_pots = missing_and_extra[1]
            extra_herbs = missing_and_extra[2]
            extra_pots = missing_and_extra[3]

            herbs = ['Pr26', 'Pr21', 'Tm5', 'Cr26', 'Rs4', 'Mn24', 'Sg2', 'Bs49', 'Dl6']
            herbs_trays = [[] for x in range(len(herbs))]

            for i in range(len(herbs)):
                for row in self.main_data:
                    if row[3] == herbs[i] and row[5] == 'default':
                        herbs_trays[i].append(row[7])

            herbs_trays = [sum(x) for x in herbs_trays]

            for i in range(len(self.herb_varieties)):
                total_herbs = 0
                total_herbs += herbs_trays[i]
                Label(self.frame_general, text = str(herbs_trays[i]), bg='#FFFFFF', width=5).place(x = (127+self.herbs_col_start), y = self.col_height+(i+1)*23)
                if missing_herbs[herbs[i].lower()] != 0:
                    Label(self.frame_general, text = missing_herbs[herbs[i].lower()], bg='#FFFFFF', fg = 'red', width=5).place(x = (127+self.herbs_col_start)+43*1, y = self.col_height+(i+1)*23)
                    total_herbs -= int(missing_herbs[herbs[i].lower()])
                if extra_herbs[herbs[i].lower()] != 0:    
                    Label(self.frame_general, text = extra_herbs[herbs[i].lower()], bg='#FFFFFF', fg = 'green', width=5).place(x = (127+self.herbs_col_start)+43*2, y = self.col_height+(i+1)*23)   
                    total_herbs += extra_herbs[herbs[i].lower()]
                Label(self.frame_general, text = str(total_herbs), bg='#E5E5E5', width=5).place(x = (127+self.herbs_col_start)+43*3, y = self.col_height+(i+1)*23)

            pots = ['Bs49', 'Cr26', 'Tm5', 'Mn24', 'Rs4', 'Che9', 'WtC2', 'Mls5', 'Or17', 'Pr26', 'Pea3']
            pots_trays = [[] for x in range(len(pots))]

            for i in range(len(pots)):
                for row in self.main_data:
                    if row[3] == pots[i] and row[5] == 'pots':
                        pots_trays[i].append(row[7])

            pots_trays = [sum(x) for x in pots_trays]

            for i in range(len(self.pot_varieties)):
                total_pots = 0
                total_pots += pots_trays[i]
                Label(self.frame_general, text = str(pots_trays[i]), bg='#FFFFFF', width=5).place(x = (127+self.pots_col_start), y = self.pots_height_start + self.col_height+(i+1)*23)
                if missing_pots[pots[i].lower()] != 0:
                    Label(self.frame_general, text = missing_pots[pots[i].lower()], bg='#FFFFFF', fg = 'red', width=5).place(x = (127+self.pots_col_start)+43*1, y = self.pots_height_start + self.col_height+(i+1)*23)
                    total_pots -= int(missing_pots[pots[i].lower()])
                if extra_pots[pots[i].lower()] != 0:    
                    Label(self.frame_general, text = extra_pots[pots[i].lower()], bg='#FFFFFF', fg = 'green', width=5).place(x = (127+self.pots_col_start)+43*2, y = self.pots_height_start + self.col_height+(i+1)*23)
                    total_pots += int(extra_pots[pots[i].lower()])
                Label(self.frame_general, text = str(total_pots), bg='#E5E5E5', width=5).place(x = (127+self.pots_col_start)+43*3, y = self.pots_height_start + self.col_height+(i+1)*23)

            empty_benches = 0
            nursery_benches = 0

            for row in self.main_data:
                if row[3] == 'nursery' or row[3] == 'NiA1':
                    nursery_benches += 1
                elif row[3] == 'E2':
                    empty_benches += 1

            specials = [empty_benches, nursery_benches]

            for i in range(len(specials)):
                Label(self.frame_general, text = str(specials[i]), bg='#FFFFFF', width=10).place(x = 132, y = self.special_height_start + self.col_height+(1+i)*23)

            self.date_selected = True
            self.message['text'] = f'Information loaded. Date: {self.date}'
            self.message['fg'] = 'green'
            self.message['font'] = ('Helvetica', 8, 'bold')

            self.clear_checkboxes()
        
        except:
            self.message['text'] = f'Harvest Labels not available for that date'
            self.message['fg'] = 'red'
            self.message['font'] = ('Helvetica', 8, 'bold')
        
    def open_labels(self):
        if self.date_selected == True:
            try:
                pdf_file_path = os.getcwd() + r'\HarvestLabels\Harvest labels - %s - IGC Copenhagen.pdf' % self.date
                os.startfile(pdf_file_path)

                self.message['text'] = f'File opened. Date: {self.date}'
                self.message['fg'] = 'green'
                self.message['font'] = ('Helvetica', 8, 'bold')
            except:
                self.message['text'] = "Please, close PDF file"
                self.message['fg'] = 'red'
                self.message['font'] = ('Helvetica', 8, 'bold')
        else:
            self.message['text'] = "Please, select date and load Harvest Labels"
            self.message['fg'] = 'red'
            self.message['font'] = ('Helvetica', 8, 'bold')

    def order_per_acre(self):
        if self.date_selected == True:
            try:    
                sorted_main_data = sorted(self.main_data, key = lambda x: x[1])
                sorted_pdf_page_list = [x[8] for x in sorted_main_data]
                sorted_pdf_page_list = list(dict.fromkeys(sorted_pdf_page_list))

                LabelGenerator(self.date, sorted_pdf_page_list)
            
                self.message['text'] = f'File opened. Date: {self.date}'
                self.message['fg'] = 'green'
                self.message['font'] = ('Helvetica', 8, 'bold')
            except:
                self.message['text'] = "Please, close PDF file"
                self.message['fg'] = 'red'
                self.message['font'] = ('Helvetica', 8, 'bold')

        else:
            self.message['text'] = "Please, select date and load Harvest Labels"
            self.message['fg'] = 'red'
            self.message['font'] = ('Helvetica', 8, 'bold')

    def custom_varieties(self):
        if self.date_selected == True:
            try:
                all_checkbuttons = [self.h_bs49.get(), self.p_bs49.get(), self.h_cr26.get(), self.p_cr26.get(), self.h_tm5.get(), self.p_tm5.get(), self.h_mn24.get(), self.p_mn24.get(), self.h_rs4.get(), self.p_rs4.get(), self.h_pr26.get(), self.p_pr26.get(), self.h_pr21.get(), self.h_sg2.get(), self.h_dl6.get(), self.p_che9.get(), self.p_wtc2.get(), self.p_mls5.get(), self.p_or17.get(), self.p_pea3.get(), self.empty_special.get(), self.nia1_special.get()]

                search_for = [[] for x in range(len(all_checkbuttons))]
                herbs_pots = {'H':'default','P':'pots','E':'empty'}

                for i in range(len(all_checkbuttons)):
                    if all_checkbuttons[i] != '0':
                        search_for[i].append(all_checkbuttons[i][2:])
                        search_for[i].append(herbs_pots[all_checkbuttons[i][0]])

                search_for = [x for x in search_for if x != []]
                
                for i in range(len(search_for)):
                    if search_for[i][0] == 'Nia1':
                        search_for[i][0] = 'nursery'
                custom_pdf_pages = []

                for i in range(len(search_for)):
                    for row in self.main_data:
                        if row[3] == search_for[i][0] and row[5] == search_for[i][1]:
                            custom_pdf_pages.append(row[8])

                custom_pdf_pages = list(dict.fromkeys(custom_pdf_pages))

                if LabelGenerator(self.date, custom_pdf_pages) == 'EmptyList':
                    self.message['text'] = f'No batches for that variety'
                    self.message['fg'] = 'black'
                    self.message['font'] = ('Helvetica', 8, 'bold')

                else:
                    self.message['text'] = f'File opened. Date: {self.date}'
                    self.message['fg'] = 'green'
                    self.message['font'] = ('Helvetica', 8, 'bold')

            except:
                self.message['text'] = "Please, close PDF file"
                self.message['fg'] = 'red'
                self.message['font'] = ('Helvetica', 8, 'bold')

        else:
            self.message['text'] = "Please, select date and load Harvest Labels"
            self.message['fg'] = 'red'
            self.message['font'] = ('Helvetica', 8, 'bold')

    def clear_columns(self):

        for i in range(len(self.herb_varieties)):
            Label(self.frame_general, text = self.herb_varieties[i], bg='#CCCCCC', width=17).place(x = 5, y = self.col_height+(i+1)*23)
            for j in range(4):
                Label(self.frame_general, text = '', bg='#FFFFFF', width=5).place(x = ((127+self.herbs_col_start)+43*j), y = self.col_height+(i+1)*23)

        for i in range(len(self.pot_varieties)):
            Label(self.frame_general, text = self.pot_varieties[i], bg='#CCCCCC', width=17).place(x = self.pots_col_start, y = self.pots_height_start + self.col_height+(i+1)*23)
            for j in range(4):
                Label(self.frame_general, text = '', bg='#FFFFFF', width=5).place(x = ((127+self.pots_col_start)+43*j), y = self.pots_height_start + self.col_height+(i+1)*23)

        for i in range(2):
            Label(self.frame_general, text = '', bg='#FFFFFF', width=10).place(x = 132, y = self.col_height+(24+i)*23)

    def clear_checkboxes(self):
        self.h_bs49.set(0)
        self.p_bs49.set(0)
        self.h_cr26.set(0)
        self.p_cr26.set(0)
        self.h_tm5.set(0)
        self.p_tm5.set(0)
        self.h_mn24.set(0)
        self.p_mn24.set(0)
        self.h_rs4.set(0)
        self.p_rs4.set(0)
        self.h_pr26.set(0)
        self.p_pr26.set(0)
        self.h_pr21.set(0)
        self.h_sg2.set(0)
        self.h_dl6.set(0)
        self.p_che9.set(0)
        self.p_wtc2.set(0)
        self.p_mls5.set(0)
        self.p_or17.set(0)
        self.p_pea3.set(0)
        self.empty_special.set(0)
        self.nia1_special.set(0)

    def callback(self, url):
        webbrowser.open_new_tab(url)

if __name__ == '__main__':
    window = Tk()
    application = FarmboardAssistant(window)
    window.mainloop()