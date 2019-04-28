#!/usr/bin/env python2
# -*- coding: utf-8 -*-
"""
Created on Sun Apr 28 16:20:06 2019

@author: erkamozturk
"""

from Tkinter import *
import docclass
from selenium import webdriver
import tkFileDialog
from xlrd import open_workbook,cellname
import time

class GuessGrade(Frame):

    def __init__(self, root):
        Frame.__init__(self, root)
        self.root = root
        self.widgets()
        self.geometricDesign()
        self.electives = {}

    def widgets(self):

        self.frame1 = Frame(self.root)
        self.frame2 = Frame(self.root)
        self.frame3 = Frame(self.root)
        self.title = Label(self.frame1, text="Guess My Grade! v1.0", font="times 15 bold ",
                           bg="black", fg="white", width=75,height=1)
        self.please = Label(self.frame1,text="Please upload your curriculum file with the grades:",font=("Verdana",12),fg="blue")
        self.browse = Button(self.frame1,text="Browse",bg="red",width=15,command=self.browse,font=("Verdana",12),fg="white")
        self.enter = Label(self.frame1,text="Enter urls for course descriptions",font=("Verdana",12))
        self.entry = Text(self.frame1,bg="white",width=60, height=5)
        # self.entry.delete(0, END)

        self.entry.insert('1.0', "http://www.sehir.edu.tr/en/Pages/Academic/Bolum.aspx?BID=12"
                                 "\nhttp://www.sehir.edu.tr/en/Pages/Academic/Bolum.aspx?BID=32"
                                 "\nhttp://www.sehir.edu.tr/en/Pages/Academic/Bolum.aspx?BID=13"
                                 "\nhttp://www.sehir.edu.tr/en/Pages/Academic/Bolum.aspx?BID=14")
        self.key = Label(self.frame1,text="Key:",font=("Verdana bold",12))
        self.A = Label(self.frame2,text="A",bg="green",width=10,height=1,fg="white",font="Times 16")
        self.B = Label(self.frame2,text="B",bg="lightgreen",width=10,height=1,fg="white",font="Times 16")
        self.C = Label(self.frame2,text="C",bg="yellow",width=10,height=1,fg="white",font="Times 16")
        self.D = Label(self.frame2,text="D",bg="red",width=10,height=1,fg="white",font="Times 16")
        self.F = Label(self.frame2,text="F",bg="black",width=10,height=1,fg="white",font="Times 16")
        self.predict = Button(self.frame2,text="Predict Grades",bg="red",width=15,font=("Verdana",12),fg="white", command
                            = self.fetch)
        self.dots = Label(self.frame2,text="."*270)
        self.predicted = Label(self.frame3,text="Predicted Grades",font=("Verdana",12))
        self.scrollbar = Scrollbar(self.frame3)
        self.text = Text(self.frame3, bg="white", height=16, yscrollcommand=self.scrollbar.set)
        self.scrollbar.config(command=self.text.yview)
        self.file_opt = options = {}  # settings of browse
        options['defaultextension'] = '.xlsx'
        options['filetypes'] = [('all files', '.*'), ('excel files', '.xlsx')]
        options['initialdir'] = '/Users/erkamozturk/Desktop/licence/IV.Donem/ENGR 212/MINI PROJECTS/6/MP6'
        options['initialfile'] = 'cs.xlsx'
        options['parent'] = self.root
        options['title'] = 'Choose a file'

    def browse(self):

        try:
            self.allSemesters = ['Semester I', 'Semester II','Semester III', 'Semester IV', 'Semester V', 'Semester VI',
                                 'Semester VII','Semester VIII']  # all semesters
            selected_file = tkFileDialog.askopenfile(mode='r', **self.file_opt)  # select one
            self.fileName = selected_file.name  # name of selected
            self.save_to_db()  # it is for get info and fill it in dicts
        except:
            None


    def save_to_db(self):

        curriculum = open_workbook(self.fileName)  # go into excel
        sheet = curriculum.sheet_by_index(0)  # sheet 0th
        self.taken = {}  # and create the dict
        self.willTake = {}  # and create the dict

        allSemesters = ['Semester I', 'Semester II', 'Semester III', 'Semester IV', 'Semester V', 'Semester VI',
                        'Semester VII', 'Semester VIII']
        for col in range(sheet.ncols): # all cols
            for row in range(sheet.nrows): # all rows
                if sheet.cell(row, col).value in allSemesters: # if in any semester
                    # print sheet.cell(row, col).value, col, row
                    semester = sheet.cell(row, col).value # get value
                    row_read = row + 2 # go 2 line down to first course
                    for i in range(9):  # nine is the biggest number student can take in one semester
                        if len(sheet.cell(row_read, col).value) == 0: # if course not, go other
                            continue
                        code = sheet.cell(row_read, col).value  # get code
                        title = sheet.cell(row_read, col + 1).value  # get title
                        letterGrade = sheet.cell(row_read, col + 6).value  # get grade
                        row_read += 1  # go up ecery each line
                        if len(letterGrade) == 0:
                            self.willTake[code] = {'Title':title, 'Semester':semester}  # it means it didnt take
                        else:
                            self.taken[code] = {'Title':title, 'Semester':semester, 'Grade':letterGrade[:1]} # it means it taken



    def fetch(self):
        driver = webdriver.Firefox()  # go mozilla
        links = self.entry.get('1.0', "end-1c").split('\n')  # the links we provided or user will provide
        for i in links:  # for each link
            try:  # it is for departmant links. if uni one selected, we will doing expect
                driver.get(i)  # open page
                elem = driver.find_element_by_link_text("Course Descriptions")  # find specific location
                elem.click()  # click one it
                time.sleep(2)
                html = driver.find_element_by_tag_name("body").text  # get the source
                text = html.split('\n')  # separete it to new line
                for i in range(len(text)):  # for whole text
                    if len(text[i]) == 0:  # if line is nothing, go on
                        continue
                    if 'ECTS' in text[i]:  # if ects in line, it means it is our code
                        a = text[i].find(' ', text[i].find(' ') + 1)
                        # print text[i][:a],'=',text[i+1]
                        code = text[i][:a]
                        description = text[i + 1]  # and description of our code

                        if code in self.taken:
                            self.taken[code]['Description'] = description  # save it in dict format
                        if code in self.willTake:
                            self.willTake[code]['Description'] = description
                        else:
                            self.willTake[code] = {'Description': description, 'Semester': 'Departmental'}


            except:  # for uni codes, do same thing
                driver.get(i)
                html = driver.find_element_by_id("icsayfa_sag").text
                text = html.split('\n')

                for i in range(len(text)):
                    if len(text[i]) == 0 :
                        continue
                    if 'ECTS' in text[i]:
                        a = text[i].find(' ', text[i].find(' ') + 1)
                        code = text[i][:a]
                        description = text[i+2]

                        if code in self.taken:
                            self.taken[code]['Description'] = description
                        if code in self.willTake:
                            self.willTake[code]['Description'] = description
                        else:
                            self.willTake[code] = {'Description': description, 'Semester': 'Electives'}

        print '**', self.willTake
        self.prediction()

        driver.close()





    def prediction(self):

        cl = docclass.naivebayes(docclass.getwords)  # naive bayes process
        # get the grades
        for lecture in self.taken:
            # print '!*-=-=-=',self.taken[lecture]['Description']
            try:
                cl.train(self.taken[lecture]['Description'], self.taken[lecture]['Grade'])  # train it
            except:
                None

        for lecture in self.willTake:
            try:
                self.willTake[lecture]['Grade'] = cl.classify(self.willTake[lecture]['Description'], default='unkown') # get result
            except:
                None


        #tags for colouring the text

        print self.willTake
        self.text.tag_config("Semester", underline=1, font="Verdana 14")
        self.text.tag_config("A", background="green",font="Verdana 14")
        self.text.tag_config("B", background="light green",font="Verdana 14")
        self.text.tag_config("C", background="yellow",font="Verdana 14")
        self.text.tag_config("D", background="red",foreground="white",font="Verdana 14")
        self.text.tag_config("F", background="black",foreground="white",font="Verdana 14")


        # what we predict
        semesterlist=  ['Semester III', 'Semester IV', 'Semester V', 'Semester VI',
                                 'Semester VII','Semester VIII', 'Electives', 'Departmental']
        self.text.delete("1.0", END)
        for semester in semesterlist:
            self.text.insert(END, semester+"\n\n", "Semester")
            for i in self.willTake:
                if semester == self.willTake[i]['Semester']:
                    if 'Grade' not in self.willTake[i]:
                        continue
                    if self.willTake[i]['Semester']== {}:
                        continue

                    line = i + '-->' + self.willTake[i]['Grade'] + '\n'
                    tag= str(self.willTake[i]['Grade'])
                    self.text.insert(END, line, tag)
            self.text.insert(END,"\n")




    def geometricDesign(self):

        self.title.grid(row=0,column=0,columnspan=5)
        self.please.grid(row=1,column=0,sticky=W,columnspan=2,pady=10,padx=5)
        self.browse.grid(row=1,column=2,sticky=W)
        self.enter.grid(row=2,column=0,sticky=W,pady=(10,0),padx=5)
        self.entry.grid(row=3,column=0,columnspan=5,sticky=W,ipady=30,padx=5)
        self.key.grid(row=4,column=0,sticky=W,pady=5,padx=5)
        self.A.grid(row=0,column=0,sticky=W,pady=10,padx=5)
        self.B.grid(row=0,column=1,sticky=W,padx=5)
        self.C.grid(row=0,column=2,sticky=W,padx=5)
        self.D.grid(row=0,column=3,sticky=W,padx=5)
        self.F.grid(row=0,column=4,sticky=W,padx=5)
        self.predict.grid(row=0,column=5,padx=10)
        self.dots.grid(row=1,column=0,columnspan=7,padx=5)
        self.predicted.grid(row=0,column=0,sticky=W,padx=5)
        self.scrollbar.grid(row=1,column=1,rowspan=2,sticky=W+NS)
        self.text.grid(row=1,column=0,sticky=EW,padx=(5,0))



        self.frame1.grid(sticky=W)
        self.frame2.grid(sticky=W)
        self.frame3.grid(sticky=W)









def main():

    root = Tk()

    root.title("Guess My Grade")

    # root.geometry("850x500+150+100")

    app = GuessGrade(root)

    root.mainloop()











if __name__ == '__main__':

    main()






