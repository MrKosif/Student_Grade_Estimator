from tkinter import *
import tkinter as Tkinter
import tkinter.ttk as ttk
import xlrd
import PyPDF2
from tkinter import filedialog
from recommendations import *

#NOTES TO ASSISTANT AND PROFFESOR: PDF DATA EXTRACTING IS LIKE TORCHER SO I MADE MISTAKES IT ONLY WORKS WITH MY TRANSCRIPT
# I COULDN'T ABLE TO CONNECT .TXT FİLE AND GUI SO THIS FILE AND .TXT FILE SHOULD BE IN THE SAME FOLDER.
# FOR PDF MY NUMBER STARTS WITH THIS IS MY SECOND YEAR IN THE UNIVERSITY AND I DID MY BEST PLEASE GRADE ME AS I AM A HUMAN
# BEING

class GUI(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)
        #Here I write list and dict varaibles for my code and stringvars
        self.all_credits = 0
        self.all_courses2 = {}
        self.gpa = 0
        self.mistake_big_list = []
        self.list_for_gpa_updating = []
        self.king_list = []
        self.result = []
        self.listsd = []
        self.var2 = StringVar()
        self.var = StringVar()
        self.var3 = StringVar()
        self.result = []
        self.letter_to_number = {}
        self.number_to_letter = {}
        self.final_data = {}
        self.listbox1_list = []
        self.dict_for_rec = {}
        self.pack(fill=X)
        self.initUI()

    def select_file(self, type):
        #This part is for file selection due to my lack of time (yeah lack of time) I couldent able to work out
        #letter_grade_system.txt file so while selecting do not select it because it will open manuelly.
        self.filename = filedialog.askopenfilename(initialdir="/", title="Select file", filetypes=(
        ("jpeg files", "*.jpg"), ("all files", "*.*"), ("txt file", "*.txt"), ("pdf file", "*.pdf"),
        ("excel file", "*.xlsx")))
        global transcript
        transcript = ""
        global excel_file
        excel_file = ""
        global grade_table
        grade_table = ""
        #when you select file it will start related function with it.
        if type == "Excel":
            excel_file = self.filename
            self.excel_reading()
        if type == "pdf":
            transcript = self.filename
            self.pdf_data()
            self.item_inserting()
        if type == "txt":
            grade_table = self.filename
    #This is all of the GUI here
    def initUI(self):
        self.frame1 = Frame(self, relief=GROOVE)
        self.frame1.pack(side=TOP, fill=X, expand=True)

        self.frame2 = Frame(self, relief=GROOVE)
        self.frame2.pack(side=TOP, fill=BOTH)

        self.frame3 = Frame(self, relief=GROOVE, borderwidth=2, padx=90, pady=8)
        self.frame3.pack(side=TOP, pady=(20, 0))

        self.frame4 = Frame(self, relief=GROOVE, borderwidth=2, padx=80, pady=12)
        self.frame4.pack(side=TOP)

        #########################################

        self.title_lable = Label(self.frame1, text="Smart Advisor - Your Intelligent Agent", bg="DodgerBlue2", fg="white",
                                 anchor=CENTER, font=('', '20'))
        self.title_lable.pack(side=TOP, fill=X)
        # There is the 3 button in the top
        self.button_grade = Button(self.frame2, text="Upload Letter Grade Data", height=2, width=20, command=lambda: self.select_file("txt"))
        self.button_grade.pack(side=LEFT, padx=(70, 15), pady=15)

        self.button_course = Button(self.frame2, text="Upload Past Course Data", height=2, command=lambda: self.select_file("Excel"))
        self.button_course.pack(side=LEFT, padx=20, pady=15)

        self.button_transcript = Button(self.frame2, text="Upload Transcript", height=2, command=lambda: self.select_file("pdf"))
        self.button_transcript.pack(side=LEFT, padx=20, pady=15)

        ##########################################

        self.label2 = Label(self.frame3, text="Recommendation Filters:", anchor=CENTER)
        self.label2.grid(row=0, column=1, padx=15, columnspan=4)

        self.label_lb = Label(self.frame3, text="Subjects:")
        self.label_lb.grid(row=1, column=0)

        self.listbox = Listbox(self.frame3, width=11, height=6, selectmode='multiple')
        self.listbox.grid(row=2, column=0)
        self.listbox.bind("<<ListboxSelect>>", self.on_select)

        self.scrollbar = Scrollbar(self.frame3, command=self.listbox.yview)
        self.scrollbar.grid(row=2, column=1, sticky="ns", pady=15)
        self.listbox.configure(yscrollcommand=self.scrollbar.set)

        # ------------------------------------------
        self.label_weird = Label(self.frame3, text="""       Estimated
        grade should
        be at least:""")
        self.label_weird.grid(row=2, column=2, sticky=N)

        self.lb2 = ttk.Combobox(self.frame3, width=5)
        self.lb2.grid(row=2, column=3, pady=20, padx=20, sticky=N)

        self.recommand_button = Button(self.frame3, text="Get Recommendations", command=self.filtering)
        self.recommand_button.grid(row=2, column=2, columnspan=2, rowspan=2, padx=15, pady=70)

        # ################## FRAME 4 ##################


        self.label_lb2 = Label(self.frame4, text="Courses & Est. Grades")
        self.label_lb2.grid(row=1, column=0, columnspan=2)

        self.listbox2 = Listbox(self.frame4, width=11, height=6, selectmode="multiple")
        self.listbox2.grid(row=2, column=0, sticky=E)
        self.listbox2.bind("<<ListboxSelect>>", self.on_select2)

        self.scrollbar2 = Scrollbar(self.frame4, command=self.listbox.yview)
        self.scrollbar2.grid(row=2, column=1, sticky="ns", pady=10)
        self.listbox2.config(yscrollcommand=self.scrollbar2.set)
        # self.listbox = Listbox(self.frame3, width=8, height=6, selectmode='multiple')
        # self.listbox.grid(row=2, column=0)
        # self.listbox.bind("<<ListboxSelect>>", self.on_select)
        #
        # self.scrollbar = Scrollbar(self.frame3, command=self.listbox.yview)
        # self.scrollbar.grid(row=2, column=1, sticky="ns")
        # self.listbox.configure(yscrollcommand=self.scrollbar.set)

        # ------------------------------------------
        #This is the place where gpa is reflected
        self.label_3 = Label(self.frame4, text="Current Gpa:")
        self.label_3.grid(row=2, column=2, sticky=N, padx=5, rowspan=2)

        self.label_4 = Label(self.frame4, text="New Gpa:")
        self.label_4.grid(row=2, column=3, sticky=N, padx=5, rowspan=2)

        self.current_gpa_data = Label(self.frame4, text="", textvariable=self.var)
        self.current_gpa_data.grid(row=2, column=2, padx=5, rowspan=2, sticky=N, pady=20)

        self.new_gpa_data = Label(self.frame4, text="", textvariable=self.var2)
        self.new_gpa_data.grid(row=2, column=3, padx=5, rowspan=2, sticky=N, pady=20)

        self.note_label = Label(self.frame4, text="""        Select some
        courses to see
        the change in
        your GPA""")
        self.note_label.grid(row=2, column=2, padx=5, pady=40, sticky=S, rowspan=2)

        self.grade_increase_label = Label(self.frame4, text="", textvariable=self.var3)
        self.grade_increase_label.grid(row=2, column=3, padx=5, pady=40, rowspan=2)

    #This is the function whenever you clicked somewhere in first listbox
    def on_select(self, val):
        self.sender = val.widget
        self.idx = self.sender.curselection()

    #This is same with second listbox
    def on_select2(self, val):
        self.list_for_gpa_updating = []
        self.sender2 = val.widget
        self.idx2 = self.sender2.curselection()
        all_points = float(self.all_credits) * float(self.gpa)
        new_credits = 0
        addeble = 0
        ##### THIS ALL CODE IS FOR CALCULATİNG GPA
        for values in self.idx2:
            self.value2 = self.sender2.get(values)
            self.list_for_gpa_updating.append(self.value2)
        for element in self.list_for_gpa_updating:
            step1 = element.split("  ")
            step2 = step1[0]
            for list in self.all_courses2.keys():
                if list == step2:
                    new_credits += float(self.all_courses2[step2].credits)
                    addeble += float(self.letter_to_number[step1[1]]) * float(self.all_courses2[step2].credits)
                    new_gpa = (float(all_points) + float(addeble)) / (float(new_credits) + float(self.all_credits))
        try:
            self.var2.set(round(new_gpa, 2))
            change = float((new_gpa * 100) / self.gpa) - 100
            if new_gpa == 2.9:
                self.var3.set("%0")
            if new_gpa < self.gpa:
                self.var3.set("-%" + str(round(change, 2)))
            else:
                self.var3.set("+%" + str(round(change, 2)))
        except:
            self.var2.set(self.gpa)

    def item_inserting(self):
        # THIS PLACE FOR INSERTING ITEMS TO THE LISTBOXES
        for item in self.listbox1_list:
            self.listbox.insert(END, item)

        with open("letter_grade_system.txt") as grade_s:
            listy = []
            for line in grade_s:
                list_form = line.split()
                listy.append(list_form[0])
            self.lb2.configure(values=listy)
    # I WROTE THIS FUNCTION FOR ROUNDING GPA'S TO NUMBERS IT'S THE ONLY WAY
    def rounder(self, number):
        if number == 4.1:
            return "A+"
        elif number >= 4.0 and number <= 4.09:
            return "A"
        elif number >= 3.7 and number <= 3.99:
            return "A-"
        elif number >= 3.3 and number <= 3.69:
            return "B+"
        elif number >= 3.0 and number <= 3.29:
            return "B"
        elif number >= 2.7 and number <= 2.99:
            return "B-"
        elif number >= 2.3 and number <= 2.69:
            return "C+"
        elif number >= 2.0 and number <= 2.29:
            return "C"
        elif number >= 1.7 and number <= 1.99:
            return "C-"
        elif number >= 1.3 and number <= 1.69:
            return "D+"
        elif number >= 1.0 and number <= 1.29:
            return "D"
        elif number >= 0.5 and number <= 0.99:
            return "D-"
        elif number == 0.0:
            return "F"

    #THIS IS THE RECOMANDATION SYSTEM PART
    def redomandation_system(self):
        with open("letter_grade_system.txt") as grade_s:
            for line in grade_s:
                list_form = line.split()
                self.letter_to_number[list_form[0]] = float(list_form[1])
                self.number_to_letter[float(list_form[1])] = list_form[0]

        for student in self.final_data:
            temporary_list = []
            temporary_dict = {}
            student_lists = self.final_data[student]
            #REASON FOR TRY EXCEPT IS WHILE I AM MERGING THE EXCEL DATA WITH PDF DATA I MADE MISTAKES IN DATA BASE SO I DONT KNOW EIGHTER
            for item in student_lists:
                try:
                    temporary_list.append(item[1][0])
                except:
                    temporary_list.append(item[1])
            for item in temporary_list:
                try:
                    temporary_dict[item.subject + " " + item.code] = self.letter_to_number[item.letter_grade]
                except:
                    pass
            self.dict_for_rec[student] = temporary_dict

        final_result = getRecommendations(self.dict_for_rec, 218333201, similarity=sim_distance)
        self.user_based_cf = getRecommendations(self.dict_for_rec, 218333201, similarity=sim_distance)
        self.king_list.append(self.user_based_cf)
        itemMatch = calculateSimilarItems(self.dict_for_rec)
        #item_based_cf = getRecommendedItems(self.dict_for_rec, itemMatch, 218333201)

    # THIS IS FOR FILTERING WHEN WE SELECT GRADE OR COURSE IT WILL FILTER
    def filtering(self):
        self.listbox2.delete("0", "end")
        self.filter1 = self.letter_to_number[self.lb2.get()]
        for values in self.idx:
            self.value = self.sender.get(values)
            self.listsd.append(self.value)
        for item in self.king_list[0]:
            sliced = item[1].split()
            if sliced[0] in self.listsd and item[0] > self.filter1:
                self.result.append(str(item[1]) + "  " + str(self.rounder(item[0])))
        my_courses = []
        for list in self.mistake_big_list:
            my_courses.append(list[0])
        for data in self.result:
            if data not in my_courses:
                self.listbox2.insert(END, data)

        self.listsd = []
        self.result = []

    #THIS IS THE ALL CODE THAT I TOOK DATA FROM EXCEL
    def excel_reading(self):
        loc = excel_file
        wb = xlrd.open_workbook(loc)
        sheet = wb.sheet_by_index(0)
        students = {}
        grade_system = {}
        students_with_objects = {}

        for i in range(sheet.nrows - 1):
            #HERE I TOOK THEM AS RAW
            course = sheet.cell_value(i + 1, 0)
            splited_course = course.split()
            subject = splited_course[0]
            subject_code = splited_course[1]
            grade = sheet.cell_value(i + 1, 1)
            student_number = sheet.cell_value(i + 1, 2)
            course_credit = sheet.cell_value(i + 1, 3)
            full_class_name = subject + " " + subject_code
            # --------------------------------------------------------------

            personal_course_object = Course(subject, subject_code, course_credit, grade)

            if student_number not in students:
                students[int(student_number)] = [[full_class_name, personal_course_object]]

            elif student_number in students:
                students[int(student_number)].append([full_class_name, personal_course_object])

        # -------------------------------------------------------------
        with open("letter_grade_system.txt") as grade_s:
            for line in grade_s:
                list_form = line.split()
                grade_system[list_form[0]] = float(list_form[1])
        # IN HERE I MADE THEM SHAPES
        for real_students in list(students.keys()):
            all_courses = students[real_students]
            all_points = 0
            all_credit = 0
            for objects_and_names in all_courses:
                objectt = objects_and_names[1]
                gpa_effect = objectt.credits
                all_points += gpa_effect * float(grade_system[objectt.letter_grade])
                all_credit += gpa_effect
            gpa = all_points / all_credit
            for lists in students[real_students]:
                self.all_courses2[lists[0]] = lists[1]
            student_course_object = Student(real_students, students[real_students], round(gpa, 2))
            students_with_objects[student_course_object.id] = student_course_object
        # --------------------------------------------------------------------------
        global grade_data
        grade_data = CourseGradeData(students_with_objects, grade_system)
        for student in grade_data.students.keys():
            self.final_data[student] = grade_data.students[student].courses

    #HERE I MADE ALL OF THE PDF DATA COLLECTING
    def pdf_data(self):
        pdfFileObj = open(transcript, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
        pageObj = pdfReader.getPage(0)

        a = pageObj.extractText()
        b = []
        main_list = []
        b.append(a)
        splity = b[0].split()

        check = False
        list_index = 0
        counter = 0
        new_list = []
        semester_count = 0
        current_semester = 0
        for item in splity:
            if item[5:] == "Cr.CmECTSCr.":
                semester_count += 1

        for item in splity:
            if item[5:] == "Cr.CmECTSCr.":
                current_semester += 1
                check = False

            if current_semester == semester_count:
                check = False

            if item[:-4] == "TitleCreditECTSGrade":
                course_in_a_list = item.split("TitleCreditECTSGrade")
                new_list.append(course_in_a_list[1])

            # Letter Grade part ex: C+
            if list_index != 0 and counter == 1:
                if item[5] == "+" or item[5] == "-":
                    main_list[list_index - 1].append(item[4:6])
                    gpa_grade = item[4:6]
                else:
                    main_list[list_index - 1].append(item[4])
                    gpa_grade = item[4]

            # Course name part ex: MATH
            if check == True and item[:-4] != "TitleCreditECTSGrade":
                if counter == 1:
                    a = item.split(gpa_grade)
                    new_list.append(a[1])
                else:
                    new_list.append(item)

            # content of the new list appending to the new list and get reset.
            if type(item) == str and current_semester != semester_count:
                try:
                    int(item)
                    main_list.append(new_list)
                    new_list = []
                    counter = 0
                    list_index += 1
                except:
                    pass

            if item == "CodeCourse" and current_semester != semester_count:
                check = True

            counter += 1

        main_list[0][0] = main_list[0][0][-4:]
        pdfFileObj.close()

        cleared_list = []
        for list in main_list:
            new_list = []
            new_list.append(list[0])
            new_list.append(list[1][:3])
            new_list.append(list[-2])
            new_list.append(list[-1])
            cleared_list.append((new_list))
        user_id = int(splity[29].split(":")[4])
        user_gpa = splity[-3][:4]
        self.gpa = float(user_gpa)
        self.all_credits = splity[-5]
        course_for_student_object = {}
        for course in cleared_list:
            personal_course_object = Course(course[0], course[1], course[2], course[3])
            course_for_student_object[personal_course_object.subject + personal_course_object.code] = personal_course_object
            mistake_list = [course[0] + " " + course[1], [personal_course_object]]
            self.mistake_big_list.append(mistake_list)

        student_object_transcript = Student(user_id, course_for_student_object, user_gpa)
        self.final_data[student_object_transcript.id] = self.mistake_big_list
        for student in self.final_data:
            student_lists = self.final_data[student]
            for item in student_lists:
                try:
                    if item[1][0].subject not in self.listbox1_list:
                        self.listbox1_list.append(item[1][0].subject)
                except:
                    if item[1].subject not in self.listbox1_list:
                        self.listbox1_list.append(item[1].subject)

        self.var.set(str(user_gpa))
        self.redomandation_system()
        # THIS CONFUSING PART IS MERGING BOTH EXCEL DATA AND PDF DATA

# THESE ARE 3 CLASSES THAT I USE IN MY PROJECT
class Course:
    def __init__(self, subject, code, creditss, letter_grade):
        self.subject = subject
        self.code = code
        self.credits = creditss
        self.letter_grade = letter_grade

class Student:
    def __init__(self, id, courses, gpa):
        self.id = id
        self.courses = courses
        self.gpa = gpa

class CourseGradeData:
    def __init__(self, students, letter_grade_mapping):
        self.students = students
        self.letter_grade_mapping = letter_grade_mapping

# THIS IS THE FUNCTION THAT STARTS EVERYTHİNG
def main():
    root = Tkinter.Tk()
    root.geometry("680x600+380+100")
    gui = GUI(root)
    root.mainloop()

main()
