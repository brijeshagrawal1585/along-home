from datetime import datetime, timezone
import time
from tkinter import filedialog
from tkinter import*
from PIL import Image,ImageTk #pip install pilow
from tkinter import ttk,messagebox
import pymysql
import pywhatkit as kit
from tkinter import messagebox
from tkcalendar import Calendar
from tkcalendar import DateEntry  # Import DateEntry from tkcalendar
import sqlite3
from tkinter import filedialog, messagebox
import pandas as pd
import re
from reportlab.lib.pagesizes import A4, landscape, A5, A3, A2, A1, portrait
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Image as RLImage
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
import os
from fpdf import FPDF, XPos, YPos
class staffClass:
    def __init__(self,root):
        self.root=root
        self.root.geometry("1350x700+0+0")
        self.root.title("Along Home Healthcare | Developed by Brijesh | July | 9998002712")
        self.root.config(bg="white")
        self.root.focus_force()
        # All Variables================
        self.var_searchby=StringVar()
        self.var_searchtxt=StringVar()

        self.var_s_id=StringVar()
        self.var_new_staff_name=StringVar()
        self.var_birth_date=StringVar()
        self.var_age=StringVar()
        
        self.var_contact=StringVar()
        self.var_other_contact=StringVar()
        self.var_email_id=StringVar()
        self.var_cast=StringVar()
        
        self.var_marital_status=StringVar()
        self.var_gender=StringVar()
        self.var_religion=StringVar()
        self.var_nationality=StringVar()
        
        self.var_experience=StringVar()
        self.var_present_address=StringVar()
        self.var_permanent_address=StringVar()
        self.var_education=StringVar()
        
        self.var_idproofcomb=StringVar()
        self.var_id_proof=StringVar()
        self.var_nursing_certificate=StringVar()
        self.var_exp_letter=StringVar()
        self.var_light_bill=StringVar()
        
        self.var_last_work_place=StringVar()
        self.var_duty=StringVar()
        #self.var_reference_no=StringVar()
        self.var_financial_year=StringVar()
        #self.var_photo = StringVar()  # Variable to store the image1 path
        self.var_img_path_1 = StringVar()  # Variable to store image1 path
        self.default_img_1 = Image.open("./image/along_home_logo_1.png")  # Default image1 placeholder
        self.default_img_1 = self.default_img_1.resize((100, 100), Image.LANCZOS)
        self.default_img_1 = ImageTk.PhotoImage(self.default_img_1)            
        
        #self.var_photo = StringVar()  # Variable to store the image2 path
        self.var_img_path_2 = StringVar()  # Variable to store image2 path
        self.default_img_2 = Image.open("./image/along_home_logo_2.png")  # Default image2 placeholder
        self.default_img_2 = self.default_img_2.resize((100, 100), Image.LANCZOS)
        self.default_img_2 = ImageTk.PhotoImage(self.default_img_2)            
        #============Search Frame==================
        SearchFrame=LabelFrame(self.root,text="Search Staff Data",font=("helvetica",12,"bold"),bd=2,relief=RIDGE,bg="white")
        SearchFrame.place(x=30,y=20,width=700,height=70)

        cmb_search=ttk.Combobox(SearchFrame,textvariable=self.var_searchby,values=("Select","S_ID","NEW_STAFF_NAME","CONTACT","2ND CONTACT","FINANCIAL_YEAR"),state='readonly',justify=CENTER,font=("helvetica",15))
        cmb_search.place(x=10,y=10,width=100)
        cmb_search.current(0)

        txt_search=Entry(SearchFrame,textvariable=self.var_searchtxt,font=("helvetica",15),bg="lightyellow")
        txt_search.place(x=120,y=10,width=100)
        btn_search=Button(SearchFrame,text="Search",command=self.search,font=("helvetica",15),bg="black",fg="white",cursor="hand2").place(x=240,y=9,width=100,height=30)
        
        btn_export_pdf = Button(SearchFrame, text="Export Search to PDF", command=self.export_searched_staff_to_pdf, font=("helvetica", 12,"bold"), bg="red", fg="white",cursor="hand2")
        btn_export_pdf.place(x=350,y=9,width=170,height=30)  # Set proper coordinates

        btn_export=Button(SearchFrame, text="Export from Search", command=self.export_from_search, font=("helvetica", 12,"bold"), bg="blue", fg="white",cursor="hand2")
        btn_export.place(x=530,y=9,width=160,height=30)        
        
        ExportFrame=LabelFrame(self.root,text="Export Data",font=("helvetica",12,"bold"),bd=2,relief=RIDGE,bg="white")
        ExportFrame.place(x=740,y=20,width=580,height=70)
        
        btn_export = Button(ExportFrame, text="Export to Excel", command=self.export_to_excel, font=("helvetica", 15, "bold"), bg="yellow", fg="black", cursor="hand2")
        btn_export.place(x=10, y=10, width=160, height=28)        
        
        btn_export_pdf = Button(ExportFrame, text="Export to PDF", command=self.export_to_pdf, font=("helvetica", 15, "bold"), bg="purple", fg="white", cursor="hand2")
        btn_export_pdf.place(x=180, y=10, width=150, height=28)
        
        btn_import=Button(ExportFrame, text="Import from Excel", command=self.import_excel_to_mysql, font=("helvetica", 15,"bold"), bg="brown", fg="white", cursor="hand2")
        btn_import.place(x=340, y=10,width=200,height=28)
        #===========Title===================
        #title=Label(self.root,text="new_staff_name_part Details",font=("helvetica",15),bg="#0f4d7d",fg="white").place(x=50,y=100,width=1000)
        #=========Clock==============
        self.lbl_clock=Label(self.root,text="Welcome to Along Home healthcare\t\t Date: DD-MM-YYYY\t\t Time: HH:MM:SS",font=("times new roman",15),bg="#0f4d7d",fg="white")
        self.lbl_clock.place(x=20,y=100,width=1300,height=30)
        #===========Content================
        #=========Row 1===========
        lbl_s_id=Label(self.root,text="S ID",font=("helvetica",13),bg="white").place(x=50,y=150)
        self.txt_s_id=Entry(self.root,textvariable=self.var_s_id,justify=CENTER,font=("helvetica",13),bg="lightyellow")
        self.txt_s_id.place(x=180,y=150,width=100)
        
        lbl_new_staff_name=Label(self.root,text="New Staff Name",font=("helvetica",13),bg="white").place(x=290,y=150)
        txt_new_staff_name=Entry(self.root,textvariable=self.var_new_staff_name,justify=CENTER,font=("helvetica",12),bg="lightyellow").place(x=440,y=150,width=110)
        
        lbl_birth_date=Label(self.root,text="Birth Date",font=("helvetica",13),bg="white").place(x=560,y=150)
        # Date picker widget
        txt_birth_date = DateEntry(self.root,textvariable=self.var_birth_date,justify=CENTER,font=("helvetica", 12),bg="lightyellow",date_pattern="dd/mm/yyyy")  # Set the format to dd/mm/yyyy
        txt_birth_date.place(x=690,y=150,width=150)
        # Display the default date in the variable
        self.var_birth_date.set(txt_birth_date.get_date().strftime("%d-%m-%Y"))
        
        lbl_age=Label(self.root,text="Age(Yrs)",font=("helvetica",13),bg="white").place(x=860,y=150)
        txt_age=Entry(self.root,textvariable=self.var_age,justify=CENTER,font=("helvetica",12),bg="lightyellow").place(x=1000,y=150,width=110)
        #===========Row 2===========
        lbl_contact=Label(self.root,text="Contact",font=("helvetica",13),bg="white").place(x=50,y=190)
        txt_contact=Entry(self.root,textvariable=self.var_contact,justify=CENTER,font=("helvetica",12),bg="lightyellow").place(x=180,y=190,width=100)
        
        lbl_other_contact=Label(self.root,text="2nd Contact",font=("helvetica",13),bg="white").place(x=290,y=190)
        txt_other_contact=Entry(self.root,textvariable=self.var_other_contact,justify=CENTER,font=("helvetica",12),bg="lightyellow").place(x=440,y=190,width=110)
        
        lbl_email_id=Label(self.root,text="Email ID",font=("helvetica",13),bg="white").place(x=560,y=190)
        txt_email_id=Entry(self.root,textvariable=self.var_email_id,justify=CENTER,font=("helvetica",12),bg="lightyellow").place(x=690,y=190,width=150)
        
        lbl_cast=Label(self.root,text="Category(Cast)",font=("helvetica",13),bg="white").place(x=860,y=190)
        cmb_cast=ttk.Combobox(self.root,textvariable=self.var_cast,values=("Select","NA","Tribal","SC","ST","OBC","General"),state='readonly',justify=CENTER,font=("helvetica",12))
        cmb_cast.place(x=1000,y=190,width=110)
        cmb_cast.current(0)
        #===========Row 3==========
        lbl_marital_status=Label(self.root,text="Marital Status",font=("helvetica",13),bg="white").place(x=50,y=230)
        cmb_marital_status=ttk.Combobox(self.root,textvariable=self.var_marital_status,values=("Select","NA","Single","Married","Divorced","Widow","Widower"),state='readonly',justify=CENTER,font=("helvetica",12))
        cmb_marital_status.place(x=180,y=230,width=100)
        cmb_marital_status.current(0)
        
        lbl_gender=Label(self.root,text="Gender",font=("helvetica",13),bg="white").place(x=290,y=230)
        cmb_gender=ttk.Combobox(self.root,textvariable=self.var_gender,values=("Select","NA","Male","Female"),state='readonly',justify=CENTER,font=("helvetica",12))
        cmb_gender.place(x=440,y=230,width=110)
        cmb_gender.current(0)
        
        lbl_religion=Label(self.root,text="Religion",font=("helvetica",13),bg="white").place(x=560,y=230)
        cmb_religion=ttk.Combobox(self.root,textvariable=self.var_religion,values=("Select","NA","Hindu","Jain","Sikh","Muslim","Christian"),state='readonly',justify=CENTER,font=("helvetica",12))
        cmb_religion.place(x=690,y=230,width=150)
        cmb_religion.current(0)
                
        lbl_nationality=Label(self.root,text="Nationality",font=("helvetica",13),bg="white").place(x=860,y=230)
        cmb_nationality=ttk.Combobox(self.root,textvariable=self.var_nationality,values=("Select","NA","Indian","NRI"),state='readonly',justify=CENTER,font=("helvetica",12))
        cmb_nationality.place(x=1000,y=230,width=110)
        cmb_nationality.current(0)
        #===========Row 4==========
        lbl_experience=Label(self.root,text="Experience(Yrs)",font=("helvetica",13),bg="white").place(x=50,y=270)
        txt_experience=Entry(self.root,textvariable=self.var_experience,justify=CENTER,font=("helvetica",12),bg="lightyellow").place(x=180,y=270,width=100)
        
        lbl_present_address=Label(self.root,text="Present Add",font=("helvetica",13),bg="white").place(x=290,y=270)
        txt_present_address=Entry(self.root,textvariable=self.var_present_address,justify=CENTER,font=("helvetica",12),bg="lightyellow").place(x=440,y=270,width=110)
        
        lbl_permanent_address=Label(self.root,text="Permanent Add",font=("helvetica",13),bg="white").place(x=560,y=270)
        txt_permanent_address=Entry(self.root,textvariable=self.var_permanent_address,justify=CENTER,font=("helvetica",12),bg="lightyellow").place(x=690,y=270,width=150)
        
        lbl_education=Label(self.root,text="Education",font=("helvetica",13),bg="white").place(x=860,y=270)
        txt_education=Entry(self.root,textvariable=self.var_education,justify=CENTER,font=("helvetica",12),bg="lightyellow").place(x=1000,y=270,width=110)
        #===========Row 5==========
        combo_txt_proof=ttk.Combobox(self.root,textvariable=self.var_idproofcomb,state="readonly",
                                                           font=("helvetica",12,"bold"),width=8)
        combo_txt_proof["value"]=("Select ID","NA","PAN CARD","ADHAR CARD","DRIVING LICENCE","ELECTION CARD")
        combo_txt_proof.place(x=50,y=310,width=110)
        combo_txt_proof.current(0)
        
        txt_id_proof=Entry(self.root,textvariable=self.var_id_proof,justify=CENTER,font=("helvetica",12),bg="lightyellow").place(x=180,y=310,width=100)
        
        lbl_nursing_certificate=Label(self.root,text="Nursing Certificate",font=("helvetica",13),bg="white").place(x=290,y=310)
        txt_nursing_certificate=Entry(self.root,textvariable=self.var_nursing_certificate,justify=CENTER,font=("helvetica",12),bg="lightyellow").place(x=440,y=310,width=110)
        
        lbl_exp_letter=Label(self.root,text="Exp Letter",font=("helvetica",13),bg="white").place(x=560,y=310)
        txt_exp_letter=Entry(self.root,textvariable=self.var_exp_letter,justify=CENTER,font=("helvetica",12),bg="lightyellow").place(x=690,y=310,width=150)
        
        lbl_light_bill=Label(self.root,text="Light Bill",font=("helvetica",13),bg="white").place(x=860,y=310)
        txt_light_bill=Entry(self.root,textvariable=self.var_light_bill,justify=CENTER,font=("helvetica",12),bg="lightyellow").place(x=1000,y=310,width=110)
        #===========Row 6==========
        lbl_last_work_place=Label(self.root,text="Last Work Place",font=("helvetica",13),bg="white").place(x=50,y=350)
        txt_last_work_place=Entry(self.root,textvariable=self.var_last_work_place,justify=CENTER,font=("helvetica",12),bg="lightyellow").place(x=180,y=350,width=100)
        
        lbl_duty=Label(self.root,text="Duty(Hrs)",font=("helvetica",13),bg="white").place(x=290,y=350)
        cmb_duty=ttk.Combobox(self.root,textvariable=self.var_duty,values=("Select","NA","Day","Night","24 Hrs","Visit"),justify=CENTER,font=("helvetica",12))
        cmb_duty.place(x=440,y=350,width=110)
        cmb_duty.current(0)

        #lbl_reference_no=Label(self.root,text="Reference No",font=("helvetica",13),bg="white").place(x=560,y=350)
        #txt_reference_no=Entry(self.root,textvariable=self.var_reference_no,justify=CENTER,font=("helvetica",12),bg="lightyellow").place(x=690,y=350,width=150)
        # ==== Financial Year Dropdown ====
        lbl_financial_year = Label(self.root, text="Financial Year", font=("helvetica", 15,"bold"), bg="white")
        lbl_financial_year.place(x=860, y=350)

        self.cmb_financial_year = ttk.Combobox(self.root, textvariable=self.var_financial_year,justify=CENTER, values=self.get_financial_year_list(), state='readonly', font=("helvetica", 14))
        self.cmb_financial_year.place(x=1000, y=350, width=120)
        self.var_financial_year.set(self.get_current_financial_year())     
        #==========Upload Photo=========================
        self.var_img_path_1 = StringVar()  # Variable to store image1 path
        self.default_img_1 = Image.open("./image/along_home_logo_1.png")  # Default image1 placeholder
        self.default_img_1 = self.default_img_1.resize((100, 100), Image.LANCZOS)
        self.default_img_1 = ImageTk.PhotoImage(self.default_img_1)   
        
        self.lbl_image_1 = Label(self.root, image=self.default_img_1, bd=2, relief=RIDGE)
        self.lbl_image_1.place(x=1140, y=140, width=150, height=120)
        #==========Button to upload photo=================
        btn_upload_1 = Button(self.root, text="Upload Photo", command=self.upload_image1, font=("helvetica", 12,"bold"), bg="maroon", fg="white", cursor="hand2")
        btn_upload_1.place(x=1140, y=260, width=120, height=30)        
        #==========Upload Image=========================
        self.var_img_path_2 = StringVar()  # Variable to store image2 path
        self.default_img_2 = Image.open("./image/along_home_logo_2.png")  # Default image2 placeholder
        self.default_img_2 = self.default_img_2.resize((100, 100), Image.LANCZOS)
        self.default_img_2 = ImageTk.PhotoImage(self.default_img_2)   
        
        self.lbl_image_2 = Label(self.root, image=self.default_img_2, bd=2, relief=RIDGE)
        self.lbl_image_2.place(x=1140, y=290, width=150, height=120)
        #==========Button to upload image=================
        btn_upload_2 = Button(self.root, text="Upload Image", command=self.upload_image2, font=("helvetica", 12,"bold"), bg="maroon", fg="white", cursor="hand2")
        btn_upload_2.place(x=1140, y=410, width=120, height=30)        
        #==========Buttons=========
        self.btn_add=Button(self.root,text="Save",command=self.add,font=("helvetica",15),bg="#2196f3",fg="white",cursor="hand2")
        self.btn_add.place(x=50,y=400,width=110,height=28)
        self.btn_update=Button(self.root,text="Update",command=self.update,font=("helvetica",15),bg="#4caf50",fg="white",cursor="hand2")
        self.btn_update.place(x=170,y=400,width=110,height=28)
        self.btn_delete=Button(self.root,text="Delete",command=self.delete,font=("helvetica",15),bg="#f44336",fg="white",cursor="hand2")
        self.btn_delete.place(x=290,y=400,width=110,height=28)
        self.btn_clear=Button(self.root,text="Clear",command=self.clear,font=("helvetica",15),bg="#607d8b",fg="white",cursor="hand2")
        self.btn_clear.place(x=410,y=400,width=110,height=28)
        
        self.btn_add.config(state='normal')
        self.btn_update.config(state='disabled')
        self.btn_delete.config(state='disabled')
        #self.btn_export_search.config(state='disabled')
        
        btn_whatsapp = Button(self.root, text="Send WhatsApp", command=self.send_whatsapp_message, font=("helvetica", 15,"bold"), bg="green", fg="white", cursor="hand2")
        btn_whatsapp.place(x=530, y=400, width=170, height=28)        
        #=========Footer==============
        lbl_footer=Label(self.root,text="Along Home Healthcare|59/2, Ground Floor, Govt. H. Colony, B/h. Laxmi Ganthiya Rath, Nehrunagar Cross Road, Ahmedabad-380015\tContact:+91 9904110283",font=("times new roman",15),bg="#4d636d",fg="white").pack(side=BOTTOM,fill=X)
        #=========new_staff_name_part Details========
        staff_frame=LabelFrame(self.root,text="Staff List",font=("helvetica",12,"bold"),bd=3,relief=RIDGE,bg="white")
        staff_frame.place(x=0,y=440,relwidth=1,height=220)

        scrolly=Scrollbar(staff_frame,orient=VERTICAL)
        scrollx=Scrollbar(staff_frame,orient=HORIZONTAL) 
        
        self.staff_Table=ttk.Treeview(staff_frame,columns=("s_id","new_staff_name","birth_date","age","contact","other_contact","email_id","cast","marital_status","gender","religion","nationality","experience","present_address","permanent_address","education","idproofcomb","id_proof","nursing_certificate","exp_letter","light_bill","last_work_place","duty","financial_year","image_1","image_2"),yscrollcommand=scrolly.set,xscrollcommand=scrollx.set)
        scrollx.pack(side=BOTTOM,fill=X)
        scrolly.pack(side=RIGHT,fill=Y)
        scrollx.config(command=self.staff_Table.xview)
        scrolly.config(command=self.staff_Table.yview)

        self.staff_Table.heading("s_id",text="S ID")
        self.staff_Table.heading("new_staff_name",text="New Staff Name")
        self.staff_Table.heading("birth_date",text="Birth Date")
        self.staff_Table.heading("age",text="Age")
        self.staff_Table.heading("contact",text="Contact")
        self.staff_Table.heading("other_contact",text="2nd Contact")
        self.staff_Table.heading("email_id",text="Email ID")
        self.staff_Table.heading("cast",text="Cast")
        self.staff_Table.heading("marital_status",text="Martial Status")
        self.staff_Table.heading("gender",text="Gender")
        self.staff_Table.heading("religion",text="religion")
        self.staff_Table.heading("nationality",text="Nationality")
        self.staff_Table.heading("experience",text="Experience")
        self.staff_Table.heading("present_address",text="Present Address")
        self.staff_Table.heading("permanent_address",text="Permanent Address")
        self.staff_Table.heading("education",text="Education")
        self.staff_Table.heading("idproofcomb",text="ID Type")
        self.staff_Table.heading("id_proof",text="ID Proof")
        self.staff_Table.heading("nursing_certificate",text="Nursing Certificate")
        self.staff_Table.heading("exp_letter",text="Exp Letter")
        self.staff_Table.heading("light_bill",text="Light Bill")
        self.staff_Table.heading("last_work_place",text="Last Work Place")
        self.staff_Table.heading("duty",text="Duty")
        #self.staff_Table.heading("reference_no",text="Reference No")
        self.staff_Table.heading("financial_year",text="Financial Year")
        self.staff_Table.heading("image_1",text="Image 1")
        self.staff_Table.heading("image_2",text="Image 2")

        self.staff_Table["show"]="headings"

        self.staff_Table.column("s_id",width=30)
        self.staff_Table.column("new_staff_name",width=50)
        self.staff_Table.column("birth_date",width=50)
        self.staff_Table.column("age",width=30)
        self.staff_Table.column("contact",width=50)
        self.staff_Table.column("other_contact",width=50)
        self.staff_Table.column("email_id",width=50)
        self.staff_Table.column("cast",width=50)
        self.staff_Table.column("marital_status",width=50)
        self.staff_Table.column("gender",width=50)
        self.staff_Table.column("religion",width=50)
        self.staff_Table.column("nationality",width=50)
        self.staff_Table.column("experience",width=30)
        self.staff_Table.column("present_address",width=100)
        self.staff_Table.column("permanent_address",width=100)
        self.staff_Table.column("education",width=50)
        self.staff_Table.column("idproofcomb",width=50)
        self.staff_Table.column("id_proof",width=50)
        self.staff_Table.column("nursing_certificate",width=50)
        self.staff_Table.column("exp_letter",width=50)
        self.staff_Table.column("light_bill",width=50)
        self.staff_Table.column("last_work_place",width=50)
        self.staff_Table.column("duty",width=50)
        #self.staff_Table.column("reference_no",width=50)
        self.staff_Table.column("financial_year",width=50)
        self.staff_Table.column("image_1",width=100)
        self.staff_Table.column("image_2",width=100)
        self.staff_Table.pack(fill=BOTH,expand=1)
        self.staff_Table.bind("<ButtonRelease-1>",self.get_data)

        self.update_content()
        self.show()
#=======================================================================================================================================================================
    # ==== Helper Methods ====
    def get_current_financial_year(self):
        today = datetime.today()
        year = today.year
        if today.month >= 4:
            return f"{year}-{year+1}"
        else:
            return f"{year-1}-{year}"

    def get_financial_year_list(self, start_year=2019):
        current_year = datetime.today().year
        if datetime.today().month < 4:
            current_year -= 1
        return [f"{y}-{y+1}" for y in range(start_year, current_year + 2)]    
    
    def add(self):
        con=pymysql.connect(host='localhost',user='root',password='',db='along_home_db')
        #con = sqlite3.connect(database=r'./ah.db')
        cur = con.cursor()
        try:
            if self.var_s_id.get() == "":
                messagebox.showerror("Error", "s_id must be required", parent=self.root)

            elif self.var_new_staff_name.get() == "":
                messagebox.showerror("Error", "gender must be required", parent=self.root)

            # ✅ Format check for gender (example: MH12AB1234)
            #elif not re.match(r'^[A-Z]{2}[0-9]{2}[A-Z]{1,2}[0-9]{4}$', self.var_gender.get()):
                #messagebox.showerror("Error", "Invalid gender format (e.g. MH12AB1234)", parent=self.root)

            # ✅ Uniqueness check for gender
            else:
                cur.execute("SELECT * FROM staff WHERE new_staff_name=%s", (self.var_new_staff_name.get(),))
                existing_new_staff_name = cur.fetchone()
                if existing_new_staff_name:
                    messagebox.showerror("Error", "This new_staff_name is already assigned, try different", parent=self.root)
                    return

                # Check s_id duplication
                cur.execute("SELECT * FROM staff WHERE s_id=%s", (self.var_s_id.get(),))
                row = cur.fetchone()
                if row:
                    messagebox.showerror("Error", "This s_id is already assigned, try different", parent=self.root)
                    return

                # Insert record
                cur.execute("INSERT INTO staff(s_id,new_staff_name,birth_date,age,contact,other_contact,email_id,cast,marital_status,gender,religion,nationality,experience,present_address,permanent_address,education,idproofcomb,id_proof,nursing_certificate,exp_letter,light_bill,last_work_place,duty,financial_year,image_1,image_2) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)", (
                    self.var_s_id.get(),
                    self.var_new_staff_name.get(),
                    self.var_birth_date.get(),
                    self.var_age.get(),
                    self.var_contact.get(),
                    self.var_other_contact.get(),
                    self.var_email_id.get(),
                    self.var_cast.get(),
                    self.var_marital_status.get(),
                    self.var_gender.get(),
                    self.var_religion.get(),
                    self.var_nationality.get(),
                    self.var_experience.get(),
                    self.var_present_address.get(),
                    self.var_permanent_address.get(),
                    self.var_education.get(),
                    self.var_idproofcomb.get(),
                    self.var_id_proof.get(),
                    self.var_nursing_certificate.get(),
                    self.var_exp_letter.get(),
                    self.var_light_bill.get(),
                    self.var_last_work_place.get(),
                    self.var_duty.get(),
                    #self.var_reference_no.get(),
                    self.var_financial_year.get(),
                    self.var_img_path_1.get(),  # Save the image1 path
                    self.var_img_path_2.get(),  # Save the image2 path
                ))
                con.commit()
                messagebox.showinfo("Success", "Staff Added Successfully", parent=self.root)
                self.clear()
                self.show()
        except Exception as ex:
            messagebox.showerror("Error", f"Error due to: {str(ex)}", parent=self.root)
            
    def upload_image1(self):
        file_path_1 = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
        if file_path_1:
            self.var_img_path_1.set(file_path_1)  # Store selected image1 path
            img_1 = Image.open(file_path_1)
            img_1 = img_1.resize((100, 100), Image.LANCZOS)
            img_1 = ImageTk.PhotoImage(img_1)
            self.lbl_image_1.config(image=img_1)
            self.lbl_image_1.image = img_1  # Keep reference to avoid garbage collection    
    
    def upload_image2(self):
        file_path_2 = filedialog.askopenfilename(filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
        if file_path_2:
            self.var_img_path_2.set(file_path_2)  # Store selected image2 path
            img_2 = Image.open(file_path_2)
            img_2 = img_2.resize((100, 100), Image.LANCZOS)
            img_2 = ImageTk.PhotoImage(img_2)
            self.lbl_image_2.config(image=img_2)
            self.lbl_image_2.image = img_2  # Keep reference to avoid garbage collection    
    
    def show(self):
        con=pymysql.connect(host='localhost',user='root',password='',db='along_home_db')
        #con=sqlite3.connect(database=r'./ah.db')
        cur=con.cursor()
        try:
            cur.execute("select * from staff")
            rows=cur.fetchall()
            self.staff_Table.delete(*self.staff_Table.get_children())
            for row in rows:
                self.staff_Table.insert('',END,values=row)
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :{str(ex)}",parent=self.root)

    def get_data(self, ev):
        f = self.staff_Table.focus()
        content = self.staff_Table.item(f)
        row = content['values']
        
        self.var_s_id.set(row[0])
        self.var_new_staff_name.set(row[1])
        self.var_birth_date.set(row[2])
        self.var_age.set(row[3])
        self.var_contact.set(row[4])
        self.var_other_contact.set(row[5])
        self.var_email_id.set(row[6])
        self.var_cast.set(row[7])
        self.var_marital_status.set(row[8])
        self.var_gender.set(row[9])
        self.var_religion.set(row[10])
        self.var_nationality.set(row[11])
        self.var_experience.set(row[12])
        self.var_present_address.set(row[13])
        self.var_permanent_address.set(row[14])
        self.var_education.set(row[15])
        self.var_idproofcomb.set(row[16])
        self.var_id_proof.set(row[17])
        self.var_nursing_certificate.set(row[18])
        self.var_exp_letter.set(row[19])
        self.var_light_bill.set(row[20])
        self.var_last_work_place.set(row[21])
        self.var_duty.set(row[22])
        #self.var_reference_no.set(row[23])
        self.var_financial_year.set(row[23])
        # Load and display the saved image
        img_path_1 = row[24] if row[24] else "./image/along_home_logo_1.png"  # Use default image1 if empty
        self.var_img_path_1.set(img_path_1)
        img_1 = Image.open(img_path_1)
        img_1 = img_1.resize((100, 100), Image.LANCZOS)
        img_1 = ImageTk.PhotoImage(img_1)
        self.lbl_image_1.config(image=img_1)
        self.lbl_image_1.image = img_1
        img_path_2 = row[25] if row[25] else "./image/along_home_logo_2.png"  # Use default image2 if empty
        self.var_img_path_2.set(img_path_2)
        img_2 = Image.open(img_path_2)
        img_2 = img_2.resize((100, 100), Image.LANCZOS)
        img_2 = ImageTk.PhotoImage(img_2)
        self.lbl_image_2.config(image=img_2)
        self.lbl_image_2.image = img_2
    
        # Button control
        self.btn_add.config(state='disabled')
        self.btn_update.config(state='normal')
        self.btn_delete.config(state='normal')
        #self.btn_export_search.config(state='normal')
    
    def update(self):
        con=pymysql.connect(host='localhost',user='root',password='',db='along_home_db')
        #con=sqlite3.connect(database=r'./ah.db')
        cur=con.cursor()
        try:
            if self.var_s_id.get()=="":
                messagebox.showerror("Error","s_id Must be required",parent=self.root)
            else:
                cur.execute("Select * from staff where s_id=%s",(self.var_s_id.get(),))
                row=cur.fetchone()
                if row==None:
                    messagebox.showerror("Error","Invalid s_id",parent=self.root)
                else:
                    cur.execute("Update staff set new_staff_name=%s,birth_date=%s,age=%s,contact=%s,other_contact=%s,email_id=%s,cast=%s,marital_status=%s,gender=%s,religion=%s,nationality=%s,experience=%s,present_address=%s,permanent_address=%s,education=%s,Idproofcomb=%s,id_proof=%s,nursing_certificate=%s,exp_letter=%s,light_bill=%s,last_work_place=%s,duty=%s,financial_year=%s,image_1=%s,image_2=%s where s_id=%s",(
                                        self.var_new_staff_name.get(),  
                                        self.var_birth_date.get(),
                                        self.var_age.get(),
                                        self.var_contact.get(),
                                        self.var_other_contact.get(),
                                        self.var_email_id.get(),
                                        self.var_cast.get(),
                                        self.var_marital_status.get(),
                                        self.var_gender.get(),
                                        self.var_religion.get(),
                                        self.var_nationality.get(),
                                        self.var_experience.get(),
                                        self.var_present_address.get(),
                                        self.var_permanent_address.get(),
                                        self.var_education.get(),
                                        self.var_idproofcomb.get(),
                                        self.var_id_proof.get(),
                                        self.var_nursing_certificate.get(),
                                        self.var_exp_letter.get(),
                                        self.var_light_bill.get(),
                                        self.var_last_work_place.get(),
                                        self.var_duty.get(),
                                        #self.var_reference_no.get(),
                                        self.var_financial_year.get(),
                                        self.var_img_path_1.get(),  # Update image1 path
                                        self.var_img_path_2.get(),  # Update image2 path
                                        self.var_s_id.get()
                    ))
                    con.commit()
                    messagebox.showinfo("Success","Staff Updated Successfully",parent=self.root)
                    self.clear() # Clear fields after updating
                    self.show()
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to : {str(ex)}",parent=self.root)

    def delete(self):
        if self.var_s_id.get() == "":
            messagebox.showerror("Error", "s_id must be required", parent=self.root)
        else:
            con=pymysql.connect(host='localhost',user='root',password='',db='along_home_db')
            #con = sqlite3.connect(database=r'./ah.db')
            cur = con.cursor()
            cur.execute("SELECT * FROM staff WHERE s_id=%s", (self.var_s_id.get(),))
            row = cur.fetchone()
            if row is None:
                messagebox.showerror("Error", "Invalid s_id", parent=self.root)
            else:
                op = messagebox.askyesno("Confirm", "Do you really want to delete%s", parent=self.root)
                if op:
                    cur.execute("DELETE FROM staff WHERE s_id=%s", (self.var_s_id.get(),))
                    con.commit()
                    self.clear()
                    messagebox.showinfo("Delete", "Staff deleted successfully", parent=self.root)
    
    def clear(self):
        self.var_s_id.set("")
        self.var_new_staff_name.set("")
        self.var_birth_date.set("Select")
        self.var_age.set("")
        self.var_contact.set("")
        self.var_other_contact.set("")
        self.var_email_id.set("")
        self.var_cast.set("Select")
        self.var_marital_status.set("Select")
        self.var_gender.set("Select")
        self.var_religion.set("Select")
        self.var_nationality.set("Select")
        self.var_experience.set("")
        self.var_present_address.set("")
        self.var_permanent_address.set("")
        self.var_education.set("")
        self.var_idproofcomb.set("Select ID")
        self.var_id_proof.set("")
        self.var_nursing_certificate.set("")
        self.var_exp_letter.set("")
        self.var_light_bill.set("")
        self.var_last_work_place.set("")
        self.var_duty.set("Select")
        #self.var_reference_no.set("")
        self.var_financial_year.set("Select")
        self.root.after(100, lambda: self.txt_s_id.focus_set())
        self.var_searchtxt.set("")
        self.var_searchby.set("Select")        
        # Reset the image to default
        self.var_img_path_1.set("")
        img_1 = Image.open("./image/along_home_logo_1.png")  # Load default image1
        img_1 = img_1.resize((100, 100), Image.LANCZOS)
        img_1 = ImageTk.PhotoImage(img_1)        
        self.lbl_image_1.config(image=img_1)
        self.lbl_image_1.image = img_1
        # Reset the image to default
        self.var_img_path_2.set("")
        img_2 = Image.open("./image/along_home_logo_2.png")  # Load default image2
        img_2 = img_2.resize((100, 100), Image.LANCZOS)
        img_2 = ImageTk.PhotoImage(img_2)        
        self.lbl_image_2.config(image=img_2)
        self.lbl_image_2.image = img_2
        # Button control
        self.btn_add.config(state='normal')
        self.btn_update.config(state='disabled')
        self.btn_delete.config(state='disabled')
        #self.btn_export_search.config(state='disabled')
        
        self.show()
    
    def search(self):  
        con=pymysql.connect(host='localhost',user='root',password='',db='along_home_db')
        #con=sqlite3.connect(database=r'./ah.db')
        cur=con.cursor()
        try:
            if self.var_searchby.get()=="Select":
                messagebox.showerror("Error","Select Search by option",parent=self.root)
            elif self.var_searchtxt.get()=="":
                messagebox.showerror("Error","Search input should be required",parent=self.root)
            
            else:
                cur.execute("select * from staff where "+self.var_searchby.get()+" LIKE '%"+self.var_searchtxt.get()+"%'")
                rows=cur.fetchall()
                if len(rows)!=0:
                    self.staff_Table.delete(*self.staff_Table.get_children())
                    for row in rows:
                        self.staff_Table.insert('',END,values=row)
                else:
                    messagebox.showerror("Error","No record found!!!",parent=self.root)
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :{str(ex)}",parent=self.root)
   
    def send_whatsapp_message(self):
        phone_number = self.var_contact.get()  # Get staff phone number
        message = f"Hello {self.var_new_staff_name.get()}, thank you for being our valued customer!"

        if phone_number == "":
            messagebox.showerror("Error", "customer phone number is required!", parent=self.root)
            return

        try:
            # Send WhatsApp message instantly
            kit.sendwhatmsg_instantly(f"+91{phone_number}", message, wait_time=10)
            messagebox.showinfo("Success", "WhatsApp message sent successfully!", parent=self.root)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send WhatsApp message\n{str(e)}", parent=self.root)    
    
    def export_searched_staff_to_pdf(self):
        # Validate search
        if self.var_searchby.get() == "Select":
            messagebox.showerror("Error", "Select Search by option", parent=self.root)
            return
        if self.var_searchtxt.get() == "":
            messagebox.showerror("Error", "Search input should be required", parent=self.root)
            return
        
        rows = [self.staff_Table.item(child)['values'] for child in self.staff_Table.get_children()]
        if not rows:
            messagebox.showwarning("No Data", "There is no searched data to export.")
            return

        # Ask where to save
        pdf_file = filedialog.asksaveasfilename(defaultextension="./data/Staff Data/export.pdf", filetypes=[("PDF files", "*.pdf")])
        if not pdf_file:
            return

        os.makedirs(os.path.dirname(pdf_file), exist_ok=True)

        # Estimate table width
        column_widths = [32] * 25 + [40]
        total_width = sum(column_widths)
        portrait_width, _ = A4
        page_size = landscape(A4) if total_width > portrait_width else portrait(A4)

        doc = SimpleDocTemplate(pdf_file, pagesize=page_size, rightMargin=5, leftMargin=5, topMargin=10, bottomMargin=10)
        elements = []
        styles = getSampleStyleSheet()
        wrap_style = ParagraphStyle(name='Wrap', fontSize=6, alignment=1)

        # Fetch firm details
        try:
            con=pymysql.connect(host='localhost',user='root',password='',db='along_home_db')
            #con = sqlite3.connect(database=r"./ah.db")
            cur = con.cursor()
            cur.execute("SELECT name, contact, address, email, gst, bank, account_holder_name, account_no, branch_ifs_code FROM firm LIMIT 1")
            firm = cur.fetchone()
            if firm:
                firm_name, contact, address, email, gst, bank, acc_holder, acc_no, ifsc = firm
            else:
                firm_name = contact = address = email = gst = bank = acc_holder = acc_no = ifsc = "N/A"
            con.close()
        except Exception as e:
            firm_name = contact = address = email = gst = bank = acc_holder = acc_no = ifsc = "Error loading firm"

        # Add logo
        logo_path = "./image/along_home_logo_1.png"
        if os.path.exists(logo_path):
            elements.append(RLImage(logo_path, width=150, height=100))
            elements.append(Spacer(1, 14))

        # Firm Details
        elements.append(Paragraph(f"<b>{firm_name}</b>", styles['Title']))
        elements.append(Spacer(1, 10))
        elements.append(Paragraph(f"Contact: {contact}", styles['Normal']))
        elements.append(Paragraph(f"Email: {email}", styles['Normal']))
        elements.append(Paragraph(f"Address: {address}", styles['Normal']))
        elements.append(Paragraph(f"GST No: {gst}", styles['Normal']))
        elements.append(Spacer(1, 4))

        # Report Title
        elements.append(Paragraph("Searched Staff Data", styles['Heading2']))
        elements.append(Spacer(1, 6))
        elements.append(Paragraph(f"As of Date: {datetime.now().strftime('%d-%m-%Y')}", styles['Normal']))
        elements.append(Spacer(1, 8))

        # Table Header
        headers = ["S Id", "Staff Name", "Birth Date", "Age", "Contact", "2nd Contact", "Email ID", "Cast",
                "Marital Status", "Gender", "Religion", "Nationality", "Experience", "Present Address",
                "Permanent Address", "Education", "ID Type", "ID Proof", "Nursing Certificate", "Exp Letter",
                "Light Bill", "Last Work Place", "Duty", "Financial Year", "Image 1", "Image 2"]
        data = [headers]

        # Add data
        for row in rows:
            row_data = []
            for i, cell in enumerate(row):
                cell_str = str(cell).strip()
                if i == 24 and os.path.exists(cell_str):  # Image 1
                    try:
                        img = RLImage(cell_str, width=30, height=30)
                        row_data.append(img)
                    except:
                        row_data.append(Paragraph("Image error", wrap_style))
                elif i == 25 and os.path.exists(cell_str):  # Image 2
                    try:
                        img = RLImage(cell_str, width=30, height=30)
                        row_data.append(img)
                    except:
                        row_data.append(Paragraph("Image error", wrap_style))
                else:
                    row_data.append(Paragraph(cell_str, wrap_style))
            data.append(row_data)

        # Table styling
        table = Table(data, repeatRows=1, colWidths=column_widths)
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-2, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ])

        # Zebra striping
        for i in range(1, len(data)):
            bg_color = colors.whitesmoke if i % 2 else colors.lightgrey
            style.add('BACKGROUND', (0, i), (-1, i), bg_color)

        table.setStyle(style)
        elements.append(table)

        # Build PDF
        try:
            doc.build(elements)
            messagebox.showinfo("Success", f"PDF exported to:\n{pdf_file}", parent=self.root)
        except Exception as ex:
            messagebox.showerror("Error", f"Error during export:\n{str(ex)}", parent=self.root)
    
    def export_from_search(self):
        # Get all rows from Treeview
        rows = []
        for child in self.staff_Table.get_children():
            rows.append(self.staff_Table.item(child)['values'])

        if not rows:
            messagebox.showwarning("No Data", "There is no data to export.")
            return

        # Create DataFrame
        df = pd.DataFrame(rows, columns=["S Id","New Staff Name","Birth Date","Age","Contact","2nd Contact","Email ID","Cast","Marital Status","Gender","Religion","Nationality","Experience","Present Address","Permanent Address","Education","ID Type","ID Proof","Nursing Certificate","Exp Letter","Light Bill","Last Work Place","Duty","Financial Year","Image 1","Image 2"])
        # Add "As of Date" column
        as_of_date = datetime.now().strftime("%d-%m-%Y")  # Current date in YYYY-MM-DD format
        df["As of Date"] = as_of_date            # Export to Excel

        try:
            output_file =r"./data/Staff Data/searched_staff_data.xlsx"
            df.to_excel(output_file, index=False,engine='openpyxl')
            messagebox.showinfo("Success", f"Data exported to {output_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data\n{str(e)}")
        
        output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if output_file:
            df.to_excel(output_file, index=False)    
        
    def export_to_excel(self):
        con=pymysql.connect(host='localhost',user='root',password='',db='along_home_db')
        #con = sqlite3.connect(database=r'./ah.db')
        cur = con.cursor()
        try:
            cur.execute("SELECT * FROM staff")
            rows = cur.fetchall()
            if not rows:
                messagebox.showerror("Error", "No data to export", parent=self.root)
                return

            # Create a DataFrame from the fetched data
            df = pd.DataFrame(rows, columns=["S Id","New Staff Name","Birth Date","Age","Contact","2nd Contact","Email ID","Cast","Marital Status","Gender","Religion","Nationality","Experience","Present Address","Permanent Address","Education","ID Type","ID Proof","Nursing Certificate","Exp Letter","Light Bill","Last Work Place","Duty","Financial Year","Image 1","Image 2"])

            # Add "As of Date" column
            as_of_date = datetime.now().strftime("%Y-%m-%d")  # Current date in YYYY-MM-DD format
            df["As of Date"] = as_of_date            # Export to Excel
            
            output_file = r"./data/Staff Data/staff_data.xlsx"
            df.to_excel(output_file, index=False, engine='openpyxl')

            messagebox.showinfo("Success", f"Data exported to {output_file}", parent=self.root)

        except Exception as ex:
            messagebox.showerror("Error", f"Error due to : {str(ex)}", parent=self.root)        
        
    def export_to_pdf(self):
        con=pymysql.connect(host='localhost',user='root',password='',db='along_home_db')
        #con = sqlite3.connect(database=r'./ah.db')
        cur = con.cursor()
        try:
            cur.execute("SELECT * FROM staff")
            rows = cur.fetchall()
            if not rows:
                messagebox.showerror("Error", "No data to export", parent=self.root)
                return

            # Ask where to save
            pdf_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
            if not pdf_file:
                return

            os.makedirs(os.path.dirname(pdf_file), exist_ok=True)

            # Estimate table width to decide orientation
            column_widths = [32] * 25 + [40]
            total_width = sum(column_widths)

            # A4 sizes in points (1 pt = 1/72 inch)
            portrait_width, portrait_height = A4
            landscape_width, landscape_height = landscape(A4)

            # Choose orientation
            if total_width > portrait_width:
                page_size = landscape(A4)
            else:
                page_size = portrait(A4)

            doc = SimpleDocTemplate(pdf_file, pagesize=page_size, rightMargin=2, leftMargin=2)
            
            elements = []
            styles = getSampleStyleSheet()
            wrap_style = ParagraphStyle(name='Wrap', fontSize=6, alignment=1)

            # Fetch firm details
            try:
                con=pymysql.connect(host='localhost',user='root',password='',db='along_home_db')
                #con = sqlite3.connect(database=r"./ah.db")
                cur = con.cursor()
                cur.execute("SELECT name, contact, address, email, gst, bank, account_holder_name, account_no, branch_ifs_code FROM firm LIMIT 1")
                firm = cur.fetchone()
                if firm:
                    firm_name, contact, address, email, gst, bank, acc_holder, acc_no, ifsc = firm
                else:
                    firm_name = contact = address = email = gst = bank = acc_holder = acc_no = ifsc = "N/A"
                con.close()
            except Exception as e:
                firm_name = contact = address = email = gst = bank = acc_holder = acc_no = ifsc = "Error loading firm"

            # Add logo
            logo_path = "./image/along_home_logo_1.png"
            if os.path.exists(logo_path):
                elements.append(RLImage(logo_path, width=150, height=100))
                elements.append(Spacer(1, 14))

            # Firm Details
            elements.append(Paragraph(f"<b>{firm_name}</b>", styles['Title']))
            elements.append(Spacer(1, 10))
            elements.append(Paragraph(f"Contact: {contact}", styles['Normal']))
            elements.append(Paragraph(f"Email: {email}", styles['Normal']))
            elements.append(Paragraph(f"Address: {address}", styles['Normal']))
            elements.append(Paragraph(f"GST No: {gst}", styles['Normal']))
            elements.append(Spacer(1, 4))

            # Report Title
            elements.append(Paragraph("Searched Staff Data", styles['Heading2']))
            elements.append(Spacer(1, 6))
            elements.append(Paragraph(f"As of Date: {datetime.now().strftime('%d-%m-%Y')}", styles['Normal']))
            elements.append(Spacer(1, 8))

            # Header
            headers = ["S Id", "Staff Name", "Birth Date", "Age", "Contact", "2nd Contact", "Email ID", "Cast",
                    "Marital Status", "Gender", "Religion", "Nationality", "Experience", "Present Address",
                    "Permanent Address", "Education", "ID Type", "ID Proof", "Nursing Certificate", "Exp Letter",
                    "Light Bill", "Last Work Place", "Duty", "Financial Year", "Image 1", "Image 2"]

            data = [headers]

            for row in rows:
                row_data = []
                for i, cell in enumerate(row):
                    if i == 24 and cell and os.path.exists(cell):  # image1
                        try:
                            img_1 = RLImage(cell, width=30, height=30)
                            row_data.append(img_1)
                        except:
                            row_data.append(Paragraph("Image error", wrap_style))
                    elif i == 25 and cell and os.path.exists(cell):  # image2
                        try:
                            img_2 = RLImage(cell, width=30, height=30)
                            row_data.append(img_2)
                        except:
                            row_data.append(Paragraph("Image error", wrap_style))
                    else:
                        row_data.append(Paragraph(str(cell), wrap_style))
                data.append(row_data)
            
            table = Table(data, repeatRows=1, colWidths=[32]*25 + [40])  # image1 column slightly wider
            style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-2, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ])

            # Zebra striping
            for i in range(1, len(data)):
                bg = colors.whitesmoke if i % 2 else colors.lightgrey
                style.add('BACKGROUND', (0, i), (-1, i), bg)

            table.setStyle(style)
            elements.append(table)

            doc.build(elements)
            messagebox.showinfo("Success", f"PDF exported to:\n{pdf_file}", parent=self.root)

        except Exception as ex:
            messagebox.showerror("Error", f"Error due to: {str(ex)}", parent=self.root)    
    
    def import_excel_to_mysql(file_path):
        try:
            file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel Files", "*.xlsx")])
            if not file_path:
                return

            # Ensure correct engine for .xlsx
            df = pd.read_excel(file_path, engine='openpyxl')

            # Replace NaN with None for MySQL
            df = df.where(pd.notnull(df), None)

            # Connect to MySQL
            con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
            cur = con.cursor()

            for _, row in df.iterrows():
                cur.execute("""
                    INSERT INTO staff(new_staff_name, birth_date, age, contact, other_contact, email_id, cast, marital_status,
                                    gender, religion, nationality, experience, present_address, permanent_address,
                                    education, idproofcomb, id_proof, nursing_certificate, exp_letter, light_bill,
                                    last_work_place, duty, financial_year, image_1, image_2)
                    VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, (
                    row['new_staff_name'], row['birth_date'], row['age'], row['contact'], row['other_contact'],
                    row['email_id'], row['cast'], row['marital_status'], row['gender'], row['religion'],
                    row['nationality'], row['experience'], row['present_address'], row['permanent_address'],
                    row['education'], row['idproofcomb'], row['id_proof'], row['nursing_certificate'], row['exp_letter'],
                    row['light_bill'], row['last_work_place'], row['duty'], row['financial_year'],
                    row['image_1'], row['image_2']
                ))

            con.commit()
            con.close()
            messagebox.showinfo("Success", "Staff data imported successfully.")

        except Exception as ex:
            messagebox.showerror("Error", f"Failed to import Excel: {str(ex)}")
    
    def update_content(self):
        con=pymysql.connect(host='localhost',user='root',password='',db='along_home_db')
        #con=sqlite3.connect(database=r'./ah.db')
        cur=con.cursor()
        try:
            time_=time.strftime("%I:%M:%S %p")
            date_=time.strftime("%d-%m-%Y")
            self.lbl_clock.config(text=f"  ALONG HOME HEALTHCARE\t\t New Staff Hiring Form(Recruitment)\t\t\t Date: {str(date_)}\t\t Time: {str(time_)}")
            self.lbl_clock.after(200,self.update_content)
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :   {str(ex)}",parent=self.root)
        
if __name__=="__main__":
    root=Tk()
    obj=staffClass(root)
    root.mainloop()