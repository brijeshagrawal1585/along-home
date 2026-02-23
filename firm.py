from datetime import datetime
import os
from tkinter import*
from PIL import Image,ImageTk #pip install pilow
from tkinter import ttk, messagebox
from fpdf import FPDF, XPos, YPos
import pymysql
from tkcalendar import DateEntry  # Import DateEntry from tkintercalendar
import time
import re
import sqlite3
import pandas as pd
import pyodbc
class firmClass:
    def __init__(self,root):
        self.root=root
        self.root.geometry("1100x600+220+50")
        self.root.title("Along Home HealthCare | Developed by Brijesh | August | 9998002712")
        self.root.config(bg="white")
        self.root.focus_force()
        # All Variables================
        self.var_searchby=StringVar()
        self.var_searchtxt=StringVar()

        self.var_f_id=StringVar()
        self.var_name=StringVar()
        self.var_contact=StringVar()
        self.var_address=StringVar()
        self.var_email=StringVar()
        self.var_gst=StringVar()
        self.var_bank=StringVar()
        self.var_account_holder_name=StringVar()
        self.var_account_no=StringVar()
        self.var_branch_ifs_code=StringVar()
        
        self.var_img_path_1 = StringVar()  # Variable to store image1 path
        self.default_img_1 = Image.open("./image/along_home_logo_1.png")  # Default image1 placeholder
        self.default_img_1 = self.default_img_1.resize((100, 100), Image.LANCZOS)
        self.default_img_1 = ImageTk.PhotoImage(self.default_img_1)            
        #============Search Frame==================
        SearchFrame=LabelFrame(self.root,text="Search Firm Data",font=("helvetica",12,"bold"),bd=2,relief=RIDGE,bg="white")
        SearchFrame.place(x=530,y=200,width=360,height=65)

        cmb_search=ttk.Combobox(SearchFrame,textvariable=self.var_searchby,values=("Select","F_ID","NAME"),state='readonly',justify=CENTER,font=("helvetica",15))
        cmb_search.place(x=10,y=10,width=120)
        cmb_search.current(0)

        txt_search=Entry(SearchFrame,textvariable=self.var_searchtxt,font=("helvetica",15),bg="lightyellow").place(x=140,y=10,width=100)
        btn_search=Button(SearchFrame,text="Search",command=self.search,font=("helvetica",15),bg="black",fg="white",cursor="hand2")
        btn_search.place(x=250,y=9,width=100,height=30)
        #===========Title===================
        title=Label(self.root,text="Firm Details",font=("helvetica",20,"bold"),bg="#0f4d7d",fg="white").place(x=50,y=10,width=1000,height=40)
        #===========Content================
        lbl_name=Label(self.root,text="Name",font=("helvetica",15),bg="white").place(x=50,y=70)
        self.txt_name=Entry(self.root,textvariable=self.var_name,font=("helvetica",15),bg="lightyellow")
        self.txt_name.place(x=200,y=70,width=180)
        
        lbl_contact=Label(self.root,text="Contact",font=("helvetica",15),bg="white")
        lbl_contact.place(x=50,y=110)
        txt_contact=Entry(self.root,textvariable=self.var_contact,font=("helvetica",15),bg="lightyellow")
        txt_contact.place(x=200,y=110,width=180)
        #txt_contact.bind("<FocusOut>", self.validate_contact)  # Call validate_phone when focus is lost        
        
        lbl_address=Label(self.root,text="Address",font=("helvetica",15),bg="white").place(x=50,y=150)
        self.txt_address=Text(self.root,font=("helvetica",15),bg="lightyellow")
        self.txt_address.place(x=200,y=150,width=300,height=70)
                
        lbl_email=Label(self.root,text="Email ID",font=("helvetica",15),bg="white").place(x=50,y=240)
        txt_email=Entry(self.root,textvariable=self.var_email,font=("helvetica",15),bg="lightyellow").place(x=200,y=240,width=180)
        
        lbl_gst=Label(self.root,text="GST No.",font=("helvetica",15),bg="white").place(x=50,y=280)
        txt_gst=Entry(self.root,textvariable=self.var_gst,font=("helvetica",15),bg="lightyellow").place(x=200,y=280,width=180)
        
        lbl_bank=Label(self.root,text="Bank Name",font=("helvetica",15),bg="white").place(x=50,y=360)
        txt_bank=Entry(self.root,textvariable=self.var_bank,font=("helvetica",15),bg="lightyellow").place(x=200,y=360,width=180)
        
        lbl_account_holder_name=Label(self.root,text="A/C Name",font=("helvetica",15),bg="white").place(x=50,y=400)
        txt_account_holder_name=Entry(self.root,textvariable=self.var_account_holder_name,font=("helvetica",15),bg="lightyellow").place(x=200,y=400,width=180)
        
        lbl_account_no=Label(self.root,text="A/C No.",font=("helvetica",15),bg="white").place(x=50,y=450)
        txt_account_no=Entry(self.root,textvariable=self.var_account_no,font=("helvetica",15),bg="lightyellow").place(x=200,y=450,width=180)
        
        lbl_branch_ifs_code=Label(self.root,text="Branch/IFS Code",font=("helvetica",15),bg="white").place(x=50,y=490)
        txt_branch_ifs_code=Entry(self.root,textvariable=self.var_branch_ifs_code,font=("helvetica",15),bg="lightyellow").place(x=200,y=490,width=180)
        #==========Upload Photo=========================
        self.var_img_path_1 = StringVar()  # Variable to store image1 path
        self.default_img_1 = Image.open("./image/along_home_logo_1.png")  # Default image1 placeholder
        self.default_img_1 = self.default_img_1.resize((100, 100), Image.LANCZOS)
        self.default_img_1 = ImageTk.PhotoImage(self.default_img_1)   
        
        self.lbl_image_1 = Label(self.root, image=self.default_img_1, bd=2, relief=RIDGE)
        self.lbl_image_1.place(x=530, y=60, width=150, height=120)        
        #==========Buttons=========
        self.btn_add=Button(self.root,text="Save",command=self.add,font=("helvetica",15),bg="#2196f3",fg="white",cursor="hand2")
        self.btn_add.place(x=50,y=530,width=110,height=35)
        
        self.btn_update=Button(self.root,text="Update",command=self.update,font=("helvetica",15),bg="#4caf50",fg="white",cursor="hand2")
        self.btn_update.place(x=170,y=530,width=110,height=35)
        
        self.btn_delete=Button(self.root,text="Delete",command=self.delete,font=("helvetica",15),bg="#f44336",fg="white",cursor="hand2")
        self.btn_delete.place(x=290,y=530,width=110,height=35)
        
        self.btn_clear=Button(self.root,text="Clear",command=self.clear,font=("helvetica",15),bg="#607d8b",fg="white",cursor="hand2")
        self.btn_clear.place(x=410,y=530,width=110,height=35)

        self.btn_add.config(state='normal')
        self.btn_update.config(state='disabled')
        self.btn_delete.config(state='disabled')
        #self.btn_export.config(state='disabled')
        
        self.btn_export_to_pdf = Button(self.root, text="Export to PDF", command=self.export_to_pdf, font=("helvetica", 15), bg="red", fg="white", cursor="hand2")
        self.btn_export_to_pdf.place(x=900, y=190, width=150, height=30)
        
        self.btn_export = Button(self.root, text="Export", command=self.export, font=("helvetica", 15), bg="yellow", fg="black", cursor="hand2")
        self.btn_export.place(x=900, y=230, width=150, height=30)
        #=========firm Details========
        firm_frame=Frame(self.root,bd=3,relief=RIDGE,bg="green")
        firm_frame.place(x=530,y=270,width=550,height=300)

        scrolly=Scrollbar(firm_frame,orient=VERTICAL)
        scrollx=Scrollbar(firm_frame,orient=HORIZONTAL) 
        
        self.firmTable=ttk.Treeview(firm_frame,columns=("f_id","name","contact","address","email","gst","bank","account_holder_name","account_no","branch_ifs_code"),yscrollcommand=scrolly.set,xscrollcommand=scrollx.set)
        scrollx.pack(side=BOTTOM,fill=X)
        scrolly.pack(side=RIGHT,fill=Y)
        scrollx.config(command=self.firmTable.xview)
        scrolly.config(command=self.firmTable.yview)

        self.firmTable.heading("f_id",text="F ID")
        self.firmTable.heading("name",text="Name")
        self.firmTable.heading("contact",text="Contact")
        self.firmTable.heading("address",text="Address")
        self.firmTable.heading("email",text="Email ID")
        self.firmTable.heading("gst",text="GST No.")
        self.firmTable.heading("bank",text="Bank Name")
        self.firmTable.heading("account_holder_name",text="A/C Name")
        self.firmTable.heading("account_no",text="A/C No")
        self.firmTable.heading("branch_ifs_code",text="Branch & IFS Code")

        self.firmTable["show"]="headings"
        
        self.firmTable.column("f_id",width=30)
        self.firmTable.column("name",width=100)
        self.firmTable.column("contact",width=60)
        self.firmTable.column("address",width=100)
        self.firmTable.column("email",width=50)
        self.firmTable.column("gst",width=50)
        self.firmTable.column("bank",width=30)
        self.firmTable.column("account_holder_name",width=30)
        self.firmTable.column("account_no",width=30)
        self.firmTable.column("branch_ifs_code",width=30)
        self.firmTable.pack(fill=BOTH,expand=1)
        self.firmTable.bind("<ButtonRelease-1>",self.get_data)

        self.show()
#=======================================================================================================================================================================
    def add(self):
        conn = pyodbc.connect(
                r"DRIVER={ODBC Driver 17 for SQL Server};"
                r"SERVER=DESKTOP-9VTU0IO\SQLEXPRESS;"
                r"DATABASE=ALONGHOME;"
                r"Trusted_Connection=yes;"
                r"Encrypt=no;"
        )
        cur = conn.cursor()
        #con=sqlite3.connect(database=r'./ah.db')
        #cur=con.cursor()
        try:
            if self.var_bank.get()=="Select" or self.var_name.get()=="Select" or self.var_contact.get()=="Empty": 
                messagebox.showerror("Error","bank must be required",parent=self.root)
            else:
                cur.execute("Select * from firm where bank=? and name=? and contact=? and gst=?",(self.var_bank.get(),self.var_name.get(),self.var_contact.get(),self.var_gst.get(),))
                row=cur.fetchone()
                if row!=None:
                    messagebox.showerror("Error","bank is already assigned, try diffrent",parent=self.root)
                else:
                    cur.execute("Insert into firm(name,contact,address,email,gst,bank,account_holder_name,account_no,branch_ifs_code) values(?,?,?,?,?,?,?,?,?)",(
                                        self.var_name.get(),
                                        self.var_contact.get(),
                                        self.txt_address.get('1.0',END),
                                        self.var_email.get(),
                                        self.var_gst.get(),
                                        self.var_bank.get(),
                                        self.var_account_holder_name.get(),
                                        self.var_account_no.get(),
                                        self.var_branch_ifs_code.get(),
                    ))
                    conn.commit()
                    messagebox.showinfo("Success","Firm Added Successfully",parent=self.root)
                    self.clear()   # Clear fields after saving 
                    self.show()
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :{str(ex)}",parent=self.root)

    #def validate_contact(self, event=None):
        #contact_number = self.var_contact.get()
        
        # Regular expression for a valid Indian phone number
        #pattern = r"^[7-9][0-9]{9}$"  # Starts with 7-9, followed by 9 digits
        
        #if re.match(pattern, contact_number):  # If the number matches the pattern
            #print("Valid phone number!")
        #else:
            # Show an error message if the number is not valid
            #messagebox.showerror("Invalid Phone Number", "Please enter a valid 10-digit mobile number starting with 7, 8, or 9.")
            #self.var_contact.set("")  # Clear the invalid phone number field
        
    def show(self):
        conn = pyodbc.connect(
                r"DRIVER={ODBC Driver 17 for SQL Server};"
                r"SERVER=DESKTOP-9VTU0IO\SQLEXPRESS;"
                r"DATABASE=ALONGHOME;"
                r"Trusted_Connection=yes;"
                r"Encrypt=no;"
        )
        cur = conn.cursor()
        try:
            cur.execute("Select * from firm")
            rows=cur.fetchall()
            self.firmTable.delete(*self.firmTable.get_children())
            for row in rows:
                self.firmTable.insert('',END,values=row)

        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :{str(ex)}",parent=self.root)

    def get_data(self,ev):
        f=self.firmTable.focus()
        content=(self.firmTable.item(f))
        row=content['values']
        # print(row)  
        self.var_f_id.set(row[0])
        self.var_name.set(row[1])
        self.var_contact.set(row[2])
        self.txt_address.delete('1.0',END)   
        self.txt_address.insert(END,row[3])
        self.var_email.set(row[4])
        self.var_gst.set(row[5])
        self.var_bank.set(row[6])
        self.var_account_holder_name.set(row[7])
        self.var_account_no.set(row[8])
        self.var_branch_ifs_code.set(row[9])

        # Button control
        self.btn_add.config(state='disabled')
        self.btn_update.config(state='normal')
        self.btn_delete.config(state='normal')
        #self.btn_export.config(state='normal')
    
    def update(self):
        conn = pyodbc.connect(
                r"DRIVER={ODBC Driver 17 for SQL Server};"
                r"SERVER=DESKTOP-9VTU0IO\SQLEXPRESS;"
                r"DATABASE=ALONGHOME;"
                r"Trusted_Connection=yes;"
                r"Encrypt=no;"
        )
        cur = conn.cursor()
        #con=sqlite3.connect(database=r'./ah.db')
        #cur=con.cursor()
        try:
            if self.var_f_id.get()=="": 
                messagebox.showerror("Error","f_id Must be required",parent=self.root)
            else:
                cur.execute("Select * from firm where f_id=?",(self.var_f_id.get(),))
                row=cur.fetchone()
                if row==None:
                    messagebox.showerror("Error","Invalid f_id",parent=self.root)
                else:
                    cur.execute("Update firm set name=?,contact=?,address=?,email=?,gst=?,bank=?,account_holder_name=?,account_no=?,branch_ifs_code=? where f_id=?",(                                   
                                        self.var_name.get(),
                                        self.var_contact.get(),
                                        self.txt_address.get('1.0',END), 
                                        self.var_email.get(),
                                        self.var_gst.get(),
                                        self.var_bank.get(),
                                        self.var_account_holder_name.get(),
                                        self.var_account_no.get(),
                                        self.var_branch_ifs_code.get(),
                                        self.var_f_id.get(),
                    ))
                    conn.commit()
                    messagebox.showinfo("Success","Firm Updated Successfully",parent=self.root)
                    self.clear()   # Clear fields after saving 
                    self.show()
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :{str(ex)}",parent=self.root)
    
    def delete(self):
        conn = pyodbc.connect(
                r"DRIVER={ODBC Driver 17 for SQL Server};"
                r"SERVER=DESKTOP-9VTU0IO\SQLEXPRESS;"
                r"DATABASE=ALONGHOME;"
                r"Trusted_Connection=yes;"
                r"Encrypt=no;"
        )
        cur = conn.cursor()
        try:
            if self.var_f_id.get()=="":
                messagebox.showerror("Error","f_id must be required",parent=self.root)
            else:
                cur.execute("Select * from firm where f_id=?",(self.var_f_id.get(),))
                row=cur.fetchone()
                if row==None:
                    messagebox.showerror("Error","Invalid f_id",parent=self.root)
                else:
                    op=messagebox.askyesno("Confirm","Do you really want to delete?",parent=self.root)
                    if op==True:
                        cur.execute("Delete from firm where f_id=?",(self.var_f_id.get(),))
                        conn.commit()
                    messagebox.showinfo("Delete","Firm Deleted Successfully",parent=self.root)
                    self.clear()
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to : {str(ex)}",parent=self.root)                          

        self.clear()   # Clear fields after saving 
    
    def clear(self):
        self.var_name.set("")
        self.var_contact.set("")
        self.txt_address.delete('1.0',END) 
        self.var_email.set("")
        self.var_gst.set("")
        self.var_bank.set("")
        self.var_account_holder_name.set("")
        self.var_account_no.set("")
        self.var_branch_ifs_code.set("")
        self.var_f_id.set("")
        self.var_searchtxt.set("")
        self.var_searchby.set("Select")        
        self.root.after(100, lambda: self.txt_name.focus_set())        
        # Button control
        self.btn_add.config(state='normal')
        self.btn_update.config(state='disabled')
        self.btn_delete.config(state='disabled')
        #self.btn_export.config(state='disabled')
               
        self.show()

    def search(self):  
        conn = pyodbc.connect(
                r"DRIVER={ODBC Driver 17 for SQL Server};"
                r"SERVER=DESKTOP-9VTU0IO\SQLEXPRESS;"
                r"DATABASE=ALONGHOME;"
                r"Trusted_Connection=yes;"
                r"Encrypt=no;"
        )
        cur = conn.cursor()
        try:
            if self.var_searchby.get()=="Select":
                messagebox.showerror("Error","Select Search by option",parent=self.root)
            elif self.var_searchtxt.get()=="":
                messagebox.showerror("Error","Search input should be required",parent=self.root)
            
            else:
                cur.execute("select * from firm where "+self.var_searchby.get()+" LIKE '%"+self.var_searchtxt.get()+"%'")
                rows=cur.fetchall()
                if len(rows)!=0:
                    self.firmTable.delete(*self.firmTable.get_children())
                    for row in rows:
                        self.firmTable.insert('',END,values=row)
                else:
                    messagebox.showerror("Error","No record found!!!",parent=self.root)
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :{str(ex)}",parent=self.root)
    
    def export_to_pdf(self):
        conn = pyodbc.connect(
                r"DRIVER={ODBC Driver 17 for SQL Server};"
                r"SERVER=DESKTOP-9VTU0IO\SQLEXPRESS;"
                r"DATABASE=ALONGHOME;"
                r"Trusted_Connection=yes;"
                r"Encrypt=no;"
        )
        cur = conn.cursor()
        try:
            # Fetch one firm only
            cur.execute("SELECT * FROM firm LIMIT 1")
            row = cur.fetchone()
            if not row:
                messagebox.showerror("Error", "No firm data found", parent=self.root)
                return

            # Define labels (skip firm_id if it's first)
            labels = [
                "Firm Name", "Contact No", "Address", "Email",
                "GST No", "Bank Name", "Account Holder Name",
                "Account No", "Branch IFSC Code"
            ]

            values = list(row)[1:]  # skip firm_id

            pdf = FPDF(orientation='P', unit='mm', format='A4')
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)

            # Header
            pdf.set_font("helvetica", 'B', 14)
            #pdf.cell(0, 10, "Firm Data Report", ln=True, align='C')
            pdf.cell(0, 10, "Firm Data Report", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
            #pdf.set_font("helvetica", '', 10)
            #pdf.cell(0, 10, f"As of Date: {datetime.now().strftime('%d-%m-%Y')}", ln=True, align='R')
            pdf.ln(5)

            # Each field on 2 lines: label and value
            for label, value in zip(labels, values):
                value = str(value) if value is not None else ""
                pdf.set_font("helvetica", 'B', 11)
                pdf.cell(0, 7, label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                pdf.set_font("helvetica", '', 11)
                pdf.multi_cell(0, 7, value)
                pdf.ln(2)

            # Save to file
            output_dir = r"./data/Firm Data"
            os.makedirs(output_dir, exist_ok=True)
            output_file = os.path.join(output_dir, "Firm Data.pdf")
            pdf.output(output_file)

            messagebox.showinfo("Success", f"Firm data exported to:\n{output_file}", parent=self.root)

        except Exception as ex:
            messagebox.showerror("Error", f"Error due to: {str(ex)}", parent=self.root)

    
    def export(self):
        conn = pyodbc.connect(
                r"DRIVER={ODBC Driver 17 for SQL Server};"
                r"SERVER=DESKTOP-9VTU0IO\SQLEXPRESS;"
                r"DATABASE=ALONGHOME;"
                r"Trusted_Connection=yes;"
                r"Encrypt=no;"
        )
        cur = conn.cursor()
        try:
            cur.execute("SELECT * FROM firm")
            rows = cur.fetchall()
            if not rows:
                messagebox.showerror("Error", "No data to export", parent=self.root)
                return

            # Create a DataFrame from the fetched data
            df = pd.DataFrame(rows, columns=["f_id","name","contact","address","email","gst","bank","account_holder_name","account_no","branch_ifs_code"])

            # Add "As of Date" column
            as_of_date = datetime.now().strftime("%d-%m-%Y")  # Current date in YYYY-MM-DD format
            df["As of Date"] = as_of_date
            # Export to Excel
            output_file = r"./data/Firm Data/export.xlsx"
            df.to_excel(output_file, index=False, engine='openpyxl')

            messagebox.showinfo("Success", f"Data exported to {output_file}", parent=self.root)

        except Exception as ex:
            messagebox.showerror("Error", f"Error due to : {str(ex)}", parent=self.root)

if __name__=="__main__":
    root=Tk()
    obj=firmClass(root)
    root.mainloop()