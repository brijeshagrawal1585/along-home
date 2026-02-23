from datetime import datetime, timezone
import time
from tkinter import*
from PIL import Image,ImageTk #pip install pilow
from tkinter import ttk,messagebox
import pymysql
import pywhatkit as kit
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
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
import os
from fpdf import FPDF, XPos, YPos  # pip install fpdf
from collections import defaultdict
from num2words import num2words  # pip install num2words

class cash_invoiceClass:
    def __init__(self,root):
        self.root=root
        self.root.geometry("1350x700+0+0")
        self.root.title("Along Home Healthcare | Developed by Brijesh | July | 9998002712")
        self.root.config(bg="white")
        self.root.focus_force()
        # All Variables================
        vcmd = (self.root.register(self.validate_length), '%P')
        
        self.var_searchby=StringVar()
        self.var_searchtxt=StringVar()
        
        self.customer_list=[] 
        self.fetch_customer() 

        self.var_c_id=StringVar(value="0")                   #c_id
        self.var_bill_no=StringVar()                   #date
        self.var_date=StringVar()                   #date
        self.var_customer=StringVar()
        self.var_service_name=StringVar()               #truck_no     service_name 
        self.var_hsn=StringVar()                 #challan_no    hsn_sac
        
        #self.var_period=StringVar()
        self.var_quantity=StringVar(value="0")             #diesel_challan_no       quantity
        self.var_unit=StringVar()                 #challan_no    hsn_sac
        self.var_rate=StringVar(value="0")        
        self.var_cgst_rate=StringVar(value="0")        
        self.var_cgst_amount=StringVar(value="0")        
        
        self.var_sgst_rate=StringVar(value="0")        
        self.var_sgst_amount=StringVar(value="0")        
        self.var_amount=StringVar(value="0")         #amount_due     amount
        #============Search Frame==================
        SearchFrame=LabelFrame(self.root,text="Search Cash Invoice Data",font=("Arial",12,"bold"),bd=2,relief=RIDGE,bg="white")
        SearchFrame.place(x=10,y=20,width=750,height=70)

        cmb_search=ttk.Combobox(SearchFrame,textvariable=self.var_searchby,values=("Select","C_ID","BILL_NO","DATE","CUSTOMER"),state='readonly',justify=CENTER,font=("Arial",15))
        cmb_search.place(x=5,y=10,width=120)
        cmb_search.current(0)

        txt_search=Entry(SearchFrame,textvariable=self.var_searchtxt,font=("Arial",15),bg="lightyellow").place(x=140,y=10,width=100)
        btn_search=Button(SearchFrame,text="Search",command=self.search,font=("Arial",15),bg="black",fg="white",cursor="hand2").place(x=260,y=9,width=80,height=30)
                
        self.btn_export_search = Button(SearchFrame, text="Export Search to PDF", command=self.export_searched_cash_invoice_to_pdf, font=("Arial", 12,"bold"), bg="red", fg="white",cursor="hand2")
        self.btn_export_search.place(x=470,y=7,width=250)
        
        ExportFrame=LabelFrame(self.root,text="Export Data",font=("Arial",12,"bold"),bd=2,relief=RIDGE,bg="white")
        ExportFrame.place(x=770,y=20,width=550,height=70)
        
        btn_export_excel = Button(ExportFrame, text="Export to Excel", command=self.export_to_excel, font=("Arial", 15, "bold"), bg="yellow", fg="black", cursor="hand2")
        btn_export_excel.place(x=5, y=10, width=160, height=28)        
        
        btn_export_pdf = Button(ExportFrame, text="Export to PDF", command=self.export_to_pdf, font=("Arial", 15, "bold"), bg="purple", fg="white", cursor="hand2")
        btn_export_pdf.place(x=180, y=10, width=150, height=28)
        
        btn_import=Button(ExportFrame, text="Import from Excel", command=self.import_excel_to_db, font=("Arial", 15,"bold"), bg="brown", fg="white", cursor="hand2")
        btn_import.place(x=340, y=10,width=200,height=28)
        #===========Title===================
        #title=Label(self.root,text="vendor_code_part Details",font=("Arial",15),bg="#0f4d7d",fg="white").place(x=50,y=100,width=1000)
        #=========Clock==============
        self.lbl_clock=Label(self.root,text="Welcome to ALONG HOME HEALTHCARE \t\t Date: DD-MM-YYYY\t\t Time: HH:MM:SS",font=("times new roman",15),bg="#0f4d7d",fg="white")
        self.lbl_clock.place(x=20,y=100,width=1300,height=30)
        #===========Content================
        #=========Row 1===============================================================================================
        lbl_c_id=Label(self.root,text="C ID",font=("Arial",13),bg="white").place(x=20,y=150)
        self.txt_c_id=Entry(self.root,textvariable=self.var_c_id,justify=CENTER,state='readonly',font=("Arial",13),bg="lightyellow",validate="key", validatecommand=vcmd)
        self.txt_c_id.place(x=130,y=150,width=80)
        
        lbl_bill_no=Label(self.root,text="Bill No",font=("Arial",13),bg="white").place(x=240,y=150)
        self.txt_bill_no=Entry(self.root,textvariable=self.var_bill_no,justify=CENTER,font=("Arial",13),bg="lightyellow",validate="key", validatecommand=vcmd)
        self.txt_bill_no.place(x=330,y=150,width=110)
        
        lbl_date=Label(self.root,text="Date",font=("Arial",13),bg="white").place(x=460,y=150)
        # Date picker widget
        txt_date = DateEntry(self.root,textvariable=self.var_date,justify=CENTER,font=("Arial", 12),bg="lightyellow",date_pattern="dd-mm-yyyy")  # Set the format to dd/mm/yyyy
        txt_date.place(x=570,y=150,width=110)
        # Display the default date in the variable
        self.var_date.set(txt_date.get_date().strftime("%d-%m-%Y"))
        
        lbl_customer=Label(self.root,text="Customer",font=("Arial",13),bg="white").place(x=700,y=150)
        #txt_customer=Entry(self.root,textvariable=self.var_customer,justify=CENTER,font=("Arial",12),bg="lightyellow").place(x=420,y=150,width=110)
        cmb_customer=ttk.Combobox(self.root,textvariable=self.var_customer,values=self.customer_list,state='readonly',justify=CENTER,font=("Arial",12))
        cmb_customer.place(x=870,y=150,width=90)
        cmb_customer.current(0)        
        
        lbl_service_name=Label(self.root,text="Service Name",font=("Arial",13),bg="white").place(x=970,y=150)
        self.txt_service_name=Entry(self.root,textvariable=self.var_service_name,justify=CENTER,font=("Arial",12),bg="lightyellow",validate="key", validatecommand=vcmd)
        self.txt_service_name.place(x=1140,y=150,width=180)
        self.txt_service_name.bind("<KeyRelease>", self.show_service_name_suggestions)  # When focus leaves the service_name field
        #===========Row 2============================================================================================
        lbl_hsn=Label(self.root,text="HSN/SAC",font=("Arial",13),bg="white").place(x=20,y=190)
        self.txt_hsn=Entry(self.root,textvariable=self.var_hsn,justify=CENTER,font=("Arial",12),bg="lightyellow",validate="key", validatecommand=vcmd)
        self.txt_hsn.place(x=130,y=190,width=80)
        self.txt_hsn.bind("<KeyRelease>", self.show_hsn_suggestions)  # When focus leaves the hsn field
        
        lbl_quantity=Label(self.root,text="Quantity",font=("Arial",13),bg="white").place(x=240,y=190)
        #txt_quantity=Entry(self.root,textvariable=self.var_quantity,justify=CENTER,font=("Arial",12),bg="lightyellow",validate="key", validatecommand=vcmd).place(x=110,y=190,width=80)
        self.txt_quantity = Entry(self.root, textvariable=self.var_quantity,justify=CENTER,font=("Arial",12),bg="lightyellow",validate="key", validatecommand=vcmd)
        self.txt_quantity.place(x=330,y=190,width=110)
        self.txt_quantity.bind("<KeyRelease>", self.recalculate_tax_and_total)
        
        lbl_unit=Label(self.root,text="Unit",font=("Arial",13),bg="white").place(x=460,y=190) 
        self.txt_unit=Entry(self.root,textvariable=self.var_unit,justify=CENTER,font=("Arial",12),bg="lightyellow",validate="key", validatecommand=vcmd)
        self.txt_unit.place(x=570,y=190,width=110)   
        self.txt_unit.bind("<KeyRelease>", self.show_unit_suggestions)  # When focus leaves the service_name field
        
        lbl_rate=Label(self.root,text="Rate (RS)",font=("Arial",13),bg="white").place(x=700,y=190)
        #txt_rate=Entry(self.root,textvariable=self.var_rate,justify=CENTER,font=("Arial",12),bg="lightyellow",validate="key", validatecommand=vcmd).place(x=570,y=190,width=110)
        self.txt_rate = Entry(self.root, textvariable=self.var_rate,justify=CENTER,font=("Arial",12),bg="lightyellow",validate="key", validatecommand=vcmd)
        self.txt_rate.place(x=870,y=190,width=90)
        self.txt_rate.bind("<KeyRelease>", self.recalculate_tax_and_total)        
        
        lbl_cgst_rate=Label(self.root,text="CGST Rate",font=("Arial",13),bg="white").place(x=970,y=190)
        self.txt_cgst_rate = Entry(self.root, textvariable=self.var_cgst_rate,justify=CENTER,font=("Arial", 12),bg="lightyellow",validate="key", validatecommand=vcmd)
        self.txt_cgst_rate.place(x=1140,y=190,width=80)
        self.txt_cgst_rate.bind("<KeyRelease>", self.recalculate_tax_and_total)        
        #===========Row 3============================================================================================
        lbl_cgst_amount=Label(self.root,text="CGST Amount",font=("Arial",13),bg="white").place(x=20,y=230)
        self.txt_cgst_amount=Entry(self.root,textvariable=self.var_cgst_amount,justify=CENTER,font=("Arial",12),bg="lightyellow",validate="key", validatecommand=vcmd)
        self.txt_cgst_amount.place(x=130,y=230,width=80)
        
        lbl_sgst_rate=Label(self.root,text="SGST Rate",font=("Arial",13),bg="white").place(x=240,y=230)
        self.txt_sgst_rate = Entry(self.root, textvariable=self.var_sgst_rate,justify=CENTER,font=("Arial",12),bg="lightyellow",validate="key", validatecommand=vcmd)
        self.txt_sgst_rate.place(x=330,y=230,width=110)
        self.txt_sgst_rate.bind("<KeyRelease>", self.recalculate_tax_and_total)        
        
        lbl_sgst_amount=Label(self.root,text="SGST Amount",font=("Arial",13),bg="white").place(x=460,y=230)
        self.txt_sgst_amount=Entry(self.root,textvariable=self.var_sgst_amount,justify=CENTER,font=("Arial",12),bg="lightyellow",validate="key", validatecommand=vcmd)
        self.txt_sgst_amount.place(x=570,y=230,width=110)
                
        lbl_amount=Label(self.root,text="Amount (RS)",font=("Arial",13),bg="white").place(x=700,y=230)
        txt_amount=Entry(self.root,textvariable=self.var_amount,justify=CENTER,font=("Arial",12),bg="lightyellow",validate="key", validatecommand=vcmd).place(x=870,y=230,width=90)
        #==========Buttons==================================================================================================
        self.btn_add=Button(self.root,text="Save",command=self.add,font=("Arial",15),bg="#2196f3",fg="white",cursor="hand2")
        self.btn_add.place(x=50,y=270,width=110,height=28)
        
        self.btn_update=Button(self.root,text="Update",command=self.update,font=("Arial",15),bg="#4caf50",fg="white",cursor="hand2")
        self.btn_update.place(x=170,y=270,width=110,height=28)
        
        self.btn_delete=Button(self.root,text="Delete",command=self.delete,font=("Arial",15),bg="#f44336",fg="white",cursor="hand2")
        self.btn_delete.place(x=290,y=270,width=110,height=28)
        
        self.btn_clear=Button(self.root,text="Clear",command=self.clear,font=("Arial",15),bg="#607d8b",fg="white",cursor="hand2")
        self.btn_clear.place(x=410,y=270,width=110,height=28)
        
        self.btn_add.config(state='normal')
        self.btn_update.config(state='disabled')
        self.btn_delete.config(state='disabled')
        #self.btn_export_search.config(state='disabled')

        btn_whatsapp = Button(self.root, text="Send WhatsApp", command=self.send_whatsapp_message, font=("Arial", 15,"bold"), bg="green", fg="white", cursor="hand2")
        btn_whatsapp.place(x=530, y=270, width=170, height=28)        
        #=========Footer==============
        lbl_footer=Label(self.root,text="ALONG HOME HEALTHCARE | 59/2, Ground Floor, Govt. H. Colony, B/h. Laxmi Ganthiya Rath, Nehrunagar Cross Road, Ahmedabad-380015 Contact:+91 9904110283",font=("times new roman",15),bg="#4d636d",fg="white").pack(side=BOTTOM,fill=X)
        #=========vendor_code_part Details========
        cash_invoice_frame=LabelFrame(self.root,text="Cash Invoice List",font=("Arial",12,"bold"),bd=3,relief=RIDGE,bg="white")
        cash_invoice_frame.place(x=0,y=310,relwidth=1,height=350)

        scrolly=Scrollbar(cash_invoice_frame,orient=VERTICAL)
        scrollx=Scrollbar(cash_invoice_frame,orient=HORIZONTAL) 
        
        self.cash_invoice_Table=ttk.Treeview(cash_invoice_frame,columns=("c_id","bill_no","date","customer","service_name","hsn","quantity","unit","rate","cgst_rate","cgst_amount","sgst_rate","sgst_amount","amount"),yscrollcommand=scrolly.set,xscrollcommand=scrollx.set)
        scrollx.pack(side=BOTTOM,fill=X)
        scrolly.pack(side=RIGHT,fill=Y)
        scrollx.config(command=self.cash_invoice_Table.xview)
        scrolly.config(command=self.cash_invoice_Table.yview)

        self.cash_invoice_Table.heading("c_id",text="C ID")
        self.cash_invoice_Table.heading("bill_no",text="Bill No")
        self.cash_invoice_Table.heading("date",text="Date")
        self.cash_invoice_Table.heading("customer",text="Customer")
        self.cash_invoice_Table.heading("service_name",text="Service Name")
        self.cash_invoice_Table.heading("hsn",text="HSN/SAC")
        self.cash_invoice_Table.heading("quantity",text="Quantity")
        self.cash_invoice_Table.heading("unit",text="Unit")
        self.cash_invoice_Table.heading("rate",text="Rate (RS)")
        self.cash_invoice_Table.heading("cgst_rate",text="CGST Rate (RS)")
        self.cash_invoice_Table.heading("cgst_amount",text="CGST Amount (RS)")
        self.cash_invoice_Table.heading("sgst_rate",text="SGST Rate (RS)")
        self.cash_invoice_Table.heading("sgst_amount",text="SGST Amount (RS)")
        self.cash_invoice_Table.heading("amount",text="Amount (RS)")

        self.cash_invoice_Table["show"]="headings"

        self.cash_invoice_Table.column("c_id",width=30)
        self.cash_invoice_Table.column("bill_no",width=30)
        self.cash_invoice_Table.column("date",width=50)
        self.cash_invoice_Table.column("customer",width=50)
        self.cash_invoice_Table.column("service_name",width=100)
        self.cash_invoice_Table.column("hsn",width=30)
        self.cash_invoice_Table.column("quantity",width=50)
        self.cash_invoice_Table.column("unit",width=50)
        self.cash_invoice_Table.column("rate",width=50)
        self.cash_invoice_Table.column("cgst_rate",width=50)
        self.cash_invoice_Table.column("cgst_amount",width=50)
        self.cash_invoice_Table.column("sgst_rate",width=50)
        self.cash_invoice_Table.column("sgst_amount",width=50)
        self.cash_invoice_Table.column("amount",width=50)
        self.cash_invoice_Table.pack(fill=BOTH,expand=1)
        self.cash_invoice_Table.bind("<ButtonRelease-1>",self.get_data)

        self.update_content()
        self.show()
    #=======================================================================================================================================================================
    def validate_length(self, new_value):
        return len(new_value) <= 32  # allow only up to 32 characters
    
    def fetch_customer(self):
        self.customer_list.append("Empty")
        con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
        #con=sqlite3.connect(database=r'./ah.db')
        cur=con.cursor()
        try:
            cur.execute("Select Name from customer")
            customer=cur.fetchall()
            if len(customer)>0:
                del self.customer_list[:]
                self.customer_list.append("Select")
                for i in customer:
                    self.customer_list.append(i[0])    
        
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :{str(ex)}",parent=self.root)
    
    def recalculate_tax_and_total(self, *args):
        try:
            quantity = float(self.var_quantity.get())
            rate = float(self.var_rate.get())
            taxable = quantity * rate

            cgst_rate = float(self.var_cgst_rate.get())
            sgst_rate = float(self.var_sgst_rate.get())

            cgst_amt = taxable * cgst_rate / 100
            sgst_amt = taxable * sgst_rate / 100
            total_amt = taxable + cgst_amt + sgst_amt

            self.var_cgst_amount.set(f"{cgst_amt:.2f}")
            self.var_sgst_amount.set(f"{sgst_amt:.2f}")
            self.var_amount.set(f"{total_amt:.2f}")
        except:
            # Ignore empty/invalid values
            pass
    
    def add(self):
        con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
        #con = sqlite3.connect(database=r'./ah.db')
        cur = con.cursor()
        try:
            # Validate required fields
            if self.var_service_name.get() == "":
                messagebox.showerror("Error", "Service Name must be required", parent=self.root)
                return

            # Convert values safely
            quantity = float(self.var_quantity.get())
            rate = float(self.var_rate.get())
            taxable_amount = quantity * rate

            # Set CGST and SGST rates
            cgst_rate = 0
            sgst_rate = 0

            # Calculate tax amounts
            cgst_amount = taxable_amount * (cgst_rate / 100)
            sgst_amount = taxable_amount * (sgst_rate / 100)

            # Total amount = taxable + both taxes
            total_amount = taxable_amount + cgst_amount + sgst_amount

            # Insert into cash_invoice table
            cur.execute("""
                INSERT INTO cash_invoice(
                    bill_no, date, customer, service_name, hsn, quantity, unit, rate,
                    cgst_rate, cgst_amount, sgst_rate, sgst_amount, amount
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                self.var_bill_no.get(),
                self.var_date.get(),
                self.var_customer.get(),
                self.var_service_name.get(),
                self.var_hsn.get(),
                quantity,
                self.var_unit.get(),
                rate,
                cgst_rate,
                cgst_amount,
                sgst_rate,
                sgst_amount,
                total_amount
            ))
            con.commit()
            messagebox.showinfo("Success", "Cash Invoice entry added successfully!", parent=self.root)
            self.clear()
            self.show()

        except Exception as ex:
            messagebox.showerror("Error", f"Error due to: {str(ex)}", parent=self.root)
        finally:
            con.close()
        
    def show(self):
        con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
        #con=sqlite3.connect(database=r'./ah.db')
        cur=con.cursor()
        try:
            cur.execute("select * from cash_invoice")
            rows=cur.fetchall()
            self.cash_invoice_Table.delete(*self.cash_invoice_Table.get_children())
            for row in rows:
                self.cash_invoice_Table.insert('',END,values=row)
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :{str(ex)}",parent=self.root)

    def get_data(self, ev):
        f = self.cash_invoice_Table.focus()
        content = self.cash_invoice_Table.item(f)
        row = content['values']
        
        self.var_c_id.set(row[0])
        self.var_bill_no.set(row[1])
        self.var_date.set(row[2])
        self.var_customer.set(row[3])
        self.var_service_name.set(row[4])
        self.var_hsn.set(row[5])
        self.var_quantity.set(row[6])
        self.var_unit.set(row[7])
        self.var_rate.set(row[8])
        self.var_cgst_rate.set(row[9])
        self.var_cgst_amount.set(row[10])
        self.var_sgst_rate.set(row[11])
        self.var_sgst_amount.set(row[12])
        self.var_amount.set(row[13])
    
        # Button control
        self.btn_add.config(state='disabled')
        self.btn_update.config(state='normal')
        self.btn_delete.config(state='normal')
        #self.btn_export_search.config(state='normal')
    
    def update(self):
        con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
        #con=sqlite3.connect(database=r'./ah.db')
        cur=con.cursor()
        try:
            if self.var_c_id.get()=="":
                messagebox.showerror("Error","c_id Must be required",parent=self.root)
            else:
                cur.execute("Select * from cash_invoice where c_id=?",(self.var_c_id.get(),))
                row=cur.fetchone()
                if row==None:
                    messagebox.showerror("Error","Invalid c_id",parent=self.root)
                else:
                    cur.execute("Update cash_invoice set bill_no=?,date=?,customer=?,service_name=?,hsn=?,quantity=?,unit=?,rate=?,cgst_rate=?,cgst_amount=?,sgst_rate=?,sgst_amount=?,amount=? where c_id=?",(
                                        self.var_bill_no.get(),
                                        self.var_date.get(),
                                        self.var_customer.get(),
                                        self.var_service_name.get(),
                                        self.var_hsn.get(),
                                        self.var_quantity.get(),
                                        self.var_unit.get(),
                                        self.var_rate.get(),
                                        self.var_cgst_rate.get(),
                                        self.var_cgst_amount.get(),
                                        self.var_sgst_rate.get(),
                                        self.var_sgst_amount.get(),
                                        float(self.var_quantity.get())*float(self.var_rate.get())+float(self.var_cgst_amount.get())+float(self.var_sgst_amount.get()),                    
                                        #self.var_amount.get(),
                                        self.var_c_id.get()
                    ))
                    con.commit()
                    messagebox.showinfo("Success","Cash Invoice Updated Successfully",parent=self.root)
                    self.clear() # Clear fields after updating
                    self.show()
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to : {str(ex)}",parent=self.root)

    def delete(self):
        if self.var_c_id.get() == "":
            messagebox.showerror("Error", "c_id must be required", parent=self.root)
        else:
            con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
            #con = sqlite3.connect(database=r'./ah.db')
            cur = con.cursor()
            cur.execute("SELECT * FROM cash_invoice WHERE c_id=?", (self.var_c_id.get(),))
            row = cur.fetchone()
            if row is None:
                messagebox.showerror("Error", "Invalid c_id", parent=self.root)
            else:
                op = messagebox.askyesno("Confirm", "Do you really want to delete?", parent=self.root)
                if op:
                    cur.execute("DELETE FROM cash_invoice WHERE c_id=?", (self.var_c_id.get(),))
                    con.commit()
                    messagebox.showinfo("Delete", "Cash Invoice Deleted Successfully", parent=self.root)
                    self.clear()    
    
    def clear(self):
        self.var_c_id.set("0")
        self.var_bill_no.set("")
        self.var_date.set("Select")
        self.var_customer.set("Select")
        self.var_service_name.set("")
        self.var_hsn.set("")
        self.var_quantity.set("0")
        self.var_unit.set("")
        self.var_rate.set("0")
        self.var_cgst_rate.set("0")
        self.var_cgst_amount.set("0")
        self.var_sgst_rate.set("0")
        self.var_sgst_amount.set("0")
        self.var_amount.set("0")
        self.root.after(100, lambda: self.txt_c_id.focus_set())
        self.var_searchtxt.set("")
        self.var_searchby.set("Select")        
        # Button control
        self.btn_add.config(state='normal')
        self.btn_update.config(state='disabled')
        self.btn_delete.config(state='disabled')
        #self.btn_export_search.config(state='disabled')
       
        self.show()
    
    def show_service_name_suggestions(self, event=None):
        typed = self.var_service_name.get()

        if typed == "":
            if hasattr(self, 'service_name_listbox') and self.service_name_listbox.winfo_exists():
                self.service_name_listbox.destroy()
            return

        con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
        #con = sqlite3.connect('./ah.db')
        cur = con.cursor()
        cur.execute("SELECT DISTINCT service_name FROM cash_invoice WHERE service_name LIKE ?", (typed + "%",))
        results = [r[0] for r in cur.fetchall()]
        con.close()

        if hasattr(self, 'service_name_listbox') and self.service_name_listbox.winfo_exists():
            self.service_name_listbox.destroy()

        if results:
            self.service_name_listbox = Listbox(self.root, height=4, bg="white", font=("Arial", 10))
            x = self.txt_service_name.winfo_x()
            y = self.txt_service_name.winfo_y() + self.txt_service_name.winfo_height()
            self.service_name_listbox.place(x=x, y=y, width=self.txt_service_name.winfo_width())

            for r in results:
                self.service_name_listbox.insert(END, r)

            self.service_name_listbox.bind("<<ListboxSelect>>", self.fill_service_name)
    
    def show_hsn_suggestions(self, event=None):
        typed = self.var_hsn.get()

        if typed == "":
            if hasattr(self, 'hsn_listbox') and self.hsn_listbox.winfo_exists():
                self.hsn_listbox.destroy()
            return

        con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
        #con = sqlite3.connect('./ah.db')
        cur = con.cursor()
        cur.execute("SELECT DISTINCT hsn FROM cash_invoice WHERE hsn LIKE ?", (typed + "%",))
        results = [r[0] for r in cur.fetchall()]
        con.close()

        if hasattr(self, 'hsn_listbox') and self.hsn_listbox.winfo_exists():
            self.hsn_listbox.destroy()

        if results:
            self.hsn_listbox = Listbox(self.root, height=4, bg="white", font=("Arial", 10))
            x = self.txt_hsn.winfo_x()
            y = self.txt_hsn.winfo_y() + self.txt_hsn.winfo_height()
            self.hsn_listbox.place(x=x, y=y, width=self.txt_hsn.winfo_width())

            for r in results:
                self.hsn_listbox.insert(END, r)

            self.hsn_listbox.bind("<<ListboxSelect>>", self.fill_hsn)
    
    def show_unit_suggestions(self, event=None):
        typed = self.var_unit.get()

        if typed == "":
            if hasattr(self, 'unit_listbox') and self.unit_listbox.winfo_exists():
                self.unit_listbox.destroy()
            return

        con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
        #con = sqlite3.connect('./ah.db')
        cur = con.cursor()
        cur.execute("SELECT DISTINCT unit FROM cash_invoice WHERE unit LIKE ?", (typed + "%",))
        results = [r[0] for r in cur.fetchall()]
        con.close()

        if hasattr(self, 'unit_listbox') and self.unit_listbox.winfo_exists():
            self.unit_listbox.destroy()

        if results:
            self.unit_listbox = Listbox(self.root, height=4, bg="white", font=("Arial", 10))
            x = self.txt_unit.winfo_x()
            y = self.txt_unit.winfo_y() + self.txt_unit.winfo_height()
            self.unit_listbox.place(x=x, y=y, width=self.txt_unit.winfo_width())

            for r in results:
                self.unit_listbox.insert(END, r)

            self.unit_listbox.bind("<<ListboxSelect>>", self.fill_unit)
    
    def fill_service_name(self, event):
        if self.service_name_listbox.curselection():
            selected = self.service_name_listbox.get(self.service_name_listbox.curselection())
            self.var_service_name.set(selected)
            self.service_name_listbox.destroy()
    
    def fill_hsn(self, event):
        if self.hsn_listbox.curselection():
            selected = self.hsn_listbox.get(self.hsn_listbox.curselection())
            self.var_hsn.set(selected)
            self.hsn_listbox.destroy()
    
    def fill_unit(self, event):
        if self.unit_listbox.curselection():
            selected = self.unit_listbox.get(self.unit_listbox.curselection())
            self.var_unit.set(selected)
            self.unit_listbox.destroy()
    
    def search(self):  
        con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
        #con=sqlite3.connect(database=r'./ah.db')
        cur=con.cursor()
        try:
            if self.var_searchby.get()=="Select":
                messagebox.showerror("Error","Select Search by option",parent=self.root)
            elif self.var_searchtxt.get()=="":
                messagebox.showerror("Error","Search input should be required",parent=self.root)
            
            else:
                cur.execute("select * from cash_invoice where "+self.var_searchby.get()+" LIKE '%"+self.var_searchtxt.get()+"%'")
                rows=cur.fetchall()
                if len(rows)!=0:
                    self.cash_invoice_Table.delete(*self.cash_invoice_Table.get_children())
                    for row in rows:
                        self.cash_invoice_Table.insert('',END,values=row)
                else:
                    messagebox.showerror("Error","No record found!!!",parent=self.root)
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :{str(ex)}",parent=self.root)
   
    def send_whatsapp_message(self):
        phone_number = self.var_hsn.get()  # Get cash_invoice phone number
        message = f"Hello {self.var_service_name.get()}, thank you for being our valued customer!"

        if phone_number == "":
            messagebox.showerror("Error", "customer phone number is required!", parent=self.root)
            return

        try:
            # Send WhatsApp message instantly
            kit.sendwhatmsg_instantly(f"+91{phone_number}", message, wait_time=10)
            messagebox.showinfo("Success", "WhatsApp message sent successfully!", parent=self.root)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send WhatsApp message\n{str(e)}", parent=self.root)    
            
    def export_searched_cash_invoice_to_pdf(self):
            # Validate search input
            if self.var_searchby.get() == "Select":
                messagebox.showerror("Error", "Select Search by option", parent=self.root)
                return
            if self.var_searchtxt.get() == "":
                messagebox.showerror("Error", "Search input should be required", parent=self.root)
                return

            # Fetch displayed (searched) data from Treeview
            rows = []
            for child in self.cash_invoice_Table.get_children():
                rows.append(self.cash_invoice_Table.item(child)['values'])

            if not rows:
                messagebox.showwarning("No Data", "There is no data to export.")
                return

            # Convert to DataFrame
            columns = ["C ID", "Bill No", "Date", "Customer", "Service Name", "HSN/SAC", "Quantity", "Unit",
                "Rate (RS)", "CGST Rate (RS)", "CGST Amount (RS)",
                "SGST Rate (RS)", "SGST Amount (RS)", "Amount (RS)"]
            df = pd.DataFrame(rows, columns=columns)

            # Compute HSN Summary
            hsn_summary = defaultdict(lambda: {"taxable": 0.0, "cgst": 0.0, "sgst": 0.0})
            for _, row in df.iterrows():
                try:
                    hsn = str(row["HSN/SAC"])
                    qty = float(row["Quantity"])
                    rate = float(row["Rate (RS)"])
                    cgst_amt = float(row["CGST Amount (RS)"])
                    sgst_amt = float(row["SGST Amount (RS)"])
                    taxable = qty * rate
                    hsn_summary[hsn]["taxable"] += taxable
                    hsn_summary[hsn]["cgst"] += cgst_amt
                    hsn_summary[hsn]["sgst"] += sgst_amt
                except:
                    continue

            # Fetch firm details
            con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
            #con = sqlite3.connect("ah.db")
            cur = con.cursor()
            cur.execute("SELECT name, contact, address, email, gst, account_holder_name, bank, account_no, branch_ifs_code FROM firm LIMIT 1")
            firm = cur.fetchone()
            con.close()

            firm_name = firm[0] if firm else "FIRM NAME"
            firm_contact = firm[1] if firm else "Contact"
            firm_address = firm[2] if firm else "Address"
            firm_email = firm[3] if firm else "Email"
            firm_gst = firm[4] if firm else "GST Number"
            account_holder = firm[5] if firm else ""
            bank_name = firm[6] if firm else ""
            account_no = firm[7] if firm else ""
            ifsc = firm[8] if firm else ""

            # Ask where to save PDF
            pdf_file = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                    filetypes=[("PDF files", "*.pdf")],
                                                    initialfile="Search_Cash_Invoice.pdf")
            if not pdf_file:
                return

            # Setup PDF
            pdf = FPDF(orientation='P', unit='mm', format='A4')
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)

            pdf.set_font("helvetica", 'B', 14)
            pdf.cell(0, 10, "CASH INVOICE", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="C")

            # --- Insert logo below CASH INVOICE ---
            logo_path = "./image/along_home_logo_1.png"  # Set your actual logo path here
            if os.path.exists(logo_path):
                # x, y is position (set y after CASH INVOICE)
                pdf.image(logo_path, x=80, y=pdf.get_y(), w=50)  # center aligned logo (adjust w/h if needed)
                pdf.ln(30)  # space below logo

            
            pdf.set_font("helvetica", 'B', 12)
            pdf.cell(0, 6, firm_name, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.set_font("helvetica", '', 10)
            pdf.multi_cell(0, 5,
                f"Mobile: {firm_contact}\n"
                f"{firm_address.strip()}\n"
                f"Email: {firm_email}\n"
                f"GSTIN/UIN: {firm_gst}\n"
                f"State Name: Gujarat, Code: 24"
            )
            pdf.ln(2)

            # customer details
            customer_name = df.iloc[0]["Customer"]
            con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
            #con = sqlite3.connect("ah.db")
            cur = con.cursor()
            cur.execute("SELECT address, contact, gst, email FROM customer WHERE name=?", (customer_name,))
            customer = cur.fetchone()
            con.close()

            pdf.set_font("helvetica", 'B', 10)
            pdf.cell(95, 6, "Customer (Bill to)", border=1)
            pdf.cell(95, 6, "Invoice Details", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.set_font("helvetica", '', 10)
            pdf.cell(95, 6, customer_name, border=1)
            pdf.cell(95, 6, f"Invoice No.: {df.iloc[0]['Bill No']}    Dated: {df.iloc[0]['Date']}", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            if customer:
                pdf.cell(95, 6, f"Contact: {customer[1]}", border=1)
                pdf.cell(95, 6, f"GST/UIN: {customer[2]}", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                pdf.cell(95, 6, f"Email: {customer[3]}", border=1)
            else:
                pdf.cell(95, 6, "GST/UIN: -", border=1)
            pdf.cell(95, 6, "", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.multi_cell(190, 5, f"Address: {customer[0] if customer else 'N/A'}", border=1)
            pdf.cell(190, 6, "State Name: Gujarat, Code: 24", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.ln(2)

            # Product Table Header
            table_cols = ["Sl", "Service Name", "HSN/SAC", "Qty", "Unit", "Rate", "CGST%", "CGST Amt", "SGST%", "SGST Amt", "Amount"]
            col_widths = [7, 40, 18, 14, 12, 15, 14, 19, 14, 19, 20]

            def print_table_header():
                pdf.set_font("helvetica", 'B', 9)
                for col, w in zip(table_cols, col_widths):
                    pdf.cell(w, 6, col, border=1, align='C')
                pdf.ln()
                pdf.set_font("helvetica", '', 9)

            print_table_header()

            # Table Rows
            total_qty = total_cgst = total_sgst = total_amount = 0
            line_height = 6

            for idx, row in df.iterrows():
                x_left = pdf.get_x()
                y_top = pdf.get_y()

                # Handle wrapped height
                desc_lines = pdf.multi_cell(col_widths[1], line_height, str(row["Service Name"]), dry_run=True, output="LINES")
                row_height = line_height * len(desc_lines)

                if y_top + row_height > 195:
                    pdf.add_page()
                    print_table_header()
                    x_left = pdf.get_x()
                    y_top = pdf.get_y()

                # Row Cells
                pdf.set_xy(x_left, y_top)
                pdf.cell(col_widths[0], row_height, str(idx + 1), border=1, align='C')
                pdf.set_xy(x_left + col_widths[0], y_top)
                pdf.multi_cell(col_widths[1], line_height, str(row["Service Name"]), border=1)
                #pdf.set_xy(x_left + sum(col_widths[:2]), y_top)
                #pdf.cell(col_widths[2], row_height, str(row["Serial No"]), border=1)
                pdf.set_xy(x_left + sum(col_widths[:2]), y_top)
                pdf.cell(col_widths[2], row_height, str(row["HSN/SAC"]), border=1)
                try:
                    qty = float(row["Quantity"])
                    total_qty += qty
                    pdf.set_xy(x_left + sum(col_widths[:3]), y_top)
                    pdf.cell(col_widths[3], row_height, f"{qty:.2f}", border=1, align='R')
                except:
                    pass
                pdf.set_xy(x_left + sum(col_widths[:4]), y_top)
                pdf.cell(col_widths[4], row_height, str(row["Unit"]), border=1)

                try:
                    rate = float(row["Rate (RS)"])
                    pdf.set_xy(x_left + sum(col_widths[:5]), y_top)
                    pdf.cell(col_widths[5], row_height, f"{rate:.2f}", border=1, align='R')
                except:
                    pass

                pdf.set_xy(x_left + sum(col_widths[:6]), y_top)
                pdf.cell(col_widths[6], row_height, str(row["CGST Rate (RS)"]), border=1, align='R')

                try:
                    cgst_amt = float(row["CGST Amount (RS)"])
                    total_cgst += cgst_amt
                    pdf.set_xy(x_left + sum(col_widths[:7]), y_top)
                    pdf.cell(col_widths[7], row_height, f"{cgst_amt:.2f}", border=1, align='R')
                except:
                    pass

                pdf.set_xy(x_left + sum(col_widths[:8]), y_top)
                pdf.cell(col_widths[8], row_height, str(row["SGST Rate (RS)"]), border=1, align='R')

                try:
                    sgst_amt = float(row["SGST Amount (RS)"])
                    total_sgst += sgst_amt
                    pdf.set_xy(x_left + sum(col_widths[:9]), y_top)
                    pdf.cell(col_widths[9], row_height, f"{sgst_amt:.2f}", border=1, align='R')
                except:
                    pass

                try:
                    amt = float(row["Amount (RS)"])
                    total_amount += amt
                    pdf.set_xy(x_left + sum(col_widths[:10]), y_top)
                    pdf.cell(col_widths[10], row_height, f"{amt:.2f}", border=1, align='R')
                except:
                    pass

                pdf.set_y(y_top + row_height)

            # Totals
            pdf.set_font("helvetica", 'B', 9)
            pdf.cell(sum(col_widths[:3]), 6, "Total", border=1)
            pdf.cell(col_widths[3], 6, f"{total_qty:.2f}", border=1, align='R')
            for i in range(4, 7):
                pdf.cell(col_widths[i], 6, "", border=1)
            pdf.cell(col_widths[7], 6, f"{total_cgst:.2f}", border=1, align='R')
            pdf.cell(col_widths[8], 6, "", border=1)
            pdf.cell(col_widths[9], 6, f"{total_sgst:.2f}", border=1, align='R')
            pdf.cell(col_widths[10], 6, f"{total_amount:.2f}", border=1, align='R')
            pdf.ln(8)

            # Amount in Words
            # --- Round to 2 decimal places for paise conversion ---
            rounded_amount = round(total_amount, 2)
            
            # --- Split into rupees and paise ---
            rupees = int(rounded_amount)
            paise = int(round((rounded_amount - rupees) * 100))
            
            # --- Convert to words ---
            amt_words = f"{num2words(rupees, lang='en_IN').title()} Rupees"
            if paise > 0:
                amt_words += f" and {num2words(paise, lang='en_IN').title()} Paise"
            amt_words += " Only"        
        
            # --- For PDF ---
            pdf.set_font("helvetica", '', 9)
            pdf.cell(0, 6, f"Amount Chargeable (in words): INR {amt_words}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            # --- Also show round figure (if needed) ---
            pdf.cell(0, 6, f"Rounded Total: INR {round(rounded_amount)}.00", new_x=XPos.LMARGIN, new_y=YPos.NEXT)        
            
            # --- HSN-wise Tax Summary ---
            pdf.set_font("helvetica", 'B', 10)
            pdf.cell(0, 6, "HSN-wise Tax Summary", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.set_font("helvetica", 'B', 9)
            pdf.cell(40, 6, "HSN/SAC", border=1)
            pdf.cell(40, 6, "Taxable Value", border=1)
            pdf.cell(30, 6, "CGST", border=1)
            pdf.cell(30, 6, "SGST", border=1)
            pdf.cell(40, 6, "Total Tax", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            # --- Totals ---
            grand_taxable = 0.0
            grand_cgst = 0.0
            grand_sgst = 0.0

            pdf.set_font("helvetica", '', 9)
            for hsn, values in hsn_summary.items():
                taxable = values["taxable"]
                cgst = values["cgst"]
                sgst = values["sgst"]
                total_tax = cgst + sgst

                pdf.cell(40, 6, hsn, border=1)
                pdf.cell(40, 6, f"{taxable:.2f}", border=1)
                pdf.cell(30, 6, f"{cgst:.2f}", border=1)
                pdf.cell(30, 6, f"{sgst:.2f}", border=1)
                pdf.cell(40, 6, f"{total_tax:.2f}", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

                grand_taxable += taxable
                grand_cgst += cgst
                grand_sgst += sgst

            # --- Grand Total Row ---
            pdf.set_font("helvetica", 'B', 9)
            pdf.cell(40, 6, "Grand Total", border=1)
            pdf.cell(40, 6, f"{grand_taxable:.2f}", border=1)
            pdf.cell(30, 6, f"{grand_cgst:.2f}", border=1)
            pdf.cell(30, 6, f"{grand_sgst:.2f}", border=1)
            pdf.cell(40, 6, f"{grand_cgst + grand_sgst:.2f}", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            # --- GST in Words ---
            total_gst = grand_cgst + grand_sgst
            gst_rounded = round(total_gst, 2)
            gst_rupees = int(gst_rounded)
            gst_paise = int(round((gst_rounded - gst_rupees) * 100))

            gst_in_words = f"{num2words(gst_rupees, lang='en_IN').title()} Rupees"
            if gst_paise > 0:
                gst_in_words += f" and {num2words(gst_paise, lang='en_IN').title()} Paise"
            gst_in_words += " Only"

            pdf.ln(3)
            pdf.set_font("helvetica", '', 9)
            pdf.cell(0, 6, f"Total GST (in words): INR {gst_in_words}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            # ✅ Rounded GST Total
            pdf.cell(0, 6, f"Rounded GST Total: INR {round(gst_rounded)}.00", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            
            # Bank Details
            pdf.set_font("helvetica", '', 9)
            pdf.ln(4)
            pdf.cell(0, 6, "Company's Bank Details:", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(0, 6, f"A/c Holder's Name : {account_holder}", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(0, 6, f"Bank Name : {bank_name}", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(0, 6, f"A/c No. : {account_no}", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            pdf.cell(0, 6, f"Branch & IFS Code: {ifsc}", border=1, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

            # Signature & Notes
            pdf.ln(12)
            pdf.set_font("helvetica", '', 12)
            pdf.cell(0, 6, f"for {firm_name}", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="R")
            pdf.ln(6)
            pdf.cell(0, 6, "(Authorised Signatory)", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align="R")
            pdf.ln(6)

            pdf.set_font("helvetica", 'I', 8)
            pdf.cell(0, 5, "SUBJECT TO AHMEDABAD JURISDICTION", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
            pdf.cell(0, 5, "This is a Computer Generated Invoice", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')

            # Save PDF
            try:
                pdf.output(pdf_file)
                messagebox.showinfo("Success", f"Invoice exported to:\n{pdf_file}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save PDF:\n{e}")
        
    
    def export_to_excel(self):
            con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
            #con = sqlite3.connect(database=r'./ah.db')
            cur = con.cursor()
            try:
                cur.execute("SELECT * FROM cash_invoice")
                rows = cur.fetchall()
                if not rows:
                    messagebox.showerror("Error", "No data to export", parent=self.root)
                    return

                # Create a DataFrame from the fetched data
                df = pd.DataFrame(rows, columns=["C ID", "Bill No","Date","Customer","Service Name","HSN/SAC","Quantity","Unit","Rate (RS)","CGST Rate (RS)","CGST Amount (RS)","SGST Rate (RS)","SGST Amount (RS)","Amount Due (RS)"])

                # Add "As of Date" column
                as_of_date = datetime.now().strftime("%d-%m-%Y")  # Current date in YYYY-MM-DD format
                df["As of Date"] = as_of_date            # Export to Excel
                
                output_file = "./data/Cash Invoice Data/Report.xlsx"
                df.to_excel(output_file, index=False, engine='openpyxl')

                messagebox.showinfo("Success", f"Data exported to {output_file}", parent=self.root)

            except Exception as ex:
                messagebox.showerror("Error", f"Error due to : {str(ex)}", parent=self.root)     
    
    def export_to_pdf(self):
            con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
            #con = sqlite3.connect(database=r'./ah.db')
            cur = con.cursor()
            cur.execute("SELECT name, address, contact, gst, email FROM firm LIMIT 1")
            firm = cur.fetchone()
            try:
                cur.execute("SELECT * FROM cash_invoice")
                rows = cur.fetchall()
                if not rows:
                    messagebox.showerror("Error", "No data to export", parent=self.root)
                    return

                pdf_file = filedialog.asksaveasfilename(
                    defaultextension="./data/Cash Invoice Data/Report.pdf",
                    filetypes=[("PDF files", "*.pdf")]
                )
                if not pdf_file:
                    return
                os.makedirs(os.path.dirname(pdf_file), exist_ok=True)

                # Header titles
                headers = ["C ID", "Bill No", "Date", "Customer", "Service Name", "HSN/SAC", "Quantity","Unit", "Rate (RS)", "CGST Rate (RS)", "CGST Amount (RS)", "SGST Rate (RS)", "SGST Amount (RS)", "Amount Due (RS)"]

                # Auto-fit header widths
                font_name = 'Helvetica-Bold'
                font_size = 6
                padding = 22
                column_widths = [stringWidth(h, font_name, font_size) + padding for h in headers]
                total_width = sum(column_widths)

                # Page orientation
                portrait_width, portrait_height = A4
                page_size = landscape(A4) if total_width > portrait_width else portrait(A4)

                # PDF document setup
                doc = SimpleDocTemplate(pdf_file, pagesize=page_size, rightMargin=1, leftMargin=1)
                elements = []
                styles = getSampleStyleSheet()
                wrap_style = ParagraphStyle(name='Wrap', fontSize=5, alignment=1)

                # Header Info
                styles = getSampleStyleSheet()

                elements.append(Paragraph("<b>CASH INVOICE REPORT</b>", styles['Title']))
                elements.append(Spacer(1, 12))  # Extra space below the title
                
                if firm:
                    elements.append(RLImage("./image/along_home_logo_1.png", width=100, height=40) if os.path.exists("./image/nurse1.png") else Spacer(1, 2))
                    elements.append(Paragraph(f"{firm[0]}", styles['Title']))  # Firm Name
                    elements.append(Paragraph(f"Office: {firm[1]}", styles['Normal']))  # Address
                    elements.append(Paragraph(f"M : {firm[2]}", styles['Normal']))  # Contact
                    elements.append(Paragraph(f"GSTIN/UIN : {firm[3]}", styles['Normal']))  # GST
                    elements.append(Paragraph(f"E : {firm[4]}", styles['Normal']))  # Email
                else:
                    elements.append(Paragraph("Firm details not available", styles['Normal']))

                elements.append(Spacer(1, 4))
                elements.append(Paragraph("=" * 141, styles['Normal']))
                elements.append(Spacer(1, 4))
            
                data = [headers]
                grouped = defaultdict(list)

                for row in rows:
                    customer = row[3]  # assuming 3 is "customer"
                    grouped[customer].append(row)

                grand_totals = defaultdict(float)
                numeric_cols = [6, 10, 12, 13]  # Quantity, CGST Amt, SGST Amt, Amount

                for customer, group_rows in grouped.items():
                    for row in group_rows:
                        row_data = []
                        for i, cell in enumerate(row[:15]):  # Limit to first 15 fields
                            cell_str = str(cell)
                            row_data.append(Paragraph(cell_str, wrap_style))
                            if i in numeric_cols:
                                try:
                                    grand_totals[i] += float(cell)
                                except:
                                    pass
                        data.append(row_data)

                    # Subtotal Row
                    subtotal_row = []
                    for i in range(len(headers)):
                        if i == 0:
                            subtotal_row.append(Paragraph(f"<b>{customer} TOTAL</b>", wrap_style))
                        elif i in numeric_cols:
                            subtotal = sum(float(r[i]) for r in group_rows if isinstance(r[i], (int, float, str)) and str(r[i]).replace('.', '', 1).isdigit())
                            subtotal_row.append(Paragraph(f"{subtotal:.2f}", wrap_style))
                        else:
                            subtotal_row.append(Paragraph("", wrap_style))
                    data.append(subtotal_row)

                # Final GRAND TOTAL
                total_row = []
                for i in range(len(headers)):
                    if i == 0:
                        total_row.append(Paragraph("<b>GRAND TOTAL</b>", wrap_style))
                    elif i in grand_totals:
                        total_row.append(Paragraph(f"{grand_totals[i]:.2f}", wrap_style))
                    else:
                        total_row.append(Paragraph("", wrap_style))
                data.append(total_row)

                table = Table(data, repeatRows=1, colWidths=column_widths)
                style = TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-2, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), font_name),
                    ('FONTSIZE', (0, 0), (-1, -1), font_size),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 5),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                ])
                for i in range(1, len(data)):
                    bg = colors.whitesmoke if i % 2 else colors.lightgrey
                    style.add('BACKGROUND', (0, i), (-1, i), bg)
                table.setStyle(style)
                elements.append(table)

                elements.append(Spacer(1, 20))
                elements.append(Paragraph("Prepared By", styles['Normal']))
                elements.append(Spacer(1, 2))
                elements.append(Paragraph("------------", styles['Normal']))
                doc.build(elements)
                messagebox.showinfo("Success", f"PDF exported to:\n{pdf_file}", parent=self.root)

            except Exception as ex:
                messagebox.showerror("Error", f"Error due to: {str(ex)}", parent=self.root)
            finally:
                con.close()
        
    def import_excel_to_db(self):
        # Ask user to select Excel file
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return
        try:
            # Read Excel into DataFrame
            df = pd.read_excel(file_path)

            # Connect to database
            con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
            #con = sqlite3.connect('ah.db')
            cur = con.cursor()

            for index, row in df.iterrows():
                cur.execute("""INSERT OR IGNORE INTO cash_invoice(c_id, bill_no, date, customer, service_name, hsn, quantity, unit, rate, cgst_rate, cgst_amount, sgst_rate, sgst_amount, amount
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (
                    row['c_id'],
                    row['bill_no'],
                    row['date'],
                    row['customer'],
                    row['service_name'],
                    row['hsn'],
                    row['quantity'],
                    row['unit'],
                    row['rate'],
                    row['cgst_rate'],
                    row['cgst_amount'],
                    row['sgst_rate'],
                    row['sgst_amount'],
                    row['amount']
                ))

            con.commit()
            con.close()
            messagebox.showinfo("Success", "Excel data imported successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import Excel:\n{e}")
    
    def update_content(self):
        con = pymysql.connect(host='localhost', user='root', password='', db='along_home_db')
        #con=sqlite3.connect(database=r'./ah.db')
        cur=con.cursor()
        try:
            time_=time.strftime("%I:%M:%S %p")
            date_=time.strftime("%d-%m-%Y")
            self.lbl_clock.config(text=f"  ALONG HOME HEALTHCARE \t\t CASH INVOICE DATA\t\t\t Date: {str(date_)}\t\t Time: {str(time_)}")
            self.lbl_clock.after(200,self.update_content)
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :   {str(ex)}",parent=self.root)
        
if __name__=="__main__":
    root=Tk()
    obj=cash_invoiceClass(root)
    root.mainloop()