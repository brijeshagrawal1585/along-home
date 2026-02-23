from datetime import datetime, timezone
import time
from tkinter import*
from PIL import Image,ImageTk #pip install pilow
from tkinter import ttk,messagebox
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
import os
from fpdf import FPDF  # pip install fpdf

from collections import defaultdict
from num2words import num2words  # pip install num2words

class tax_invoiceClass:
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

        self.var_t_id=StringVar(value="0")                   #t_id
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
        SearchFrame=LabelFrame(self.root,text="Search Tax Invoice Data",font=("Arial",12,"bold"),bd=2,relief=RIDGE,bg="white")
        SearchFrame.place(x=10,y=20,width=750,height=70)

        cmb_search=ttk.Combobox(SearchFrame,textvariable=self.var_searchby,values=("Select","T_ID","BILL_NO","DATE","CUSTOMER"),state='readonly',justify=CENTER,font=("Arial",15))
        cmb_search.place(x=5,y=10,width=120)
        cmb_search.current(0)

        txt_search=Entry(SearchFrame,textvariable=self.var_searchtxt,font=("Arial",15),bg="lightyellow").place(x=140,y=10,width=100)
        btn_search=Button(SearchFrame,text="Search",command=self.search,font=("Arial",15),bg="black",fg="white",cursor="hand2").place(x=260,y=9,width=80,height=30)
                
        self.btn_export_search = Button(SearchFrame, text="Export Search to Excel & PDF", command=self.export_searched_tax_invoice_to_pdf_excel, font=("Arial", 12,"bold"), bg="red", fg="white",cursor="hand2")
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
        lbl_t_id=Label(self.root,text="T ID",font=("Arial",13),bg="white").place(x=20,y=150)
        self.txt_t_id=Entry(self.root,textvariable=self.var_t_id,justify=CENTER,state='readonly',font=("Arial",13),bg="lightyellow",validate="key", validatecommand=vcmd)
        self.txt_t_id.place(x=130,y=150,width=80)
        
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
        self.btn_export_search.config(state='disabled')

        btn_whatsapp = Button(self.root, text="Send WhatsApp", command=self.send_whatsapp_message, font=("Arial", 15,"bold"), bg="green", fg="white", cursor="hand2")
        btn_whatsapp.place(x=530, y=270, width=170, height=28)        
        #=========Footer==============
        lbl_footer=Label(self.root,text="ALONG HOME HEALTHCARE | 59/2, Ground Floor, Govt. H. Colony, B/h. Laxmi Ganthiya Rath, Nehrunagar Cross Road, Ahmedabad-380015 Contact:+91 9904110283",font=("times new roman",15),bg="#4d636d",fg="white").pack(side=BOTTOM,fill=X)
        #=========vendor_code_part Details========
        tax_invoice_frame=LabelFrame(self.root,text="Tax Invoice List",font=("Arial",12,"bold"),bd=3,relief=RIDGE,bg="white")
        tax_invoice_frame.place(x=0,y=310,relwidth=1,height=350)

        scrolly=Scrollbar(tax_invoice_frame,orient=VERTICAL)
        scrollx=Scrollbar(tax_invoice_frame,orient=HORIZONTAL) 
        
        self.tax_invoice_Table=ttk.Treeview(tax_invoice_frame,columns=("t_id","bill_no","date","customer","service_name","hsn","quantity","unit","rate","cgst_rate","cgst_amount","sgst_rate","sgst_amount","amount"),yscrollcommand=scrolly.set,xscrollcommand=scrollx.set)
        scrollx.pack(side=BOTTOM,fill=X)
        scrolly.pack(side=RIGHT,fill=Y)
        scrollx.config(command=self.tax_invoice_Table.xview)
        scrolly.config(command=self.tax_invoice_Table.yview)

        self.tax_invoice_Table.heading("t_id",text="T ID")
        self.tax_invoice_Table.heading("bill_no",text="Bill No")
        self.tax_invoice_Table.heading("date",text="Date")
        self.tax_invoice_Table.heading("customer",text="Customer")
        self.tax_invoice_Table.heading("service_name",text="Service Name")
        self.tax_invoice_Table.heading("hsn",text="HSN/SAC")
        self.tax_invoice_Table.heading("quantity",text="Quantity")
        self.tax_invoice_Table.heading("unit",text="Unit")
        self.tax_invoice_Table.heading("rate",text="Rate (RS)")
        self.tax_invoice_Table.heading("cgst_rate",text="CGST Rate (RS)")
        self.tax_invoice_Table.heading("cgst_amount",text="CGST Amount (RS)")
        self.tax_invoice_Table.heading("sgst_rate",text="SGST Rate (RS)")
        self.tax_invoice_Table.heading("sgst_amount",text="SGST Amount (RS)")
        self.tax_invoice_Table.heading("amount",text="Amount (RS)")

        self.tax_invoice_Table["show"]="headings"

        self.tax_invoice_Table.column("t_id",width=30)
        self.tax_invoice_Table.column("bill_no",width=30)
        self.tax_invoice_Table.column("date",width=50)
        self.tax_invoice_Table.column("customer",width=50)
        self.tax_invoice_Table.column("service_name",width=100)
        self.tax_invoice_Table.column("hsn",width=30)
        self.tax_invoice_Table.column("quantity",width=50)
        self.tax_invoice_Table.column("unit",width=50)
        self.tax_invoice_Table.column("rate",width=50)
        self.tax_invoice_Table.column("cgst_rate",width=50)
        self.tax_invoice_Table.column("cgst_amount",width=50)
        self.tax_invoice_Table.column("sgst_rate",width=50)
        self.tax_invoice_Table.column("sgst_amount",width=50)
        self.tax_invoice_Table.column("amount",width=50)
        self.tax_invoice_Table.pack(fill=BOTH,expand=1)
        self.tax_invoice_Table.bind("<ButtonRelease-1>",self.get_data)

        self.update_content()
        self.show()
    #=======================================================================================================================================================================
    def validate_length(self, new_value):
        return len(new_value) <= 32  # allow only up to 32 characters
    
    def fetch_customer(self):
        self.customer_list.append("Empty")
        con=sqlite3.connect(database=r'./ah.db')
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
        con = sqlite3.connect(database=r'./ah.db')
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

            # Insert into tax_invoice table
            cur.execute("""
                INSERT INTO tax_invoice(
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
            messagebox.showinfo("Success", "Tax Invoice entry added successfully!", parent=self.root)
            self.clear()
            self.show()

        except Exception as ex:
            messagebox.showerror("Error", f"Error due to: {str(ex)}", parent=self.root)
        finally:
            con.close()
        
    def show(self):
        con=sqlite3.connect(database=r'./ah.db')
        cur=con.cursor()
        try:
            cur.execute("select * from tax_invoice")
            rows=cur.fetchall()
            self.tax_invoice_Table.delete(*self.tax_invoice_Table.get_children())
            for row in rows:
                self.tax_invoice_Table.insert('',END,values=row)
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :{str(ex)}",parent=self.root)

    def get_data(self, ev):
        f = self.tax_invoice_Table.focus()
        content = self.tax_invoice_Table.item(f)
        row = content['values']
        
        self.var_t_id.set(row[0])
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
        self.btn_export_search.config(state='normal')
    
    def update(self):
        con=sqlite3.connect(database=r'./ah.db')
        cur=con.cursor()
        try:
            if self.var_t_id.get()=="":
                messagebox.showerror("Error","t_id Must be required",parent=self.root)
            else:
                cur.execute("Select * from tax_invoice where t_id=?",(self.var_t_id.get(),))
                row=cur.fetchone()
                if row==None:
                    messagebox.showerror("Error","Invalid t_id",parent=self.root)
                else:
                    cur.execute("Update tax_invoice set bill_no=?,date=?,customer=?,service_name=?,hsn=?,quantity=?,unit=?,rate=?,cgst_rate=?,cgst_amount=?,sgst_rate=?,sgst_amount=?,amount=? where t_id=?",(
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
                                        self.var_t_id.get()
                    ))
                    con.commit()
                    messagebox.showinfo("Success","Tax Invoice Updated Successfully",parent=self.root)
                    self.clear() # Clear fields after updating
                    self.show()
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to : {str(ex)}",parent=self.root)

    def delete(self):
        if self.var_t_id.get() == "":
            messagebox.showerror("Error", "t_id must be required", parent=self.root)
        else:
            con = sqlite3.connect(database=r'./ah.db')
            cur = con.cursor()
            cur.execute("SELECT * FROM tax_invoice WHERE t_id=?", (self.var_t_id.get(),))
            row = cur.fetchone()
            if row is None:
                messagebox.showerror("Error", "Invalid t_id", parent=self.root)
            else:
                op = messagebox.askyesno("Confirm", "Do you really want to delete?", parent=self.root)
                if op:
                    cur.execute("DELETE FROM tax_invoice WHERE t_id=?", (self.var_t_id.get(),))
                    con.commit()
                    messagebox.showinfo("Delete", "Tax Invoice Deleted Successfully", parent=self.root)
                    self.clear()    
    
    def clear(self):
        self.var_t_id.set("0")
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
        self.root.after(100, lambda: self.txt_t_id.focus_set())
        self.var_searchtxt.set("")
        self.var_searchby.set("Select")        
        # Button control
        self.btn_add.config(state='normal')
        self.btn_update.config(state='disabled')
        self.btn_delete.config(state='disabled')
        self.btn_export_search.config(state='disabled')
       
        self.show()
    
    def show_service_name_suggestions(self, event=None):
        typed = self.var_service_name.get()

        if typed == "":
            if hasattr(self, 'service_name_listbox') and self.service_name_listbox.winfo_exists():
                self.service_name_listbox.destroy()
            return

        con = sqlite3.connect('./ah.db')
        cur = con.cursor()
        cur.execute("SELECT DISTINCT service_name FROM tax_invoice WHERE service_name LIKE ?", (typed + "%",))
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

        con = sqlite3.connect('./ah.db')
        cur = con.cursor()
        cur.execute("SELECT DISTINCT hsn FROM tax_invoice WHERE hsn LIKE ?", (typed + "%",))
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

        con = sqlite3.connect('./ah.db')
        cur = con.cursor()
        cur.execute("SELECT DISTINCT unit FROM tax_invoice WHERE unit LIKE ?", (typed + "%",))
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
        con=sqlite3.connect(database=r'./ah.db')
        cur=con.cursor()
        try:
            if self.var_searchby.get()=="Select":
                messagebox.showerror("Error","Select Search by option",parent=self.root)
            elif self.var_searchtxt.get()=="":
                messagebox.showerror("Error","Search input should be required",parent=self.root)
            
            else:
                cur.execute("select * from tax_invoice where "+self.var_searchby.get()+" LIKE '%"+self.var_searchtxt.get()+"%'")
                rows=cur.fetchall()
                if len(rows)!=0:
                    self.tax_invoice_Table.delete(*self.tax_invoice_Table.get_children())
                    for row in rows:
                        self.tax_invoice_Table.insert('',END,values=row)
                else:
                    messagebox.showerror("Error","No record found!!!",parent=self.root)
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :{str(ex)}",parent=self.root)
   
    def send_whatsapp_message(self):
        phone_number = self.var_hsn.get()  # Get tax_invoice phone number
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
            
    def export_searched_tax_invoice_to_pdf_excel(self):
        rows = []
        for child in self.tax_invoice_Table.get_children():
            rows.append(self.tax_invoice_Table.item(child)['values'])

        if not rows:
            messagebox.showwarning("No Data", "There is no data to export.")
            return

        columns = ["T ID", "Bill No", "Date", "Customer", "Service Name", "HSN/SAC", "Quantity", "Unit",
                "Rate (RS)", "CGST Rate (RS)", "CGST Amount (RS)",
                "SGST Rate (RS)", "SGST Amount (RS)", "Amount (RS)"]

        df = pd.DataFrame(rows, columns=columns)

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
        
        # Save Excel
        excel_file = filedialog.asksaveasfilename(defaultextension="./data/Tax Invoice Data/Export Data/export.xlsx",
                                                filetypes=[("Excel files", "*.xlsx")])
        if not excel_file:
            return
        df.to_excel(excel_file, index=False)

        # Save PDF
        # Fetch firm details
        con = sqlite3.connect("ah.db")
        cur = con.cursor()
        cur.execute("SELECT name, contact, address, email, gst, bank, account_holder_name, account_no, branch_ifs_code FROM firm LIMIT 1")
        firm = cur.fetchone()
        con.close()

        # Default fallbacks in case table is empty
        firm_name = firm[0] if firm else "FIRM NAME"
        firm_contact = firm[1] if firm else "Contact"
        firm_address = firm[2] if firm else "Firm Address"
        firm_email = firm[3] if firm else "Firm Email"
        firm_gst = firm[4] if firm else "GST Number"
        #firm_licence = firm[5] if firm else "Licence Number"
        
        pdf_file = filedialog.asksaveasfilename(defaultextension="./data/Tax Invoice Data/Export Data/export.pdf",
                                                filetypes=[("PDF files", "*.pdf")])
        if not pdf_file:
            return

        pdf = FPDF()
        pdf.add_page()

        pdf.set_font("Arial", 'B', 14)
        pdf.cell(0, 10, "TAX INVOICE", ln=1, align="C")

        # --- Insert logo below TAX INVOICE ---
        logo_path = "./image/along_home_logo_1.png"  # Set your actual logo path here
        if os.path.exists(logo_path):
            # x, y is position (set y after TAX INVOICE)
            pdf.image(logo_path, x=80, y=pdf.get_y(), w=50)  # center aligned logo (adjust w/h if needed)
            pdf.ln(30)  # space below logo
        
        pdf.set_font("Arial", 'B', 12)
        pdf.cell(0, 6, firm_name, ln=1)

        pdf.set_font("Arial", '', 10)
        pdf.multi_cell(0, 5,
            f"M: {firm_contact}\n"
            f"{firm_address.strip()}\n"
            f"Email: {firm_email}\n"
            f"GSTIN/UIN: {firm_gst}\n"
            f"State Name : Gujarat, Code : 24\n"
            #f"FOOD LICENCE NO.: {firm_licence}"
        )
        pdf.ln(2)

        customer_name = df.iloc[0]["Customer"]
        con = sqlite3.connect("./ah.db")
        cur = con.cursor()
        cur.execute("SELECT address, contact, gst, email, 'Gujarat' as state FROM customer WHERE name=?", (customer_name,))
        customer = cur.fetchone()
        con.close()

        pdf.set_font("Arial", 'B', 10)
        pdf.cell(95, 6, "Buyer (Bill to)", border=1)
        pdf.cell(95, 6, "Invoice Details", border=1, ln=1)
        pdf.set_font("Arial", '', 10)
        pdf.cell(95, 6, str(customer_name), border=1)
        pdf.cell(95, 6, f"Invoice No.: {df.iloc[0]['Bill No']}" " " " " " " " " " " " " " " " " " " " " " " " " " " " " f"Dated: {df.iloc[0]['Date']}", border=1, ln=1)

        if customer:
            pdf.cell(95, 6, f"Contact : {customer[1]}", border=1)
            pdf.cell(95, 6, f"GST/UIN:: {customer[2]}", border=1, ln=1)
        else:
            pdf.cell(95, 6, f"GST/UIN:: {customer[2]}", border=1)
        pdf.cell(95, 6, f"Email: {customer[3]}", border=1, ln=1)

        pdf.cell(190, 6, f"Address: {customer[0]}", border=1, ln=1)
        pdf.cell(190, 6, " State Name: Gujarat, Code: 24", border=1, ln=1)
        pdf.ln(2)

        pdf.set_font("Arial", 'B', 9)
        table_cols = ["Sl", "Service Name", "HSN/SAC", "Qty", "Unit", "Rate", "CGST%", "CGST Amt", "SGST%", "SGST Amt", "Amount"]
        col_widths = [7, 40, 16, 14, 9, 17, 14, 18, 14, 18, 24]
        for col, w in zip(table_cols, col_widths):
            pdf.cell(w, 6, col, border=1, align='C')
        pdf.ln()

        pdf.set_font("Arial", '', 9)
        total_qty = total_cgst = total_sgst = total_amount = 0

        for idx, row in df.iterrows():
            pdf.cell(col_widths[0], 6, str(idx + 1), border=1)
            pdf.cell(col_widths[1], 6, str(row["Service Name"]), border=1)
            pdf.cell(col_widths[2], 6, str(row["HSN/SAC"]), border=1)

            # Quantity
            try:
                qty = float(row["Quantity"])
                qty_str = f"{qty:.2f}"
                total_qty += qty
            except:
                qty_str = str(row["Quantity"])
            pdf.cell(col_widths[3], 6, qty_str, border=1, align='R')

            # Unit
            pdf.cell(col_widths[4], 6, str(row["Unit"]), border=1)

            # Rate
            try:
                rate = float(row["Rate (RS)"])
                rate_str = f"Rs. {rate:.2f}"
            except:
                rate_str = str(row["Rate (RS)"])
            pdf.cell(col_widths[5], 6, rate_str, border=1, align='R')

            # CGST%
            pdf.cell(col_widths[6], 6, str(row["CGST Rate (RS)"]), border=1, align='R')

            # CGST Amount
            try:
                cgst_amt = float(row["CGST Amount (RS)"])
                cgst_amt_str = f"Rs. {cgst_amt:.2f}"
                total_cgst += cgst_amt
            except:
                cgst_amt_str = str(row["CGST Amount (RS)"])
            pdf.cell(col_widths[7], 6, cgst_amt_str, border=1, align='R')

            # SGST%
            pdf.cell(col_widths[8], 6, str(row["SGST Rate (RS)"]), border=1, align='R')

            # SGST Amount
            try:
                sgst_amt = float(row["SGST Amount (RS)"])
                sgst_amt_str = f"Rs. {sgst_amt:.2f}"
                total_sgst += sgst_amt
            except:
                sgst_amt_str = str(row["SGST Amount (RS)"])
            pdf.cell(col_widths[9], 6, sgst_amt_str, border=1, align='R')

            # Amount
            try:
                amt = float(row["Amount (RS)"])
                amt_str = f"Rs. {amt:.2f}"
                total_amount += amt
            except:
                amt_str = str(row["Amount (RS)"])
            pdf.cell(col_widths[10], 6, amt_str, border=1, align='R')

            pdf.ln()

        # Total row
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(col_widths[0], 6, "", border=1)
        pdf.cell(col_widths[1], 6, "", border=1)
        pdf.cell(col_widths[2], 6, "Total", border=1)
        pdf.cell(col_widths[3], 6, f"{total_qty:.2f}", border=1, align='R')
        pdf.cell(col_widths[4], 6, "", border=1, align='R')
        pdf.cell(col_widths[5], 6, "", border=1)
        pdf.cell(col_widths[6], 6, "", border=1)
        pdf.cell(col_widths[7], 6, f"Rs. {total_cgst:.2f}", border=1, align='R')
        pdf.cell(col_widths[8], 6, "", border=1)
        pdf.cell(col_widths[9], 6, f"Rs. {total_sgst:.2f}", border=1, align='R')
        pdf.cell(col_widths[10], 6, f"Rs. {total_amount:.2f}", border=1, align='R')
        pdf.ln(8)

        # Amount in Words
        amt_words = num2words(total_amount, lang='en_IN').title() + " Only"
        pdf.set_font("Arial", '', 9)
        pdf.cell(0, 6, f"Amount Chargeable (in words): INR {amt_words}", ln=1)

        # HSN Summary
        # HSN Tax Summary Title
        pdf.set_font("Arial", 'B', 10)
        pdf.cell(0, 6, "HSN-wise Tax Summary", ln=1)
        pdf.set_font("Arial", 'B', 9)

        # Table Header
        pdf.cell(40, 6, "HSN/SAC", border=1)
        pdf.cell(40, 6, "Taxable Value", border=1)
        pdf.cell(30, 6, "CGST", border=1)
        pdf.cell(30, 6, "SGST", border=1)
        pdf.cell(40, 6, "Total Tax", border=1, ln=1)

        # Table Rows
        pdf.set_font("Arial", '', 9)
        grand_taxable = 0
        grand_cgst = 0
        grand_sgst = 0

        for hsn, values in hsn_summary.items():
            taxable = values["taxable"]
            cgst = values["cgst"]
            sgst = values["sgst"]
            total_tax = cgst + sgst

            pdf.cell(40, 6, hsn, border=1)
            pdf.cell(40, 6, f"{taxable:.2f}", border=1)
            pdf.cell(30, 6, f"{cgst:.2f}", border=1)
            pdf.cell(30, 6, f"{sgst:.2f}", border=1)
            pdf.cell(40, 6, f"{total_tax:.2f}", border=1, ln=1)

            grand_taxable += taxable
            grand_cgst += cgst
            grand_sgst += sgst

        # Grand Totals
        pdf.set_font("Arial", 'B', 9)
        pdf.cell(40, 6, "Grand Total", border=1)
        pdf.cell(40, 6, f"{grand_taxable:.2f}", border=1)
        pdf.cell(30, 6, f"{grand_cgst:.2f}", border=1)
        pdf.cell(30, 6, f"{grand_sgst:.2f}", border=1)
        pdf.cell(40, 6, f"{grand_cgst + grand_sgst:.2f}", border=1, ln=1)

        # Tax in Words
        tax_total_words = num2words(grand_cgst + grand_sgst, lang='en_IN').title() + " Only"
        pdf.ln(3)
        pdf.set_font("Arial", '', 9)
        pdf.cell(0, 6, f"Total GST (in words): INR {tax_total_words}", ln=1)
        
        # Declaration
        pdf.set_font("Arial", '', 8)
        pdf.multi_cell(0, 5, "Declaration:\nWe declare that this invoice shows the actual price of the goods described and that all particulars are true and correct.")
        pdf.ln(2)
       
        # Bank Details Section
        pdf.set_font("Arial", '', 9)
        pdf.cell(0, 6, "Company's Bank Details:", border=1, ln=1)

        if firm:
            pdf.cell(0, 6, f"A/c Holder's Name : {firm[6]}", border=1, ln=1)
            pdf.cell(0, 6, f"Bank Name : {firm[5]}", border=1, ln=1)
            pdf.cell(0, 6, f"A/c No. : {firm[7]}", border=1, ln=1)
            pdf.cell(0, 6, f"Branch & IFS Code: {firm[8]}", border=1, ln=1)
        else:
            pdf.cell(0, 6, "Firm bank details not found", border=1, ln=1)
        
        # Signature
        pdf.set_y(-50)
        pdf.set_font("Arial", '', 12)
        pdf.cell(0, 6, "for ALONG HOME HEALTHCARE", ln=1, align="R")
        pdf.cell(0, 6, "(Authorised Signatory)", ln=1, align="R")

        # Notes
        pdf.set_font("Arial", 'I', 8)
        pdf.cell(0, 5, "SUBJECT TO AHMEDABAD JURISDICTION", ln=1, align='C')
        pdf.cell(0, 5, "This is a Computer Generated Invoice", ln=1, align='C')
        
        try:
            pdf.output(pdf_file)
            messagebox.showinfo("Success", f"Invoice exported to:\n{pdf_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save PDF:\n{e}")
    
    def export_to_excel(self):
            con = sqlite3.connect(database=r'./ah.db')
            cur = con.cursor()
            try:
                cur.execute("SELECT * FROM tax_invoice")
                rows = cur.fetchall()
                if not rows:
                    messagebox.showerror("Error", "No data to export", parent=self.root)
                    return

                # Create a DataFrame from the fetched data
                df = pd.DataFrame(rows, columns=["T ID", "Bill No","Date","Customer","Service Name","HSN/SAC","Quantity","Unit","Rate (RS)","CGST Rate (RS)","CGST Amount (RS)","SGST Rate (RS)","SGST Amount (RS)","Amount Due (RS)"])

                # Add "As of Date" column
                as_of_date = datetime.now().strftime("%d-%m-%Y")  # Current date in YYYY-MM-DD format
                df["As of Date"] = as_of_date            # Export to Excel
                
                output_file = "./data/Tax Invoice Data/Export Data/export_exe.xlsx"
                df.to_excel(output_file, index=False, engine='openpyxl')

                messagebox.showinfo("Success", f"Data exported to {output_file}", parent=self.root)

            except Exception as ex:
                messagebox.showerror("Error", f"Error due to : {str(ex)}", parent=self.root)     
    
    def export_to_pdf(self):
        con = sqlite3.connect(database=r'./ah.db')
        cur = con.cursor()
        # Fetch firm details
        cur.execute("SELECT name, address, contact, email, gst FROM firm WHERE f_id = 1")
        firm = cur.fetchone()
        try:
            cur.execute("SELECT * FROM tax_invoice")
            rows = cur.fetchall()
            if not rows:
                messagebox.showerror("Error", "No data to export", parent=self.root)
                return

            pdf_file = filedialog.asksaveasfilename(defaultextension="./data/Tax Invoice Data/Export Data/pdf.pdf", filetypes=[("PDF files", "*.pdf")])
            if not pdf_file:
                return
            os.makedirs(os.path.dirname(pdf_file), exist_ok=True)

            column_widths = [39] * 14 + [40]
            total_width = sum(column_widths)
            portrait_width, portrait_height = A4
            page_size = landscape(A4) if total_width > portrait_width else portrait(A4)

            doc = SimpleDocTemplate(pdf_file, pagesize=page_size, rightMargin=1, leftMargin=1)
            elements = []
            styles = getSampleStyleSheet()
            wrap_style = ParagraphStyle(name='Wrap', fontSize=5, alignment=1)

            # Logo
            logo_path_1 = "./image/along_home_logo_1.png"
            if os.path.exists(logo_path_1):
                elements.append(RLImage(logo_path_1, width=100, height=40))
                elements.append(Spacer(1, 2))

            # Header
            if firm:
                name, address, contact, gst, email = firm
                elements.append(Paragraph(name, styles['Title']))
                elements.append(Spacer(1, 2))
                elements.append(Paragraph(f"Office : {address}", styles['Normal']))
                elements.append(Spacer(2, 2))
                elements.append(Paragraph(f"M : {contact}", styles['Normal']))
                elements.append(Spacer(2, 2))
                elements.append(Paragraph(f"E : {email}", styles['Normal']))
                elements.append(Spacer(2, 2))
                elements.append(Paragraph(f"GSTIN/UIN : {gst}", styles['Normal']))
                elements.append(Spacer(2, 2))
                elements.append(Paragraph("="*55, styles['Title']))
                elements.append(Spacer(2, 2))
            else:
                print("Firm details not found.")
            
            headers = ["T ID", "Bill No", "Date", "Customer", "Service Name", "HSN/SAC", "Quantity","Unit", "Rate (RS)", "CGST Rate (RS)", "CGST Amount (RS)", "SGST Rate (RS)", "SGST Amount (RS)", "Amount Due (RS)"]
            data = [headers]

            # Group rows by service_name (assuming column index 2 is service_name)
            grouped = defaultdict(list)
            for row in rows:
                service_name = row[3]
                grouped[service_name].append(row)

            # Track grand totals
            grand_totals = defaultdict(float)
            numeric_indices_to_total = [6, 10, 12, 13]  # Rate, cgst amount, sgst amount, Amount 

            for service_name, service_name_rows in grouped.items():
                for row in service_name_rows:
                    row_data = []
                    for i, cell in enumerate(row):
                        if i == 14 and cell and os.path.exists(cell):
                            try:
                                img_1 = RLImage(cell, width=30, height=30)
                                row_data.append(img_1)
                            except:
                                row_data.append(Paragraph("Image error", wrap_style))
                        elif i == 15 and cell and os.path.exists(cell):
                            try:
                                img_2 = RLImage(cell, width=30, height=30)
                                row_data.append(img_2)
                            except:
                                row_data.append(Paragraph("Image error", wrap_style))
                        else:
                            row_data.append(Paragraph(str(cell), wrap_style))
                    data.append(row_data)

                    # Add to grand totals
                    for idx in numeric_indices_to_total:
                        try:
                            val = float(row[idx]) if str(row[idx]).replace('.', '', 1).replace('-', '').isdigit() else 0.0
                            grand_totals[idx] += val
                        except:
                            pass

                # Subtotal row for each service_name
                subtotal_row = []
                for i in range(len(headers)):
                    if i == 0:
                        subtotal_row.append(Paragraph(f"<b>{service_name} TOTAL</b>", wrap_style))
                    elif i in numeric_indices_to_total:
                        subtotal = sum(float(r[i]) for r in service_name_rows if str(r[i]).replace('.', '', 1).replace('-', '').isdigit())
                        subtotal_row.append(Paragraph(f"{subtotal:.2f}", wrap_style))
                    else:
                        subtotal_row.append(Paragraph("", wrap_style))
                data.append(subtotal_row)

            # Add final GRAND TOTAL row
            total_row = []
            for i in range(len(headers)):
                if i == 0:
                    total_row.append(Paragraph("<b>GRAND TOTAL</b>", wrap_style))
                elif i in numeric_indices_to_total:
                    total_row.append(Paragraph(f"{grand_totals[i]:.2f}", wrap_style))
                else:
                    total_row.append(Paragraph("", wrap_style))
            data.append(total_row)

            table = Table(data, repeatRows=1, colWidths=[45]*15 + [40])
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

            # Signatures
            elements.append(Paragraph("Prepared By", styles['Normal']))
            elements.append(Spacer(1, 2))

            elements.append(Paragraph("------------", styles['Normal']))
            elements.append(Spacer(1, 2))
                        
            doc.build(elements)
            messagebox.showinfo("Success", f"PDF exported to:\n{pdf_file}", parent=self.root)

        except Exception as ex:
            messagebox.showerror("Error", f"Error due to: {str(ex)}", parent=self.root)
    
    def import_excel_to_db(self):
        # Ask user to select Excel file
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not file_path:
            return
        try:
            # Read Excel into DataFrame
            df = pd.read_excel(file_path)

            # Connect to database
            con = sqlite3.connect('ah.db')
            cur = con.cursor()

            for index, row in df.iterrows():
                cur.execute("""INSERT OR IGNORE INTO tax_invoice(t_id, bill_no, date, customer, service_name, hsn, quantity, unit, rate, cgst_rate, cgst_amount, sgst_rate, sgst_amount, amount
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""", (
                    row['t_id'],
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
        con=sqlite3.connect(database=r'./ah.db')
        cur=con.cursor()
        try:
            time_=time.strftime("%I:%M:%S %p")
            date_=time.strftime("%d-%m-%Y")
            self.lbl_clock.config(text=f"  ALONG HOME HEALTHCARE \t\t TAX INVOICE DATA\t\t\t Date: {str(date_)}\t\t Time: {str(time_)}")
            self.lbl_clock.after(200,self.update_content)
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :   {str(ex)}",parent=self.root)
        
if __name__=="__main__":
    root=Tk()
    obj=tax_invoiceClass(root)
    root.mainloop()