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
class customerClass:
    def __init__(self,root):
        self.root=root
        self.root.geometry("1100x500+220+130")
        self.root.title("Along Home Healthcare | Developed by Brijesh | July | 9998002712")
        self.root.config(bg="white")
        self.root.focus_force()
        # All Variables================
        self.var_searchby=StringVar()
        self.var_searchtxt=StringVar()

        self.var_c_id=StringVar()
        self.var_serial_no=StringVar()
        self.var_name=StringVar()
        self.var_contact=StringVar()
        self.var_email=StringVar()
        self.var_gst=StringVar()
        self.var_address=StringVar()
        #============Search Frame==================
        SearchFrame=LabelFrame(self.root,text="Search Customer Data",font=("helvetica",12,"bold"),bd=2,relief=RIDGE,bg="white")
        SearchFrame.place(x=530,y=50,width=550,height=65)

        cmb_search=ttk.Combobox(SearchFrame,textvariable=self.var_searchby,values=("Select","C_ID","SERIAL_NO","NAME","CONTACT"),state='readonly',justify=CENTER,font=("helvetica",15))
        cmb_search.place(x=10,y=10,width=110)
        cmb_search.current(0)

        txt_search=Entry(SearchFrame,textvariable=self.var_searchtxt,font=("helvetica",15),bg="lightyellow").place(x=130,y=10,width=90)
        btn_search=Button(SearchFrame,text="Search",command=self.search,font=("helvetica",15),bg="black",fg="white",cursor="hand2").place(x=230,y=9,width=80,height=30)
        
        self.btn_export_to_search=Button(SearchFrame,text="Export Search to PDF",command=self.export_searched_customer_to_pdf,font=("helvetica",15),bg="red",fg="white",cursor="hand2")
        self.btn_export_to_search.place(x=320,y=9,width=200,height=30)
        #===========Title===================
        title=Label(self.root,text="Customer Details",font=("helvetica",20,"bold"),bg="#0f4d7d",fg="white").place(x=50,y=10,width=1000,height=40)
        #===========Content================
        lbl_serial_no=Label(self.root,text="Serial No",font=("helvetica",15),bg="white").place(x=50,y=80)
        self.txt_serial_no=Entry(self.root,textvariable=self.var_serial_no,font=("helvetica",15),bg="lightyellow")
        self.txt_serial_no.place(x=180,y=80,width=180)

        lbl_name=Label(self.root,text="Name",font=("helvetica",15),bg="white").place(x=50,y=130)
        txt_name=Entry(self.root,textvariable=self.var_name,font=("helvetica",15),bg="lightyellow").place(x=180,y=130,width=180)
        
        lbl_contact=Label(self.root,text="Contact",font=("helvetica",15),bg="white")
        lbl_contact.place(x=50,y=180)
        txt_contact=Entry(self.root,textvariable=self.var_contact,font=("helvetica",15),bg="lightyellow")
        txt_contact.place(x=180,y=180,width=180)
        #txt_contact.bind("<FocusOut>", self.validate_contact)  # Call validate_phone when focus is lost        
        
        lbl_email=Label(self.root,text="Email ID",font=("helvetica",15),bg="white").place(x=50,y=230)
        txt_email=Entry(self.root,textvariable=self.var_email,font=("helvetica",15),bg="lightyellow").place(x=180,y=230,width=180)
        
        lbl_gst=Label(self.root,text="GST No.",font=("helvetica",15),bg="white").place(x=50,y=280)
        txt_gst=Entry(self.root,textvariable=self.var_gst,font=("helvetica",15),bg="lightyellow").place(x=180,y=280,width=180)

        lbl_address=Label(self.root,text="Address",font=("helvetica",15),bg="white").place(x=50,y=330)
        self.txt_address=Text(self.root,font=("helvetica",15),bg="lightyellow")
        self.txt_address.place(x=180,y=330,width=320,height=70)
        #==========Buttons=========
        self.btn_add=Button(self.root,text="Save",command=self.add,font=("helvetica",15),bg="#2196f3",fg="white",cursor="hand2")
        self.btn_add.place(x=50,y=420,width=110,height=35)
        self.btn_update=Button(self.root,text="Update",command=self.update,font=("helvetica",15),bg="#4caf50",fg="white",cursor="hand2")
        self.btn_update.place(x=170,y=420,width=110,height=35)
        self.btn_delete=Button(self.root,text="Delete",command=self.delete,font=("helvetica",15),bg="#f44336",fg="white",cursor="hand2")
        self.btn_delete.place(x=290,y=420,width=110,height=35)
        self.btn_clear=Button(self.root,text="Clear",command=self.clear,font=("helvetica",15),bg="#607d8b",fg="white",cursor="hand2")
        self.btn_clear.place(x=410,y=420,width=110,height=35)

        self.btn_add.config(state='normal')
        self.btn_update.config(state='disabled')
        self.btn_delete.config(state='disabled')
        #self.btn_export_search.config(state='disabled')
        
        self.btn_export_to_pdf = Button(self.root, text="Export to pdf", command=self.export_to_pdf, font=("helvetica", 15), bg="red", fg="white", cursor="hand2")
        self.btn_export_to_pdf.place(x=530, y=420, width=150, height=35)
        
        btn_export = Button(self.root, text="Export to Excel", command=self.export_to_excel, font=("helvetica", 15), bg="yellow", fg="black", cursor="hand2")
        btn_export.place(x=690, y=420, width=150, height=35)
        #=========customer Details========
        customer_frame=Frame(self.root,bd=3,relief=RIDGE)
        customer_frame.place(x=530,y=120,width=550,height=290)

        scrolly=Scrollbar(customer_frame,orient=VERTICAL)
        scrollx=Scrollbar(customer_frame,orient=HORIZONTAL) 
        
        self.customerTable=ttk.Treeview(customer_frame,columns=("c_id","serial_no","name","contact","email","gst","address"),yscrollcommand=scrolly.set,xscrollcommand=scrollx.set)
        scrollx.pack(side=BOTTOM,fill=X)
        scrolly.pack(side=RIGHT,fill=Y)
        scrollx.config(command=self.customerTable.xview)
        scrolly.config(command=self.customerTable.yview)

        self.customerTable.heading("c_id",text="c_id")
        self.customerTable.heading("serial_no",text="Sr No")
        self.customerTable.heading("name",text="Name")
        self.customerTable.heading("contact",text="Contact")
        self.customerTable.heading("email",text="Email ID")
        self.customerTable.heading("gst",text="GST No.")
        self.customerTable.heading("address",text="Address")

        self.customerTable["show"]="headings"
        
        self.customerTable.column("c_id",width=30)
        self.customerTable.column("serial_no",width=30)
        self.customerTable.column("name",width=100)
        self.customerTable.column("contact",width=60)
        self.customerTable.column("email",width=50)
        self.customerTable.column("gst",width=50)
        self.customerTable.column("address",width=100)
        self.customerTable.pack(fill=BOTH,expand=1)
        self.customerTable.bind("<ButtonRelease-1>",self.get_data)

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
        try:
            if self.var_serial_no.get() == "Select" or self.var_name.get() == "Select" or self.var_contact.get() == "Empty":
                messagebox.showerror("Error", "Serial No must be required", parent=self.root)
            else:
                # Use ? for MySQL placeholders
                cur.execute("SELECT * FROM customer WHERE serial_no=? AND name=? AND contact=? AND gst=?", (
                    self.var_serial_no.get(),
                    self.var_name.get(),
                    self.var_contact.get(),
                    self.var_gst.get(),
                ))
                row = cur.fetchone()
                if row is not None:
                    messagebox.showerror("Error", "Serial No is already assigned, try different", parent=self.root)
                else:
                    cur.execute("INSERT INTO customer(serial_no, name, contact, email, gst, address) VALUES (?, ?, ?, ?, ?, ?)", (
                        self.var_serial_no.get(),
                        self.var_name.get(),
                        self.var_contact.get(),
                        self.var_email.get(),
                        self.var_gst.get(),
                        self.txt_address.get('1.0', END).strip()  # Remove trailing newline
                    ))
                    conn.commit()
                    messagebox.showinfo("Success", "Customer Added Successfully", parent=self.root)
                    self.clear()
                    self.show()
        except Exception as ex:
            messagebox.showerror("Error", f"Error due to : {str(ex)}", parent=self.root)
        finally:
            conn.close()

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
            cur.execute("Select * from customer")
            rows=cur.fetchall()
            self.customerTable.delete(*self.customerTable.get_children())
            for row in rows:
                self.customerTable.insert('',END,values=row)

        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :{str(ex)}",parent=self.root)

    def get_data(self,ev):
        f=self.customerTable.focus()
        content=(self.customerTable.item(f))
        row=content['values']
        # print(row)  
        self.var_c_id.set(row[0])
        self.var_serial_no.set(row[1])
        self.var_name.set(row[2])
        self.var_contact.set(row[3])
        self.var_email.set(row[4])
        self.var_gst.set(row[5])
        self.txt_address.delete('1.0',END)   
        self.txt_address.insert(END,row[6])                                  

        # Button control
        self.btn_add.config(state='disabled')
        self.btn_update.config(state='normal')
        self.btn_delete.config(state='normal')
        #self.btn_export_search.config(state='normal')
    
    def update(self):
        conn = pyodbc.connect(
                r"DRIVER={ODBC Driver 17 for SQL Server};"
                r"SERVER=DESKTOP-9VTU0IO\SQLEXPRESS;"
                r"DATABASE=ALONGHOME;"
                r"Trusted_Connection=yes;"
                r"Encrypt=no;"
        )
        cur = conn.cursor()
        try:
            if self.var_c_id.get() == "":
                messagebox.showerror("Error", "Customer ID must be required", parent=self.root)
            else:
                # Use ? placeholders for MySQL
                cur.execute("SELECT * FROM customer WHERE c_id=?", (self.var_c_id.get(),))
                row = cur.fetchone()
                if row is None:
                    messagebox.showerror("Error", "Invalid Customer ID", parent=self.root)
                else:
                    cur.execute("""
                        UPDATE customer SET 
                            serial_no=?,
                            name=?, 
                            contact=?,
                            email=?, 
                            gst=?,                                  
                            address=?, 
                        WHERE c_id=?
                    """, (
                        self.var_serial_no.get(),
                        self.var_name.get(),
                        self.var_contact.get(),
                        self.txt_address.get('1.0', END).strip(),  # Remove trailing newline
                        self.var_email.get(),
                        self.var_gst.get(),
                        int(self.var_c_id.get())  # ensure c_id is int if DB column is int
                    ))

                    # cur.execute("""
                    #     UPDATE customer SET 
                    #         serial_no=?,
                    #         name=?,
                    #         contact=?,
                    #         email=?,
                    #         gst=?,
                    #         address=?
                    #     WHERE c_id=?
                    # """, (
                    #     self.var_serial_no.get(),
                    #     self.var_name.get(),
                    #     self.var_contact.get(),
                    #     self.var_email.get(),
                    #     self.var_gst.get(),
                    #     self.txt_address.get("1.0", END).strip(),
                    #     self.var_c_id.get()
                    # ))
                    conn.commit()
                    messagebox.showinfo("Success", "Customer updated successfully", parent=self.root)
                    self.clear()
                    self.show()
        except Exception as ex:
            messagebox.showerror("Error", f"Error due to : {str(ex)}", parent=self.root)
        finally:
            conn.close()

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
            if self.var_c_id.get()=="":
                messagebox.showerror("Error","serial_no must be required",parent=self.root)
            else:
                cur.execute("Select * from customer where c_id=?",(self.var_c_id.get(),))
                row=cur.fetchone()
                if row==None:
                    messagebox.showerror("Error","Invalid serial_no",parent=self.root)
                else:
                    op=messagebox.askyesno("Confirm","Do you really want to delete?",parent=self.root)
                    if op==True:
                        cur.execute("Delete from customer where c_id=?",(self.var_c_id.get(),))
                        conn.commit()
                    messagebox.showinfo("Delete","customer Deleted Successfully",parent=self.root)
                    self.clear()   # Clear fields after saving 
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to : {str(ex)}",parent=self.root)                          

    def clear(self):
        self.var_serial_no.set("")
        self.var_name.set("")
        self.var_contact.set("")
        self.var_email.set("")
        self.var_gst.set("")
        self.txt_address.delete('1.0',END)                                    
        self.var_c_id.set("")
        self.var_searchtxt.set("")
        self.var_searchby.set("Select")        
        self.root.after(100, lambda: self.txt_serial_no.focus_set())        
        # Button control
        self.btn_add.config(state='normal')
        self.btn_update.config(state='disabled')
        self.btn_delete.config(state='disabled')
        #self.btn_export_search.config(state='disabled')
        
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
                cur.execute("select * from customer where "+self.var_searchby.get()+" LIKE '%"+self.var_searchtxt.get()+"%'")
                rows=cur.fetchall()
                if len(rows)!=0:
                    self.customerTable.delete(*self.customerTable.get_children())
                    for row in rows:
                        self.customerTable.insert('',END,values=row)
                else:
                    messagebox.showerror("Error","No record found!!!",parent=self.root)
        except Exception as ex:
            messagebox.showerror("Error",f"Error due to :{str(ex)}",parent=self.root)
    
    def export_searched_customer_to_pdf(self):
        conn = pyodbc.connect(
                r"DRIVER={ODBC Driver 17 for SQL Server};"
                r"SERVER=DESKTOP-9VTU0IO\SQLEXPRESS;"
                r"DATABASE=ALONGHOME;"
                r"Trusted_Connection=yes;"
                r"Encrypt=no;"
        )
        cur = conn.cursor()
        try:
            if self.var_searchby.get() == "Select":
                messagebox.showerror("Error", "Select Search by option", parent=self.root)
                return
            if self.var_searchtxt.get() == "":
                messagebox.showerror("Error", "Search input should be required", parent=self.root)
                return

            # Search customer
            search_query = f"SELECT * FROM customer WHERE {self.var_searchby.get()} LIKE ?"
            cur.execute(search_query, ('%' + self.var_searchtxt.get() + '%',))
            rows = cur.fetchall()

            if not rows:
                messagebox.showerror("Error", "No matching Customer records found!", parent=self.root)
                return

            # Labels (adjust to match your table)
            labels = [
                "Customer Name", "Contact No", "Email",
                "GST No", "Address"
            ]

            pdf = FPDF(orientation='P', unit='mm', format='A4')
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()

            pdf.set_font("helvetica", 'B', 14)
            pdf.cell(0, 10, "Searched Customer Report", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
            pdf.set_font("helvetica", '', 10)
            pdf.cell(0, 10, f"As of Date: {datetime.now().strftime('%d-%m-%Y')}", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='R')
            pdf.ln(5)

            for idx, row in enumerate(rows):
                values = list(row)[2:]  # skip ID
                pdf.set_font("helvetica", 'B', 12)
                pdf.cell(0, 8, f"customer #{idx + 1}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                pdf.ln(1)

                for label, value in zip(labels, values):
                    value = str(value) if value else ""
                    pdf.set_font("helvetica", 'B', 11)
                    pdf.cell(0, 7, label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf.set_font("helvetica", '', 11)
                    pdf.multi_cell(0, 7, value)
                    pdf.ln(1)

                # Line separator
                pdf.set_draw_color(180, 180, 180)
                pdf.line(10, pdf.get_y(), 200, pdf.get_y())
                pdf.ln(5)

                # Auto new page
                if pdf.get_y() > 260:
                    pdf.add_page()

            # Save PDF
            output_dir = r"./data/Customer Data"
            os.makedirs(output_dir, exist_ok=True)
            output_file = os.path.join(output_dir, "Search Customer Data.pdf")
            pdf.output(output_file)

            messagebox.showinfo("Success", f"Searched customer data exported to:\n{output_file}", parent=self.root)

        except Exception as ex:
            messagebox.showerror("Error", f"Error due to: {str(ex)}", parent=self.root)
    
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
            # Fetch all customers
            cur.execute("SELECT * FROM customer")
            rows = cur.fetchall()
            if not rows:
                messagebox.showerror("Error", "No Customer data found", parent=self.root)
                return

            # Define field labels (excluding ID)
            labels = [
                "Customer Name", "Contact No", "Email",
                "GST No", "Address"
            ]

            pdf = FPDF(orientation='P', unit='mm', format='A4')
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.add_page()

            pdf.set_font("helvetica", 'B', 14)
            pdf.cell(0, 10, "Customer Data Report", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
            pdf.set_font("helvetica", '', 10)
            pdf.cell(0, 10, f"As of Date: {datetime.now().strftime('%d-%m-%Y')}", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='R')
            pdf.ln(5)

            for idx, row in enumerate(rows):
                values = list(row)[2:]  # skip customer_id

                pdf.set_font("helvetica", 'B', 12)
                pdf.cell(0, 8, f"Customer #{idx + 1}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                pdf.ln(1)

                for label, value in zip(labels, values):
                    value = str(value) if value is not None else ""
                    pdf.set_font("helvetica", 'B', 11)
                    pdf.cell(0, 7, label, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
                    pdf.set_font("helvetica", '', 11)
                    pdf.multi_cell(0, 7, value)
                    pdf.ln(1)

                # Add separation between customers
                pdf.ln(3)
                pdf.set_draw_color(180, 180, 180)
                pdf.line(10, pdf.get_y(), 200, pdf.get_y())
                pdf.ln(5)

                # Add new page if near bottom
                if pdf.get_y() > 260:
                    pdf.add_page()

            # Save file
            output_dir = r"./data/Customer Data"
            os.makedirs(output_dir, exist_ok=True)
            output_file = os.path.join(output_dir, "All Customers Data.pdf")
            pdf.output(output_file)

            messagebox.showinfo("Success", f"All customer data exported to:\n{output_file}", parent=self.root)

        except Exception as ex:
            messagebox.showerror("Error", f"Error due to: {str(ex)}", parent=self.root)
    
    def export_to_excel(self):
        conn = pyodbc.connect(
                r"DRIVER={ODBC Driver 17 for SQL Server};"
                r"SERVER=DESKTOP-9VTU0IO\SQLEXPRESS;"
                r"DATABASE=ALONGHOME;"
                r"Trusted_Connection=yes;"
                r"Encrypt=no;"
        )
        cur = conn.cursor()
        try:
            cur.execute("SELECT * FROM customer")
            rows = cur.fetchall()
            if not rows:
                messagebox.showerror("Error", "No data to export", parent=self.root)
                return

            # Create a DataFrame from the fetched data
            df = pd.DataFrame(rows, columns=["c_id","serial_no","name","contact","email","gst","address"])

            # Add "As of Date" column
            as_of_date = datetime.now().strftime("%d-%m-%Y")  # Current date in YYYY-MM-DD format
            df["As of Date"] = as_of_date
            # Export to Excel
            output_file = r"./data/Customer Data/all_export.xlsx"
            df.to_excel(output_file, index=False, engine='openpyxl')

            messagebox.showinfo("Success", f"Data exported to {output_file}", parent=self.root)

        except Exception as ex:
            messagebox.showerror("Error", f"Error due to : {str(ex)}", parent=self.root)

if __name__=="__main__":
    root=Tk()
    obj=customerClass(root)
    root.mainloop()