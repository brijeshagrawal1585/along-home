import time
import pymysql
import pyodbc
from tkinter import *
from tkinter import messagebox
from PIL import Image, ImageTk
from nurse import nurseClass
from staff import staffClass
from cash_invoice import cash_invoiceClass
from customer import customerClass
from firm import firmClass
from tax_invoice import tax_invoiceClass

class ASMS:
    def __init__(self, root):
        self.root = root
        self.root.state('zoomed')  # Full screen
        self.root.title("Along Home Healthcare | Developed by Brijesh | August | 9998002712")
        self.root.config(bg="white")

        # ===== Title =====
        self.icon_title = PhotoImage(file="./image/nurse1.png")
        title = Label(
            self.root,
            text="Along Home Healthcare",
            image=self.icon_title,
            compound=LEFT,
            font=("helvetica", 40, "bold"),
            bg="#010c48",
            fg="white",
            padx=20,
            pady=10
        )
        title.pack(fill=X)

        # ===== Clock =====
        self.lbl_clock = Label(
            self.root,
            text="Welcome to Along Home Healthcare    Date: --    Time: --",
            font=("helvetica", 15),
            bg="purple",
            fg="white",
            pady=6
        )
        self.lbl_clock.pack(fill=X)

        # ===== Gap After Clock =====
        Frame(self.root, height=15, bg="white").pack(fill=X)

        # ===== Left Menu =====
        left_frame = Frame(self.root, bg="white", bd=2, relief=RIDGE)
        left_frame.place(relx=0, rely=0.16, relwidth=0.16, relheight=0.74)  # Slightly lower and taller

        img = Image.open("./image/along_home_logo_1.png").resize((200, 160), Image.LANCZOS)
        self.MenuLogo = ImageTk.PhotoImage(img)
        Label(left_frame, image=self.MenuLogo).pack(fill=X)

        Label(left_frame, text="Menu", font=("helvetica", 20), bg="#009688").pack(fill=X)
        self.icon_side = PhotoImage(file="./image/side.png")

        for txt, cmd in [("NURSE", self.nurse), ("STAFF", self.staff), ("CUSTOMER", self.customer),
                         ("FIRM", self.firm), ("CASH INVOICE", self.cash_invoice), ("TAX INVOICE", self.tax_invoice)]:
            Button(
                left_frame, text=txt, image=self.icon_side, compound=LEFT, anchor="w",
                font=("helvetica", 14, "bold"), bg="white", bd=2, cursor="hand2",
                command=cmd
            ).pack(fill=X, pady=3, padx=5)

        # ===== Content Panels (Increased Width) =====
        panel_specs = [
            ("Total Nurse", "red", 0.18, 0.20),
            ("Total Staff", "blue", 0.43, 0.20),
            ("Total Customer", "green", 0.68, 0.20),
            ("Total Firm", "maroon", 0.18, 0.42),
            ("Total Cash Invoice", "violet", 0.43, 0.42),
            ("Total Tax Invoice", "navy", 0.68, 0.42),
        ]

        self.panels = {}
        for text, bg, relx, rely in panel_specs:
            lbl = Label(
                self.root,
                text=f"{text}\n[ 0 ]",
                bd=4,
                relief=RIDGE,
                bg=bg,
                fg="white",
                font=("helvetica", 24, "bold"),
                justify=CENTER
            )
            lbl.place(relx=relx, rely=rely, relwidth=0.24, relheight=0.18)
            self.panels[text] = lbl

        # ===== Footer =====
        footer = Label(
            self.root,
            text="Along Home Healthcare | 59/2, Ground Floor, Govt. H. Colony, B/h. Laxmi Ganthiya Rath, Nehrunagar Cross Road, Ahmedabad-380015    Contact: +91 9904110283",
            font=("helvetica", 12),
            bg="black",
            fg="yellow",
            pady=5
        )
        footer.pack(side=BOTTOM, fill=X)

        self.update_content()

    def nurse(self): self._open_window(nurseClass)
    def staff(self): self._open_window(staffClass)
    def customer(self): self._open_window(customerClass)
    def firm(self): self._open_window(firmClass)
    def cash_invoice(self): self._open_window(cash_invoiceClass)
    def tax_invoice(self): self._open_window(tax_invoiceClass)

    def _open_window(self, cls):
        if hasattr(self, 'new_win') and self.new_win.winfo_exists():
            self.new_win.destroy()
        self.new_win = Toplevel(self.root)
        cls(self.new_win)

    def update_content(self):
        try:
            conn = pyodbc.connect(
                        r"DRIVER={ODBC Driver 17 for SQL Server};"
                        r"SERVER=DESKTOP-9VTU0IO\SQLEXPRESS;"
                        r"DATABASE=default;"
                        r"Trusted_Connection=yes;"
                        r"Encrypt=no;"
            )
            cur = conn.cursor()            
            
            for table, label in [
                ("nurse", "Total Nurse"),
                ("staff", "Total Staff"),
                ("customer", "Total Customer"),
                ("firm", "Total Firm"),
                ("cash_invoice", "Total Cash Invoice"),
                ("tax_invoice", "Total Tax Invoice"),
            ]:
                cur.execute(f"SELECT COUNT(*) FROM {table}")
                count = cur.fetchone()[0]
                self.panels[label].config(text=f"{label}\n[ {count} ]")
            conn.close()
        except Exception as ex:
            messagebox.showerror("Error", f"Error due to: {ex}", parent=self.root)

        # Update clock
        # date_ = time.strftime("%d-%m-%Y")
        # time_ = time.strftime("%I:%M:%S %p")
        # self.lbl_clock.config(text=f"Welcome to Along Home Healthcare    Date: {date_}    Time: {time_}")
        # self.root.after(1000, self.update_content)
        time_=time.strftime("%I:%M:%S %p")
        date_=time.strftime("%d-%m-%Y")
        self.lbl_clock.config(text=f"Welcome to Along Home Healthcare\t\t\t Date: {str(date_)}\t\t\t Time: {str(time_)}")
        self.lbl_clock.after(200,self.update_content)
    

if __name__ == "__main__":
    root = Tk()
    ASMS(root)
    root.mainloop()
