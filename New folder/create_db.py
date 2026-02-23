import sqlite3
import pymysql
def create_db():
    #con=pymysql.connect(host='localhost',user='root',password='',db='along_home_db')
    con=sqlite3.connect(database=r'./ah.db')
    cur=con.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS staff(s_id INTEGER PRIMARY KEY AUTOINCREMENT,new_staff_name text,birth_date text,age text,contact text,other_contact text,email_id text,cast text,marital_status text,gender text,religion text,nationality text,experience text,present_address text,permanent_address text,education text,idproofcomb text,id_proof text,nursing_certificate text,exp_letter text,light_bill text,last_work_place text,duty text,financial_year text,image_1 text,image_2 text)")
    con.commit()

    cur.execute("CREATE TABLE IF NOT EXISTS nurse(n_id INTEGER PRIMARY KEY AUTOINCREMENT,new_nurse_name text,birth_date text,age text,contact text,other_contact text,email_id text,cast text,marital_status text,gender text,religion text,nationality text,experience text,present_address text,permanent_address text,education text,idproofcomb text,id_proof text,nursing_certificate text,exp_letter text,light_bill text,last_work_place text,duty text,financial_year text,image_1 text,image_2 text)")
    con.commit()
        
    cur.execute("CREATE TABLE IF NOT EXISTS tax_invoice(t_id INTEGER PRIMARY KEY AUTOINCREMENT,bill_no text,date text,customer text,service_name text,hsn text,quantity text,unit text,rate text,cgst_rate text,cgst_amount text,sgst_rate text,sgst_amount text,amount text)")
    con.commit()
        
    cur.execute("CREATE TABLE IF NOT EXISTS cash_invoice(c_id INTEGER PRIMARY KEY AUTOINCREMENT,bill_no text,date text,customer text,service_name text,hsn text,quantity text,unit text,rate text,cgst_rate text,cgst_amount text,sgst_rate text,sgst_amount text,amount text)")
    con.commit()
    
    cur.execute("CREATE TABLE IF NOT EXISTS firm(f_id INTEGER PRIMARY KEY AUTOINCREMENT,name text,contact text,address text,email text,gst text,bank text,account_holder_name text,account_no text,branch_ifs_code text)")
    con.commit()        

    cur.execute("CREATE TABLE IF NOT EXISTS customer(c_id INTEGER PRIMARY KEY AUTOINCREMENT,serial_no text,name text,contact text,email text,gst text,address text)")
    con.commit()    

create_db()