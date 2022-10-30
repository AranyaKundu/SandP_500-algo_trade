from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import os, subprocess, time, openpyxl
fpath = "E:\\Python Learning\\Quant Finance\\algorithmic-trading-python-master\\starter_files\\data.xlsx"

def button1():

    def del_data():
        fname_entry.delete(0, END)
        lname_entry.delete(0, END)
        pf_entry.delete(0, END)
        rqst_val.set("Radio")

    def isChecked():
        if yesno_box.get() == 1:
            mynest_Button1['state'] = 'active'
        else:
            mynest_Button1['state'] = 'disable'
            messagebox.showwarning(title="Warning", message="Please agree to the terms and conditions to proceed!")

    def eqwt_run(fpath):
        portfolio_size = pf_entry.get()

        try:
            val = float(portfolio_size)
            fname_ew = fname_entry.get()
            lname_ew= lname_entry.get()
            if fname_ew == "" or lname_ew == "":
                messagebox.askokcancel("Want to continue?", "You didn't write your fullname.")
            else:
                strategy = rqst_val.get()
                if strategy == "":
                    strategy = "Batch"
                if not os.path.exists(fpath):
                    wb = openpyxl.Workbook()
                    sheet = wb.active
                    my_cols = ["Sl.No.", "First Name", "Last Name", "Portfolio Size", "Strategy Type"]
                    sheet.append(my_cols)
                    cols = ['A', 'B', 'C', 'D', 'E']
                    for i in cols:
                        sheet.column_dimensions[i].width = 18
                    wb.save(fpath)
                wb = openpyxl.load_workbook(fpath)
                sheet = wb.active
                maxrow = sheet.max_row
                sheet.append([maxrow, fname_ew, lname_ew, portfolio_size, strategy])
                wb.save(fpath)
                del_data()
                time.sleep(2)
                subprocess.run(["python", "e:\\Python Learning\\Quant Finance\\algorithmic-trading-python-master\\starter_files\\equal_weights.py"], shell=True)
                win.destroy()
        except ValueError:
            messagebox.showwarning(title = "Warning", message = "Please agree to the terms and conditions to proceed!")


    win = Toplevel()
    win.geometry("350x350")
    win.resizable(height = False, width = False)
    win.title("Equal Weights Recommendation")
    win.wm_iconbitmap("e:\\Python Learning\\Quant Finance\\algorithmic-trading-python-master\\starter_files\\qt.ico")
    l1 = ['Individual', 'Batch', 'Specific']

    user_entry_frame = LabelFrame(win, text = "User Information")
    user_entry_frame.grid(row = 0, column = 0, padx = 20, pady = (5, 20), sticky = "news")
    firstname_label = Label(user_entry_frame, text = "First Name")
    firstname_label.grid(row = 0, column = 0, sticky = "news", padx=5, pady=5)
    lastname_label = Label(user_entry_frame, text = "Last Name")
    lastname_label.grid(row=1, column = 0, sticky = "news", padx = 5, pady = 5)
    fname_entry = Entry(user_entry_frame, width = 35)
    fname_entry.grid(row = 0, column = 1, sticky = "news", padx = 5, pady = 5)
    lname_entry = Entry(user_entry_frame, width = 35)
    lname_entry.grid(row = 1, column = 1, sticky = "news", padx = 5, pady = 5)

    user_input_info = LabelFrame(win, text = "Your portfolio details")
    user_input_info.grid(row = 1, column = 0, padx = 20, pady = (5, 20), sticky = "news")
    pf_size = Label(user_input_info, text = "Portfolio Size")
    pf_size.grid(row=0, column = 0, sticky = "news")
    pf_entry = Entry(user_input_info, width = 20)
    pf_entry.grid(row = 0, column = 1, columnspan = 3, sticky = "news", padx = 5, pady = 5)
    rqt_entry = Label(user_input_info, text = "Request Type")
    rqt_entry.grid(row = 1, column = 0, sticky = "news", padx = 5, pady = 5)
    rqst_val = StringVar()
    rqst_val.set("Radio")
    for i in range(1, 4):
        rqt_type = Radiobutton(user_input_info, text=l1[i-1], variable = rqst_val, value=l1[i-1])
        rqt_type.grid(row = 1, column=i, sticky="news")
    yesno_box = IntVar()
    tnc1 = Checkbutton(win, text="I agree to the terms and conditions", variable=yesno_box, onvalue = 1, offvalue = 0, command = isChecked)
    tnc1.grid(row = 2, column = 0, sticky = "news", pady = (5, 20))

    mynest_Button1 = Button(win, text = "Run", bg = '#333333', fg='#ffffff', font = 'Helvetica', state = 'disable', command = lambda: eqwt_run(fpath))
    mynest_Button1.grid(row = 3, column = 0)


def button2():

    def del_data():
        firstname_entry.delete(0, END)
        lastname_entry.delete(0, END)
        pf_entry.delete(0, END)
        strategy_type.current(0)

    def isChecked():
        if accept_var.get() == 1:
            myqms_Button1['state'] = 'active'
        else:
            mynest_Button1['state'] = 'disable'
            messagebox.showwarning(title = "Warning", message = "Please agree to the terms and conditions to proceed!")

    def qms(fpath):
        portfolio_size = pf_entry.get()

        try:
            val = float(portfolio_size)
            fname = firstname_entry.get()
            lname = lastname_entry.get()
            if fname == "" or lname == "":
                messagebox.askokcancel("askokcancel", "Want to continue?")  
            else:
                strategy = strategy_type.get()
                if strategy == "":
                    strategy = "Mean"
                if not os.path.exists(fpath):
                    wb = openpyxl.Workbook()
                    sheet = wb.active
                    my_cols = ['Sl. No.', 'First Name', 'Last Name', 'Portfolio Size', 'Strategy Type']
                    sheet.append(my_cols)
                    cols =['A', 'B', 'C', 'D', 'E']
                    for i in cols:
                        sheet.column_dimensions[i].width = 18
                    wb.save(fpath)
                wb = openpyxl.load_workbook(fpath)
                sheet = wb.active
                maxrow = sheet.max_row
                sheet.append([maxrow, fname, lname, portfolio_size, strategy])
                wb.save(fpath)
                del_data()
                subprocess.run(["python", "e:\\Python Learning\\Quant Finance\\algorithmic-trading-python-master\\starter_files\\quant_momentum.py"], shell=True, 
                env=os.environ)
                time.sleep(4)
                window.destroy()

        except ValueError: messagebox.showwarning(title = "Warning", message = "Please enter a positive integer for your portfolio size.")


    window = Toplevel()
    window.geometry("300x350")
    window.resizable(height = False, width = False)
    window.title("Quantitative Momentum Strategy")
    window.wm_iconbitmap("e:\\Python Learning\\Quant Finance\\algorithmic-trading-python-master\\starter_files\\qt.ico")

    user_entry_fr = LabelFrame(window, text = "User Information")
    user_entry_fr.grid(row = 0, column = 0, sticky = "news", padx = 15, pady = 20)
    fname_label = Label(user_entry_fr, text = "First Name") 
    fname_label.grid(row = 0, column = 0, sticky = "news", padx=5, pady=5)
    lastname_label = Label(user_entry_fr, text = "Last Name")
    lastname_label.grid(row=1, column = 0, sticky = "news", padx = 5, pady = 5)
    firstname_entry = Entry(user_entry_fr, width = 30)
    firstname_entry.grid(row = 0, column = 1, sticky = "news", padx = 5, pady = 5)
    lastname_entry = Entry(user_entry_fr, width = 30)
    lastname_entry.grid(row = 1, column = 1, sticky = "news", padx = 5, pady = 5)

    user_input = LabelFrame(window, text = "Your portfolio details")
    user_input.grid(row = 1, column = 0, padx = 15, pady = (5, 20), sticky = "news")
    user_pf_size = Label(user_input, text = "Portfolio Size")
    user_pf_size.grid(row=0, column = 0, sticky = "news")
    pf_entry = Entry(user_input, width = 28)
    pf_entry.grid(row = 0, column = 1, sticky = "news", padx = 5, pady = 5)
    rqst_entry = Label(user_input, text = "Strategy Type")
    rqst_entry.grid(row = 1, column = 0, sticky = "news", padx = 5, pady = 5)
    strategy_type = ttk.Combobox(user_input, values=["Mean", "Median"])
    strategy_type.grid(row = 1, column = 1, sticky = "w", padx = (5, 0))
    
    accept_var = IntVar()
    terms_check = Checkbutton(window, text= "I agree to the terms and conditions.", variable=accept_var, onvalue=1, offvalue=0, command=isChecked)
    terms_check.grid(row=2, column=0)
    myqms_Button1 = Button(window, text = "Run", bg = '#333333', fg='#ffffff', font = 'Helvetica', state ='disable', command = lambda: qms(fpath))
    myqms_Button1.grid(row = 3, column = 0)


def button3():

    def del_data():
        fname.delete(0, END)
        lname.delete(0, END)
        user_pf_size.delete(0, END)
        strategy_type.current(0)

    def isChecked():
        if yesno_var.get() == 1:
            my_qvs_Button1['state'] = 'active'
        else:
            mynest_Button1['state'] = 'dsiable'
            messagebox.showwarning(title="Warning", message="Please agree to the terms and conditions to proceed!")
    
    def qvs(fpath):
        portfolio_size = user_pf_size.get()

        try:
            val = float(portfolio_size)
            fname_qvs = fname.get()
            lname_qvs= lname.get()
            if fname_qvs == "" or lname_qvs == "":
                messagebox.askokcancel("Want to continue?", "You didn't write your fullname.")
            else:
                strategy = strategy_type.get()
                if strategy == "":
                    strategy = "Mean"
                if not os.path.exists(fpath):
                    wb = openpyxl.Workbook()
                    sheet = wb.active
                    my_cols = ["Sl. No.", "First Name", "Last Name", "Portfolio Size", "Strategy Type"]
                    sheet.append(my_cols)
                    cols = ['A', 'B', 'C', 'D', 'E']
                    for i in cols:
                        sheet.column_dimensions[i].width = 18
                    wb.save(fpath)
                wb = openpyxl.load_workbook(fpath)
                sheet = wb.active
                maxrow = sheet.max_row
                sheet.append([maxrow, fname_qvs, lname_qvs, portfolio_size, strategy])
                wb.save(fpath)
                del_data()
                time.sleep(2)
                subprocess.run(["python", "e:\\Python Learning\\Quant Finance\\algorithmic-trading-python-master\\starter_files\\quant_value.py"], shell=True)
                rwm.destroy()
        except ValueError:
            messagebox.showwarning("Warning", "Please enter a positive integer for your portfolio size.")  
        

    rwm = Toplevel()
    rwm.geometry("300x350")
    rwm.resizable(height = False, width = False)
    rwm.title("Quantitative Value Strategy")
    rwm.wm_iconbitmap("e:\\Python Learning\\Quant Finance\\algorithmic-trading-python-master\\starter_files\\qt.ico")

    user_entry = LabelFrame(rwm, text = "User Information")
    user_entry.grid(row = 0, column = 0, padx = 15, pady = (20, 20), sticky="news")
    firstname = Label(user_entry, text = "First Name")
    firstname.grid(row = 0, column = 0, sticky = "news", padx=5, pady=5)
    lastname = Label(user_entry, text = "Last Name")
    lastname.grid(row=1, column = 0, sticky = "news", padx = 5, pady = 5)
    fname = Entry(user_entry, width = 30)
    fname.grid(row = 0, column = 1, sticky = "news", padx = 5, pady = 5)
    lname = Entry(user_entry, width = 30)
    lname.grid(row = 1, column = 1, sticky = "news", padx = 5, pady = 5)

    user_pf_details = LabelFrame(rwm, text = "Your portfolio details")
    user_pf_details.grid(row = 1, column = 0, pady = (5, 20), sticky = "news")
    user_pf = Label(user_pf_details, text = "Portfolio Size")
    user_pf.grid(row=0, column = 0, sticky = "news")
    user_pf_size = Entry(user_pf_details, width = 28)
    user_pf_size.grid(row = 0, column = 1, sticky = "news", padx = 5, pady = 5)
    user_rqst_entry = Label(user_pf_details, text = "Strategy Type")
    user_rqst_entry.grid(row = 1, column = 0, sticky = "news", padx = 5, pady = 5)
    strategy_type = ttk.Combobox(user_pf_details, values=["Mean", "Median"])
    strategy_type.grid(row = 1, column = 1, sticky = "w", padx = (5, 0))
    yesno_var = IntVar()
    terms_check = Checkbutton(rwm, text = "I agree to the terms and conditions.", variable = yesno_var, onvalue = 1, offvalue = 0, command = isChecked)
    terms_check.grid(row=2, column=0)
    my_qvs_Button1 = Button(rwm, text = "Run", bg = '#333333', fg = '#ffffff', font = 'Helvetica', state = 'disable', command = lambda: qvs(fpath))
    my_qvs_Button1.grid(row = 3, column = 0)


if __name__ == "__main__":
    root = Tk()
    root.geometry("900x600")
    root.resizable(False, False)
    root.title("Your Trading Support")
    root.wm_iconbitmap("e:\\Python Learning\\Quant Finance\\algorithmic-trading-python-master\\starter_files\\qt.ico")

    bg = PhotoImage(file = "e:\\Python Learning\\Quant Finance\\algorithmic-trading-python-master\\starter_files\\qt_background.png")
    label1 = Label(root, image = bg)
    label1.place(x = 0,y = 0)

    myButton1 = Button(root, text = "Equal Weights Strategy", bg = '#0052cc', fg = '#ffffff', font= 'Helvetica 14', command = button1)
    myButton1.config(height = 2, width = 30)
    myButton1.place(x = 80, y = 150)

    myButton2 = Button(root, text = "Quantitative Momentum Strategy", bg = '#0052cc', fg = '#ffffff', font= 'Helvetica 14', command = button2)
    myButton2.config(height = 2, width = 30)
    myButton2.place(x = 490, y = 150)

    myButton3 = Button(root, text = "Quantitative Value Strategy", bg = '#0052cc', fg = '#ffffff', font= 'Helvetica 14', command = button3)
    myButton3.config(height = 2, width = 30)
    myButton3.place(x = 275, y = 275)
    
    myButton4= Button(root, text = "Quit", bg = '#0052cc', fg = '#ffffff', font= 'Helvetica 14', command = root.destroy)
    myButton4.config(height = 2, width = 15)
    myButton4.place(x = 375, y = 400)


    root.mainloop()