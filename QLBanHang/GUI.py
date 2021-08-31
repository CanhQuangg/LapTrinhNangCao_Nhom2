from pathlib import Path
import re

from tkinter import Tk, Canvas, Entry, Text, Button, PhotoImage, messagebox, END, filedialog, ttk, LabelFrame, \
    Scrollbar, Toplevel, StringVar, ttk, Event

from openpyxl import *
import pandas as pd
from Class import *

OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path("./assets")

#open database
global wb
wb = load_workbook("Database.xlsx")

# hàm để lấy đường dẫn của file
def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

#cửa sổ đăng nhập
class GUI_Login(Tk):
    def __init__(self, window):
        self.window = window
        self.window.title("Moring Coffee")
        self.window.geometry("1250x800")
        self.window.configure(bg="#FFFFFF")

        self.canvas = Canvas(
            self.window,
            bg = "#FFFFFF",
            height = 800,
            width = 1250,
            bd = 0,
            highlightthickness = 0,
            relief = "ridge"
        )
        self.canvas.place(x=0, y=0)

        #Background
        self.image_LoginBG = PhotoImage(
            file=relative_to_assets("image_1.png"))
        self.Login_Background = self.canvas.create_image(
            625.0,
            400.0,
            image = self.image_LoginBG
        )

        self.image_image_2 = PhotoImage(
            file=relative_to_assets("image_2.png"))
        self.image_2 = self.canvas.create_image(
            625.0,
            361.0,
            image=self.image_image_2
        )

        #Button
        ##Button login with staff account
        self.button_Login_Staff_Image = PhotoImage(
            file=relative_to_assets("Login_Staff.png"))
        self.button_login_staff = Button(
            image=self.button_Login_Staff_Image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonLoginStaff,
            relief="flat"
        )
        self.button_login_staff.place(
            x=625.0,
            y=609.0,
            width=265.0,
            height=77.0
        )

        ##Button login with manager account
        self.button_login_manager_image_2 = PhotoImage(
            file=relative_to_assets("Login_Manager.png"))
        self.button_login_manager = Button(
            image=self.button_login_manager_image_2,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonLoginManager,
            relief="flat"
        )
        self.button_login_manager.place(
            x=360.0,
            y=609.0,
            width=265.0,
            height=77.0
        )

        #Entry
        ##Entry username
        ###Underline entry box
        self.image_image_4 = PhotoImage(
            file=relative_to_assets("image_4.png"))
        self.image_4 = self.canvas.create_image(
            633.0,
            422.0,
            image=self.image_image_4
        )
        ###entry box
        self.entry_username_image = PhotoImage(
            file=relative_to_assets("entry_username.png"))
        self.entry_bg_2 = self.canvas.create_image(
            633,
            377,
            image=self.entry_username_image
        )

        self.entry_username = Entry(
            bd=0,
            bg="#BFBFBF",
            font=("Helvetica", 15),
            highlightthickness=0
        )
        self.entry_username.place(
            x=430.0,
            y=355,
            width=395.0,
            height=61
        )

        ##Entry password
        ### Underline entry password
        self.image_image_3 = PhotoImage(
            file=relative_to_assets("image_3.png"))
        image_3 = self.canvas.create_image(
            633.0,
            541.0,
            image=self.image_image_3
        )

        ###Entry box
        self.entry_password_image = PhotoImage(
            file=relative_to_assets("entry_password.png"))
        self.entry_bg_1 = self.canvas.create_image(
            633.5,
            495.5,
            image=self.entry_password_image
        )
        self.entry_password = Entry(
            bd=0,
            bg="#BFBFBF",
            show="*",
            width=30,
            highlightthickness=0
        )
        self.entry_password.place(
            x=430.0,
            y=475,
            width=395.0,
            height=61
        )

        #Label "Login"
        self.image_image_5 = PhotoImage(
            file=relative_to_assets("label_login.png"))
        self.image_5 = self.canvas.create_image(
            625.0,
            297.0,
            image=self.image_image_5
        )

        self.image_image_6 = PhotoImage(
            file=relative_to_assets("label_brand.png"))
        self.image_6 = self.canvas.create_image(
            625.0,
            210.0,
            image=self.image_image_6
        )
        self.window.resizable(False, False)
        self.window.mainloop()

    #message
    def message_wrong(self):
        messagebox.showerror("Cannot Login", "Username or Password is incorrect")

    def message_wrong_role(self):
        messagebox.showerror("Cannot Login", "Canot login with that role")

    #Login
    def check_Account(self):
        ws_Login = wb['LoginAccount']

        if self.entry_username.get() == '':
            messagebox.showwarning("Cannot Login", "Enter your Username")
            return False
        if self.entry_password.get() == '':
            messagebox.showwarning("Cannot Login", "Enter your Password")
            return False

        for row in ws_Login.values:
            if self.entry_username.get() == str(row[0]):
                if self.entry_password.get() == str(row[1]):
                    if self.role == row[2]:
                        wb.close()
                        return True
                    else:
                        self.message_wrong_role()
                        return False
                else:
                    self.message_wrong()
                    self.entry_password.delete(0, END)
                    return False
            else:
                continue

        #
        self.message_wrong()
        self.entry_username.delete(0, END)
        self.entry_password.delete(0, END)
        return False

    #run button
    def buttonLoginManager(self):
        self.role = "Manager"

        check = self.check_Account()
        if check == True:
            self.entry_username.delete(0, END)
            self.entry_password.delete(0, END)

            self.canvas.destroy()
            self.entry_password.destroy()
            self.entry_username.destroy()
            self.button_login_manager.destroy()
            self.button_login_staff.destroy()

            #self.window.destroy()
            GUI_Manager(self.window)

    def buttonLoginStaff(self):
        self.role = "Staff"

        check = self.check_Account()
        if check == True:
            self.entry_username.delete(0, END)
            self.entry_password.delete(0, END)


# cửa sổ làm việc với vai trò quản lý
class GUI_Manager(Tk):
    def __init__(self,window):
        self.window = window
        self.status = "Overview"
        self.canvas = Canvas(
            #self.manager_window,
            self.window,
            bg="#FFFFFF",
            height=800,
            width=1250,
            bd=0,
            highlightthickness=0,
            relief="ridge"
        )

        # Navigation bar
        self.canvas.place(x=0, y=0)
        self.canvas.create_rectangle(
            0.0,
            0.0,
            1250.0,
            65.0,
            fill="#85603F",
            outline="")

        # Logout button
        self.button_logout_image = PhotoImage(
            file=relative_to_assets("button_logout.png"))
        self.button_logout = Button(
            image=self.button_logout_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonLogout,
            relief="flat"
        )
        self.button_logout.place(
            x=1049.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # Report button
        self.button_report_image = PhotoImage(
            file=relative_to_assets("button_report.png"))
        self.button_report = Button(
            image=self.button_report_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonReport,
            relief="flat"
        )
        self.button_report.place(
            x=883.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        #button_customers
        self.button_customers_image = PhotoImage(
            file=relative_to_assets("button_customers.png"))
        self.button_customers = Button(
            image=self.button_customers_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonCustomers,
            relief="flat"
        )
        self.button_customers.place(
            x=718.0,
            y=0.0,
            width=145.0,
            height=65.0
        )

        #button_cashBook
        self.button_cashBook_image = PhotoImage(
            file=relative_to_assets("button_cashBook.png"))
        self.button_cashBook = Button(
            image=self.button_cashBook_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonCashBook,
            relief="flat"
        )
        self.button_cashBook.place(
            x=552.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        #button_staffs
        self.button_staffs_image = PhotoImage(
            file=relative_to_assets("button_staffs.png"))
        self.button_staffs = Button(
            image=self.button_staffs_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonStaffs,
            relief="flat"
        )
        self.button_staffs.place(
            x=387.0,
            y=0.0,
            width=145.0,
            height=65.0
        )

        #button_warehouse
        self.button_warehouse_image = PhotoImage(
            file=relative_to_assets("button_warehouse.png"))
        self.button_warehouse = Button(
            image=self.button_warehouse_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonWarehouse,
            relief="flat"
        )
        self.button_warehouse.place(
            x=221.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        #button_overview
        self.button_overview_image = PhotoImage(
            file=relative_to_assets("button_overview_show.png"))
        self.button_overview = Button(
            image=self.button_overview_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonOverview,
            relief="flat"
        )
        self.button_overview.place(
            x=55.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        self.frame_overview()

        self.window.mainloop()

    #----------------------------

    #searchBar event
    def search_material(self, Event):
        check = self.entry_searchBar.get()
        list = []
        if check == "":
            self.frame_treeview_data("Warehouse", self.tv3)
        else:
            for row in self.ws_warehouse.values:
                for value in row:
                    if check.lower() in str(value).lower():
                        list.append(row)
                        break

            self.clean_data(self.tv3)

            for row in list:
                self.tv3.insert("", "end", values=row)

    def search_staff(self, Event):
        check = self.staff_entry_searchBar.get()
        list = []
        if check == "":
            self.frame_treeview_data("StaffList", self.tv4)
        else:
            for row in self.ws_staffList.values:
                for value in row:
                    if check.lower() in str(value).lower():
                        list.append(row)
                        break

            self.clean_data(self.tv4)

            for row in list:
                self.tv4.insert("", "end", values=row)

    def search_customer(self, Event):
        check = self.customer_entry_searchBar.get()
        list = []
        if check == "":
            self.frame_treeview_data("CustomerList", self.tv5)
        else:
            for row in self.ws_CustomerList.values:
                for value in row:
                    if check.lower() in str(value).lower():
                        list.append(row)
                        break

            self.clean_data(self.tv5)

            for row in list:
                self.tv5.insert("", "end", values=row)

    def search_bill(self, Event):
        check = self.bill_entry_searchBar.get()
        list = []
        if check == "":
            self.frame_treeview_data("Bill", self.tv6)
        else:
            for row in self.ws_cashBook.values:
                for value in row:
                    if check.lower() in str(value).lower():
                        list.append(row)
                        break

            self.clean_data(self.tv6)

            for row in list:
                self.tv6.insert("", "end", values=row)

    #frame
    def frame_overview(self):
        #Pie chart showing
        self.image_image_pieChart = PhotoImage(
            file=relative_to_assets("PieChart.png"))
        self.image_pieChart = self.canvas.create_image(
            221.0,
            309.0,
            image=self.image_image_pieChart
        )

        self.pieChartBox = self.canvas.create_rectangle(
            442.0,
            144.0,
            1195.0,
            636.0,
            fill="#FFFFFF",
            outline="")

        # buttun_options
        ##show_data_warehouse
        self.button_image_8 = PhotoImage(
            file=relative_to_assets("button_8.png"))
        self.button_8 = Button(
            image=self.button_image_8,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: self.frame_treeview_data("Warehouse",self.tv1),
            relief="flat"
        )
        self.button_8.place(
            x=939.0,
            y=678.0,
            width=146.0,
            height=40.0
        )

        ##show_data_bill
        self.button_image_9 = PhotoImage(
            file=relative_to_assets("button_9.png"))
        self.button_9 = Button(
            image=self.button_image_9,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: self.frame_treeview_data("Bill",self.tv1),
            relief="flat"
        )
        self.button_9.place(
            x=1105.0,
            y=678.0,
            width=90.0,
            height=40.0
        )

        ##show_data_customers
        self.button_image_10 = PhotoImage(
            file=relative_to_assets("button_10.png"))
        self.button_10 = Button(
            image=self.button_image_10,
            borderwidth=0,
            highlightthickness=0,
            command=lambda : self.frame_treeview_data("CustomerList",self.tv1),
            relief="flat"
        )
        self.button_10.place(
            x=774.0,
            y=678.0,
            width=145.0,
            height=40.0
        )

        ##show_data_menu
        self.button_image_11 = PhotoImage(
            file=relative_to_assets("button_11.png"))
        self.button_11 = Button(
            image=self.button_image_11,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: self.frame_treeview_data("Menu",self.tv1),
            relief="flat"
        )
        self.button_11.place(
            x=608.0,
            y=678.0,
            width=145.0,
            height=40.0
        )

        ##show_data_staffs
        self.button_image_12 = PhotoImage(
            file=relative_to_assets("button_12.png"))
        self.button_12 = Button(
            image=self.button_image_12,
            borderwidth=0,
            highlightthickness=0,
            command=lambda : self.frame_treeview_data("StaffList",self.tv1),
            relief="flat"
        )
        self.button_12.place(
            x=442.0,
            y=678.0,
            width=146.0,
            height=40.0
        )

        # frame_showData
        self.frame_showData = LabelFrame(
            self.window,
            background="white",
        )
        self.frame_showData.place(
            x=442,
            y=137,
            width=753,
            height=449
        )

        ## Treeview Widget revenue
        self.tv1 = ttk.Treeview(self.frame_showData)
        self.tv1.place(relheight=1,
                       relwidth=1)

        self.treescrolly = Scrollbar(
            self.frame_showData, orient="vertical",
            command=self.tv1.yview)
        self.treescrollx = Scrollbar(
            self.frame_showData, orient="horizontal",
            command=self.tv1.xview)
        self.tv1.configure(xscrollcommand=self.treescrollx.set,
                           yscrollcommand=self.treescrolly.set)
        self.treescrollx.pack(side="bottom", fill="x")
        self.treescrolly.pack(side="right", fill="y")

        #show data in treeview
        self.frame_treeview_data("StaffList",self.tv1)

        ##Statistical_data
        ### frame_show_statistical_data
        self.frame_statistical_data = LabelFrame(
            self.window,
            background="white",
        )
        self.frame_statistical_data.place(
            x=55,
            y=535,
            width=312,
            height=182
        )

        self.tv2 = ttk.Treeview(self.frame_statistical_data)
        self.tv2.place(relheight=1,
                       relwidth=1)

        self.treescrolly2 = Scrollbar(
            self.frame_statistical_data, orient="vertical",
            command=self.tv2.yview)
        self.treescrollx2 = Scrollbar(
            self.frame_statistical_data, orient="horizontal",
            command=self.tv2.xview)
        self.tv2.configure(xscrollcommand=self.treescrollx2.set,yscrollcommand=self.treescrolly2.set)
        self.treescrollx2.pack(side="bottom", fill="x")
        self.treescrolly2.pack(side="right", fill="y")

        self.clean_data(self.tv2)
        self.df2 = pd.read_excel("Database.xlsx", sheet_name="Statistical")
        self.tv2["column"] = list(self.df2.columns)
        self.tv2.column("#0", width=120, minwidth=20)
        self.tv2.column("Revenue", width=70, minwidth=60)
        self.tv2.column("Cost", width=70, minwidth=60)
        self.tv2.column("Profit", width=70, minwidth=60)
        self.tv2["show"] = "headings"
        for column in self.tv2["columns"]:
            self.tv2.heading(column, text=column)

        self.df2_rows = self.df2.to_numpy().tolist()
        for row in self.df2_rows:
            self.tv2.insert("", "end", values=row)

    def frame_warehouse(self):
        self.ws_warehouse = wb["Warehouse"]

        # warehouse_showData
        self.warehouse_frame_showData = LabelFrame(
            self.window,
            background="white",
        )
        self.warehouse_frame_showData.place(
            x=55,
            y=246,
            width=1140,
            height=472
        )

        ## Treeview Widget
        self.tv3 = ttk.Treeview(self.warehouse_frame_showData)
        self.tv3.place(relheight=1,
                       relwidth=1)

        self.treescrolly3 = Scrollbar(
            self.warehouse_frame_showData, orient="vertical",
            command=self.tv3.yview)
        self.treescrollx3 = Scrollbar(
            self.warehouse_frame_showData, orient="horizontal",
            command=self.tv3.xview)
        self.tv3.configure(xscrollcommand=self.treescrollx3.set,
                           yscrollcommand=self.treescrolly3.set)
        self.treescrollx3.pack(side="bottom", fill="x")
        self.treescrolly3.pack(side="right", fill="y")

        # show data in treeview
        self.frame_treeview_data("Warehouse",self.tv3)

        #entry_searchBar
        self.entry_image_searchBar = PhotoImage(
            file=relative_to_assets("searchBar.png"))
        self.entry_bg_searchBar = self.canvas.create_image(
            359.0,
            195.0,
            image=self.entry_image_searchBar
        )
        self.entry_searchBar = Entry(
            bd=0,
            bg="#C4C4C4",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_searchBar.place(
            x=130.0,
            y=164.0,
            width=608.0,
            height=60.0
        )
        self.entry_searchBar.bind("<KeyRelease>",self.search_material)

        #button
        ##button_insert
        self.button_image_insert = PhotoImage(
            file=relative_to_assets("button_insert.png"))
        self.button_insert = Button(
            image=self.button_image_insert,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: InsertMaterial(),
            relief="flat"
        )
        self.button_insert.place(
            x=1134.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_save
        self.button_image_save = PhotoImage(
            file=relative_to_assets("button_save.png"))
        self.button_save = Button(
            image=self.button_image_save,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: messagebox.showinfo("Save","File was saved"),
            relief="flat"
        )
        self.button_save.place(
            x=1043.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_delete
        self.button_image_delete = PhotoImage(
            file=relative_to_assets("button_delete.png"))
        self.button_delete = Button(
            image=self.button_image_delete,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: DeleteMaterial(),
            relief="flat"
        )
        self.button_delete.place(
            x=952.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_edit
        self.button_image_edit = PhotoImage(
            file=relative_to_assets("button_edit.png"))
        self.button_edit = Button(
            image=self.button_image_edit,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: EditMaterial(),
            relief="flat"
        )
        self.button_edit.place(
            x=861.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_search
        self.button_image_search = PhotoImage(
            file=relative_to_assets("button_search.png"))
        self.button_search = Button(
            image=self.button_image_search,
            borderwidth=0,
            highlightthickness=0,
            command=lambda : print("Đã search trong entry_searchBar"),
            relief="flat"
        )
        self.button_search.place(
            x=663.0,
            y=164.0,
            width=90.0,
            height=62.0
        )

    def frame_staffs(self):
        self.ws_staffList = wb["StaffList"]

        # entry_searchBar
        self.staff_entry_image_searchBar = PhotoImage(
            file=relative_to_assets("searchBar.png"))
        self.staff_entry_bg_searchBar = self.canvas.create_image(
            359.0,
            195.0,
            image=self.staff_entry_image_searchBar
        )
        self.staff_entry_searchBar = Entry(
            bd=0,
            bg="#C4C4C4",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.staff_entry_searchBar.place(
            x=130.0,
            y=164.0,
            width=608.0,
            height=60.0
        )
        self.staff_entry_searchBar.bind("<KeyRelease>", self.search_staff)

        # button
        ##button_insert
        self.staff_button_image_insert = PhotoImage(
            file=relative_to_assets("button_insert.png"))
        self.staff_button_insert = Button(
            image=self.staff_button_image_insert,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: InsertStaff(),
            relief="flat"
        )
        self.staff_button_insert.place(
            x=1134.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_save
        self.staff_button_image_save = PhotoImage(
            file=relative_to_assets("button_save.png"))
        self.staff_button_save = Button(
            image=self.staff_button_image_save,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: messagebox.showinfo("Save", "File was saved"),
            relief="flat"
        )
        self.staff_button_save.place(
            x=1043.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_delete
        self.staff_button_image_delete = PhotoImage(
            file=relative_to_assets("button_delete.png"))
        self.staff_button_delete = Button(
            image=self.staff_button_image_delete,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: DeleteStaff(),
            relief="flat"
        )
        self.staff_button_delete.place(
            x=952.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_edit
        self.staff_button_image_edit = PhotoImage(
            file=relative_to_assets("button_edit.png"))
        self.staff_button_edit = Button(
            image=self.staff_button_image_edit,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: EditStaff(),
            relief="flat"
        )
        self.staff_button_edit.place(
            x=861.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_search
        self.staff_button_image_search = PhotoImage(
            file=relative_to_assets("button_search.png"))
        self.staff_button_search = Button(
            image=self.staff_button_image_search,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: print("Đã search trong searchBar"),
            relief="flat"
        )
        self.staff_button_search.place(
            x=663.0,
            y=164.0,
            width=90.0,
            height=62.0
        )

        # staff_showData
        self.staff_frame_showData = LabelFrame(
            self.window,
            background="white",
        )
        self.staff_frame_showData.place(
            x=55,
            y=246,
            width=1140,
            height=472
        )

        ## Treeview Widget
        self.tv4 = ttk.Treeview(self.staff_frame_showData)
        self.tv4.place(relheight=1,
                       relwidth=1)

        self.treescrolly4 = Scrollbar(
            self.staff_frame_showData, orient="vertical",
            command=self.tv4.yview)
        self.treescrollx4 = Scrollbar(
            self.staff_frame_showData, orient="horizontal",
            command=self.tv4.xview)
        self.tv4.configure(xscrollcommand=self.treescrollx4.set,
                           yscrollcommand=self.treescrolly4.set)
        self.treescrollx4.pack(side="bottom", fill="x")
        self.treescrolly4.pack(side="right", fill="y")

        # show data in treeview
        self.frame_treeview_data("StaffList", self.tv4)

    def frame_cashBook(self):
        self.ws_cashBook = wb["Bill"]

        # entry_searchBar
        self.bill_entry_image_searchBar = PhotoImage(
            file=relative_to_assets("searchBar.png"))
        self.bill_entry_bg_searchBar = self.canvas.create_image(
            359.0,
            195.0,
            image=self.bill_entry_image_searchBar
        )
        self.bill_entry_searchBar = Entry(
            bd=0,
            bg="#C4C4C4",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.bill_entry_searchBar.place(
            x=130.0,
            y=164.0,
            width=608.0,
            height=60.0
        )
        self.bill_entry_searchBar.bind("<KeyRelease>", self.search_bill)

        # button
        ##button_insert
        self.bill_button_image_insert = PhotoImage(
            file=relative_to_assets("button_insert.png"))
        self.bill_button_insert = Button(
            image=self.bill_button_image_insert,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: print("Chưa hoàn thiện"),
            relief="flat"
        )
        self.bill_button_insert.place(
            x=1134.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_save
        self.bill_button_image_save = PhotoImage(
            file=relative_to_assets("button_save.png"))
        self.bill_button_save = Button(
            image=self.bill_button_image_save,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: messagebox.showinfo("Save", "File was saved"),
            relief="flat"
        )
        self.bill_button_save.place(
            x=1043.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_delete
        self.bill_button_image_delete = PhotoImage(
            file=relative_to_assets("button_delete.png"))
        self.bill_button_delete = Button(
            image=self.bill_button_image_delete,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: print("Chưa hoàn thiện"),
            relief="flat"
        )
        self.bill_button_delete.place(
            x=952.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_edit
        self.bill_button_image_edit = PhotoImage(
            file=relative_to_assets("button_edit.png"))
        self.bill_button_edit = Button(
            image=self.bill_button_image_edit,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: print("chưa hoàn thiện"),
            relief="flat"
        )
        self.bill_button_edit.place(
            x=861.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_search
        self.bill_button_image_search = PhotoImage(
            file=relative_to_assets("button_search.png"))
        self.bill_button_search = Button(
            image=self.bill_button_image_search,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: print("Đã search trong searchBar"),
            relief="flat"
        )
        self.bill_button_search.place(
            x=663.0,
            y=164.0,
            width=90.0,
            height=62.0
        )

        # staff_showData
        self.bill_frame_showData = LabelFrame(
            self.window,
            background="white",
        )
        self.bill_frame_showData.place(
            x=55,
            y=246,
            width=1140,
            height=472
        )

        ## Treeview Widget
        self.tv6 = ttk.Treeview(self.bill_frame_showData)
        self.tv6.place(relheight=1,
                       relwidth=1)

        self.treescrolly6 = Scrollbar(
            self.bill_frame_showData, orient="vertical",
            command=self.tv6.yview)
        self.treescrollx6 = Scrollbar(
            self.bill_frame_showData, orient="horizontal",
            command=self.tv6.xview)
        self.tv6.configure(xscrollcommand=self.treescrollx6.set,
                           yscrollcommand=self.treescrolly6.set)
        self.treescrollx6.pack(side="bottom", fill="x")
        self.treescrolly6.pack(side="right", fill="y")

        # show data in treeview
        self.frame_treeview_data("Bill", self.tv6)


    def frame_customers(self):
        self.ws_CustomerList = wb["CustomerList"]

        # entry_searchBar
        self.customer_entry_image_searchBar = PhotoImage(
            file=relative_to_assets("searchBar.png"))
        self.customer_entry_bg_searchBar = self.canvas.create_image(
            359.0,
            195.0,
            image=self.customer_entry_image_searchBar
        )
        self.customer_entry_searchBar = Entry(
            bd=0,
            bg="#C4C4C4",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.customer_entry_searchBar.place(
            x=130.0,
            y=164.0,
            width=608.0,
            height=60.0
        )
        self.customer_entry_searchBar.bind("<KeyRelease>", self.search_customer)

        # button
        ##button_insert
        self.customer_button_image_insert = PhotoImage(
            file=relative_to_assets("button_insert.png"))
        self.customer_button_insert = Button(
            image=self.customer_button_image_insert,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: InsertCustomer(),
            relief="flat"
        )
        self.customer_button_insert.place(
            x=1134.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_save
        self.customer_button_image_save = PhotoImage(
            file=relative_to_assets("button_save.png"))
        self.customer_button_save = Button(
            image=self.customer_button_image_save,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: messagebox.showinfo("Save", "File was saved"),
            relief="flat"
        )
        self.customer_button_save.place(
            x=1043.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_delete
        self.customer_button_image_delete = PhotoImage(
            file=relative_to_assets("button_delete.png"))
        self.customer_button_delete = Button(
            image=self.customer_button_image_delete,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: DeleteCustomer(),
            relief="flat"
        )
        self.customer_button_delete.place(
            x=952.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_edit
        self.customer_button_image_edit = PhotoImage(
            file=relative_to_assets("button_edit.png"))
        self.customer_button_edit = Button(
            image=self.customer_button_image_edit,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: EditCustomer(),
            relief="flat"
        )
        self.customer_button_edit.place(
            x=861.0,
            y=164.0,
            width=61.0,
            height=62.0
        )

        ##button_search
        self.customer_button_image_search = PhotoImage(
            file=relative_to_assets("button_search.png"))
        self.customer_button_search = Button(
            image=self.customer_button_image_search,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: print("Đã search trong searchBar"),
            relief="flat"
        )
        self.customer_button_search.place(
            x=663.0,
            y=164.0,
            width=90.0,
            height=62.0
        )

        # staff_showData
        self.customer_frame_showData = LabelFrame(
            self.window,
            background="white",
        )
        self.customer_frame_showData.place(
            x=55,
            y=246,
            width=1140,
            height=472
        )

        ## Treeview Widget
        self.tv5 = ttk.Treeview(self.customer_frame_showData)
        self.tv5.place(relheight=1,
                       relwidth=1)

        self.treescrolly5 = Scrollbar(
            self.customer_frame_showData, orient="vertical",
            command=self.tv5.yview)
        self.treescrollx5 = Scrollbar(
            self.customer_frame_showData, orient="horizontal",
            command=self.tv5.xview)
        self.tv5.configure(xscrollcommand=self.treescrollx5.set,
                           yscrollcommand=self.treescrolly5.set)
        self.treescrollx5.pack(side="bottom", fill="x")
        self.treescrolly5.pack(side="right", fill="y")

        # show data in treeview
        self.frame_treeview_data("CustomerList", self.tv5)

    # def frame_report(self):
    #     pass


    def clean_data(self,treeview):
        treeview.delete(*treeview.get_children())

    def clean_window(self):
        if self.status == "Overview":
            self.canvas.delete(self.image_pieChart)
            self.frame_showData.destroy()
            self.frame_statistical_data.destroy()
            self.button_8.destroy()
            self.button_9.destroy()
            self.button_10.destroy()
            self.button_11.destroy()
            self.button_12.destroy()

        if self.status == "Warehouse":
            self.canvas.delete(self.entry_bg_searchBar)
            self.entry_searchBar.destroy()
            self.warehouse_frame_showData.destroy()
            self.button_insert.destroy()
            self.button_save.destroy()
            self.button_edit.destroy()
            self.button_delete.destroy()
            self.button_search.destroy()

        if self.status == "Staffs":
            self.canvas.delete(self.staff_entry_bg_searchBar)
            self.staff_entry_searchBar.destroy()
            self.staff_frame_showData.destroy()
            self.staff_button_insert.destroy()
            self.staff_button_save.destroy()
            self.staff_button_edit.destroy()
            self.staff_button_delete.destroy()
            self.staff_button_search.destroy()

        if self.status == "Customers":
            self.canvas.delete(self.customer_entry_bg_searchBar)
            self.customer_entry_searchBar.destroy()
            self.customer_frame_showData.destroy()
            self.customer_button_insert.destroy()
            self.customer_button_save.destroy()
            self.customer_button_edit.destroy()
            self.customer_button_delete.destroy()
            self.customer_button_search.destroy()

        if self.status == "Cashbook":
            self.canvas.delete(self.bill_entry_bg_searchBar)
            self.bill_entry_searchBar.destroy()
            self.bill_frame_showData.destroy()
            self.bill_button_insert.destroy()
            self.bill_button_save.destroy()
            self.bill_button_edit.destroy()
            self.bill_button_delete.destroy()
            self.bill_button_search.destroy()


    #change color navbar button
    def change_color_button_overview(self):
        # Change color of navbar
        ## Report button
        self.clean_window()
        self.button_report.destroy()
        self.button_report_image = PhotoImage(
            file=relative_to_assets("button_report.png"))
        self.button_report = Button(
            image=self.button_report_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonReport,
            relief="flat"
        )
        self.button_report.place(
            x=883.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        ## button_staffs
        self.button_staffs.destroy()
        self.button_staffs_image = PhotoImage(
            file=relative_to_assets("button_staffs.png"))
        self.button_staffs = Button(
            image=self.button_staffs_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonStaffs,
            relief="flat"
        )
        self.button_staffs.place(
            x=387.0,
            y=0.0,
            width=145.0,
            height=65.0
        )

        ##button_cashBook
        self.button_cashBook.destroy()
        self.button_cashBook_image = PhotoImage(
            file=relative_to_assets("button_cashBook.png"))
        self.button_cashBook = Button(
            image=self.button_cashBook_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonCashBook,
            relief="flat"
        )
        self.button_cashBook.place(
            x=552.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        ##button_customers
        self.button_customers.destroy()
        self.button_customers_image = PhotoImage(
            file=relative_to_assets("button_customers.png"))
        self.button_customers = Button(
            image=self.button_customers_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonCustomers,
            relief="flat"
        )
        self.button_customers.place(
            x=718.0,
            y=0.0,
            width=145.0,
            height=65.0
        )

        ##button_warehouse
        self.button_warehouse.destroy()
        self.button_warehouse_image = PhotoImage(
            file=relative_to_assets("button_warehouse.png"))
        self.button_warehouse = Button(
            image=self.button_warehouse_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonWarehouse,
            relief="flat"
        )
        self.button_warehouse.place(
            x=221.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        ## button_overview
        self.button_overview.destroy()
        self.button_overview_image = PhotoImage(
            file=relative_to_assets("button_overview_show.png"))
        self.button_overview = Button(
            image=self.button_overview_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonOverview,
            relief="flat"
        )
        self.button_overview.place(
            x=55.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

    def change_color_button_warehouse(self):
        self.clean_window()
        # Report button
        self.button_report.destroy()
        self.button_report_image = PhotoImage(
            file=relative_to_assets("button_report.png"))
        self.button_report = Button(
            image=self.button_report_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonReport,
            relief="flat"
        )
        self.button_report.place(
            x=883.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_overview
        self.button_overview.destroy()
        self.button_overview_image = PhotoImage(
            file=relative_to_assets("button_overview.png"))
        self.button_overview = Button(
            image=self.button_overview_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonOverview,
            relief="flat"
        )
        self.button_overview.place(
            x=55.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_staffs
        self.button_staffs.destroy()
        self.button_staffs_image = PhotoImage(
            file=relative_to_assets("button_staffs.png"))
        self.button_staffs = Button(
            image=self.button_staffs_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonStaffs,
            relief="flat"
        )
        self.button_staffs.place(
            x=387.0,
            y=0.0,
            width=145.0,
            height=65.0
        )

        # button_cashBook
        self.button_cashBook.destroy()
        self.button_cashBook_image = PhotoImage(
            file=relative_to_assets("button_cashBook.png"))
        self.button_cashBook = Button(
            image=self.button_cashBook_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonCashBook,
            relief="flat"
        )
        self.button_cashBook.place(
            x=552.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_customers
        self.button_customers.destroy()
        self.button_customers_image = PhotoImage(
            file=relative_to_assets("button_customers.png"))
        self.button_customers = Button(
            image=self.button_customers_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonCustomers,
            relief="flat"
        )
        self.button_customers.place(
            x=718.0,
            y=0.0,
            width=145.0,
            height=65.0
        )

        # button_warehouse
        self.button_warehouse.destroy()
        self.button_warehouse_image = PhotoImage(
            file=relative_to_assets("button_warehouse_show.png"))
        self.button_warehouse = Button(
            image=self.button_warehouse_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonWarehouse,
            relief="flat"
        )
        self.button_warehouse.place(
            x=221.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

    def change_color_button_staff(self):

        self.clean_window()

        # Change color of navbar
        # Report button
        self.button_report.destroy()
        self.button_report_image = PhotoImage(
            file=relative_to_assets("button_report.png"))
        self.button_report = Button(
            image=self.button_report_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonReport,
            relief="flat"
        )
        self.button_report.place(
            x=883.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_cashBook
        self.button_cashBook.destroy()
        self.button_cashBook_image = PhotoImage(
            file=relative_to_assets("button_cashBook.png"))
        self.button_cashBook = Button(
            image=self.button_cashBook_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonCashBook,
            relief="flat"
        )
        self.button_cashBook.place(
            x=552.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_customers
        self.button_customers.destroy()
        self.button_customers_image = PhotoImage(
            file=relative_to_assets("button_customers.png"))
        self.button_customers = Button(
            image=self.button_customers_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonCustomers,
            relief="flat"
        )
        self.button_customers.place(
            x=718.0,
            y=0.0,
            width=145.0,
            height=65.0
        )

        # button_warehouse
        self.button_warehouse.destroy()
        self.button_warehouse_image = PhotoImage(
            file=relative_to_assets("button_warehouse.png"))
        self.button_warehouse = Button(
            image=self.button_warehouse_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonWarehouse,
            relief="flat"
        )
        self.button_warehouse.place(
            x=221.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_overview
        self.button_overview.destroy()
        self.button_overview_image = PhotoImage(
            file=relative_to_assets("button_overview.png"))
        self.button_overview = Button(
            image=self.button_overview_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonOverview,
            relief="flat"
        )
        self.button_overview.place(
            x=55.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_staffs
        self.button_staffs.destroy()
        self.button_staffs_image = PhotoImage(
            file=relative_to_assets("button_staffs_show.png"))
        self.button_staffs = Button(
            image=self.button_staffs_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonStaffs,
            relief="flat"
        )
        self.button_staffs.place(
            x=387.0,
            y=0.0,
            width=145.0,
            height=65.0
        )

    def change_color_button_cashbook(self):

        self.clean_window()

        # Change color of navbar
        # Report button
        self.button_report.destroy()
        self.button_report_image = PhotoImage(
            file=relative_to_assets("button_report.png"))
        self.button_report = Button(
            image=self.button_report_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonReport,
            relief="flat"
        )
        self.button_report.place(
            x=883.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_staffs
        self.button_staffs.destroy()
        self.button_staffs_image = PhotoImage(
            file=relative_to_assets("button_staffs.png"))
        self.button_staffs = Button(
            image=self.button_staffs_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonStaffs,
            relief="flat"
        )
        self.button_staffs.place(
            x=387.0,
            y=0.0,
            width=145.0,
            height=65.0
        )

        # button_customers
        self.button_customers.destroy()
        self.button_customers_image = PhotoImage(
            file=relative_to_assets("button_customers.png"))
        self.button_customers = Button(
            image=self.button_customers_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonCustomers,
            relief="flat"
        )
        self.button_customers.place(
            x=718.0,
            y=0.0,
            width=145.0,
            height=65.0
        )

        # button_warehouse
        self.button_warehouse.destroy()
        self.button_warehouse_image = PhotoImage(
            file=relative_to_assets("button_warehouse.png"))
        self.button_warehouse = Button(
            image=self.button_warehouse_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonWarehouse,
            relief="flat"
        )
        self.button_warehouse.place(
            x=221.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_overview
        self.button_overview.destroy()
        self.button_overview_image = PhotoImage(
            file=relative_to_assets("button_overview.png"))
        self.button_overview = Button(
            image=self.button_overview_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonOverview,
            relief="flat"
        )
        self.button_overview.place(
            x=55.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_cashBook
        self.button_cashBook.destroy()
        self.button_cashBook_image = PhotoImage(
            file=relative_to_assets("button_cashBook_show.png"))
        self.button_cashBook = Button(
            image=self.button_cashBook_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonCashBook,
            relief="flat"
        )
        self.button_cashBook.place(
            x=552.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

    def change_color_button_customers(self):
        self.clean_window()

        # Change color of navbar
        # Report button
        self.button_report.destroy()
        self.button_report_image = PhotoImage(
            file=relative_to_assets("button_report.png"))
        self.button_report = Button(
            image=self.button_report_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonReport,
            relief="flat"
        )
        self.button_report.place(
            x=883.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_staffs
        self.button_staffs.destroy()
        self.button_staffs_image = PhotoImage(
            file=relative_to_assets("button_staffs.png"))
        self.button_staffs = Button(
            image=self.button_staffs_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonStaffs,
            relief="flat"
        )
        self.button_staffs.place(
            x=387.0,
            y=0.0,
            width=145.0,
            height=65.0
        )

        # button_warehouse
        self.button_warehouse.destroy()
        self.button_warehouse_image = PhotoImage(
            file=relative_to_assets("button_warehouse.png"))
        self.button_warehouse = Button(
            image=self.button_warehouse_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonWarehouse,
            relief="flat"
        )
        self.button_warehouse.place(
            x=221.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_overview
        self.button_overview.destroy()
        self.button_overview_image = PhotoImage(
            file=relative_to_assets("button_overview.png"))
        self.button_overview = Button(
            image=self.button_overview_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonOverview,
            relief="flat"
        )
        self.button_overview.place(
            x=55.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_cashBook
        self.button_cashBook.destroy()
        self.button_cashBook_image = PhotoImage(
            file=relative_to_assets("button_cashBook.png"))
        self.button_cashBook = Button(
            image=self.button_cashBook_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonCashBook,
            relief="flat"
        )
        self.button_cashBook.place(
            x=552.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_customers
        self.button_customers.destroy()
        self.button_customers_image = PhotoImage(
            file=relative_to_assets("button_customers_show.png"))
        self.button_customers = Button(
            image=self.button_customers_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonCustomers,
            relief="flat"
        )
        self.button_customers.place(
            x=718.0,
            y=0.0,
            width=145.0,
            height=65.0
        )

    def change_color_button_report(self):
        self.clean_window()
        # button_staffs
        self.button_staffs.destroy()
        self.button_staffs_image = PhotoImage(
            file=relative_to_assets("button_staffs.png"))
        self.button_staffs = Button(
            image=self.button_staffs_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonStaffs,
            relief="flat"
        )
        self.button_staffs.place(
            x=387.0,
            y=0.0,
            width=145.0,
            height=65.0
        )

        # button_warehouse
        self.button_warehouse.destroy()
        self.button_warehouse_image = PhotoImage(
            file=relative_to_assets("button_warehouse.png"))
        self.button_warehouse = Button(
            image=self.button_warehouse_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonWarehouse,
            relief="flat"
        )
        self.button_warehouse.place(
            x=221.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_overview
        self.button_overview.destroy()
        self.button_overview_image = PhotoImage(
            file=relative_to_assets("button_overview.png"))
        self.button_overview = Button(
            image=self.button_overview_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonOverview,
            relief="flat"
        )
        self.button_overview.place(
            x=55.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_cashBook
        self.button_cashBook.destroy()
        self.button_cashBook_image = PhotoImage(
            file=relative_to_assets("button_cashBook.png"))
        self.button_cashBook = Button(
            image=self.button_cashBook_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonCashBook,
            relief="flat"
        )
        self.button_cashBook.place(
            x=552.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

        # button_customers
        self.button_customers.destroy()
        self.button_customers_image = PhotoImage(
            file=relative_to_assets("button_customers.png"))
        self.button_customers = Button(
            image=self.button_customers_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonCustomers,
            relief="flat"
        )
        self.button_customers.place(
            x=718.0,
            y=0.0,
            width=145.0,
            height=65.0
        )

        # Report button
        self.button_report.destroy()
        self.button_report_image = PhotoImage(
            file=relative_to_assets("button_report_show.png"))
        self.button_report = Button(
            image=self.button_report_image,
            borderwidth=0,
            highlightthickness=0,
            command=self.buttonReport,
            relief="flat"
        )
        self.button_report.place(
            x=883.0,
            y=0.0,
            width=146.0,
            height=65.0
        )

    #file
    def frame_treeview_data(self, sheet_name,treeview):
        self.clean_data(treeview)
        self.df = pd.read_excel("Database.xlsx", sheet_name=sheet_name)
        treeview["column"] = list(self.df.columns)
        treeview["show"] = "headings"
        for column in treeview["columns"]:
            treeview.heading(column, text=column)

        self.df_rows = self.df.to_numpy().tolist()
        for row in self.df_rows:
            treeview.insert("", "end", values=row)

    #run_button
    def buttonOverview(self):
        self.change_color_button_overview()
        #frame_overview
        self.frame_overview()

        self.status = "Overview"

    def buttonWarehouse(self):
        self.change_color_button_warehouse()
        self.frame_warehouse()

        self.status = "Warehouse"

    def buttonStaffs(self):
        self.change_color_button_staff()
        self.frame_staffs()
        self.status = "Staffs"

    def buttonCashBook(self):
        self.change_color_button_cashbook()
        self.frame_cashBook()
        self.status = "Cashbook"

    def buttonCustomers(self):
        self.change_color_button_customers()
        self.frame_customers()
        self.status = "Customers"

    def buttonReport(self):

        self.change_color_button_report()
        self.status = "Report"


    def buttonLogout(self):
        self.canvas.destroy()
        GUI_Login(self.window)

#Warehouse
class InsertMaterial():
    def __init__(self):
        self.window_insert = Toplevel()
        self.window_insert.title("Insert Material in Warehouse")
        self.window_insert.geometry("502x461")
        self.window_insert.configure(bg="#FFFFFF")

        self.ws_Material = pd.read_excel("Database.xlsx", sheet_name="Warehouse")

        self.canvas_insert = Canvas(
            self.window_insert,
            bg="#FFFFFF",
            height=461,
            width=502,
            bd=0,
            highlightthickness=0,
            relief="ridge"
        )

        self.canvas_insert.place(x=0, y=0)

        #button insert
        self.button_image_insertMaterial = PhotoImage(
            file=relative_to_assets("button_insertData.png"))
        self.button_insertMaterial = Button(
            self.window_insert,
            image=self.button_image_insertMaterial,
            borderwidth=0,
            highlightthickness=0,
            command=self.button_Insert,
            relief="flat"
        )
        self.button_insertMaterial.place(
            x=203.99999999999997,
            y=381.0,
            width=94.0,
            height=40.0
        )

        #entry
        self.entry_image_shortData = PhotoImage(
            file=relative_to_assets("entry_insertShortData.png"))

        self.type = ["Trà", "Bột", "Cafe", "Syrup", "Đường", "Sữa"]

        self.entry_insertType = ttk.Combobox(
            self.window_insert,
            value=self.type,
            font=("Helvetica", 10)
        )
        self.entry_insertType.place(
            x=298,
            y=60,
            width=129,
            height=40
        )
        self.entry_insertType.current(0)
        self.entry_insertType.bind("<<ComboboxSelected>>",self.autoFillIdAndUnit)

        self.entry_bg_id = self.canvas_insert.create_image(
            362.5,
            140.0,
            image=self.entry_image_shortData
        )
        self.entry_insertId = Entry(
            self.window_insert,
            bd=0,
            bg="#DFDFDF",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_insertId.place(
            x=348.0,
            y=120.0,
            width=69.0,
            height=38.0
        )

        self.entry_image_insertLongData = PhotoImage(
            file=relative_to_assets("entry_insertLongData.png"))

        self.entry_bg_insertName = self.canvas_insert.create_image(
            325.5,
            200.5,
            image=self.entry_image_insertLongData
        )
        self.entry_insertName = Entry(
            self.window_insert,
            bd=0,
            bg="#DFDFDF",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_insertName.place(
            x=233.99999999999997,
            y=180.0,
            width=183.0,
            height=39.0
        )

        self.entry_bg_quantity = self.canvas_insert.create_image(
            362.5,
            261.0,
            image=self.entry_image_shortData
        )
        self.entry_insertQuantity = Entry(
            self.window_insert,
            bd=0,
            bg="#DFDFDF",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_insertQuantity.place(
            x=308.0,
            y=241.0,
            width=109.0,
            height=38.0
        )

        self.entry_bg_unit = self.canvas_insert.create_image(
            362.5,
            321.0,
            image=self.entry_image_shortData
        )
        self.unit = ["hộp", "bịch", "chai", "Kg"]

        self.entry_insertUnit = ttk.Combobox(
            self.window_insert,
            value=self.unit,
            font=("Helvetica", 10)
        )
        self.entry_insertUnit.place(
            x=298.0,
            y=301.0,
            width=129,
            height=40
        )
        self.entry_insertUnit.current(0)
        self.entry_insertUnit.bind("<<ComboboxSelected>>")



        self.image_image_type = PhotoImage(
            file=relative_to_assets("image_type.png"))
        self.image_1 = self.canvas_insert.create_image(
            94.99999999999997,
            80.0,
            image=self.image_image_type
        )

        self.image_image_id = PhotoImage(
            file=relative_to_assets("image_id.png"))
        self.image_2 = self.canvas_insert.create_image(
            83.99999999999997,
            141.0,
            image=self.image_image_id
        )

        self.image_image_name = PhotoImage(
            file=relative_to_assets("image_name.png"))
        self.image_3 = self.canvas_insert.create_image(
            94.99999999999997,
            201.0,
            image=self.image_image_name
        )

        self.image_image_quantity = PhotoImage(
            file=relative_to_assets("image_quantity.png"))
        self.image_4 = self.canvas_insert.create_image(
            105.99999999999997,
            262.0,
            image=self.image_image_quantity
        )

        self.image_image_unit = PhotoImage(
            file=relative_to_assets("image_unit.png"))
        self.image_5 = self.canvas_insert.create_image(
            91.99999999999997,
            323.0,
            image=self.image_image_unit
        )
        self.window_insert.resizable(False, False)
        self.window_insert.mainloop()


    #auto fill
    def autoFillIdAndUnit(self, Event):
        self.searchList_material = self.ws_Material.to_numpy().tolist()
        self.searchList_material.reverse()

        if self.entry_insertType.get() == "Trà":
            for row in self.searchList_material:
                if "T" in row[0]:
                    num = int(row[0][1:len(row)]) + 1
                    newId = "T"+str(num)
                    self.entry_insertId.delete(0, END)
                    self.entry_insertId.insert(0, newId)
                    self.entry_insertUnit.delete(0, END)
                    self.entry_insertUnit.insert(0, "hộp")
                    break

        elif self.entry_insertType.get() == "Bột":
            for row in self.searchList_material:
                if "F" in row[0]:
                    num = int(row[0][1:len(row)]) + 1
                    newId = "F" + str(num)
                    self.entry_insertId.delete(0, END)
                    self.entry_insertId.insert(0, newId)
                    self.entry_insertUnit.delete(0, END)
                    self.entry_insertUnit.insert(0, "bịch")
                    break

        elif self.entry_insertType.get() == "Cafe":
            for row in self.searchList_material:
                if "C" in row[0]:
                    num = int(row[0][1:len(row)]) + 1
                    newId = "C" + str(num)
                    self.entry_insertId.delete(0, END)
                    self.entry_insertId.insert(0, newId)
                    self.entry_insertUnit.delete(0, END)
                    self.entry_insertUnit.insert(0, "bịch")
                    break

        elif self.entry_insertType.get() == "Syrup":
            for row in self.searchList_material:  # đọc giữ liệu theo chiều ngược lại
                if "SY" in row[0]:
                    num = int(row[0][2:len(row)]) + 1
                    newId = "SY" + str(num)
                    self.entry_insertId.delete(0, END)
                    self.entry_insertId.insert(0, newId)
                    self.entry_insertUnit.delete(0, END)
                    self.entry_insertUnit.insert(0, "chai")
                    break

        elif self.entry_insertType.get() == "Sữa":
            for row in self.searchList_material:  # đọc giữ liệu theo chiều ngược lại
                if "M" in row[0]:
                    num = int(row[0][1:len(row)]) + 1
                    newId = "M" + str(num)
                    self.entry_insertId.delete(0, END)
                    self.entry_insertId.insert(0, newId)
                    self.entry_insertUnit.delete(0, END)
                    self.entry_insertUnit.insert(0, "hộp")
                    break

        else:
            for row in self.searchList_material:  # đọc giữ liệu theo chiều ngược lại. item Đường
                if "SU" in row[0]:
                    num = int(row[0][2:len(row)]) + 1
                    newId = "SU" + str(num)
                    self.entry_insertId.delete(0, END)
                    self.entry_insertId.insert(0, newId)
                    self.entry_insertUnit.delete(0, END)
                    self.entry_insertUnit.insert(0, "Kg")
                    break


    #run button
    def button_Insert(self):
        check = True
        if self.entry_insertName.get() == "":
            messagebox.showerror("Do not have Material name", "Please enter material name")
            check = False

        if self.entry_insertQuantity.get() == "":
            messagebox.showerror("Do not have Quantity", "Please enter quantity")
            check = False

        if check == True:
            for row in self.ws_Material.values:
                if self.entry_insertName.get().lower() == row[1].lower():
                    self.entry_insertName.delete(0, END)
                    messagebox.showwarning("Exist Data", "Material name already existed")
                    check = False
                    break

            try:
                self.quantity = 0
                if self.entry_insertUnit.get() == "Kg":
                    self.quantity = float(self.entry_insertQuantity.get())
                else:
                    self.quantity = int(self.entry_insertQuantity.get())

                if self.quantity <= 0:
                    self.entry_insertQuantity.delete(0, END)
                    messagebox.showerror("Incorrect quantity", "Please enter number greater than 0")
                    check = False

            except:
                    self.entry_insertQuantity.delete(0, END)
                    messagebox.showerror("Incorrect type of quantity", "Please enter Interger of Float number")
                    check = False


        if check == True:
            self.ws_Material_openpyxl = wb["Warehouse"]
            self.ws_Material_openpyxl.append([self.entry_insertId.get(),
                                              self.entry_insertName.get().capitalize(),
                                              self.quantity,
                                              self.entry_insertUnit.get()
                                              ])

            wb.save("Database.xlsx")
            self.closeInsertWindow()

    def closeInsertWindow(self):
        self.canvas_insert.destroy()
        self.window_insert.destroy()

class DeleteMaterial():
    def __init__(self):
        self.window_delete = Toplevel()
        self.window_delete.title("Delete Material in Warehouse")
        self.window_delete.geometry("502x461")
        self.window_delete.configure(bg="#FFFFFF")

        self.ws_Material = pd.read_excel("Database.xlsx", sheet_name="Warehouse")

        self.canvas_delete = Canvas(
            self.window_delete,
            bg="#FFFFFF",
            height=461,
            width=502,
            bd=0,
            highlightthickness=0,
            relief="ridge"
        )

        self.canvas_delete.place(x=0, y=0)

        # button delete
        self.button_image_deleteMaterial = PhotoImage(
            file=relative_to_assets("button_deleteData.png"))
        self.button_deleteMaterial = Button(
            self.window_delete,
            image=self.button_image_deleteMaterial,
            borderwidth=0,
            highlightthickness=0,
            command=self.button_Delete,
            relief="flat"
        )
        self.button_deleteMaterial.place(
            x=203.99999999999997,
            y=381.0,
            width=94.0,
            height=40.0
        )

        # entry
        self.entry_image_shortData = PhotoImage(
            file=relative_to_assets("entry_insertShortData.png"))

        self.type = ["Trà", "Bột", "Cafe", "Syrup", "Đường", "Sữa"]

        self.entry_deleteType = ttk.Combobox(
            self.window_delete,
            value=self.type,
            font=("Helvetica", 10)
        )
        self.entry_deleteType.place(
            x=298,
            y=60,
            width=129,
            height=40
        )
        self.entry_deleteType.current(0)
        self.entry_deleteType.bind("<<ComboboxSelected>>", self.autoFillIdAndUnit)


        self.list_id = []
        self.entry_bg_id = self.canvas_delete.create_image(
            362.5,
            140.0,
            image=self.entry_image_shortData
        )
        self.entry_deleteId = ttk.Combobox(
            self.window_delete,
            font=("Helvetica", 12),
            value = self.list_id
        )
        self.entry_deleteId.place(
            x=298,
            y=120.0,
            width=129,
            height=40
        )


        self.entry_image_deleteLongData = PhotoImage(
            file=relative_to_assets("entry_insertLongData.png"))

        self.entry_bg_deleteName = self.canvas_delete.create_image(
            325.5,
            200.5,
            image=self.entry_image_deleteLongData
        )
        self.entry_deleteName = Entry(
            self.window_delete,
            bd=0,
            bg="#DFDFDF",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_deleteName.place(
            x=233.99999999999997,
            y=180.0,
            width=183.0,
            height=39.0
        )

        self.entry_bg_quantity = self.canvas_delete.create_image(
            362.5,
            261.0,
            image=self.entry_image_shortData
        )
        self.entry_deleteQuantity = Entry(
            self.window_delete,
            bd=0,
            bg="#DFDFDF",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_deleteQuantity.place(
            x=308.0,
            y=241.0,
            width=109.0,
            height=38.0
        )

        self.entry_bg_unit = self.canvas_delete.create_image(
            362.5,
            321.0,
            image=self.entry_image_shortData
        )
        self.entry_deleteUnit = Entry(
            self.window_delete,
            bd=0,
            bg="#DFDFDF",
            highlightthickness=0,
            font=("Helvetica", 10)
        )
        self.entry_deleteUnit.place(
            x=308.0,
            y=301.0,
            width=109.0,
            height=38.0
        )

        self.image_image_type = PhotoImage(
            file=relative_to_assets("image_type.png"))
        self.image_1 = self.canvas_delete.create_image(
            94.99999999999997,
            80.0,
            image=self.image_image_type
        )

        self.image_image_id = PhotoImage(
            file=relative_to_assets("image_id.png"))
        self.image_2 = self.canvas_delete.create_image(
            83.99999999999997,
            141.0,
            image=self.image_image_id
        )

        self.image_image_name = PhotoImage(
            file=relative_to_assets("image_name.png"))
        self.image_3 = self.canvas_delete.create_image(
            94.99999999999997,
            201.0,
            image=self.image_image_name
        )

        self.image_image_quantity = PhotoImage(
            file=relative_to_assets("image_quantity.png"))
        self.image_4 = self.canvas_delete.create_image(
            105.99999999999997,
            262.0,
            image=self.image_image_quantity
        )

        self.image_image_unit = PhotoImage(
            file=relative_to_assets("image_unit.png"))
        self.image_5 = self.canvas_delete.create_image(
            91.99999999999997,
            323.0,
            image=self.image_image_unit
        )
        self.window_delete.resizable(False, False)
        self.window_delete.mainloop()

    # auto fill
    def autoFillIdAndUnit(self, Event):
        self.searchList_material = self.ws_Material.to_numpy().tolist()

        if self.entry_deleteType.get() == "Trà":
            self.list_id = []
            for row in self.searchList_material:
                if "T" in row[0]:
                    self.list_id.append(row[0])
            self.entry_deleteId['values'] = self.list_id
            self.entry_deleteId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_deleteId.current(0)


        elif self.entry_deleteType.get() == "Bột":
            self.list_id = []
            for row in self.searchList_material:
                if "F" in row[0]:
                    self.list_id.append(row[0])
            self.entry_deleteId['values'] = self.list_id
            self.entry_deleteId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_deleteId.current(0)


        elif self.entry_deleteType.get() == "Cafe":
            self.list_id = []
            for row in self.searchList_material:
                if "C" in row[0]:
                    self.list_id.append(row[0])
            self.entry_deleteId['values'] = self.list_id
            self.entry_deleteId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_deleteId.current(0)


        elif self.entry_deleteType.get() == "Syrup":
            self.list_id = []
            for row in self.searchList_material:
                if "SY" in row[0]:
                    self.list_id.append(row[0])
            self.entry_deleteId['values'] = self.list_id
            self.entry_deleteId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_deleteId.current(0)


        elif self.entry_deleteType.get() == "Sữa":
            self.list_id = []
            for row in self.searchList_material:
                if "M" in row[0]:
                    self.list_id.append(row[0])
            self.entry_deleteId['values'] = self.list_id
            self.entry_deleteId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_deleteId.current(0)


        else:
            self.list_id = []
            for row in self.searchList_material:
                if "SU" in row[0]:
                    self.list_id.append(row[0])
            self.entry_deleteId['values'] = self.list_id
            self.entry_deleteId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_deleteId.current(0)


    def autoFillData(self, Event):
        for row in self.ws_Material.values:
            if self.entry_deleteId.get() == row[0]:
                self.existedMaterial = Material(row[0], row[1], row[2], row[3])
                break

        self.entry_deleteName.delete(0, END)
        self.entry_deleteQuantity.delete(0, END)
        self.entry_deleteUnit.delete(0, END)

        self.entry_deleteName.insert(0, self.existedMaterial.name)
        self.entry_deleteQuantity.insert(0, self.existedMaterial.quantity)
        self.entry_deleteUnit.insert(0, self.existedMaterial.unit)


    # run button
    def button_Delete(self):
        self.ws_Material_openpyxl = wb["Warehouse"]
        for row in self.ws_Material_openpyxl:
            if self.entry_deleteId.get() == row[0].value:
                self.ws_Material_openpyxl.delete_rows(row[0].row, 1)
                wb.save("Database.xlsx")
                break
            else:
                continue

        self.closeInsertWindow()

    def closeInsertWindow(self):
        self.canvas_delete.destroy()
        self.window_delete.destroy()

class EditMaterial():
    def __init__(self):
        self.window_edit = Toplevel()
        self.window_edit.title("Edit Material in Warehouse")
        self.window_edit.geometry("502x461")
        self.window_edit.configure(bg="#FFFFFF")

        self.ws_Material = pd.read_excel("Database.xlsx", sheet_name="Warehouse")

        self.canvas_edit = Canvas(
            self.window_edit,
            bg="#FFFFFF",
            height=461,
            width=502,
            bd=0,
            highlightthickness=0,
            relief="ridge"
        )

        self.canvas_edit.place(x=0, y=0)

        # button edit
        self.button_image_editMaterial = PhotoImage(
            file=relative_to_assets("button_editData.png"))
        self.button_editMaterial = Button(
            self.window_edit,
            image=self.button_image_editMaterial,
            borderwidth=0,
            highlightthickness=0,
            command=self.button_Edit,
            relief="flat"
        )
        self.button_editMaterial.place(
            x=203.99999999999997,
            y=381.0,
            width=94.0,
            height=40.0
        )

        # entry
        self.entry_image_shortData = PhotoImage(
            file=relative_to_assets("entry_insertShortData.png"))

        self.type = ["Trà", "Bột", "Cafe", "Syrup", "Đường", "Sữa"]

        self.entry_editType = ttk.Combobox(
            self.window_edit,
            value=self.type,
            font=("Helvetica", 10)
        )
        self.entry_editType.place(
            x=298,
            y=60,
            width=129,
            height=40
        )
        self.entry_editType.current(0)
        self.entry_editType.bind("<<ComboboxSelected>>", self.autoFillIdAndUnit)


        self.list_id = []
        self.entry_bg_id = self.canvas_edit.create_image(
            362.5,
            140.0,
            image=self.entry_image_shortData
        )
        self.entry_editId = ttk.Combobox(
            self.window_edit,
            font=("Helvetica", 12),
            value = self.list_id
        )
        self.entry_editId.place(
            x=298,
            y=120.0,
            width=129,
            height=40
        )
        # self.entry_editId.bind("<<ComboboxSelected>>", print("id selected"))
        self.entry_image_editLongData = PhotoImage(
            file=relative_to_assets("entry_insertLongData.png"))

        self.entry_bg_insertName = self.canvas_edit.create_image(
            325.5,
            200.5,
            image=self.entry_image_editLongData
        )
        self.entry_editName = Entry(
            self.window_edit,
            bd=0,
            bg="#DFDFDF",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_editName.place(
            x=233.99999999999997,
            y=180.0,
            width=183.0,
            height=39.0
        )

        self.entry_bg_quantity = self.canvas_edit.create_image(
            362.5,
            261.0,
            image=self.entry_image_shortData
        )
        self.entry_editQuantity = Entry(
            self.window_edit,
            bd=0,
            bg="#DFDFDF",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_editQuantity.place(
            x=308.0,
            y=241.0,
            width=109.0,
            height=38.0
        )

        self.entry_bg_unit = self.canvas_edit.create_image(
            362.5,
            321.0,
            image=self.entry_image_shortData
        )
        self.unit = ["hộp", "bịch", "chai", "Kg"]

        self.entry_editUnit = ttk.Combobox(
            self.window_edit,
            value=self.unit,
            font=("Helvetica", 10)
        )
        self.entry_editUnit.place(
            x=298.0,
            y=301.0,
            width=129,
            height=40
        )
        self.entry_editUnit.current(0)
        self.entry_editUnit.bind("<<ComboboxSelected>>")

        self.image_image_type = PhotoImage(
            file=relative_to_assets("image_type.png"))
        self.image_1 = self.canvas_edit.create_image(
            94.99999999999997,
            80.0,
            image=self.image_image_type
        )

        self.image_image_id = PhotoImage(
            file=relative_to_assets("image_id.png"))
        self.image_2 = self.canvas_edit.create_image(
            83.99999999999997,
            141.0,
            image=self.image_image_id
        )

        self.image_image_name = PhotoImage(
            file=relative_to_assets("image_name.png"))
        self.image_3 = self.canvas_edit.create_image(
            94.99999999999997,
            201.0,
            image=self.image_image_name
        )

        self.image_image_quantity = PhotoImage(
            file=relative_to_assets("image_quantity.png"))
        self.image_4 = self.canvas_edit.create_image(
            105.99999999999997,
            262.0,
            image=self.image_image_quantity
        )

        self.image_image_unit = PhotoImage(
            file=relative_to_assets("image_unit.png"))
        self.image_5 = self.canvas_edit.create_image(
            91.99999999999997,
            323.0,
            image=self.image_image_unit
        )
        self.window_edit.resizable(False, False)
        self.window_edit.mainloop()

    # auto fill
    def autoFillIdAndUnit(self, Event):
        self.searchList_material = self.ws_Material.to_numpy().tolist()

        if self.entry_editType.get() == "Trà":
            self.list_id = []
            for row in self.searchList_material:
                if "T" in row[0]:
                    self.list_id.append(row[0])
            self.entry_editId['values'] = self.list_id
            self.entry_editId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_editId.current(0)
            self.entry_editUnit.delete(0, END)
            self.entry_editUnit.insert(0, "hộp")


        elif self.entry_editType.get() == "Bột":
            self.list_id = []
            for row in self.searchList_material:
                if "F" in row[0]:
                    self.list_id.append(row[0])
            self.entry_editId['values'] = self.list_id
            self.entry_editId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_editId.current(0)
            self.entry_editUnit.delete(0, END)
            self.entry_editUnit.insert(0, "bịch")


        elif self.entry_editType.get() == "Cafe":
            self.list_id = []
            for row in self.searchList_material:
                if "C" in row[0]:
                    self.list_id.append(row[0])
            self.entry_editId['values'] = self.list_id
            self.entry_editId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_editId.current(0)
            self.entry_editUnit.delete(0, END)
            self.entry_editUnit.insert(0, "bịch")


        elif self.entry_editType.get() == "Syrup":
            self.list_id = []
            for row in self.searchList_material:
                if "SY" in row[0]:
                    self.list_id.append(row[0])
            self.entry_editId['values'] = self.list_id
            self.entry_editId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_editId.current(0)
            self.entry_editUnit.delete(0, END)
            self.entry_editUnit.insert(0, "chai")


        elif self.entry_editType.get() == "Sữa":
            self.list_id = []
            for row in self.searchList_material:
                if "M" in row[0]:
                    self.list_id.append(row[0])
            self.entry_editId['values'] = self.list_id
            self.entry_editId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_editId.current(0)
            self.entry_editUnit.delete(0, END)
            self.entry_editUnit.insert(0, "hộp")


        else:
            self.list_id = []
            for row in self.searchList_material:
                if "SU" in row[0]:
                    self.list_id.append(row[0])
            self.entry_editId['values'] = self.list_id
            self.entry_editId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_editId.current(0)
            self.entry_editUnit.delete(0, END)
            self.entry_editUnit.insert(0, "Kg")

    def autoFillData(self, Event):
        for row in self.ws_Material.values:
            if self.entry_editId.get() == row[0]:
                self.existedMaterial = Material(row[0], row[1], row[2], row[3])
                break

        self.entry_editName.delete(0, END)
        self.entry_editQuantity.delete(0, END)
        self.entry_editName.insert(0, self.existedMaterial.name)
        self.entry_editQuantity.insert(0, self.existedMaterial.quantity)


    # run button
    def button_Edit(self):
        check = True
        if self.entry_editName.get() == "":
            messagebox.showerror("Do not have Material name", "Please enter material name")
            check = False

        if self.entry_editQuantity.get() == "":
            messagebox.showerror("Do not have Quantity", "Please enter quantity")
            check = False

            try:
                self.quantity = 0
                if self.entry_editUnit.get() == "Kg":
                    self.quantity = float(self.entry_editQuantity.get())
                else:
                    self.quantity = int(self.entry_editQuantity.get())

                if self.quantity <= 0:
                    self.entry_editQuantity.delete(0, END)
                    messagebox.showerror("Incorrect quantity", "Please enter number greater than 0")
                    check = False

            except:
                self.entry_editQuantity.delete(0, END)
                messagebox.showerror("Incorrect type of quantity", "Please enter Interger of Float number")
                check = False

        if check == True:
            self.ws_Material_openpyxl = wb["Warehouse"]
            for row in self.ws_Material_openpyxl:
                if self.entry_editId.get() == row[0].value:
                    self.ws_Material_openpyxl["B"+str(row[0].row)].value = self.entry_editName.get()
                    self.ws_Material_openpyxl["C"+str(row[0].row)].value = self.entry_editQuantity.get()
                    self.ws_Material_openpyxl["D"+str(row[0].row)].value = self.entry_editUnit.get()
                    wb.save("Database.xlsx")
                    break
                else:
                    continue

            self.closeEditWindow()

    def closeEditWindow(self):
        self.canvas_edit.destroy()
        self.window_edit.destroy()

#StaffList
class InsertStaff():
    def __init__(self):
        self.window_insert = Toplevel()
        self.window_insert.title("Insert Staff")
        self.window_insert.geometry("502x600")
        self.window_insert.configure(bg="#FFFFFF")

        self.ws_StaffList = pd.read_excel("Database.xlsx", sheet_name="StaffList")

        self.canvas_insertStaff = Canvas(
            self.window_insert,
            bg="#FFFFFF",
            height=600,
            width=502,
            bd=0,
            highlightthickness=0,
            relief="ridge"
        )

        self.canvas_insertStaff.place(x=0, y=0)


        self.entry_image_1 = PhotoImage(
            file=relative_to_assets("entry_staffName.png"))
        self.entry_bg_1 = self.canvas_insertStaff.create_image(
            294.0,
            145.0,
            image=self.entry_image_1
        )
        self.entry_staffName = Entry(
            self.window_insert,
            bd=0,
            bg="#E9E9E9",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_staffName.place(
            x=150.0,
            y=124.0,
            width=300.0,
            height=40.0
        )

        self.image_image_1 = PhotoImage(
            file=relative_to_assets("image_staffName.png"))
        self.image_1 = self.canvas_insertStaff.create_image(
            88.0,
            145.0,
            image=self.image_image_1
        )

        self.entry_image_2 = PhotoImage(
            file=relative_to_assets("entry_staffRole.png"))
        self.entry_bg_2 = self.canvas_insertStaff.create_image(
            186.0,
            207.0,
            image=self.entry_image_2
        )
        self.role = ["Manager", "Staff"]

        self.entry_staffRole = ttk.Combobox(
            self.window_insert,
            value = self.role,
            font=("Helvetica", 10)
        )
        self.entry_staffRole.place(
            x=131.0,
            y=186.0,
            width=110.0,
            height=42.0
        )
        self.entry_staffRole.current(0)
        self.entry_staffRole.bind("<<ComboboxSelected>>", self.autoFillId)

        self.image_image_2 = PhotoImage(
            file=relative_to_assets("image_staffRole.png"))
        self.image_2 = self.canvas_insertStaff.create_image(
            88.0,
            207.0,
            image=self.image_image_2
        )

        self.entry_image_3 = PhotoImage(
            file=relative_to_assets("entry_staffGender.png"))
        self.entry_bg_3 = self.canvas_insertStaff.create_image(
            402.0,
            207.0,
            image=self.entry_image_3
        )
        self.gender = ["Male", "Female"]
        self.entry_staffGender = ttk.Combobox(
            self.window_insert,
            value = self.gender,
            font=("Helvetica", 10)
        )
        self.entry_staffGender.place(
            x=347.0,
            y=186.0,
            width=110,
            height=42.0
        )
        self.entry_staffGender.current(0)

        self.image_image_3 = PhotoImage(
            file=relative_to_assets("image_staffGender.png"))
        self.image_3 = self.canvas_insertStaff.create_image(
            304.0,
            207.0,
            image=self.image_image_3
        )

        self.entry_image_4 = PhotoImage(
            file=relative_to_assets("entry_staffId.png"))
        self.entry_bg_4 = self.canvas_insertStaff.create_image(
            294.0,
            269.0,
            image=self.entry_image_4
        )
        self.entry_staffId = Entry(
            self.window_insert,
            bd=0,
            bg="#E9E9E9",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_staffId.place(
            x=260,
            y=248.0,
            width=100,
            height=40.0
        )

        self.image_image_4 = PhotoImage(
            file=relative_to_assets("image_staffId.png"))
        self.image_4 = self.canvas_insertStaff.create_image(
            88.0,
            269.0,
            image=self.image_image_4
        )

        self.entry_image_5 = PhotoImage(
            file=relative_to_assets("entry_staffDate.png"))
        self.entry_bg_5 = self.canvas_insertStaff.create_image(
            186.0,
            331.0,
            image=self.entry_image_5
        )
        self.listDate = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,21]
        self.entry_staffDate = ttk.Combobox(
            self.window_insert,
            value = self.listDate,
            font=("Helvetica", 10)

        )
        self.entry_staffDate.place(
            x=131.0,
            y=310.0,
            width=110.0,
            height=42.0
        )

        self.entry_image_6 = PhotoImage(
            file=relative_to_assets("entry_staffMonth.png"))
        self.entry_bg_6 = self.canvas_insertStaff.create_image(
            294.0,
            331.0,
            image=self.entry_image_6
        )
        self.listMonth = [1,2,3,4,5,6,7,8,9,10,11,12]
        self.entry_staffMonth = ttk.Combobox(
            self.window_insert,
            value = self.listMonth,
            font=("Helvetica", 10)
        )
        self.entry_staffMonth.place(
            x=241.0,
            y=310.0,
            width=106,
            height=42.0
        )

        self.entry_image_7 = PhotoImage(
            file=relative_to_assets("entry_staffYear.png"))
        self.entry_bg_7 = self.canvas_insertStaff.create_image(
            402.0,
            331.0,
            image=self.entry_image_7
        )
        self.listYear = []
        for year in range(1960, 2003):
            self.listYear.append(year)
            year += 1

        self.entry_staffYear = ttk.Combobox(
            self.window_insert,
            value = self.listYear,
            font=("Helvetica", 10)
        )
        self.entry_staffYear.place(
            x=347.0,
            y=310.0,
            width=110,
            height=42.0
        )

        self.image_image_5 = PhotoImage(
            file=relative_to_assets("image_staffDate.png"))
        self.image_5 = self.canvas_insertStaff.create_image(
            88.0,
            331.0,
            image=self.image_image_5
        )

        self.entry_image_8 = PhotoImage(
            file=relative_to_assets("entry_staffAddress.png"))
        self.entry_bg_8 = self.canvas_insertStaff.create_image(
            294.0,
            393.0,
            image=self.entry_image_8
        )
        self.entry_staffAddress = Entry(
            self.window_insert,
            bd=0,
            bg="#E9E9E9",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_staffAddress.place(
            x=150.0,
            y=372.0,
            width=300.0,
            height=40.0
        )

        self.image_image_6 = PhotoImage(
            file=relative_to_assets("image_staffAddress.png"))
        self.image_6 = self.canvas_insertStaff.create_image(
            88.0,
            393.0,
            image=self.image_image_6
        )

        self.entry_image_9 = PhotoImage(
            file=relative_to_assets("entry_staffPhone.png"))
        self.entry_bg_9 = self.canvas_insertStaff.create_image(
            293.0,
            455.0,
            image=self.entry_image_9
        )
        self.entry_staffPhone = Entry(
            self.window_insert,
            bd=0,
            bg="#E9E9E9",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_staffPhone.place(
            x=150.0,
            y=434.0,
            width=300.0,
            height=40.0
        )

        self.image_image_7 = PhotoImage(
            file=relative_to_assets("image_staffPhone.png"))
        self.image_7 = self.canvas_insertStaff.create_image(
            88.0,
            455.0,
            image=self.image_image_7
        )

        self.button_image_1 = PhotoImage(
            file=relative_to_assets("button_staffInsert.png"))
        self.button_insertStaff = Button(
            self.window_insert,
            image=self.button_image_1,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: self.button_insert(),
            relief="flat"
        )
        self.button_insertStaff.place(
            x=196.0,
            y=516.0,
            width=110.0,
            height=42.0
        )

        self.image_image_8 = PhotoImage(
            file=relative_to_assets("image_InsertNewStaff.png"))
        self.image_8 = self.canvas_insertStaff.create_image(
            162.0,
            58.99999999999999,
            image=self.image_image_8
        )
        self.window_insert.resizable(False, False)
        self.window_insert.mainloop()

    #auto fill
    def autoFillId(self, Event):
        self.searchList_staff = self.ws_StaffList.to_numpy().tolist()
        self.searchList_staff.reverse()

        if self.entry_staffRole.get() == "Manager":
            for row in self.searchList_staff:
                if "MN" in row[0]:
                    num = int(row[0][2:len(row)]) + 1
                    newId = "MN"+str(num)
                    self.entry_staffId.delete(0, END)
                    self.entry_staffId.insert(0, newId)
                    break

        if self.entry_staffRole.get() == "Staff":
            for row in self.searchList_staff:
                if "ST" in row[0]:
                    num = int(row[0][2:len(row)]) + 1
                    newId = "ST"+str(num)
                    self.entry_staffId.delete(0, END)
                    self.entry_staffId.insert(0, newId)
                    break


    def button_insert(self):
        try:
            check = True
            if self.entry_staffName.get() == "":
                messagebox.showerror("Do not have Staff name", "Please enter staff name")
                check = False
            if self.entry_staffAddress.get() == "":
                messagebox.showerror("Do not have Staff address", "Please enter staff address")
                check = False
            if self.entry_staffPhone.get() == "":
                messagebox.showerror("Do not have Staff phone number", "Please enter staff phone number")
                check = False
            if self.entry_staffDate.get() == "":
                messagebox.showerror("Do not have Staff date", "Please enter staff date")
                check = False
            if self.entry_staffMonth.get() == "":
                messagebox.showerror("Do not have Staff month", "Please enter staff month")
                check = False
            if self.entry_staffYear.get() == "":
                messagebox.showerror("Do not have Staff year", "Please enter staff year")
                check = False
            if len(self.entry_staffPhone.get()) == 10:
                if self.entry_staffPhone.get()[0] == "0":
                    if re.search('[A-Za-z]', self.entry_staffPhone.get()):
                        messagebox.showerror("Incorrect staff phone number", "Please enter staff phone number")
                        self.entry_staffPhone.delete(0, END)
                        check = False
                else:
                    messagebox.showerror("Incorrect staff phone number", "Please enter staff phone number")
                    self.entry_staffPhone.delete(0, END)
                    check = False
            else:
                messagebox.showerror("Incorrect staff phone number", "Please enter staff phone number")
                self.entry_staffPhone.delete(0, END)
                check = False

            if check == True:
                self.ws_StaffList_openpyxl = wb["StaffList"]

                self.list_fullName = [name.capitalize() for name in self.entry_staffName.get().split()]
                self.fullName = " "
                self.fullName = self.fullName.join(self.list_fullName)

                self.date = self.entry_staffDate.get()+"/"+self.entry_staffMonth.get()+"/"+self.entry_staffYear.get()

                self.ws_StaffList_openpyxl.append([self.entry_staffId.get(),
                                                   self.fullName,
                                                   self.entry_staffGender.get(),
                                                   self.date,
                                                   self.entry_staffAddress.get(),
                                                   self.entry_staffPhone.get(),
                                                   self.entry_staffRole.get()
                                                   ])
                wb.save("Database.xlsx")
                self.closeInsertWindow()


        except:
            self.entry_staffName.delete(0, END)
            self.entry_staffRole.delete(0, END)
            self.entry_staffGender.delete(0, END)
            self.entry_staffDate.delete(0, END)
            self.entry_staffMonth.delete(0, END)
            self.entry_staffYear.delete(0, END)
            self.entry_staffAddress.delete(0, END)
            self.entry_staffPhone.delete(0, END)
            messagebox.showerror("Incorrect", "Please try again")

    def closeInsertWindow(self):
        self.canvas_insertStaff.destroy()
        self.window_insert.destroy()

class DeleteStaff():
    def __init__(self):
        self.window_delete = Toplevel()
        self.window_delete.title("Delete Staff")
        self.window_delete.geometry("502x600")
        self.window_delete.configure(bg="#FFFFFF")

        self.ws_StaffList = pd.read_excel("Database.xlsx", sheet_name="StaffList")

        self.canvas_deleteStaff = Canvas(
            self.window_delete,
            bg="#FFFFFF",
            height=600,
            width=502,
            bd=0,
            highlightthickness=0,
            relief="ridge"
        )

        self.canvas_deleteStaff.place(x=0, y=0)


        self.entry_image_1 = PhotoImage(
            file=relative_to_assets("entry_staffName.png"))
        self.entry_bg_1 = self.canvas_deleteStaff.create_image(
            294.0,
            145.0,
            image=self.entry_image_1
        )
        self.entry_staffName = Entry(
            self.window_delete,
            bd=0,
            bg="#E9E9E9",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_staffName.place(
            x=150.0,
            y=124.0,
            width=300.0,
            height=40.0
        )

        self.image_image_1 = PhotoImage(
            file=relative_to_assets("image_staffName.png"))
        self.image_1 = self.canvas_deleteStaff.create_image(
            88.0,
            145.0,
            image=self.image_image_1
        )

        self.entry_image_2 = PhotoImage(
            file=relative_to_assets("entry_staffRole.png"))
        self.entry_bg_2 = self.canvas_deleteStaff.create_image(
            186.0,
            207.0,
            image=self.entry_image_2
        )
        self.role = ["Manager", "Staff"]

        self.entry_staffRole = ttk.Combobox(
            self.window_delete,
            value = self.role,
            font=("Helvetica", 10)
        )
        self.entry_staffRole.place(
            x=131.0,
            y=186.0,
            width=110.0,
            height=42.0
        )
        self.entry_staffRole.current(0)
        self.entry_staffRole.bind("<<ComboboxSelected>>", self.autoFillId)

        self.image_image_2 = PhotoImage(
            file=relative_to_assets("image_staffRole.png"))
        self.image_2 = self.canvas_deleteStaff.create_image(
            88.0,
            207.0,
            image=self.image_image_2
        )

        self.entry_image_3 = PhotoImage(
            file=relative_to_assets("entry_staffGender.png"))
        self.entry_bg_3 = self.canvas_deleteStaff.create_image(
            402.0,
            207.0,
            image=self.entry_image_3
        )
        self.gender = ["Male", "Female"]
        self.entry_staffGender = ttk.Combobox(
            self.window_delete,
            value = self.gender,
            font=("Helvetica", 10)
        )
        self.entry_staffGender.place(
            x=347.0,
            y=186.0,
            width=110,
            height=42.0
        )
        self.entry_staffGender.current(0)

        self.image_image_3 = PhotoImage(
            file=relative_to_assets("image_staffGender.png"))
        self.image_3 = self.canvas_deleteStaff.create_image(
            304.0,
            207.0,
            image=self.image_image_3
        )

        self.entry_image_4 = PhotoImage(
            file=relative_to_assets("entry_staffId.png"))
        self.entry_bg_4 = self.canvas_deleteStaff.create_image(
            294.0,
            269.0,
            image=self.entry_image_4
        )
        self.list_id = []
        self.entry_staffId = ttk.Combobox(
            self.window_delete,
            value = self.list_id,
            font=("Helvetica", 12),
        )
        self.entry_staffId.place(
            x=131,
            y=248.0,
            width=326,
            height=42.0
        )

        self.image_image_4 = PhotoImage(
            file=relative_to_assets("image_staffId.png"))
        self.image_4 = self.canvas_deleteStaff.create_image(
            88.0,
            269.0,
            image=self.image_image_4
        )

        self.entry_image_5 = PhotoImage(
            file=relative_to_assets("entry_staffDate.png"))
        self.entry_bg_5 = self.canvas_deleteStaff.create_image(
            186.0,
            331.0,
            image=self.entry_image_5
        )
        self.listDate = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,21]
        self.entry_staffDate = ttk.Combobox(
            self.window_delete,
            value = self.listDate,
            font=("Helvetica", 10)

        )
        self.entry_staffDate.place(
            x=131.0,
            y=310.0,
            width=110.0,
            height=42.0
        )

        self.entry_image_6 = PhotoImage(
            file=relative_to_assets("entry_staffMonth.png"))
        self.entry_bg_6 = self.canvas_deleteStaff.create_image(
            294.0,
            331.0,
            image=self.entry_image_6
        )
        self.listMonth = [1,2,3,4,5,6,7,8,9,10,11,12]
        self.entry_staffMonth = ttk.Combobox(
            self.window_delete,
            value = self.listMonth,
            font=("Helvetica", 10)
        )
        self.entry_staffMonth.place(
            x=241.0,
            y=310.0,
            width=106,
            height=42.0
        )

        self.entry_image_7 = PhotoImage(
            file=relative_to_assets("entry_staffYear.png"))
        self.entry_bg_7 = self.canvas_deleteStaff.create_image(
            402.0,
            331.0,
            image=self.entry_image_7
        )
        self.listYear = []
        for year in range(1960, 2003):
            self.listYear.append(year)
            year += 1

        self.entry_staffYear = ttk.Combobox(
            self.window_delete,
            value = self.listYear,
            font=("Helvetica", 10)
        )
        self.entry_staffYear.place(
            x=347.0,
            y=310.0,
            width=110,
            height=42.0
        )

        self.image_image_5 = PhotoImage(
            file=relative_to_assets("image_staffDate.png"))
        self.image_5 = self.canvas_deleteStaff.create_image(
            88.0,
            331.0,
            image=self.image_image_5
        )

        self.entry_image_8 = PhotoImage(
            file=relative_to_assets("entry_staffAddress.png"))
        self.entry_bg_8 = self.canvas_deleteStaff.create_image(
            294.0,
            393.0,
            image=self.entry_image_8
        )
        self.entry_staffAddress = Entry(
            self.window_delete,
            bd=0,
            bg="#E9E9E9",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_staffAddress.place(
            x=150.0,
            y=372.0,
            width=300.0,
            height=40.0
        )

        self.image_image_6 = PhotoImage(
            file=relative_to_assets("image_staffAddress.png"))
        self.image_6 = self.canvas_deleteStaff.create_image(
            88.0,
            393.0,
            image=self.image_image_6
        )

        self.entry_image_9 = PhotoImage(
            file=relative_to_assets("entry_staffPhone.png"))
        self.entry_bg_9 = self.canvas_deleteStaff.create_image(
            293.0,
            455.0,
            image=self.entry_image_9
        )
        self.entry_staffPhone = Entry(
            self.window_delete,
            bd=0,
            bg="#E9E9E9",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_staffPhone.place(
            x=150.0,
            y=434.0,
            width=300.0,
            height=40.0
        )

        self.image_image_7 = PhotoImage(
            file=relative_to_assets("image_staffPhone.png"))
        self.image_7 = self.canvas_deleteStaff.create_image(
            88.0,
            455.0,
            image=self.image_image_7
        )

        self.button_image_1 = PhotoImage(
            file=relative_to_assets("button_staffDelete.png"))
        self.button_insertStaff = Button(
            self.window_delete,
            image=self.button_image_1,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: self.button_delete(),
            relief="flat"
        )
        self.button_insertStaff.place(
            x=196.0,
            y=516.0,
            width=110.0,
            height=42.0
        )

        self.image_image_8 = PhotoImage(
            file=relative_to_assets("image_DeleteStaff.png"))
        self.image_8 = self.canvas_deleteStaff.create_image(
            128.0,
            58.99999999999999,
            image=self.image_image_8
        )
        self.window_delete.resizable(False, False)
        self.window_delete.mainloop()

    def autoFillId(self, Event):
        self.searchList_staff = self.ws_StaffList.to_numpy().tolist()

        if self.entry_staffRole.get() == "Manager":
            self.list_id = []
            for row in self.searchList_staff:
                if "MN" in row[0]:
                    self.list_id.append(row[0])
            self.entry_staffId['values'] = self.list_id
            self.entry_staffId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_staffId.current(0)

        if self.entry_staffRole.get() == "Staff":
            self.list_id = []
            for row in self.searchList_staff:
                if "ST" in row[0]:
                    self.list_id.append(row[0])
            self.entry_staffId['values'] = self.list_id
            self.entry_staffId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_staffId.current(0)


    def autoFillData(self, Event):
        for row in self.ws_StaffList.values:
            if self.entry_staffId.get() == row[0]:
                self.existedStaff = Staff(row[0], row[1], row[2], row[3], row[4], row[5], row[6],)
                break

        self.entry_staffName.delete(0, END)
        self.entry_staffRole.delete(0, END)
        self.entry_staffGender.delete(0, END)
        self.entry_staffId.delete(0, END)
        self.entry_staffAddress.delete(0, END)
        self.entry_staffPhone.delete(0, END)
        self.entry_staffDate.delete(0, END)
        self.entry_staffMonth.delete(0, END)
        self.entry_staffYear.delete(0, END)

        self.entry_staffName.insert(0, self.existedStaff.name)
        self.entry_staffRole.insert(0, self.existedStaff.role)
        self.entry_staffGender.insert(0, self.existedStaff.gender)
        self.entry_staffId.insert(0, self.existedStaff.id)
        self.entry_staffAddress.insert(0, self.existedStaff.address)
        self.entry_staffPhone.insert(0, self.existedStaff.phoneNumber)

        date = re.findall(r'\d+', self.existedStaff.date)
        self.entry_staffDate.insert(0, date[0])
        self.entry_staffMonth.insert(0, date[1])
        self.entry_staffYear.insert(0, date[2])

    def closeDeleteWindow(self):
        self.canvas_deleteStaff.destroy()
        self.window_delete.destroy()

    def button_delete(self):
        self.ws_StaffList_openpyxl = wb["StaffList"]
        for row in self.ws_StaffList_openpyxl:
            if self.entry_staffId.get() == row[0].value:
                self.ws_StaffList_openpyxl.delete_rows(row[0].row, 1)
                wb.save("Database.xlsx")
                break
            else:
                continue
        self.closeDeleteWindow()

class EditStaff():
    def __init__(self):
        self.window_edit = Toplevel()
        self.window_edit.title("Delete Staff")
        self.window_edit.geometry("502x600")
        self.window_edit.configure(bg="#FFFFFF")

        self.ws_StaffList = pd.read_excel("Database.xlsx", sheet_name="StaffList")

        self.canvas_editStaff = Canvas(
            self.window_edit,
            bg="#FFFFFF",
            height=600,
            width=502,
            bd=0,
            highlightthickness=0,
            relief="ridge"
        )

        self.canvas_editStaff.place(x=0, y=0)


        self.entry_image_1 = PhotoImage(
            file=relative_to_assets("entry_staffName.png"))
        self.entry_bg_1 = self.canvas_editStaff.create_image(
            294.0,
            145.0,
            image=self.entry_image_1
        )
        self.entry_staffName = Entry(
            self.window_edit,
            bd=0,
            bg="#E9E9E9",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_staffName.place(
            x=150.0,
            y=124.0,
            width=300.0,
            height=40.0
        )

        self.image_image_1 = PhotoImage(
            file=relative_to_assets("image_staffName.png"))
        self.image_1 = self.canvas_editStaff.create_image(
            88.0,
            145.0,
            image=self.image_image_1
        )

        self.entry_image_2 = PhotoImage(
            file=relative_to_assets("entry_staffRole.png"))
        self.entry_bg_2 = self.canvas_editStaff.create_image(
            186.0,
            207.0,
            image=self.entry_image_2
        )
        self.role = ["Manager", "Staff"]

        self.entry_staffRole = ttk.Combobox(
            self.window_edit,
            value = self.role,
            font=("Helvetica", 10)
        )
        self.entry_staffRole.place(
            x=131.0,
            y=186.0,
            width=110.0,
            height=42.0
        )
        self.entry_staffRole.current(0)
        self.entry_staffRole.bind("<<ComboboxSelected>>", self.autoFillId)

        self.image_image_2 = PhotoImage(
            file=relative_to_assets("image_staffRole.png"))
        self.image_2 = self.canvas_editStaff.create_image(
            88.0,
            207.0,
            image=self.image_image_2
        )

        self.entry_image_3 = PhotoImage(
            file=relative_to_assets("entry_staffGender.png"))
        self.entry_bg_3 = self.canvas_editStaff.create_image(
            402.0,
            207.0,
            image=self.entry_image_3
        )
        self.gender = ["Male", "Female"]
        self.entry_staffGender = ttk.Combobox(
            self.window_edit,
            value = self.gender,
            font=("Helvetica", 10)
        )
        self.entry_staffGender.place(
            x=347.0,
            y=186.0,
            width=110,
            height=42.0
        )
        self.entry_staffGender.current(0)

        self.image_image_3 = PhotoImage(
            file=relative_to_assets("image_staffGender.png"))
        self.image_3 = self.canvas_editStaff.create_image(
            304.0,
            207.0,
            image=self.image_image_3
        )

        self.entry_image_4 = PhotoImage(
            file=relative_to_assets("entry_staffId.png"))
        self.entry_bg_4 = self.canvas_editStaff.create_image(
            294.0,
            269.0,
            image=self.entry_image_4
        )
        self.list_id = []
        self.entry_staffId = ttk.Combobox(
            self.window_edit,
            value = self.list_id,
            font=("Helvetica", 12),
        )
        self.entry_staffId.place(
            x=131,
            y=248.0,
            width=326,
            height=42.0
        )

        self.image_image_4 = PhotoImage(
            file=relative_to_assets("image_staffId.png"))
        self.image_4 = self.canvas_editStaff.create_image(
            88.0,
            269.0,
            image=self.image_image_4
        )

        self.entry_image_5 = PhotoImage(
            file=relative_to_assets("entry_staffDate.png"))
        self.entry_bg_5 = self.canvas_editStaff.create_image(
            186.0,
            331.0,
            image=self.entry_image_5
        )
        self.listDate = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,21]
        self.entry_staffDate = ttk.Combobox(
            self.window_edit,
            value = self.listDate,
            font=("Helvetica", 10)

        )
        self.entry_staffDate.place(
            x=131.0,
            y=310.0,
            width=110.0,
            height=42.0
        )

        self.entry_image_6 = PhotoImage(
            file=relative_to_assets("entry_staffMonth.png"))
        self.entry_bg_6 = self.canvas_editStaff.create_image(
            294.0,
            331.0,
            image=self.entry_image_6
        )
        self.listMonth = [1,2,3,4,5,6,7,8,9,10,11,12]
        self.entry_staffMonth = ttk.Combobox(
            self.window_edit,
            value = self.listMonth,
            font=("Helvetica", 10)
        )
        self.entry_staffMonth.place(
            x=241.0,
            y=310.0,
            width=106,
            height=42.0
        )

        self.entry_image_7 = PhotoImage(
            file=relative_to_assets("entry_staffYear.png"))
        self.entry_bg_7 = self.canvas_editStaff.create_image(
            402.0,
            331.0,
            image=self.entry_image_7
        )
        self.listYear = []
        for year in range(1960, 2003):
            self.listYear.append(year)
            year += 1

        self.entry_staffYear = ttk.Combobox(
            self.window_edit,
            value = self.listYear,
            font=("Helvetica", 10)
        )
        self.entry_staffYear.place(
            x=347.0,
            y=310.0,
            width=110,
            height=42.0
        )

        self.image_image_5 = PhotoImage(
            file=relative_to_assets("image_staffDate.png"))
        self.image_5 = self.canvas_editStaff.create_image(
            88.0,
            331.0,
            image=self.image_image_5
        )

        self.entry_image_8 = PhotoImage(
            file=relative_to_assets("entry_staffAddress.png"))
        self.entry_bg_8 = self.canvas_editStaff.create_image(
            294.0,
            393.0,
            image=self.entry_image_8
        )
        self.entry_staffAddress = Entry(
            self.window_edit,
            bd=0,
            bg="#E9E9E9",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_staffAddress.place(
            x=150.0,
            y=372.0,
            width=300.0,
            height=40.0
        )

        self.image_image_6 = PhotoImage(
            file=relative_to_assets("image_staffAddress.png"))
        self.image_6 = self.canvas_editStaff.create_image(
            88.0,
            393.0,
            image=self.image_image_6
        )

        self.entry_image_9 = PhotoImage(
            file=relative_to_assets("entry_staffPhone.png"))
        self.entry_bg_9 = self.canvas_editStaff.create_image(
            293.0,
            455.0,
            image=self.entry_image_9
        )
        self.entry_staffPhone = Entry(
            self.window_edit,
            bd=0,
            bg="#E9E9E9",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_staffPhone.place(
            x=150.0,
            y=434.0,
            width=300.0,
            height=40.0
        )

        self.image_image_7 = PhotoImage(
            file=relative_to_assets("image_staffPhone.png"))
        self.image_7 = self.canvas_editStaff.create_image(
            88.0,
            455.0,
            image=self.image_image_7
        )

        self.button_image_1 = PhotoImage(
            file=relative_to_assets("button_staffEdit.png"))
        self.button_insertStaff = Button(
            self.window_edit,
            image=self.button_image_1,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: self.button_edit(),
            relief="flat"
        )
        self.button_insertStaff.place(
            x=196.0,
            y=516.0,
            width=110.0,
            height=42.0
        )

        self.image_image_8 = PhotoImage(
            file=relative_to_assets("image_EditStaff.png"))
        self.image_8 = self.canvas_editStaff.create_image(
            110,
            58.99999999999999,
            image=self.image_image_8
        )
        self.window_edit.resizable(False, False)
        self.window_edit.mainloop()

    def autoFillId(self, Event):
        self.searchList_staff = self.ws_StaffList.to_numpy().tolist()

        if self.entry_staffRole.get() == "Manager":
            self.list_id = []
            for row in self.searchList_staff:
                if "MN" in row[0]:
                    self.list_id.append(row[0])
            self.entry_staffId['values'] = self.list_id
            self.entry_staffId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_staffId.current(0)

        if self.entry_staffRole.get() == "Staff":
            self.list_id = []
            for row in self.searchList_staff:
                if "ST" in row[0]:
                    self.list_id.append(row[0])
            self.entry_staffId['values'] = self.list_id
            self.entry_staffId.bind("<<ComboboxSelected>>", self.autoFillData)
            self.entry_staffId.current(0)


    def autoFillData(self, Event):
        for row in self.ws_StaffList.values:
            if self.entry_staffId.get() == row[0]:
                self.existedStaff = Staff(row[0], row[1], row[2], row[3], row[4], str(row[5]), row[6])
                break

        self.entry_staffName.delete(0, END)
        self.entry_staffRole.delete(0, END)
        self.entry_staffGender.delete(0, END)
        self.entry_staffId.delete(0, END)
        self.entry_staffAddress.delete(0, END)
        self.entry_staffPhone.delete(0, END)
        self.entry_staffDate.delete(0, END)
        self.entry_staffMonth.delete(0, END)
        self.entry_staffYear.delete(0, END)

        self.entry_staffName.insert(0, self.existedStaff.name)
        self.entry_staffRole.insert(0, self.existedStaff.role)
        self.entry_staffGender.insert(0, self.existedStaff.gender)
        self.entry_staffId.insert(0, self.existedStaff.id)
        self.entry_staffAddress.insert(0, self.existedStaff.address)
        self.entry_staffPhone.insert(0, str(self.existedStaff.phoneNumber))

        date = re.findall(r'\d+', self.existedStaff.date)
        self.entry_staffDate.insert(0, date[0])
        self.entry_staffMonth.insert(0, date[1])
        self.entry_staffYear.insert(0, date[2])

    def closeEditWindow(self):
        self.canvas_editStaff.destroy()
        self.window_edit.destroy()

    def button_edit(self):
        try:
            check = True
            if self.entry_staffName.get() == "":
                messagebox.showerror("Do not have Staff name", "Please enter staff name")
                check = False
            if self.entry_staffAddress.get() == "":
                messagebox.showerror("Do not have Staff address", "Please enter staff address")
                check = False
            if self.entry_staffPhone.get() == "":
                messagebox.showerror("Do not have Staff phone number", "Please enter staff phone number")
                check = False
            if self.entry_staffDate.get() == "":
                messagebox.showerror("Do not have Staff date", "Please enter staff date")
                check = False
            if self.entry_staffMonth.get() == "":
                messagebox.showerror("Do not have Staff month", "Please enter staff month")
                check = False
            if self.entry_staffYear.get() == "":
                messagebox.showerror("Do not have Staff year", "Please enter staff year")
                check = False
            if len(self.entry_staffPhone.get()) == 9:
                if re.search('[A-Za-z]', self.entry_staffPhone.get()):
                    messagebox.showerror("Incorrect staff phone number", "Please enter staff phone number")
                    self.entry_staffPhone.delete(0, END)
                    check = False
            else:
                messagebox.showerror("Incorrect staff phone number", "Please enter staff phone number")
                self.entry_staffPhone.delete(0, END)
                check = False

            if check == True:
                self.ws_StaffList_openpyxl = wb["StaffList"]

                self.list_fullName = [name.capitalize() for name in self.entry_staffName.get().split()]
                self.fullName = " "
                self.fullName = self.fullName.join(self.list_fullName)

                self.date = self.entry_staffDate.get()+"/"+self.entry_staffMonth.get()+"/"+self.entry_staffYear.get()

                for row in self.ws_StaffList_openpyxl:
                    if self.entry_staffId.get() == row[0].value:
                        self.ws_StaffList_openpyxl["B"+str(row[0].row)].value = self.fullName
                        self.ws_StaffList_openpyxl["C"+str(row[0].row)].value = self.entry_staffGender.get()
                        self.ws_StaffList_openpyxl["D"+str(row[0].row)].value = self.date
                        self.ws_StaffList_openpyxl["E"+str(row[0].row)].value = self.entry_staffAddress.get()
                        self.ws_StaffList_openpyxl["F"+str(row[0].row)].value = "0"+str(self.entry_staffPhone.get())
                        self.ws_StaffList_openpyxl["G"+str(row[0].row)].value = self.entry_staffRole.get()
                        wb.save("Database.xlsx")
                        break
                    else:
                        continue

                self.closeEditWindow()


        except:
            self.entry_staffName.delete(0, END)
            self.entry_staffRole.delete(0, END)
            self.entry_staffGender.delete(0, END)
            self.entry_staffDate.delete(0, END)
            self.entry_staffMonth.delete(0, END)
            self.entry_staffYear.delete(0, END)
            self.entry_staffAddress.delete(0, END)
            self.entry_staffPhone.delete(0, END)
            messagebox.showerror("Incorrect", "Please try again")

#CustomerList
class InsertCustomer():
    def __init__(self):
        self.window_insert = Toplevel()
        self.window_insert.title("Insert Customer")
        self.window_insert.geometry("502x500")
        self.window_insert.configure(bg="#FFFFFF")

        self.ws_CustomerList = pd.read_excel("Database.xlsx", sheet_name="CustomerList")

        self.canvas_insertCustomer = Canvas(
            self.window_insert,
            bg="#FFFFFF",
            height=500,
            width=502,
            bd=0,
            highlightthickness=0,
            relief="ridge"
        )

        self.canvas_insertCustomer.place(x=0, y=0)
        self.image_image_1 = PhotoImage(
            file=relative_to_assets("image_InsertNewCustomer.png"))
        self.image_1 = self.canvas_insertCustomer.create_image(
            196.0,
            49.0,
            image=self.image_image_1
        )

        self.entry_image_1 = PhotoImage(
            file=relative_to_assets("entry_customerName.png"))
        self.entry_bg_1 = self.canvas_insertCustomer.create_image(
            294.0,
            134.5,
            image=self.entry_image_1
        )
        self.entry_customerName = Entry(
            self.window_insert,
            bd=0,
            bg="#E9E9E9",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_customerName.place(
            x=150.0,
            y=116.0,
            width=290.0,
            height=35.0
        )

        self.image_image_2 = PhotoImage(
            file=relative_to_assets("image_customerName.png"))
        self.image_2 = self.canvas_insertCustomer.create_image(
            88.0,
            134.0,
            image=self.image_image_2
        )

        self.entry_image_2 = PhotoImage(
            file=relative_to_assets("entry_customerId.png"))
        self.entry_bg_2 = self.canvas_insertCustomer.create_image(
            186.0,
            192.0,
            image=self.entry_image_2
        )
        self.entry_customerId = Entry(
            self.window_insert,
            bd=0,
            bg="#E9E9E9",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_customerId.place(
            x=170.0,
            y=173.0,
            width=50,
            height=36.0
        )

        self.image_image_3 = PhotoImage(
            file=relative_to_assets("image_customerId.png"))
        self.image_3 = self.canvas_insertCustomer.create_image(
            88.0,
            192.0,
            image=self.image_image_3
        )
        self.searchList_customer = self.ws_CustomerList.to_numpy().tolist()
        self.searchList_customer.reverse()
        for row in self.searchList_customer:
            if "C" in row[0]:
                num = int(row[0][1:len(row)]) + 1
                newId = "C"+str(num)
                self.entry_customerId.delete(0, END)
                self.entry_customerId.insert(0, newId)
                break


        self.entry_image_3 = PhotoImage(
            file=relative_to_assets("entry_customerGender.png"))
        self.entry_bg_3 = self.canvas_insertCustomer.create_image(
            402.0,
            192.0,
            image=self.entry_image_3
        )
        self.gender = ["Male", "Female"]
        self.entry_customerGender = ttk.Combobox(
            self.window_insert,
            value = self.gender,
            font=("Helvetica",10)
        )
        self.entry_customerGender.place(
            x=347,
            y=173.0,
            width=110,
            height=38.0
        )
        self.entry_customerGender.current(0)

        self.image_image_4 = PhotoImage(
            file=relative_to_assets("image_customerGender.png"))
        self.image_4 = self.canvas_insertCustomer.create_image(
            304.0,
            192.0,
            image=self.image_image_4
        )

        self.image_image_5 = PhotoImage(
            file=relative_to_assets("image_customerDate.png"))
        self.image_5 = self.canvas_insertCustomer.create_image(
            88.0,
            250.0,
            image=self.image_image_5
        )

        self.entry_image_4 = PhotoImage(
            file=relative_to_assets("entry_customerDate.png"))
        self.entry_bg_4 = self.canvas_insertCustomer.create_image(
            186.0,
            250.0,
            image=self.entry_image_4
        )
        self.listDate = [1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,21]
        self.entry_customerDate = ttk.Combobox(
            self.window_insert,
            value = self.listDate,
            font=("Helvetica", 10)
        )
        self.entry_customerDate.place(
            x=131.0,
            y=231.0,
            width=110.0,
            height=36.0
        )

        self.entry_image_5 = PhotoImage(
            file=relative_to_assets("entry_customerMonth.png"))
        self.entry_bg_5 = self.canvas_insertCustomer.create_image(
            294.0,
            250.0,
            image=self.entry_image_5
        )
        self.listMonth = [1,2,3,4,5,6,7,8,9,10,11,12]
        self.entry_customerMonth = ttk.Combobox(
            self.window_insert,
            value = self.listMonth,
            font=("Helvetica", 10)
        )
        self.entry_customerMonth.place(
            x=241.0,
            y=231.0,
            width=106.0,
            height=36.0
        )

        self.entry_image_6 = PhotoImage(
            file=relative_to_assets("entry_customeryear.png"))
        self.entry_bg_6 = self.canvas_insertCustomer.create_image(
            402.0,
            250.0,
            image=self.entry_image_6
        )
        self.listYear = []
        for year in range(1930, 2017):
            self.listYear.append(year)
            year += 1

        self.entry_customerYear = ttk.Combobox(
            self.window_insert,
            value = self.listYear,
            font=("Helvetica", 10)
        )
        self.entry_customerYear.place(
            x=347.0,
            y=231.0,
            width=110.0,
            height=36.0
        )

        self.entry_image_7 = PhotoImage(
            file=relative_to_assets("entry_customerAddress.png"))
        self.entry_bg_7 = self.canvas_insertCustomer.create_image(
            294.0,
            308.0,
            image=self.entry_image_7
        )
        self.entry_customerAddress = Entry(
            self.window_insert,
            bd=0,
            bg="#E9E9E9",
            font=("Helvetica",12),
            highlightthickness=0
        )
        self.entry_customerAddress.place(
            x=150,
            y=289.0,
            width=290,
            height=36.0
        )

        self.image_image_6 = PhotoImage(
            file=relative_to_assets("image_customerAddress.png"))
        self.image_6 = self.canvas_insertCustomer.create_image(
            88.0,
            308.0,
            image=self.image_image_6
        )

        self.entry_image_8 = PhotoImage(
            file=relative_to_assets("entry_customerPhone.png"))
        self.entry_bg_8 = self.canvas_insertCustomer.create_image(
            294.0,
            365.5,
            image=self.entry_image_8
        )
        self.entry_customerPhone = Entry(
            self.window_insert,
            bd=0,
            bg="#E9E9E9",
            font=("Helvetica", 12),
            highlightthickness=0
        )
        self.entry_customerPhone.place(
            x=150,
            y=347.0,
            width=290,
            height=35.0
        )

        self.image_image_7 = PhotoImage(
            file=relative_to_assets("image_customerPhone.png"))
        self.image_7 = self.canvas_insertCustomer.create_image(
            88.0,
            365.0,
            image=self.image_image_7
        )

        self.button_image_1 = PhotoImage(
            file=relative_to_assets("button_customerInsert.png"))
        self.button_insertCustomer = Button(
            self.window_insert,
            image=self.button_image_1,
            borderwidth=0,
            highlightthickness=0,
            command=lambda: self.button_insert(),
            relief="flat"
        )
        self.button_insertCustomer.place(
            x=196.0,
            y=428.0,
            width=110.0,
            height=34.0
        )
        self.window_insert.resizable(False, False)
        self.window_insert.mainloop()

    def button_insert(self):
        try:
            check = True
            if self.entry_customerName.get() == "":
                messagebox.showerror("Do not have customer name", "Please enter customer name")
                check = False
            if self.entry_customerAddress.get() == "":
                messagebox.showerror("Do not have customer address", "Please enter customer address")
                check = False
            if self.entry_customerPhone.get() == "":
                messagebox.showerror("Do not have customer phone number", "Please enter customer phone number")
                check = False
            if self.entry_customerDate.get() == "":
                messagebox.showerror("Do not have customer date", "Please enter customer date")
                check = False
            if self.entry_customerMonth.get() == "":
                messagebox.showerror("Do not have customer month", "Please enter customer month")
                check = False
            if self.entry_customerYear.get() == "":
                messagebox.showerror("Do not have customer year", "Please enter customer year")
                check = False
            if len(self.entry_customerPhone.get()) == 10:
                if self.entry_customerPhone.get()[0] == "0":
                    if re.search('[A-Za-z]', self.entry_customerPhone.get()):
                        messagebox.showerror("Incorrect customer phone number", "Please enter customer phone number")
                        self.entry_customerPhone.delete(0, END)
                        check = False
                else:
                    messagebox.showerror("Incorrect customer phone number", "Please enter customer phone number")
                    self.entry_customerPhone.delete(0, END)
                    check = False
            else:
                messagebox.showerror("Incorrect customer phone number", "Please enter customer phone number")
                self.entry_customerPhone.delete(0, END)
                check = False

            if check == True:
                self.ws_CustomerList_openpyxl = wb["CustomerList"]

                self.list_fullName = [name.capitalize() for name in self.entry_customerName.get().split()]
                self.fullName = " "
                self.fullName = self.fullName.join(self.list_fullName)

                self.date = self.entry_customerDate.get()+"/"+self.entry_customerMonth.get()+"/"+self.entry_customerYear.get()

                self.ws_CustomerList_openpyxl.append([self.entry_customerId.get(),
                                                   self.fullName,
                                                   self.entry_customerGender.get(),
                                                   self.date,
                                                   self.entry_customerAddress.get(),
                                                   self.entry_customerPhone.get()
                                                   ])
                wb.save("Database.xlsx")
                self.closeInsertWindow()


        except:
            self.entry_customerName.delete(0, END)
            self.entry_customerGender.delete(0, END)
            self.entry_customerDate.delete(0, END)
            self.entry_customerMonth.delete(0, END)
            self.entry_customerYear.delete(0, END)
            self.entry_customerAddress.delete(0, END)
            self.entry_customerPhone.delete(0, END)
            messagebox.showerror("Incorrect", "Please try again")

    def closeInsertWindow(self):
        self.canvas_insertCustomer.destroy()
        self.window_insert.destroy()

class DeleteCustomer():
    pass

class EditCustomer():
    pass
