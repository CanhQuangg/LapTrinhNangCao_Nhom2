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