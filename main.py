from screeninfo import get_monitors
import sqlite3
import subprocess
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, font, filedialog
from tkcalendar import *
import xlsxwriter

from datetime import datetime
from datetime import date as dt

class App(tk.Tk):
    def __init__(self, date, x, y):
        super().__init__()
        self.money_box = None
        self.daily_income = None
        self.in_month = None
        self.date = date
        self.today = str(dt.today())
        self.geometry(f"+{x}+{y}")           
        self.main_font = font.Font(family="Arial", size=12)
        self.title("Kasa")
        self.resizable(height=False, width=False)

        # DATABASE
        self.conn = sqlite3.connect('database.db')
        self.cursor = self.conn.cursor()

        self.cursor.execute(""" CREATE TABLE IF NOT EXISTS sales (
                                    id INTEGER,
                                    sale REAL NOT NULL,
                                    description TEXT,
                                    date NUMERIC NOT NULL,
                                    time NUMERIC NOT NULL,
                                    PRIMARY KEY(id)
                                )""")

        self.cursor.execute(""" CREATE TABLE IF NOT EXISTS operations (
                                    id INTEGER,
                                    value REAL NOT NULL,
                                    type NOT NULL CHECK(type IN ('KW', 'KP')),
                                    comment TEXT,
                                    date NUMERIC NOT NULL,
                                    time NUMERIC NOT NULL,
                                    PRIMARY KEY(id)
                                )""")

        self.conn.commit()

        # STATIC FRAMES
        upper_left_frame = tk.Frame(master=self, height=350, width=550)
        upper_left_frame.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        upper_center_frame = tk.Frame(master=self, height=200, width=130)
        upper_center_frame.grid(row=0, column=1, sticky="sw", pady=20)
        upper_right_frame = tk.Frame(master=self, height=350, width=400, relief=tk.RIDGE, borderwidth=3)
        upper_right_frame.grid(row=0, column=2, sticky="sw", ipady=10, ipadx=10, padx=20, pady=20)
        lower_left_frame = tk.Frame(master=self, height=200, width=550)
        lower_left_frame.grid(row=2, column=0, pady=20, sticky="swn", padx=20)
        lower_center_frame = tk.Frame(master=self, height=200, width=50)
        lower_center_frame.grid(row=2, column=1, pady=20, sticky="sw")
        lower_right_frame = tk.Frame(master=self, height=200, width=400, relief=tk.RIDGE, borderwidth=3)
        lower_right_frame.grid(row=2, column=2, pady=20, padx=20, sticky="sw", ipady=10, ipadx=10)
        buttons_frame = tk.Frame(master=self, width=60)
        buttons_frame.grid(row=0, column=4, rowspan=3, sticky="ns")
        ttk.Separator(master=self, orient='vertical').grid(row=0, column=3, rowspan=3, sticky="ns")
        ttk.Separator(master=self, orient='horizontal').grid(row=1, column=0, sticky="ew", columnspan=3)

        # STATIC LABELS AND ELEMENTS
        tk.Label(master=upper_left_frame, relief=tk.GROOVE, borderwidth=2,
                 text="Sprzedaż", height=2, font=self.main_font).grid(row=0, column=1, sticky="nsew")
        tk.Label(master=upper_left_frame, relief=tk.GROOVE, borderwidth=2,
                 text="Opis", height=2, font=self.main_font).grid(row=0, column=2, sticky="nsew")
        tk.Label(master=upper_left_frame, relief=tk.GROOVE, borderwidth=2,
                 text="Godzina", height=2, font=self.main_font).grid(row=0, column=3, sticky="nsew")
        tk.Label(master=lower_left_frame, relief=tk.GROOVE, borderwidth=2,
                 text="Wartość", height=2, font=self.main_font).grid(row=0, column=1, sticky="nsew")
        tk.Label(master=lower_left_frame, relief=tk.GROOVE, borderwidth=2,
                 text="Operacja", height=2, font=self.main_font).grid(row=0, column=2, sticky="nsew")
        tk.Label(master=lower_left_frame, relief=tk.GROOVE, borderwidth=2,
                 text="Komentarz", height=2, font=self.main_font).grid(row=0, column=3, sticky="nsew")
        tk.Label(master=lower_left_frame, relief=tk.GROOVE, borderwidth=2,
                 text="Godzina", height=2, font=self.main_font).grid(row=0, column=4, sticky="nsew")
        tk.Label(master=upper_center_frame, text="Utarg", font=self.main_font).grid(row=0, column=0, sticky="sw")
        tk.Label(master=upper_center_frame, text="W miesiącu", font=self.main_font).grid(row=2, column=0, sticky="sw")
        tk.Label(master=lower_center_frame, text="W kasie", font=self.main_font).grid(row=0, column=0, sticky="sw")
        tk.Label(master=upper_right_frame, text="Sprzedaż", font=self.main_font).grid(row=0, column=0, sticky="sw", padx=10)
        tk.Label(master=upper_right_frame, text="Opis", font=self.main_font).grid(row=2, column=0, sticky="sw", padx=10)
        tk.Label(master=lower_right_frame, text="Wartość", font=self.main_font).grid(row=0, column=0,
                                                                columnspan=3, sticky="sw", padx=10)
        tk.Label(master=lower_right_frame, text="Komentarz", font=self.main_font).grid(row=3, column=0,
                                                                  columnspan=3, sticky="sw", padx=10)

        if self.date != self.today:
            widget_state = tk.DISABLED
        else:
            widget_state = tk.NORMAL
            
        # LISTS AND SCROLLS
        self.upper_vertical_scroll = tk.Scrollbar(upper_left_frame, orient=tk.VERTICAL)
        self.upper_vertical_scroll.grid(row=0, column=4, rowspan=2, sticky="ns")
        self.upper_horizontal_scroll = tk.Scrollbar(upper_left_frame, orient=tk.HORIZONTAL)
        self.upper_horizontal_scroll.grid(row=2, column=2, sticky="ew")
        self.upper_delete_list = tk.Listbox(upper_left_frame, height=14, width=3, font=self.main_font,
                                            yscrollcommand=self.upper_vertical_scroll.set,
                                            borderwidth=2, relief=tk.GROOVE, justify=tk.CENTER, fg="red",
                                            highlightthickness=0, selectbackground="SystemButtonFace",
                                            selectforeground="white", state=widget_state)
        self.upper_delete_list.grid(row=1, column=0, sticky="NSEW")
        self.sale_list = tk.Listbox(upper_left_frame, height=14, width=10, relief=tk.GROOVE, borderwidth=2,
                                    yscrollcommand=self.upper_vertical_scroll.set, font=self.main_font)
        self.sale_list.grid(row=1, column=1, sticky="NSEW", )
        self.description_list = tk.Listbox(upper_left_frame, height=14, width=40, relief=tk.GROOVE, borderwidth=2,
                                           yscrollcommand=self.upper_vertical_scroll.set,
                                           xscrollcommand=self.upper_horizontal_scroll.set, font=self.main_font)
        self.description_list.grid(row=1, column=2, sticky="NSEW")
        self.upper_hour_list = tk.Listbox(upper_left_frame, height=14, width=10, relief=tk.GROOVE, borderwidth=2,
                                          yscrollcommand=self.upper_vertical_scroll.set, font=self.main_font)
        self.upper_hour_list.grid(row=1, column=3, sticky="NSEW")

        self.lower_vertical_scroll = tk.Scrollbar(lower_left_frame, orient=tk.VERTICAL)
        self.lower_vertical_scroll.grid(row=0, column=5, rowspan=2, sticky="ns")
        self.lower_horizontal_scroll = tk.Scrollbar(lower_left_frame, orient=tk.HORIZONTAL)
        self.lower_horizontal_scroll.grid(row=2, column=3, sticky="ew")
        self.lower_delete_list = tk.Listbox(lower_left_frame, height=10, width=3, borderwidth=2, font=self.main_font,
                                            yscrollcommand=self.upper_vertical_scroll.set,
                                            relief=tk.GROOVE, justify=tk.CENTER, fg="red",
                                            highlightthickness=0, selectbackground="SystemButtonFace",
                                            selectforeground="white", state=widget_state)
        self.lower_delete_list.grid(row=1, column=0, sticky="NSEW")
        self.value_list = tk.Listbox(lower_left_frame, height=10, width=10, relief=tk.GROOVE, borderwidth=2,
                                     yscrollcommand=self.lower_vertical_scroll.set, font=self.main_font)
        self.value_list.grid(row=1, column=1, sticky="NSEW", )
        self.operation_list = tk.Listbox(lower_left_frame, height=10, width=10, relief=tk.GROOVE, borderwidth=2,
                                         yscrollcommand=self.lower_vertical_scroll.set, font=self.main_font)
        self.operation_list.grid(row=1, column=2, sticky="NSEW")
        self.comment_list = tk.Listbox(lower_left_frame, height=10, width=29, relief=tk.GROOVE, borderwidth=2,
                                       yscrollcommand=self.lower_vertical_scroll.set,
                                       xscrollcommand=self.lower_horizontal_scroll.set, font=self.main_font)
        self.comment_list.grid(row=1, column=3, sticky="NSEW")
        self.lower_hour_list = tk.Listbox(lower_left_frame, height=12, width=10, relief=tk.GROOVE, borderwidth=2,
                                          yscrollcommand=self.lower_vertical_scroll.set, font=self.main_font)
        self.lower_hour_list.grid(row=1, column=4, sticky="NSEW")

        # DYNAMIC LABELS AND ENTRIES
        date_parts = self.date.split('-')
        date_to_display = f"{date_parts[2]}.{date_parts[1]}.{date_parts[0]}"

        self.current_date_lbl = tk.Label(master=self, text=date_to_display, relief=tk.RIDGE, borderwidth=2, font=self.main_font, width=15, height=2)
        self.current_date_lbl.place(relx=0.879, rely=0.0444, anchor=tk.CENTER)
        self.daily_income_lbl = tk.Label(master=upper_center_frame, text="", relief=tk.RIDGE, borderwidth=2, width=13,
                                        font=self.main_font)
        self.daily_income_lbl.grid(row=1, column=0, sticky="sw")

        self.in_month_lbl = tk.Label(master=upper_center_frame, text="", relief=tk.RIDGE, borderwidth=2, width=13,
                                     font=self.main_font)
        self.in_month_lbl.grid(row=3, column=0, sticky="sw")

        self.money_box_lbl = tk.Label(master=lower_center_frame, text="", relief=tk.RIDGE, borderwidth=2, width=13,
                                      font=self.main_font)
        self.money_box_lbl.grid(row=1, column=0, sticky="sw")

        self.sale_entry = tk.Spinbox(master=upper_right_frame, width=10, from_=0, to=100000, increment=0.01,
                                     font=self.main_font, state=widget_state)
        self.sale_entry.grid(row=1, column=0, sticky="sw", padx=10)

        self.description_entry = tk.Entry(master=upper_right_frame, width=35, font=self.main_font, state=widget_state)
        self.description_entry.grid(row=3, column=0, sticky="sw", padx=10)

        self.upper_insert_button = tk.Button(master=upper_right_frame, text="Wprowadź", command=self.submit_upper_form,
                                             font=self.main_font, state=widget_state)
        self.upper_insert_button.grid(row=4, column=0, sticky="sw", padx=10)

        self.value_entry = tk.Spinbox(master=lower_right_frame, width=10, from_=0, to=100000, increment=0.01,
                                      font=self.main_font, state=widget_state)
        self.value_entry.grid(row=1, column=0, columnspan=3, sticky="sw", padx=10)

        self.operation = tk.StringVar()
        self.operation.set('KP')
        self.left_radio = tk.Radiobutton(master=lower_right_frame, text='KP', variable=self.operation, value='KP',
                                         font=self.main_font, state=widget_state)
        self.left_radio.grid(row=2, column=0, sticky="sw", padx=10, pady=10)

        self.right_radio = tk.Radiobutton(master=lower_right_frame, text='KW', variable=self.operation, value='KW',
                                          font=self.main_font, state=widget_state)
        self.right_radio.grid(row=2, column=1, sticky="sw", padx=10, pady=10)

        self.comment_entry = tk.Entry(master=lower_right_frame, width=35, font=self.main_font, state=widget_state)
        self.comment_entry.grid(row=4, column=0, columnspan=3, sticky="sw", padx=10)

        self.lower_insert_button = tk.Button(master=lower_right_frame, text="Wprowadź", command=self.submit_lower_form,
                                             font=self.main_font, state=widget_state)
        self.lower_insert_button.grid(row=5, column=0, columnspan=3, sticky="sw", padx=10)

        calendar_icon = tk.PhotoImage(file="calendar-icon.png")
        self.calendar_button = tk.Button(buttons_frame, image=calendar_icon, command=self.pick_date)
        self.calendar_button.grid(row=0, column=0, padx=4, pady=4, ipadx=2, ipady=2, sticky="swne")

        excel_icon = tk.PhotoImage(file="excel-icon.png")
        self.excel_button = tk.Button(buttons_frame, image=excel_icon, command=self.export_to_excel)
        self.excel_button.grid(row=1, column=0, padx=4, pady=4, ipadx=2, ipady=2, sticky="swne")

        # DELETE CONFIGURATION
        self.upper_delete_list.bind("<ButtonRelease-1>", self.delete_upper_record)
        self.lower_delete_list.bind("<ButtonRelease-1>", self.delete_lower_record)
        
        # SCROLL CONFIGURATIONS
        self.upper_vertical_scroll.config(command=self.upper_scroll_yview)
        self.upper_horizontal_scroll.config(command=self.description_list.xview)
        self.lower_vertical_scroll.config(command=self.lower_scroll_yview)
        self.lower_horizontal_scroll.config(command=self.comment_list.xview)

        # QUIT CONFIGURATION
        self.protocol("WM_DELETE_WINDOW", self.quit)

        # ENTER CONFIGURATION
        self.bind('<Return>', self.enter_clicked)

        # UPDATE UPPER AND LOWER LIST
        self.cursor.execute(""" SELECT sale, description, time FROM sales
                                WHERE date = (?)""", (self.date,))
        records = self.cursor.fetchall()

        for i, record in enumerate(records):
            self.upper_delete_list.insert(i, "❌")
            self.sale_list.insert(i, record[0])
            self.description_list.insert(i, record[1])
            self.upper_hour_list.insert(i, record[2])

        self.cursor.execute(""" SELECT value, type, comment, time FROM operations
                                WHERE date = (?)""", (self.date,))
        records = self.cursor.fetchall()

        for i, record in enumerate(records):
            self.lower_delete_list.insert(i, "❌")
            self.value_list.insert(i, record[0])
            self.operation_list.insert(i, record[1])
            self.comment_list.insert(i, record[2])
            self.lower_hour_list.insert(i, record[3])

        self.update_daily_income()
        self.update_in_month()
        self.update_money_box()

        self.mainloop()

    def upper_scroll_yview(self, *args):
        self.upper_delete_list.yview(*args)
        self.sale_list.yview(*args)
        self.description_list.yview(*args)
        self.upper_hour_list.yview(*args)

    def lower_scroll_yview(self, *args):
        self.lower_delete_list.yview(*args)
        self.value_list.yview(*args)
        self.operation_list.yview(*args)
        self.comment_list.yview(*args)
        self.lower_hour_list.yview(*args)

    def submit_upper_form(self):
        # UPDATE PROGRAM
        try:
            sale = round(float(self.sale_entry.get()), 2)
            description = self.description_entry.get()
            time = datetime.now().strftime("%H:%M")
            if sale <= 0:
                raise ValueError
        except ValueError:
            self.show_message("Błąd sprzedaży", "Wprowadzono nieprawidłową wartość.")
            return
        else:
            if len(description) > 100:
                self.show_message("Błąd sprzedaży", "Wprowadzony opis jest zbyt długi.")
                return
            insert_index = self.sale_list.size()
            self.upper_delete_list.insert(insert_index, "❌")
            self.sale_list.insert(insert_index, sale)
            self.description_list.insert(insert_index, description)
            self.upper_hour_list.insert(insert_index, time)

            self.sale_entry.delete(0, tk.END)
            self.description_entry.delete(0, tk.END)

            # UPDATE DATABASE
            self.cursor.execute(""" INSERT INTO sales (sale, description, date, time)
                                    VALUES (?, ?, ?, ?) """, (sale, description, self.date, time))
            self.conn.commit()

            self.update_daily_income()
            self.update_in_month()
            self.update_money_box()
        

    def submit_lower_form(self):
        # UPDATE PROGRAM
        try:
            value = round(float(self.value_entry.get()), 2)
            operation = self.operation.get()
            comment = self.comment_entry.get()
            time = datetime.now().strftime("%H:%M")
            if value <= 0:
                self.show_message("Błąd operacji", "Wprowadzono nieprawidłową wartość.")
                raise ValueError
            if len(comment) > 100:
                self.show_message("Błąd operacji", "Wprowadzony komentarz jest zbyt długi.")
                raise ValueError
        except ValueError:
            return
        else:
            if operation == "KW":
                if value > self.money_box:
                    self.show_message("Błąd operacji", "Za mało środków w kasie")
                    return
                value *= -1

            insert_index = self.operation_list.size()
            self.lower_delete_list.insert(insert_index, "❌")
            self.value_list.insert(insert_index, value)
            self.operation_list.insert(insert_index, operation)
            self.comment_list.insert(insert_index, comment)
            self.lower_hour_list.insert(insert_index, time)

            self.value_entry.delete(0, tk.END)
            self.comment_entry.delete(0, tk.END)

            self.cursor.execute(""" INSERT INTO operations (value, type, comment, date, time)
                                    VALUES (?, ?, ?, ?, ?) """, (value, operation, comment, self.date, time))
            self.conn.commit()

            self.update_daily_income()
            self.update_in_month()
            self.update_money_box()

    def update_daily_income(self):
        self.cursor.execute(""" SELECT IFNULL(SUM(sale), 0) FROM sales 
                                WHERE date = (?) """, (self.date,))
        self.daily_income = self.cursor.fetchone()[0]
        self.daily_income = round(self.daily_income, 2)
        self.daily_income_lbl['text'] = self.daily_income

    def update_in_month(self):
        current_month = str(self.date)[:-2] + '%'
        self.cursor.execute(""" SELECT IFNULL(SUM(sale), 0) FROM sales
                                WHERE date LIKE (?) AND date <= (?) """, (current_month, self.date))
        self.in_month = self.cursor.fetchone()[0]
        self.in_month = round(self.in_month, 2)
        self.in_month_lbl['text'] = self.in_month

    def update_money_box(self):
        self.cursor.execute(""" SELECT IFNULL(SUM(sale), 0) FROM sales 
                                WHERE date <= (?) """, (self.date,))
        sales_balance = self.cursor.fetchone()[0]

        self.cursor.execute(""" SELECT IFNULL(SUM(value), 0) FROM operations
                                WHERE date <= (?) """, (self.date,))
        operations_balance = self.cursor.fetchone()[0]

        self.money_box = sales_balance + operations_balance
        self.money_box = round(self.money_box, 2)
        self.money_box_lbl['text'] = self.money_box

    def show_message(self, title, message):
        messagebox.showerror(title, message)

    def delete_upper_record(self, event):
        selected_index = self.upper_delete_list.curselection()
        self.upper_delete_list.select_clear(0, 'end')
        self.focus_set()
        result = messagebox.askquestion("Usuń rekord",
                                        f"Czy na pewno chcesz usunąć sprzedaż nr. {selected_index[0] + 1}?", default='no')
        if result == 'no':
            return

        sale = self.sale_list.get(selected_index)
        description = self.description_list.get(selected_index)
        time = self.upper_hour_list.get(selected_index)

        self.upper_delete_list.delete(selected_index)
        self.sale_list.delete(selected_index)
        self.description_list.delete(selected_index)
        self.upper_hour_list.delete(selected_index)

        self.cursor.execute(""" DELETE FROM sales
                                WHERE sale = (?)
                                AND description = (?)
                                AND time = (?)
                                AND date = (?) """, (sale, description, time, self.date))
        self.conn.commit()

        self.update_daily_income()
        self.update_in_month()
        self.update_money_box()

    def change_focus(self, event):
        self.focus_set()

    def delete_lower_record(self, event):
        selected_index = self.lower_delete_list.curselection()
        self.lower_delete_list.select_clear(0, 'end')
        self.focus_set()
        result = messagebox.askquestion("Usuń rekord",
                                        f"Czy na pewno chcesz usunąć operacje nr. {selected_index[0] + 1}?", default='no')

        if result == 'no':
            return

        value = self.value_list.get(selected_index)
        operation = self.operation_list.get(selected_index)
        comment = self.comment_list.get(selected_index)
        time = self.lower_hour_list.get(selected_index)

        self.lower_delete_list.delete(selected_index)
        self.value_list.delete(selected_index)
        self.operation_list.delete(selected_index)
        self.comment_list.delete(selected_index)
        self.lower_hour_list.delete(selected_index)

        self.cursor.execute(""" DELETE FROM operations
                                WHERE value = (?)
                                AND type = (?)
                                AND comment = (?)
                                AND time = (?)
                                AND date = (?) """, (value, operation, comment, time, self.date))
        self.conn.commit()

        self.update_daily_income()
        self.update_in_month()
        self.update_money_box()

    def pick_date(self):
        date_window = tk.Toplevel()
        date_window.withdraw()

        self.update()
        x = self.winfo_x() + int(self.winfo_width() / 2 - 340 / 2)
        y = self.winfo_y() + int(self.winfo_height() / 2 - 320 / 2)

        date_window.geometry(f'340x320+{x}+{y}')
        date_window.resizable(height=False, width=False)
        date_window.title('Wybierz dzień')
        date_window.grab_set()

        self.cursor.execute(""" SELECT MIN(date) FROM sales """)
        check_date = self.cursor.fetchone()[0]
        if check_date:
            min_date = dt.fromisoformat(check_date)
        else:
            min_date = dt.today()
        max_date = dt.today()

        year, month, day = str(self.date).split('-')

        calendar = Calendar(date_window, year=int(year), month=int(month), day=int(day), selectmode='day', date_pattern='y-mm-dd', font=self.main_font, mindate=min_date, maxdate=max_date)
        calendar.place(relx=0.5, rely=0.38, anchor=tk.CENTER)

        submit_button = tk.Button(date_window, text="Przejdź", font=self.main_font, width=10, command=lambda: self.change_date(calendar.get_date()))
        submit_button.place(relx=0.5, rely=0.79, anchor=tk.CENTER)

        today_button = tk.Button(date_window, text="Dzisiaj", font=self.main_font, width=10, command=lambda: self.change_date(self.today))
        today_button.place(relx=0.5, rely=0.92, anchor=tk.CENTER)

        date_window.deiconify()

    def change_date(self, date):
        self.update()
        x = self.winfo_x()
        y = self.winfo_y()

        self.conn.commit()
        self.conn.close()

        self.destroy()
        App(date, x, y)

    def export_to_excel(self):
        
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[(".xlsx", "*.xlsx")], initialfile=self.date)
        if file_path:
            self.cursor.execute(f""" SELECT sale AS 'Sprzedaż', description AS 'Opis', time AS 'Godzina' FROM sales
                                     WHERE date = '{self.date}' """)
            set1 = self.cursor.fetchall()
            self.cursor.execute(f""" SELECT value AS 'Wartość', type AS 'Operacja', comment AS 'Komentarz', time AS 'Godzina' FROM operations
                                     WHERE date = '{self.date}' """)
            set2 = self.cursor.fetchall()

            workbook = xlsxwriter.Workbook(file_path)
            worksheet_data = workbook.add_worksheet('worksheet')

            headers_format = workbook.add_format(
                {   
                    "bg_color": "#AFAFAF",
                    "border": 1,
                    "border_color": "#000000",
                }
            )
            headers_format.set_bold()

            data_format = workbook.add_format(
                {
                    "border": 1,
                    "border_color": "#000000"
                }
            )

            worksheet_data.write("A1", "Sprzedaż", headers_format)
            worksheet_data.write("B1", "Opis", headers_format)
            worksheet_data.write("C1", "Godzina", headers_format)
            worksheet_data.write("E1", "Wartość", headers_format)
            worksheet_data.write("F1", "Operacja", headers_format)
            worksheet_data.write("G1", "Komentarz", headers_format)
            worksheet_data.write("H1", "Godzina", headers_format)
            worksheet_data.write("J1", "Utarg", headers_format)
            worksheet_data.write("J2", self.daily_income, data_format)
            worksheet_data.write("K1", "W miesiącu", headers_format)
            worksheet_data.write("K2", self.in_month, data_format)
            worksheet_data.write("L1", "W kasie", headers_format)
            worksheet_data.write("L2", self.money_box, data_format)

            for i, record in enumerate(set1):
                worksheet_data.write(i+1, 0, record[0], data_format)
                worksheet_data.write(i+1, 1, record[1], data_format)
                worksheet_data.write(i+1, 2, record[2], data_format)

            for i, record in enumerate(set2):
                worksheet_data.write(i+1, 4, record[0], data_format)
                worksheet_data.write(i+1, 5, record[1], data_format)
                worksheet_data.write(i+1, 6, record[2], data_format)
                worksheet_data.write(i+1, 7, record[3], data_format)

            worksheet_data.set_column("A:A", 10)
            worksheet_data.set_column("B:B", 30)
            worksheet_data.set_column("C:C", 7)

            worksheet_data.set_column("E:E", 10)
            worksheet_data.set_column("F:F", 8)
            worksheet_data.set_column("G:G", 30)
            worksheet_data.set_column("H:H", 7)

            worksheet_data.set_column("J:J", 12)
            worksheet_data.set_column("K:K", 12)
            worksheet_data.set_column("L:L", 12)
            workbook.close()

    def enter_clicked(self, event):
        
        if self.sale_entry.get() != '' and self.sale_entry.get() != '0.00':
            self.submit_upper_form()
        if self.value_entry.get() != '' and self.value_entry.get() != '0.00':
            self.submit_lower_form()

        return

    def quit(self):
        if messagebox.askokcancel("Zamknij aplikację", "Czy na pewno chcesz zamknąć aplikację?"):
            self.conn.commit()
            self.conn.close()
            self.destroy()

if __name__ == "__main__":
    executable_path = sys.argv[0]
    executable_directory = os.path.abspath(os.path.dirname(executable_path))
    os.chdir(executable_directory)

    monitor = get_monitors()[0]
    x = int(monitor.width / 2 - 1200 / 2)
    y = int(monitor.height / 2 - 750 / 2)

    date = str(dt.today())
    app = App(date, x, y)


