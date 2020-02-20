from tksheet import Sheet
from tkinter import *
from tkinter.scrolledtext import ScrolledText
from datetime import datetime
import gspread, re, os
from time import sleep
from oauth2client.service_account import ServiceAccountCredentials
from httplib2 import ServerNotFoundError
from gspread.exceptions import SpreadsheetNotFound
from pathlib import Path
from threading import Thread
from time import time


class Application:
    def __init__(self, master):
        self.master = master
        self.now = datetime.now()
        # self.master.wm_iconbitmap("logo.ico")
        self.master.eval('tk::PlaceWindow . center')
        self.master.minsize(400, 300)
        self.master.grid_columnconfigure(0, weight = 1)
        self.master.grid_rowconfigure(0, weight = 1)
        style = ttk.Style(self.master)
        style.configure("Placeholder.TEntry", foreground="#d5d5d5")
        self.master.title("Query thông tin sản phẩm")
        self.savedKw = str(Path.home()) + os.path.sep + "savedkw.txt"
        self.master_data = {}
        self.selected_sheets_data = []
        self.data_to_fill = []
        self.menubar_items = ["File", "Xem", "Công cụ", "Giúp đỡ"]
        self.sheet = None
        self.gspread_sheet = None
        self.url_input = None
        self.sheet_list = []
        self.sheet_titles = []
        self.highlighted_index = []
        # self.url_gsheet = "https://docs.google.com/spreadsheets/d/10oJamLk0Bj4ffcDbnu9-96-dn7Tf7TM0EnJ2-emSp9c/edit#gid=1900586655"
        self.url_gsheet = "Reason Code CB"
        self.error_load_sheet = False
        self.current_words = []
        self.toggle_theme = IntVar()
        self.toggle_theme.set(0)
        self.toggle_compact = IntVar()
        self.toggle_compact.set(0)
        self.logging = []
        self.current_log = []
        self.area_filter_value = IntVar()
        self.sheet_filter_value = IntVar()
        self.create_widgets()            

    def get_data_to_fill(self, selected_sheets):
        for selected_sheet in selected_sheets:
            self.selected_sheets_data.append(self.master_data[selected_sheet])
            self.sheet_titles.append(self.master_data[selected_sheet][0])
        for each_sheet in self.selected_sheets_data:
            self.data_to_fill = self.data_to_fill + each_sheet

    def get_gsheet_value(self, gsheet_url):
        scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
        # creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        self.data_to_fill = []
        self.error_load_sheet = False
        try:
            client = gspread.authorize(creds)
            if "http" in gsheet_url:
                self.gspread_sheet = client.open_by_url(gsheet_url)
            else:
                self.gspread_sheet = client.open(gsheet_url)
            worksheets = self.gspread_sheet.worksheets()
            for worksheet in worksheets:
                worksheet_name = worksheet.title
                worksheet_values = self.gspread_sheet.worksheet(worksheet_name).get_all_values()
                self.sheet_list.append(worksheet_name)
                self.master_data[worksheet_name] = worksheet_values
            self.get_data_to_fill(self.sheet_list)
        except (ServerNotFoundError, Exception) as e:
            self.error_load_sheet = True
            self.cell_value_label.configure(text="Có lỗi xảy ra, hãy thử lại", anchor="w")
            self.data_to_fill = [[]]
        except SpreadsheetNotFound as notfound:
            self.error_load_sheet = True
            self.cell_value_label.configure(text="File spreadsheet không tồn tại", anchor="w")
            self.data_to_fill = [[]]
            
    class AutocompleteEntry(ttk.Entry):
        def __init__(self, autocompleteList, placeholder, *args, **kwargs):

            # Custom matches function
            if 'matchesFunction' in kwargs:
                self.matchesFunction = kwargs['matchesFunction']
                del kwargs['matchesFunction']
            else:
                def matches(fieldValue, acListEntry):
                    pattern = re.compile('.*' + re.escape(fieldValue) + '.*', re.IGNORECASE)
                    return re.match(pattern, acListEntry)
                    
                self.matchesFunction = matches

            ttk.Entry.__init__(self, *args, style="Placeholder.TEntry", **kwargs)
            self.autocompleteList = autocompleteList
            self.placeholder = placeholder
            self.focus()
            self.insert("0", self.placeholder)
            self.bind("<FocusIn>", self.clear_placeholder)
            self.bind("<FocusOut>", self.add_placeholder)
            
            self.var = self["textvariable"]
            if self.var == '':
                self.var = self["textvariable"] = StringVar()

            self.var.trace('w', self.changed)
            self.bind("<Right>", self.selection)
            self.bind("<Up>", self.moveUp)
            self.bind("<Down>", self.moveDown)
            self.bind("<Escape>", self.hide)

            self.listboxUp = False

        def clear_placeholder(self, e):
            if self["style"] == "Placeholder.TEntry":
                self.delete("0", "end")
                self["style"] = "TEntry"

        def add_placeholder(self, e):
            if not self.get():
                self.insert("0", self.placeholder)
                self["style"] = "Placeholder.TEntry"

        def changed(self, name, index, mode):
            if self.var.get() == '':
                if self.listboxUp:
                    self.listbox.destroy()
                    self.listboxUp = False
            else:
                words = self.comparison()
                self.listboxLength = len(words)
                if words:
                    if not self.listboxUp:
                        self.listbox = Listbox(width=self["width"], height=self.listboxLength)
                        self.listbox.bind("<<ListboxSelect>>", self.clickItem)
                        self.listbox.bind("<Right>", self.selection)
                        self.listbox.place(x=self.winfo_x(), y=self.winfo_y() + self.winfo_height())
                        self.listboxUp = True
                    
                    self.listbox.delete(0, END)
                    for w in words:
                        self.listbox.insert(END,w)
                else:
                    if self.listboxUp:
                        self.listbox.destroy()
                        self.listboxUp = False
        
        def clickItem(self, event):
            if self.listboxUp:
                index = int(self.listbox.curselection()[0])
                self.var.set(self.listbox.get(index))
                self.listbox.destroy()
                self.listboxUp = False
                self.icursor(END)

        def selection(self, event):
            if self.listboxUp and self.listbox.curselection() != ():
                self.var.set(self.listbox.get(ACTIVE))
                self.listbox.destroy()
                self.listboxUp = False
                self.icursor(END)
            else:
                if self.listboxUp:
                    self.listbox.destroy()
                self.listboxUp = False
                self.icursor(END)

        def hide(self, event):
            if self.listboxUp:
                self.listbox.destroy()
                self.listboxUp = False

        def moveUp(self, event):
            if self.listboxUp:
                if self.listbox.curselection() == ():
                    index = '-2'
                else:
                    index = self.listbox.curselection()[-1]
                    
                if index != END:                
                    self.listbox.selection_clear(first=index)
                    index = str(int(index) - 1)
                    
                    self.listbox.see(index) # Scroll!
                    self.listbox.selection_set(first=index)
                    self.listbox.activate(index)

        def moveDown(self, event):
            if self.listboxUp:
                if self.listbox.curselection() == ():
                    index = '-1'
                else:
                    index = self.listbox.curselection()[-1]
                    
                if index != END:                        
                    self.listbox.selection_clear(first=index)
                    index = str(int(index) + 1)
                    
                    self.listbox.see(index) # Scroll!
                    self.listbox.selection_set(first=index)
                    self.listbox.activate(index)

        def comparison(self):
            return [ w for w in self.autocompleteList if self.matchesFunction(self.var.get(), w) ]

    def get_saved_keywords(self):
        try:
            with open(self.savedKw, "r", encoding="utf-8") as f:
                lines = f.readlines()
                for line in lines:
                    self.current_words.append(line.replace("\n",""))
        except FileNotFoundError as fe:
            pass

    def save_keywords(self, keyword):
        if keyword not in self.current_words and keyword != "":
            with open(self.savedKw, "a", encoding="utf-8") as f:
                f.write(f"{keyword}\n")

    def search_keyword(self, event):
        self.search_bar.selection(event)
        kw = self.search_bar.get()
        result = []
        titles = []
        for each_sheet in self.selected_sheets_data:
            search_result = [row for row in each_sheet if kw.lower() in str(row).lower()]
            if len(search_result) > 0:
                search_result.insert(0, each_sheet[0])
                titles.append(each_sheet[0])
            # search_result = list(set(search_result))
            result = result + search_result
        self.sheet.set_sheet_data(result, reset_col_positions=False)
        self.dehighlight_current_titles()
        self.highlight_sheet_title(result, titles)
        if kw != "":
            self.save_keywords(kw)
            log_output = "Đã tìm từ khóa: {}\n".format(kw)
            self.logging.append(log_output)
            if len(result) > 0 and kw not in self.current_words:
                self.current_words.append(kw)

    def highlight_sheet_title(self, current_data, titles):
        for index, value in enumerate(current_data):
            for i in titles:
                if i == value:
                    for i in range(len(i)):
                        self.sheet.highlight_cells(row=index, column=i, bg="#ed4337", fg="white")
            self.highlighted_index.append(index)
        self.sheet.refresh()

    def dehighlight_current_titles(self, indexes):
        for r in self.highlighted_index:
            for c in range(50):
                self.sheet.dehighlight_cells(row=r, column=c)
        self.sheet.refresh()

    def load_data(self):
        self.sheet.set_sheet_data([[]], reset_col_positions=True)
        for i in self.menubar_items:
            self.menubar.entryconfig(i, state="disabled")
        self.cell_value_label.configure(text="Đang tải dữ liệu file ...", anchor="w")
        self.get_gsheet_value(self.url_gsheet)
        self.toggle_compact.set(0)
        self.sheet.set_sheet_data(self.data_to_fill, reset_col_positions=True)
        if not self.error_load_sheet:
            self.cell_value_label.configure(text="---", anchor="w")
        for i in self.menubar_items:
            self.menubar.entryconfig(i, state="normal")
        self.highlight_sheet_title(self.data_to_fill, self.sheet_titles)

    def load_data_in_thread(self):
        t = Thread(target=self.load_data)
        t.start()

    def cell_select(self, response):
        self.cell_value_label.config(text=self.sheet.get_cell_data(response[1], response[2]))

    def load_another_gsheet(self):
        self.load_new_wd = Toplevel()
        self.load_new_wd.title("Nhập tên hoặc url của spreadsheet")
        self.load_new_wd.grab_set()
        self.load_new_wd.resizable(False, False)
        self.load_new_wd.grid_columnconfigure(0, weight = 1)
        self.load_new_wd.grid_rowconfigure(0, weight = 1)
        sub_frame = Frame(self.load_new_wd)
        sub_frame.grid(row = 0, column = 0, sticky = "nsew", padx=1, pady=1)

        self.url_input = ttk.Entry(sub_frame, width=100)
        self.url_input.focus()
        self.url_input.grid(row=0, column=0, ipady=5, sticky="we", padx=5, pady=5)
        self.url_input.bind("<Return>", self.get_new_gsheet_data)

        sub_frame.grid_columnconfigure(0, weight = 1)
        sub_frame.grid_rowconfigure(1, weight = 1)     

    def get_new_gsheet_data(self, event):
        if self.url_gsheet.strip() != "":
            self.url_gsheet = self.url_input.get()
            self.load_new_wd.destroy()
            self.load_data_in_thread()

    def switch_theme(self):
        mode = self.toggle_theme.get()
        if mode == 0:
            self.sheet.change_theme("light")
        else:
            self.sheet.change_theme("dark")

    def switch_compact(self):
        mode = self.toggle_compact.get()
        len_shd = len(self.data_to_fill)
        if mode == 0:
            self.sheet.set_column_widths([120 for c in range(len_shd)])
        else:
            self.sheet.set_column_widths([30 for c in range(len_shd)])
        self.sheet.refresh()

    def filter_sheet(self):
        self.filter_wd = Toplevel()
        self.filter_wd.title("Filter dữ liệu")
        self.filter_wd.maxsize(400, 200)
        self.filter_wd.grab_set()
        self.filter_wd.grid_columnconfigure(0, weight = 1)
        self.filter_wd.grid_rowconfigure(0, weight = 1)
        nb = ttk.Notebook(self.filter_wd)
        nb.grid(row=0, column=0, sticky="nswe")
        area_filter = Frame(self.filter_wd, padx=1, pady=1)
        sheet_filter = Frame(self.filter_wd, padx=1, pady=1)
        area_filter.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        sheet_filter.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        
        nb.add(area_filter, text="Filter theo vùng")
        nb.add(sheet_filter, text="Filter theo sheet")
        
        radiobtn1 = Radiobutton(area_filter, text="Shopee C2C", variable=self.area_filter_value, value=1)
        radiobtn1.grid(row=0, column=0)
        radiobtn2 = Radiobutton(area_filter, text="Shopee Mall", variable=self.area_filter_value, value=2)
        radiobtn2.grid(row=1, column=0)
        radiobtn3 = Radiobutton(area_filter, text="Note", variable=self.area_filter_value, value=3)
        radiobtn3.grid(row=1, column=0)
        for sheet_name in self.sheet_list:
            ttk.Checkbutton(sheet_filter, text=sheet_name).grid(sticky="w")

        nb.grid_columnconfigure(0, weight = 1)
        nb.grid_rowconfigure(0, weight = 1)

        sheet_filter.grid_columnconfigure(0, weight = 1)
        sheet_filter.grid_rowconfigure(0, weight = 1)


    # def get_logging(self):
    #     if len(self.current_log) < len(self.logging):
    #         self.log_entry.config(state=NORMAL)
    #         self.log_entry.insert(INSERT, self.logging[len(self.logging) - 1])
    #         self.log_entry.see(END)
    #         self.log_entry.config(state=DISABLED)
    #         self.current_log.append(self.logging[len(self.logging) - 1])
    #     self.log_entry.after(1000, self.get_logging)

    def open_console(self):
        self.console_wd = Toplevel()
        self.console_wd.title("Console log")
        # self.console_wd.grab_set()
        self.console_wd.minsize(500, 200)
        self.console_wd.resizable(False, False)
        self.console_wd.grid_columnconfigure(0, weight = 1)
        self.console_wd.grid_rowconfigure(0, weight = 1)
        sub_frame = Frame(self.console_wd)
        sub_frame.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        self.log_entry = ScrolledText(sub_frame)
        self.log_entry.grid(row=0, column=0, sticky="nswe", padx=5, pady=5)
        self.log_entry.insert(INSERT, "Console log của app được mở vào lúc {}\n\n".format(self.now))
        # self.get_logging()
        sub_frame.grid_columnconfigure(0, weight = 1)
        sub_frame.grid_rowconfigure(0, weight = 1)

    def create_menu_bar(self):
        self.menubar = Menu(self.master)

        filemenu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label=self.menubar_items[0], menu=filemenu)
        filemenu.add_command(label="Reload dữ liệu", command=self.load_data_in_thread)
        filemenu.add_command(label="Thoát", command=self.master.quit)

        viewmenu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label=self.menubar_items[1], menu=viewmenu)
        viewmenu.add_command(label="Filter dữ liệu", command=self.filter_sheet)
        viewmenu.add_checkbutton(label="Đổi theme trắng/đen", variable=self.toggle_theme, command=self.switch_theme)
        viewmenu.add_checkbutton(label="Chế độ compact", variable=self.toggle_compact, command=self.switch_compact)

        toolmenu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label=self.menubar_items[2], menu=toolmenu)
        toolmenu.add_command(label="Load dữ liệu khác", command=self.load_another_gsheet)

        helpmenu = Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label=self.menubar_items[3], menu=helpmenu)
        helpmenu.add_command(label="Console log", command=self.open_console)

        self.master.config(menu=self.menubar)

    def create_widgets(self):
        self.get_saved_keywords()

        #--------CREATE MENU BAR----------
        self.create_menu_bar()

        #--------CREATE NAVIGATION FRAME----------
        self.main_frame = Frame(self.master)
        self.main_frame.grid(row = 0, column = 0, sticky = "nsew", padx=1, pady=1)

        #--------CREATE SEARCH BAR AND SEARCH BUTTON WIDGET----------
        self.search_bar = self.AutocompleteEntry(self.current_words, "Nhập từ khóa cần tìm", self.main_frame)
        self.search_bar.grid(row=0, column=0, ipady=5, sticky="we", padx=5, pady=5)
        self.search_bar.bind("<Return>", self.search_keyword)

        #--------CREATE CELL DETAIL VALUE LABEL----------
        self.cell_value_label = Label(self.main_frame, text="---", anchor="w")
        self.cell_value_label.grid(row=1, column=0, padx=5, sticky="nw")

        self.main_frame.grid_columnconfigure(0, weight = 1)
        self.main_frame.grid_rowconfigure(2, weight = 1)

        #--------CREATE SHEET WIDGET----------
        self.sheet = Sheet(self.main_frame, align = "w")
        self.sheet.enable_bindings(("single_select","drag_select",
                            "column_select","row_select","column_width_resize",
                            "double_click_column_resize","row_width_resize","column_height_resize",
                            "arrowkeys","row_height_resize","double_click_row_resize","rc_select","copy","paste","undo"))
        self.sheet.extra_bindings([("cell_select", self.cell_select)])
        self.sheet.grid(row=2, column = 0, sticky = "nswe", padx=5, pady=5)
        self.load_data_in_thread()


if __name__ == '__main__':
    root = Tk()
    app = Application(root)
    root.mainloop()
