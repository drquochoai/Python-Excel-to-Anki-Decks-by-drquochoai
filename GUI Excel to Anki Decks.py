# Python program to create
# a file explorer in Tkinter
# pyinstaller --onefile --windowed --icon=assets/app.ico "GUI Excel to Anki Decks.py"
# Excel to Anki package by drquochoai v2.1

from tkinter import *
import tkinter as tk
from pathlib import Path
import os
from LIBhoaiAnki import Cloze;
import xlrd
import subprocess
import webbrowser

FILEBROWSER_PATH = os.path.join(os.getenv('WINDIR'), 'explorer.exe')
urlHuongdan = 'https://www.notion.so/hoaiump/Excel-to-anki-T-i-v-Ph-n-m-m-h-ng-d-n-0c40e2f0c27648359d606530b460b452'
def explore(path):
    # explorer would choke on forward slashes
    path = os.path.normpath(path)
    if os.path.isdir(path):
        subprocess.run([FILEBROWSER_PATH, path])
    elif os.path.isfile(path):
        subprocess.run([FILEBROWSER_PATH, '/select,', path])


# import filedialog module
from tkinter import filedialog
# Create the root window
window = Tk()
# Set window title
window.title('Excel to Anki package by drquochoai v3.1 20/05/2021')
w, h = window.winfo_screenwidth(), window.winfo_screenheight()
window.geometry("%dx%d+%d+%d" % (w/2, h/2, w/4, h/6))
# Set window size
# Set window background color
window.config(background="white")

# LOGIC CÁC HÀM
book = any
ankiName = any
hoaianki = any
dirPath = '/'
def browseFiles():
    global book
    global ankiName
    global hoaianki
    global dirPath
    filePath = filedialog.askopenfilename(initialdir= dirPath, title="Select a File", filetypes=(("Excel 97-2003 .XLS","*.xls"), ("không thể chọn khác", "*.drquochoai")))
    # Change label contents
    if filePath != "":
        fileName = Path(filePath).stem
        label_file_explorer.configure(text=fileName)
        ankiName = fileName
        dirPath = os.path.dirname(filePath)
        extension = os.path.splitext(filePath)[1]
        # print(extension) #.xls
        Lb1.delete(0, END)
        # Chọn tất cả 
        if extension == ".xls":
            #Tạo class Anki mới để chơi
            book = xlrd.open_workbook(filePath.strip("‪u202a"), formatting_info=True)
            for nsheet in range(0, book.nsheets):
                sheet = book.sheet_by_index(nsheet)
                if len(Cloze().my_model.fields) - 1 != sheet.ncols:
                    print('Lỗi: "'+ sheet.name+'" không đủ số cột theo yêu cầu → bỏ qua')
                    continue
                else:
                    Lb1.insert(nsheet, sheet.name)
            cb_selectall()
# Checkbox cb_selectall
def cb_selectall():
    if chontatca.get() == 1:
        Lb1.select_set(0, END)
    else:
        Lb1.select_clear(0, END)
# Checkbox function open folder when finish
def openWhenFinish():
    global dirPath
    global ankiName
    global openFinished
    if openFinished.get() == 1:
        explore(dirPath +"/"+ ankiName.replace(" ", "_")+".apkg")

#Xử lý và tạo anki package
def runProcessExcel():
    global ankiName
    global hoaianki
    global dirPath
    hoaianki = Cloze()
    hoaianki.resetDeck()
    for i in Lb1.curselection():
        #print(Lb1.get(i)) → Lb1.get(i) = name;  book.sheet_by_name
        sheet = book.sheet_by_name(Lb1.get(i))
        #Check nếu sheet có số cột bằng với số fields - id = 5 cột
        #print(hoaianki.my_model.fields)
        # Khi đúng là Sheet anki với đủ 5 cột, sẽ tạo deck cho sheet đó
        deckName = ankiName +"::"+ sheet.name
        hoaianki.createDeck(deckName=deckName, mota="")
        currentTitle = ""
        currentFooter = ""
        # Khi đủ yêu cầu chúng ta sẽ loop quanh từng cột trong hàng
        for row_idx in range(sheet.nrows):
            if row_idx == 0:
                continue

            noteValues = []
            # Tạo ID cho hàng đó = tên file + sheet + số hàng
            guid = deckName + "::" + str(row_idx+1)
            noteValues.append(guid)

            for col_idx in range(sheet.ncols):
                text_cell = sheet.cell_value(row_idx, col_idx)
                text_cell_xf = book.xf_list[sheet.cell_xf_index(row_idx, col_idx)]
                # print(text_cell)
                
                # skip rows where cell is empty
                # if not text_cell:
                #     noteValues.append("")
                #     continue
                
                text_cell_runlist = sheet.rich_text_runlist_map.get((row_idx, col_idx))
                htmlValue = hoaianki.htmlProcess(text_cell, text_cell_xf, text_cell_runlist, book.font_list, col_idx)
                # Copy title và footer nếu blank
                if col_idx == 2:
                    # currentTitle
                    if htmlValue != "":
                        currentTitle = htmlValue
                    noteValues.append(currentTitle)
                elif col_idx == 3:
                    if htmlValue != "":
                        currentFooter = htmlValue
                    noteValues.append(currentFooter)
                else:
                    noteValues.append(htmlValue)
                # print(noteValues)
            #  Thêm note cho Deck của sheet hiện tại
            hoaianki.addNote(noteValues, guid)
    # print(dirPath)
    hoaianki.saveAnkiPackage(dirPath +"/"+ ankiName.replace(" ", "_"))
    openWhenFinish()
def openHuongDan():
    webbrowser.open(urlHuongdan)
# Create a File Explorer label
label_file_explorer = Label(window,
                            text="Bấm nút phía trên để chọn file excel",
                            fg="blue")
button_explore = Button(window,
                        text="Browse Files (Chọn file excel)",
                        command=browseFiles, background="orange", fg="white")
button_run_ankiProccess = Button(window,
                                 text="Create Anki Package (Tạo Anki)",
                                 command=runProcessExcel, background="orange", fg="white")
label_tacgia = Label(window,
                            text="Author: Trần Quốc Hoài",
                            fg="blue")
button_huongdan = Button(window,
                        text="Instructions (hướng dẫn sử dụng)",
                        command=openHuongDan, background="green", fg="white")
def exit():
    window.quit()
button_exit = Button(window,
                     text="Exit",
                     command=exit)
# Grid method is chosen for placing
# the widgets at respective positions
# in a table like structure by
# specifying rows and columns
button_explore.grid(column=1, row=1)
label_file_explorer.grid(column=1, row=2)
button_run_ankiProccess.grid(column=1, row=50)
label_tacgia.grid(column=1, row=99)
button_huongdan.grid(column=2, row=99)
button_exit.grid(column=3, row=99)
# Let the window wait for any events

labelframe = LabelFrame(
    window, text="(Select Sheets to export (chọn sheets để xuất)")
labelframe.grid(column=1, row=20)
Lb1 = Listbox(labelframe, selectmode="multiple", height=10, width=50)
Lb1.pack()

# check All
chontatca = tk.IntVar(value=1)
c_checkall = Checkbutton(labelframe, text='All (tất cả)', variable=chontatca,
                 onvalue=1, offvalue=0, command=cb_selectall)
c_checkall.pack()

# check All
openFinished = tk.IntVar(value=1)
c_openFinished = Checkbutton(labelframe, text='Mở thư mục khi xong', variable=openFinished,
                 onvalue=1, offvalue=0, command=openWhenFinish)
c_openFinished.pack()
# window.eval('tk::PlaceWindow %s center' % window.winfo_toplevel())
window.mainloop()

# Function for opening the
# file explorer window
