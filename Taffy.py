## to-do:
##1: Check if file is open before running Taffy and prompt 'do you want to close' or add + _1 to name and allow both to be open
##3: Check if file is truly a CV or LSV
##4: Create an error log
##5: Create avg_generator and append the data to graph
##6 create CV Generator and Combined Avgs Generator

import xlwings as xw
#makes the UI less blurry
from ctypes import windll
windll.shcore.SetProcessDpiAwareness(1)
#tkinter is used for the UI
import tkinter as tk
#tkk stands for 'themed tkinter' and allows for more UI customization
import tkinter.ttk as ttk
from tkinter.filedialog import askopenfile
#importing os is required for opening the excel file
import os


# Select the same worksheet

#used to get functions from the genData python file
from genData import get_data, getTafel
from appendData import append_data
#threading is used to run the loading bar and the 'generate sheets' functions at the same time.
import threading


#creates the UI window
root = tk.Tk()
#specifies the window dimensions
root.geometry('1500x400')
#app background color
root.configure(bg='#d9d9d9')
#window icon and title
root.iconbitmap("imgs/small_taffy.ico")
root.title('Taffy 1.5')
#selects the 'alt' theme from tkinter
themey = ttk.Style()
themey.theme_names()
themey.theme_use('alt')

#i forgot what this is for
tk.Label(text=" ", background='#d9d9d9').grid(column=12, row=0, pady=3)

#sets the height of 9 rows so that the taffy image doesn't skew the UI
for i in range(9):
    tk.Label(text=" ", background='#d9d9d9').grid(column=0, row=i, pady=5)


#gets and places logo in the first visible column
logoimg = tk.PhotoImage(file="imgs/taffy.gif")
logolabel = ttk.Label(root, image=logoimg)
logolabel.grid(row=2, column=1, rowspan=3)
# removes the border around the image
themey.configure("TLabel", borderwidth=0)

#hides the columns with the inputs and shows the shrug emoticon when CVs or Combine Avgs. radiobuttons are selected
def hide_columns(m):
    global red_wip
    for widget in root.grid_slaves():
        if int(widget.grid_info()["column"]) >= 2:
            widget.grid_forget()
    if m == 'add':
      tk.Label(root, text=f'¯\_(ツ)_/¯', font=('calibri', 30, 'bold'), background='#d9d9d9').grid(row=0, rowspan=5, column=2, padx=300)


#creates radio buttons
rb_var = tk.IntVar()
rb_var.set(1)
rb1 = tk.Radiobutton(root, text="LSVs", variable=rb_var, value=1, command=lambda: [hide_columns('destroy'), create_textboxes()], background='#d9d9d9', padx=10)
rb2 = tk.Radiobutton(root, text="CVs", variable=rb_var, value=2, command=lambda: [hide_columns('add')], background='#d9d9d9', padx=10)
rb3 = tk.Radiobutton(root, text="Combine Avgs.", variable=rb_var, command=lambda: [hide_columns('add')], value=3, background='#d9d9d9', padx=10)
rb1.grid(row=5, column=1, sticky='W')
rb2.grid(row=6, column=1, sticky='W')
rb3.grid(row=7, column=1, sticky='W')



def subscr(num):
    SUB = str.maketrans("0123456789", "₀₁₂₃₄₅₆₇₈₉")
    num = str(num)
    num = num.translate(SUB)
    return num


# Labels and checkboxes
def create_textboxes():
    global textboxes
    labels = ['RPM', 'Electrode', 'Molarity', 'Temp. (°C):', 'Electrolyte', 'Expt. #']
    for i in range(len(labels)):
        tk.Label(root, text=labels[i], background='#d9d9d9').grid(row=0, column=i+2)
    checkbox_names = ['RPM_check', 'Etrode_check', 'M_check', 'Temp_check', 'Elyte_check']
    checkboxes = []
    checkbox_vars = []
    for i in range(len(checkbox_names)):
        checkbox_vars.append(tk.IntVar())
        checkboxes.append(tk.Checkbutton(root, variable=checkbox_vars[i], background='#d9d9d9'))
        checkboxes[i].grid(row=1, column=i+2)
        checkboxes[i].var = checkbox_vars[i]
    # Textboxes and arrays
    rows = 6
    columns = 6
    arr = [[0 for j in range(columns)] for i in range(rows)]
    textboxes = []

    def on_select(event):
        if event.widget.get() == "missing":
            event.widget.delete(0, tk.END)
            event.widget.config(fg='black')

    for i in range(rows):
        textboxes.append([])
        for j in range(columns):
            if j == 4:
                textboxes[i].append(tk.Entry(root, width=16))
                textboxes[i][j].grid(row=i+2, column=j+2, padx=3)
                textboxes[0][j].delete(0, tk.END)
                textboxes[0][j].insert(0, "Mg(ClO_4)_2")


            else:
                textboxes[i].append(tk.Entry(root, width=8))
                textboxes[i][j].grid(row=i+2, column=j+2, padx=3)

    for i in range(rows):
        tk.Button(root, text="Choose File... ", command=lambda row=i: openFile(row)).grid(row=i+2, column=8)
        tk.Button(root, text=" X ", foreground='#8D005F', font=('calibri', 8, 'bold'), command=lambda row=i: clearFile(row, 'y')).grid(row=i + 2,column=9,)

    tk.Button(root, text="Generate Sheet", foreground='#8D005F', font=('calibri', 10, 'bold'), command=lambda: uploadFiles()).grid(row=8, column=8, columnspan=2, padx=4, pady=10)
    ttk.Label(root, text='Type _ before char\nto  subscript',font=('calibri', 8, 'bold'), foreground='black').grid(row=8, column=6, rowspan=2)
    ttk.Label(root, text='Type RT for \nroom temp.', font=('calibri', 8, 'bold'), foreground='black').grid(row=8, column=5, rowspan=2)
    tk.Label(root, text='Check box to make constant.', font=('calibri', 8, 'bold'), background='#d9d9d9').grid(row=1, column=1)
    def on_check_change(i):
        disabled_check = checkboxes[i].var.get()

        if disabled_check == 1:
            txtbox = 1
            while txtbox <= 5:
                default = textboxes[0][i].get()
                textboxes[txtbox][i].delete(0, tk.END)
                textboxes[txtbox][i].insert(0, default)
                textboxes[txtbox][i].config(state="disabled", foreground="grey")
                txtbox += 1

        else:
            txtbox = 1
            while txtbox <= 5:
                textboxes[txtbox][i].config(state="normal", foreground="black", text="")
                textboxes[txtbox][i].delete(0, tk.END)
                txtbox += 1

    def uploadFiles():
        global data_array
        data_array = [[[], [], [], [], []], [[], [], [], [], []], [[], [], [], [], []], [[], [], [], [], []],
                      [[], [], [], [], []], [[], [], [], [], [], []]]
        global expt_data
        expnum = 0
        newnum = 0
        numexpts = 0
        for c in range(6):
            if expt_data[c][6] != '':
                numexpts += 1
        while expnum <= 5:
            #skip sets with no file uploaded
            if expt_data[expnum][6] != '':
                #get variables/constants for set.
                # warn of missing constant.
                i = 0
                pass_var = 0
                while i <= 5:
                    if textboxes[expnum][i].get() == '':
                        textboxes[expnum][i].insert(0, 'missing')
                        textboxes[expnum][i].config(fg="red")
                        textboxes[expnum][i].bind('<FocusIn>', on_select)
                    else:
                        expt_data[expnum][i] = textboxes[expnum][i].get()
                        pass_var += 1
                    i += 1

                    if pass_var == 6:
                        print(expt_data[expnum][5])
                        #cur_list, pot_list, curdens_list, log_list, OERHER
                        data_array[newnum][0], data_array[newnum][1], data_array[newnum][2], data_array[newnum][3], OERHER = get_data(expt_data[expnum][6])
                        slope, Tslope, max_range, r2, ind1, ind2, log_low, log_high = getTafel(data_array[newnum][3], data_array[newnum][1])
                        data_array[newnum][4].append(OERHER)  # 0
                        data_array[newnum][4].append(expt_data[newnum][5])  # 1
                        data_array[newnum][4].append(Tslope)  # 2
                        data_array[newnum][4].append(max_range)  # 3
                        data_array[newnum][4].append(ind1)  # 4
                        data_array[newnum][4].append(ind2)  # 5
                        data_array[newnum][4].append(log_low)  # 6
                        data_array[newnum][4].append(log_high)  # 7
                        #[8:rpm],[9:electrode],[10: molarity],[11: temp],[12:electrolyte],[13:r_squared]
                        data_array[newnum][4].append(expt_data[expnum][0])  # 8
                        data_array[newnum][4].append(expt_data[expnum][1])  # 9
                        data_array[newnum][4].append(expt_data[expnum][2])  # 10
                        data_array[newnum][4].append(expt_data[expnum][3])  # 11
                        data_array[newnum][4].append(expt_data[expnum][4])  # 12
                        data_array[newnum][4].append(r2)  # 13
                        newnum += 1

            expnum += 1
        indvar = ['', '', '', '', '', '',[]]
        if newnum > 0 and newnum == numexpts:
            loading_bar = ttk.Progressbar(root, orient='horizontal', mode='indeterminate', length=250,
                                          style='red.Horizontal.TProgressbar')
            style = ttk.Style()
            style.theme_use('clam')
            style.configure("red.Horizontal.TProgressbar", foreground='#8D005F', background='#8D005F')
            loading_bar.grid(column=10, row=10, rowspan=3)

            def appending_func():
                loading_bar.start()
                jj = 1
                while jj < newnum:
                    if data_array[jj][4][8] != data_array[0][4][8] and 'RPM' not in indvar[6]:
                        indvar[6].append('RPM')
                    if data_array[jj][4][9] != data_array[0][4][9] and 'Electrode' not in indvar[6]:
                        indvar[6].append('Electrode')
                    if data_array[jj][4][10] != data_array[0][4][10] and 'M'  not in indvar[6]:
                        indvar[6].append('M')
                    if data_array[jj][4][11] != data_array[0][4][11] and '°C' not in indvar[6]:
                        indvar[6].append('°C')
                    if data_array[jj][4][12] != data_array[0][4][12] and 'Electrolyte' not in indvar[6]:
                        indvar[6].append('Electrolyte')
                    if data_array[jj][4][0] != data_array[0][4][0] and 'OER/HER' not in indvar[6]:
                        indvar[6].append('OER/HER')
                    jj += 1
                print(f'indvar: {indvar[6]}')
                append_data(data_array, newnum, indvar[6])
                loading_bar.stop()
                ## delete the line below to stop program from automatically opening the file


                # Connect to an existing workbook
                wb_xlw = xw.Book(f'excelfiles/file_temp.xlsx')

                # Activate the sheet with your chart
                sheet_xlw = wb_xlw.sheets['temp']

                # Access the first chart
                chart_xlw = sheet_xlw.api.ChartObjects(1).Chart
                chart_count = sheet_xlw.api.ChartObjects().Count
                chart_xlw1 = sheet_xlw.api.ChartObjects(chart_count - 1).Chart
                chart_xlw2 = sheet_xlw.api.ChartObjects(chart_count).Chart

                # Delete every other legend entry
                for i in range(numexpts * 2, 0, -2):
                    chart_xlw.Legend.LegendEntries(i).Delete()

                chart_xlw1.Legend.LegendEntries(1).Delete()
                for i in range(numexpts * 2, 0, -2):
                    chart_xlw1.Legend.LegendEntries(i).Delete()
                # Iterate over every other legend entry, starting from the first one
                print('success')

                for i in range(numexpts-1):
                    chart_xlw1.Legend.LegendEntries(numexpts+1).Delete()


                wb_xlw.save()
                loading_bar.destroy()
                ttk.Label(root, text='Sheet Generated!', foreground='#8D005F').grid(column=8, row=11, columnspan=2)



            t = threading.Thread(target=appending_func)
            t.start()







    # Add the command to the checkbox
    checkboxes[0].config(command= lambda: on_check_change(0))
    checkboxes[1].config(command=lambda: on_check_change(1))
    checkboxes[2].config(command=lambda: on_check_change(2))
    checkboxes[3].config(command=lambda: on_check_change(3))
    checkboxes[4].config(command=lambda: on_check_change(4))



#[0:rpm],[1:electrode],[2: molarity],[3: temp],[4:electrolyte],[5: expt num], [6:file_path]
expt_data = ['','','','','','',''],\
    ['','','','','','',''],\
    ['','','','','','',''],\
    ['','','','','','',''],\
    ['','','','','','',''],\
    ['','','','','','','']


def openFile(row):
    global expt_data
    file = askopenfile(mode='r', filetypes=[('CSV Files', '*csv')])
    if file:
        clearFile(row, 'n')
        file_path = os.path.abspath(file.name)
        expt_data[row][6] = file_path
        file_str = expt_data[row][6][-60:]
        ttk.Label(root, text="..." + file_str).grid(row=row+2, column=10, padx=10)


def clearFile(row, yn):
    global textboxes
    for widget in root.grid_slaves():
        if int(widget.grid_info()["column"]) > 9 and int(widget.grid_info()["row"]) == row + 2:
            widget.grid_forget()
            expt_data[row][6] = ""
    if yn == 'y':
        for i in range(6):
            if textboxes[row][i]["state"] != "disabled" and i!=0:
                textboxes[row][i].delete(0, tk.END)
                textboxes[row][i].config(fg='black')

create_textboxes()
root.mainloop()