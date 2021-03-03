# note program need install xlrd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import checkdataframes
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pathlib
from tkinter import *

"""GUI code"""
# initialised the tkinter GUI
root = tk.Tk()

# Setting up root main window
root.geometry("900x600")  # set the root dimensions
root.pack_propagate(False)  # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0)  # makes the root window fixed in size.


# button onhover styling
class HoverButton(tk.Button):
    def __init__(self, master, **kw):
        tk.Button.__init__(self, master=master, **kw)
        self.defaultBackground = self["background"]
        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)

    def on_enter(self, e):
        self['background'] = self['activebackground']

    def on_leave(self, e):
        self['background'] = self.defaultBackground


# Creating Menubar
menubar = Menu(root)

# Adding File Menu and commands
file = Menu(menubar, tearoff=0)
menubar.add_cascade(label='File', menu=file)
file.add_command(label='Open...', command=lambda: [File_dialog(), Load_excel_data()])
file.add_separator()
file.add_command(label='Exit', command=root.destroy)

# Adding Edit Menu and commands
Search = Menu(menubar, tearoff=0)
menubar.add_cascade(label='Search', menu=Search)
Search.add_command(label='Search Data', command=lambda: search_cri())

# Adding Analyse Menu
Analysis = Menu(menubar, tearoff=0)
menubar.add_cascade(label='Analyse', menu=Analysis)
Analysis.add_command(label='Analyse Charts', command=lambda: analyse_data())

# Adding Help Menu
help_ = Menu(menubar, tearoff=0)
menubar.add_cascade(label='Help', menu=help_)
help_.add_command(label='Tk Help', command=None)
help_.add_command(label='Demo', command=None)
help_.add_separator()
help_.add_command(label='About Tk', command=None)

# display Menu
root.config(menu=menubar)

# Frame for TreeView to display data set
Tree_Frame = tk.LabelFrame(root, text="Excel Data")
Tree_Frame.place(height=400, width=900)

# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="Open File")
file_frame.place(height=100, width=300, rely=0.7, relx=0.01)

# Buttons to read excel files
browse_button = HoverButton(file_frame, bg="#508bc7", activebackground='#66a2de', font="Helvetica 8",
                            text="Browse A File", command=lambda: File_dialog())
browse_button.place(rely=0.65, relx=0.50)

load_button = HoverButton(file_frame, bg="#66d982", activebackground='#3fd163', font="Helvetica 8", text="Load File",
                          command=lambda: Load_excel_data())
load_button.place(rely=0.65, relx=0.30)

# The file/file path text
Selected_FLabel = ttk.Label(file_frame, text="No File Selected")
Selected_FLabel.place(rely=0, relx=0)

# Search Criterion button
search_button = HoverButton(root, bg="#b660d1", activebackground='#c976e3', text="Search Data", font="Helvetica 9",
                            command=lambda: search_cri())
search_button.pack()
search_button.place(x=300, y=525, height=50, width=80)



# quit button
quit_button = HoverButton(root, bg="#cc5867", activebackground='#de6a79', font="Helvetica 9", text="Quit",
                          command=lambda: destroy())
quit_button.pack()
quit_button.place(x=500, y=525, height=50, width=50)


# Function to end program
def destroy():
    root.destroy()


# label to open graphing tool
analyse_label = tk.LabelFrame(root, text="Analyse Data using charts")
analyse_label.place(height=100, width=200, x=600, y=415)

# button to open graphing tool
analyse_button = HoverButton(root, text="Analyse Data", bg="#66d982", activebackground='#3fd163', font="Helvetica 10",
                             command=lambda: analyse_data())
analyse_button.pack()
analyse_button.place(x=650, y=440, height=50, width=100)

# Tree view widget in my 1st frame
Tree_View1 = ttk.Treeview(Tree_Frame)
Tree_View1.place(relheight=1, relwidth=1)

# Scrolling bar to fit more data in
TVScroll_y = tk.Scrollbar(Tree_Frame, orient='vertical', command=Tree_View1.yview)  # scolling feauture for y
TVScroll_x = tk.Scrollbar(Tree_Frame, orient='horizontal', command=Tree_View1.xview)
Tree_View1.configure(xscrollcommand=TVScroll_x.set, yscrollcommand=TVScroll_y.set)
# Pack these in side
TVScroll_x.pack(side='bottom', fill='x')
TVScroll_y.pack(side='right', fill='y')

""" Open new window to search data """


# This function is used to open another Tkinter window to allow user to search through the columns
def search_cri():
    # Get the name of data file being opened
    file_path = Selected_FLabel["text"]
    # call the function get_columns that return list of columns head
    column_names = checkdataframes.get_columns(file_path)
    # Initialise Search TK window
    search_data = tk.Tk()
    search_data.title("Search data")
    search_data.geometry("800x600")  # set the root dimensions
    search_data.pack_propagate(False)  # tells the root to not let the widgets inside it determine its size.
    search_data.resizable(0, 0)  # makes the root window fixed in size.

    # Function check data to entered by the user against dataset. Return new data set with match values
    def search_now():
        # store which data is being selected by user
        selected = drop.get()
        # get Pandas Data frame
        data_frame = checkdataframes.checkfiletype(file_path)
        # get values entered by user
        searched = search_box.get()
        # match data key in to data set
        searched_df = data_frame[data_frame[selected].astype(str).str.contains(searched)]
        Tree_View2["column"] = list(searched_df.columns)
        # Show only the headings
        Tree_View2["show"] = "headings"
        for column in Tree_View2["column"]:
            Tree_View2.heading(column, text=column)  # let the column heading = column name
        # turns the dataframe into a list of lists
        searched_df_rows = searched_df.to_numpy().tolist()
        # inserts each list into the treeview
        for row in searched_df_rows:
            Tree_View2.insert("", "end", values=row)

        return None

    def clear_searcheddata():
        Tree_View2.delete(*Tree_View2.get_children())
        return None


    # Frame for TreeView Data search
    Search_Frame = tk.LabelFrame(search_data, text="Search Criterion")
    Search_Frame.place(height=100, width=500, relx=0.01)

    # Entry box to search data
    search_box = tk.Entry(Search_Frame)
    search_box.grid(row=0, column=1, padx=10, pady=10)

    # Search box label search for customer
    search_box_label = tk.Label(Search_Frame, text="Search Data: ")
    search_box_label.grid(row=0, column=0, padx=10, pady=10)

    # Entry box search Button customer
    search_button = HoverButton(Search_Frame, bg="#66d982", activebackground='#3fd163', text="Search",
                                command=search_now)  # add command
    search_button.grid(row=1, column=0, padx=10)
    # Entry box clear Button customer
    clear_button = HoverButton(Search_Frame, bg="#FF0000", activebackground='#3fd163', text="Clear",
                                command=clear_searcheddata)  # add command
    clear_button.grid(row=1, column=2, padx=10)

    # Drop down Box
    drop = ttk.Combobox(Search_Frame, values=column_names)
    drop.current(0)
    drop.grid(row=0, column=2)

    # Frame for TreeView Data search
    Tree2_Frame = tk.LabelFrame(search_data, text="Search Data")
    Tree2_Frame.place(height=370, width=785, rely=0.20, relx=0.01)

    # Tree view 2 widget in my 1st frame
    Tree_View2 = ttk.Treeview(Tree2_Frame)
    Tree_View2.place(relheight=1, relwidth=1)

    # Scrolling bar to fit more data in
    TVScroll_y = tk.Scrollbar(Tree2_Frame, orient='vertical', command=Tree_View2.yview)  # scolling feauture for y
    TVScroll_x = tk.Scrollbar(Tree2_Frame, orient='horizontal', command=Tree_View2.xview)
    Tree_View2.configure(xscrollcommand=TVScroll_x.set, yscrollcommand=TVScroll_y.set)

    # Pack these in side
    TVScroll_x.pack(side='bottom', fill='x')
    TVScroll_y.pack(side='right', fill='y')

    # Frame for TreeView Export
    Export_Frame = tk.LabelFrame(search_data, text="Export Data")
    Export_Frame.place(height=100, width=250, rely=0.82, relx=0.35)

    # Entry box to export file name
    export_box = tk.Entry(Export_Frame)
    export_box.grid(row=0, column=1, padx=5, pady=5)

    # Search box label search for customer
    export_box_label = tk.Label(Export_Frame, text="New file name: ")
    export_box_label.grid(row=0, column=0, padx=10, pady=10)

    # Export button
    export_button = HoverButton(Export_Frame, text="Export to excel", bg="#508bc7", activebackground='#66a2de',
                                command=lambda: export())  # add command
    export_button.grid(row=1, column=0, padx=5, pady=5)

    # back button
    back_button = HoverButton(Export_Frame, bg="#cc5867", activebackground='#de6a79', text="Back",
                              command=lambda: back())
    back_button.grid(row=1, column=1, padx=5, pady=5)

    # Remove search data window
    def back():
        search_data.destroy()

    # check export file for special characters or blank input.
    def checkFileName(export_to):
        # check export file for special characters or blank input.
        special_char = ".!@#$%^&*()-+"

        if export_to == "":
            print("Invalid file name! Please enter another file name, don't leave it blank this time :)")

            return False

        elif any(c in special_char for c in export_to):
            print("You have entered an invalid filename with special characters! Run the program again.")
            return False

        else:
            return True

    def checkFilePath(comparedir):
        path = pathlib.Path(comparedir)
        if path.is_file():
            print("File already exists! Please enter another file name or select another directory")
            return True
        else:
            return False

    def export():
        export_to = export_box.get()
        selected = drop.get()
        data_frame = checkdataframes.checkfiletype(file_path)
        searched = search_box.get()

        # ask for exporting directory
        askdir = filedialog.askdirectory()
        targetdir = askdir.replace("/", '\\')
        comparedir = targetdir + "\\" + export_to + ".xlsx"

        if checkFileName(export_to) == False:
            # need to find out a way to run program again.
            exit()

        elif checkFilePath(comparedir) == True:
            exit()

        else:
            # allow user to specify filename, then append to .xlsx
            searched_df = data_frame[data_frame[selected].astype(str).str.contains(searched)]
            searched_df.to_excel(targetdir + "\\" + export_to + ".xlsx")


def analyse_data():
    """ Open new window to analyse data """
    file_path = Selected_FLabel["text"]
    column_names = checkdataframes.get_columns(file_path)
    # Chart type offered by program stored in list to print in dropdown box.
    chart_names = ['pie', 'bar', 'line', 'scatter']
    analyse_data = tk.Tk()
    analyse_data.title("Search data")
    analyse_data.geometry("800x600")  # set the root dimensions
    analyse_data.pack_propagate(False)  # tells the root to not let the widgets inside it determine its size.
    analyse_data.resizable(0, 0)  # makes the root window fixed in size.

    def back():
        analyse_data.destroy()

    def popupmsg(msg):
        popup = tk.Tk()
        popup.wm_title("!")
        label = ttk.Label(popup, text=msg, font=("Helvetica", 10))
        label.pack(side="top", fill="x", pady=10)
        B1 = ttk.Button(popup, text="Okay", command=popup.destroy)
        B1.pack()
        popup.mainloop()

    # this function is used by program to plot graphs using Matplotlib
    def analyse_now():
        # get the chosen x and y axis
        selected_y = dropy.get()
        selected_x = dropx.get()
        # get chart type
        selected_chart = dropCtype.get()
        # get data frame
        data_frame = checkdataframes.checkfiletype(file_path)
        # plot graphs
        figure = plt.Figure(figsize=(5, 5), dpi=100)
        ax = figure.add_subplot(111)

        # Frame for Data chart
        Chart_Frame = tk.LabelFrame(analyse_data, text="Display Chart")
        Chart_Frame.place(height=460, width=785, rely=0.15, relx=0.01)

        chart_type = FigureCanvasTkAgg(figure, Chart_Frame)
        chart_type.get_tk_widget().pack()

        try:
            # Pie and scatter dont work with the general grp statement so use if use to plot graph
            if selected_chart == 'pie':
                df = data_frame.groupby([selected_y]).sum()
                df.plot(kind=selected_chart, y=selected_x, ax=ax)
                ax.set_title(f"{selected_y}" " Vs " f"{selected_x}")

            elif selected_chart == 'scatter':
                data_frame.plot(kind=selected_chart, x=selected_x, y=selected_y, legend=True, ax=ax)
                ax.set_title(f"{selected_y}" " Vs " f"{selected_x}")
                plt.setp(ax.get_xticklabels(), fontsize=8, rotation='30')

            else:
                df = data_frame[[selected_y, selected_x]].groupby(selected_x).sum()
                df.plot(kind=selected_chart, legend=True, ax=ax)
                ax.set_title(f"{selected_y}" " Vs " f"{selected_x}")
                plt.setp(ax.get_xticklabels(), fontsize=8, rotation='30')

        except KeyError:
            popupmsg("This chart is not graphical!")
        except TypeError:
            popupmsg("This chart is not graphical!")
        except ValueError:
            popupmsg("This chart is not graphical!")

    # Frame for analyse columns
    analyse_Frame = tk.LabelFrame(analyse_data, text="Analyse Data")
    analyse_Frame.place(height=100, width=500, relx=0.01)

    # y drop label search for customer
    y_drop_label = tk.Label(analyse_Frame, text="Y Data: ")
    y_drop_label.grid(row=0, column=0, padx=10, pady=10)

    # Drop down Box - x col
    dropy = ttk.Combobox(analyse_Frame, values=column_names)
    dropy.current(0)
    dropy.grid(row=0, column=1)

    # y drop label search for customer
    x_drop_label = tk.Label(analyse_Frame, text="X Data: ")
    x_drop_label.grid(row=1, column=0, padx=10, pady=10)

    # Drop down Box - y col
    dropx = ttk.Combobox(analyse_Frame, values=column_names)
    dropx.current(0)
    dropx.grid(row=1, column=1)

    # y drop label search for customer
    type_drop_label = tk.Label(analyse_Frame, text="Chart Type: ")
    type_drop_label.grid(row=0, column=2, padx=10, pady=10)

    # Drop down Box - y col
    dropCtype = ttk.Combobox(analyse_Frame, values=chart_names)
    dropCtype.current(0)
    dropCtype.grid(row=0, column=3)

    # Entry box analyse Button customer
    analyse_button = HoverButton(analyse_Frame, bg="#508bc7", activebackground='#66a2de', text="Analyse",
                                 command=analyse_now)  # add command
    analyse_button.grid(row=1, column=3, padx=10)

    # back button
    back_button = HoverButton(analyse_data, bg="#cc5867", activebackground='#de6a79', text="Back",
                              command=lambda: back())
    back_button.pack()
    back_button.place(x=400, y=560, height=30, width=50)


"""Functions for uploading files"""


def File_dialog():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"), ("text files", "*.txt"),
                                                    ("csv files", "*.csv"), ("json files", "*.json"),
                                                    ("All Files", "*.*")))
    Selected_FLabel["text"] = filename
    return None


def Load_excel_data():
    """If the file selected is valid this will load the file into the Treeview"""
    file_path = Selected_FLabel["text"]
    try:
        data_frame = checkdataframes.checkfiletype(file_path)
    # if the above fail
    except ValueError:
        tk.messagebox.showerror("Information", "The file chosen is invalid!")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No file named {file_path} found!")

    clear_data()
    Tree_View1["column"] = list(data_frame.columns)
    # Show only the headings
    Tree_View1["show"] = "headings"
    for column in Tree_View1["column"]:
        Tree_View1.heading(column, text=column)  # let the column heading = column name
    # turns the dataframe into a list of lists
    dataframe_rows = data_frame.to_numpy().tolist()
    # inserts each list into the treeview
    for row in dataframe_rows:
        Tree_View1.insert("", "end", values=row)
    return None


def clear_data():
    Tree_View1.delete(*Tree_View1.get_children())
    return None


# Create a loop to check for end of program
root.mainloop()
