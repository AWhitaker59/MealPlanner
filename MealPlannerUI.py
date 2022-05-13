import datetime
from tkcalendar import Calendar
from tkinter import *
import csv
import random
from calendar import monthrange
import os
import os.path
from xlwt import Workbook
import subprocess
FILEBROWSER_PATH = os.path.join(os.getenv('WINDIR'), 'explorer.exe')


def add_meal():
    """This function opens the window to add meals to the default .csv file"""

    def append_new_meal():
        """
        This nested function appends the new meal typed in the entry in the add_meal window
        to the meals.csv
        """
        with open('meals.csv', 'a', newline='') as meals:
            writer = csv.writer(meals)
            meal = [add_entry.get()]
            writer.writerow(meal)
        add_entry.delete(0, END)

    add_meal_window = Toplevel()
    add_meal_window.title("Add Meal")
    add_meal_window.geometry("200x100")
    add_meal_window.resizable(False, False)

    add_label = Label(add_meal_window, text="Enter Meal Name:")
    add_label.grid(row=0, padx=40)
    add_entry = Entry(add_meal_window)
    add_entry.grid(row=1, padx=40)
    add_button = Button(add_meal_window, text="Add Meal", command=append_new_meal)
    add_button.grid(row=2, padx=40)

    add_meal_window.mainloop()


def open_meals_spread():
    """
    This Function opens the meals.csv file in Excel
    """
    os.system("start EXCEL.EXE meals.csv")

def open_saved(self):
    """
    This function opens file explorer to the location where all previously saved
    months are located and allows the user to reopen that month in Excel.
    TODO:Also should open the saved calendar in the app with relative event-ids
    """
    subprocess.run([FILEBROWSER_PATH])


class GUI:
    def __init__(self, planner):
        """
        The below creates the entire GUI for the MealPlanner application
        """
        # main window and initial variables
        self.save_name = None
        self.planner = planner
        self.planner.title("MealRandomizer")

        # menu options
        self.menu_bar = Menu(self.planner)
        self.planner.config(menu=self.menu_bar)
        self.subMenu = Menu(self.menu_bar)
        self.menu_bar.add_cascade(label="File", menu=self.subMenu)
        self.subMenu.add_command(label="Open Meal List", command=open_meals_spread)  # Calls "open_meals_list function from Controller
        self.subMenu.add_command(label="Add Meals", command=add_meal)                # Calls "add_meal" function from Controller
        self.subMenu.add_command(label="Open Saved Month", command=self.open_saved)

        # Calendar Initialization UI
        self.calendar_name_label = Label(self.planner, text="Enter new Calendar Name(Or save to default):")
        self.calendar_name_label.grid(row=0, column=0)
        self.entry_calendar_name = Entry(self.planner)
        self.entry_calendar_name.grid(row=1, column=0)
        self.cal = Calendar(self.planner, selectmode='day', showweeknumbers=False, selectbackground='black')
        self.cal.grid(row=3, column=0)
        self.month_label = Label(self.planner, text="Select Month to plan:")
        self.month_label.grid(row=2, column=0)
        self.cal.tag_config("meal", background='green')
        # Randomize button
        self.randomize_button = Button(self.planner, text="Randomize Meals!", command=self.randomize_meals)
        self.randomize_button.grid(row=4, column=0, pady=5)
        self.get_meal_button = Button(self.planner, text="Selected Date's Meal", command=self.get_meal)
        self.get_meal_button.grid(row=5, pady=5)
        self.get_meal_return = Label(self.planner)
        self.get_meal_return.grid(row=6, pady=5)
        self.save_month_button = Button(self.planner, text="Save Month to Excel", command=self.save_month)
        self.save_month_button.grid(row=7, pady=5)
        self.no_name_label = Label(self.planner, text='')
        self.no_name_label.grid(row=8, pady=5)

    def randomize_meals(self):
        """
        This function iterates through the meals.csv, applies and event ID to
        each meal through the tkcalendar module, then assigns each event ID to a
        date in the selected month. The function then saves that month's meals to a csv
        """
        meals_list = []
        month_year = self.cal.get_displayed_month()      # Used to get the current month/year
        year = month_year[1]
        month = month_year[0]

        #  Below loop creates a list of all meals in the meals.csv
        with open('meals.csv', 'r') as meals:
            for line in meals:
                line = line.rstrip()
                meals_list.append(line)
        num_days = monthrange(year, month)[1]
        num = 1

        # Below loop iterates through month giving a random meal
        while num <= num_days:
            meal_choice = random.choice(meals_list)
            self.cal.calevent_create(datetime.date(year, month, num), meal_choice, tags=["meal"])
            num += 1

    def get_meal(self):
        """
        This function takes the selected date on the Calendar and prints the name of
        the meal assigned to that date onto the this_meal_label
        """
        date = self.cal.get_date()
        date = date.split('/')
        self.get_meal_return['text'] = self.cal.calevent_cget(int(date[1]) - 1, option='text')

    def save_month(self):
        """
        saves the month to a .csv using the name typed into the entry
        opened in Excel
        """
        if self.entry_calendar_name.get() == '':
            self.no_name_label['text'] = 'ENTER A SAVE NAME'
        else:
            try:
                month_year = self.cal.get_displayed_month()  # Used to get the current month/year
                year = month_year[1]
                month = month_year[0]
                num_days = monthrange(year, month)[1]
                day = 1
                new_cal = Workbook()
                self.cal.selection_set(date=datetime.date(year, month, 1))
                date = self.cal.get_date()
                date = date.split('/')
                date[1] = int(date[1])
                sheet1 = new_cal.add_sheet('Meals')
                while day <= num_days:
                    sheet1.write(day-1, 0, self.cal.calevent_cget(date[1] - 1, option='text'))
                    day += 1
                    date[1] += 1
                new_cal.save(f'{self.entry_calendar_name.get()}.xls')
            finally:
                if os.path.isfile(f'{self.entry_calendar_name.get()}.xls'):
                    self.no_name_label['text'] = 'SAVED!'
                else:
                    self.no_name_label['text'] = 'Randomize Days First'
