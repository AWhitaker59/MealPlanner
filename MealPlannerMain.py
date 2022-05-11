from MealPlannerUI import *


def main():
    planner = Tk()
    planner.title("MealRandomizer")
    planner.geometry("250x400")
    planner.resizable(False, False)
    widgets = GUI(planner)
    planner.mainloop()


if __name__ == "__main__":
    main()
