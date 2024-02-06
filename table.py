from tkinter import *
import pandas as pd
from tkinter import ttk


class DataFrameViewer:
    def __init__(self):
        self.window = Tk()
        self.window.geometry("1500x600")
        self.window.resizable(False,False)

        self.df = pd.read_excel("modbus.xlsx")
        self.text_widgets = [] 
        self.create_header()
        self.create_data_grid()
        self.button=ttk.Button(self.window,text="Exit",command=self.window.destroy)
        self.button.place(x=730,y=560)
        self.window.mainloop()

    def create_header(self):
        column_names = self.df.columns
        for j, col in enumerate(column_names):
            text = self.create_text_widget(0, j, bg="#9BC2E6")
            text.insert(INSERT, col)

    def create_data_grid(self):
        n_rows, n_cols = self.df.shape
        for i in range(n_rows):
            for j in range(n_cols):
                text = self.create_text_widget(i+1, j)
                text.insert(INSERT, self.df.loc[i][j])

    def create_text_widget(self, row, column, bg=None):
        text = Text(self.window, width=40, height=1, bg=bg)
        text.grid(row=row, column=column)
        return text



if __name__ == "__main__":
    obj=DataFrameViewer()
