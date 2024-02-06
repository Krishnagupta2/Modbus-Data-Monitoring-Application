from tkinter import *
import pandas as pd
from tkinter import ttk

class DataFrameViewer:
    def __init__(self):
        self.window = Tk()
        self.window.geometry("400x200")
        self.window.resizable(False, TRUE)

        self.df = pd.read_excel("modbus_data.xlsx")
        self.create_header()
        self.create_data_grid()
        self.button = ttk.Button(self.window, text="Exit", command=self.window.destroy)
        self.button.place(x=170, y=175)

        # Create a Canvas widget
        canvas = Canvas(self.window)
        canvas.grid(row=1, column=0, sticky="nsew")

        # Create a vertical scrollbar
        scrollbar = Scrollbar(self.window, orient="vertical", command=canvas.yview)
        scrollbar.grid(row=1, column=1, sticky="ns")

        # Configure the canvas to use the scrollbar
        canvas.configure(yscrollcommand=scrollbar.set)

        # Create a frame inside the canvas to hold the data grid
        frame = Frame(canvas)
        canvas.create_window((0, 0), window=frame, anchor="nw")

        # Bind the canvas to the mouse wheel to enable scrolling
        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))

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
                text = self.create_text_widget(i + 1, j)

                # Insert data into the Text widget
                text.insert(INSERT, self.df.loc[i][j])

    def create_text_widget(self, row, column, bg=None):
        text = Text(self.window, width=20, height=1, bg=bg)
        text.grid(row=row, column=column)
        return text

if __name__ == "__main__":
    obj = DataFrameViewer()
