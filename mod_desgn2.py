from tkinter import *
from tkinter import ttk
import logic
import threading
from queue import Queue
from time import sleep
import pandas as pd
import datetime
import table
import table2
from openpyxl import load_workbook



class Modbus2():

    def __init__(self):
      
        self.root = Tk()
        self.root.title("Modbus Data")
        self.root.width = 700
        self.root.Height = 500
        self.root.geometry(f"{self.root.width}x{self.root.Height}")
        self.root.resizable(False, False)
        self.Voltage_label = Label(self.root, text="...", font=("Mongolian Baiti", 22, 'bold'), bg="white", bd=0)
        self.Voltage_label.place(x=460, y=120)

        
        # self.logic_instance = logic.logic2()
        self.logic_instance = logic.logic2(gui_instance=self)
        self.Frequency_label = Label(self.root, text="..", font=("Mongolian Baiti", 22, 'bold'), bg="white", bd=0)
        self.queue = Queue() 
        self.Frequency_label.place(x=460, y=190)
        self.warning_label = ttk.Label(self.root, text="", foreground="red",font=("Mongolian Baiti", 8, 'bold') )
        self.warning_label.place(x=160,y=350)

        self.error_label = ttk.Label(self.root, text="", foreground="red", font=("Mongolian Baiti", 15, 'bold'))
        self.error_label.place(x=40, y=370)
        self.save_data_thread = threading.Thread(target=self.save_data_periodically)
        self.save_data_thread.daemon = True
        self.save_data_thread.start()
        self.connection_status_label = Label(self.root, text="Status:Not Connected", font=("Mongolian Baiti", 22), bg="white", fg="red")
        self.connection_status_label.place(x=390, y=40)
        self.stop_thread_event = threading.Event()

    def show(self):
        self.stop_thread_event.clear()
        BaudRate=int(self.Baudrate_dropdown2.get())
        print(BaudRate)
        port=self.Port_dropdown3.get()
        method=self.Method_dropdown4.get()
        parity=self.Parity_dropdown5.get()
        connected=self.logic_instance.baudrate(baudrate=BaudRate, port=port, method=method, parity=parity)
        x='modbus.xlsx'
        self.logic_instance.find_row_by_value(x,self.value_read_dropdown5.get())
        
        print(connected)
        if "Connected" in connected:
            self.connection_status_label.config(text="Status: Connected", fg="green")
            self.Baudrate_dropdown2.config(state="disabled")
            self.Port_dropdown3.config(state="disabled")
            self.Method_dropdown4.config(state="disabled")
            self.Parity_dropdown5.config(state="disabled")
            self.button1['state']=DISABLED
            self.disconnect_button1['state']='enable'
            self.disconnect_button2['state']='enable'
            # self.value_read_dropdown5.config(state="disabled")
            self.error_label.config(text='')
            thread1 = threading.Thread(target=self.logic_instance.continuously_read_voltage, args=(self.queue,self.stop_thread_event))
            thread1.start()
            self.root.after(1, self.check_voltage)
            
        else:
            self.connection_status_label.config(text="Status: not Connected", fg="red")
            self.error_label.config(text=self.logic_instance.error_message)
    def check_voltage(self):
        try:
            while not self.queue.empty():
                voltage,Frequency = self.queue.get()
                self.voltage(voltage)
                self.Frequency(Frequency)
        except RuntimeError:
            pass  
        self.root.after(900, self.check_voltage)
    def voltage(self,voltage):
        if voltage:
            print(voltage)
            self.Voltage_label.config(text=f"Voltage:{voltage}V")
        else:
            print("not come")
            self.Voltage_label.config(text="data not come from modbus")

    def Frequency(self,Frequency):
        
        if Frequency:
            print(Frequency)
            self.Frequency_label.config(text=f"value:{Frequency}")
        else:
            print("not come")
            self.Frequency_label.config(text="Data not come")
    def change(self):
        self.logic_instance.find_row_by_value('modbus.xlsx',self.value_read_dropdown5.get())

    def disconnect(self):
    
        self.stop_thread_event.set()
        if hasattr(self, 'thread1') and self.thread1.is_alive():
            self.logic_instance.close_port()
            self.thread1.join()  

        
        self.Baudrate_dropdown2.config(state="enabled")
        self.Port_dropdown3.config(state="enabled")
        self.Method_dropdown4.config(state="enabled")
        self.Parity_dropdown5.config(state="enabled")
        self.button1['state']="enable"
        self.disconnect_button2['state']="disable"
        self.disconnect_button1['state']="disable"
        

        
        self.connection_status_label.config(text="Status: Not Connected", fg="red")
    def exit_program(self):
        
        self.logic_instance.close_port()
        
        self.root.destroy()
    def validate_input(self,P):
        try:
            if P == "":
                self.warning_label.config(text="")
                return True
            int(P)
            self.warning_label.config(text="")
            return True
        except ValueError:
            self.warning_label.config(text="Please enter a valid integer")
            return False
    def write_button(self):
            user_input = int(self.entry.get())
            self.logic_instance.write_data(user_input)
            
            print("User entered:", user_input)
    def save_data_periodically(self):
        while True:
            if not self.queue.empty():
                self.save_data_to_excel()
            sleep(1)  

    def save_data_to_excel(self):
        try:
            current_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            while not self.queue.empty():
                voltage, frequency = self.queue.get_nowait()
                data = {'Time': [current_time], 'Voltage': [voltage], 'Frequency': [frequency]}
                # data['Time'].append(current_time)
                # data['Voltage'].append(voltage)
                # data['Frequency'].append(frequency)

            df = pd.DataFrame(data)
            existing_df=pd.read_excel('modbus_data.xlsx')
            updated_df=pd.concat([df,existing_df],ignore_index=True)
            # print(type(df))
            # df.to_excel('modbus_data.xlsx', index=False)
            with pd.ExcelWriter('modbus_data.xlsx', engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                updated_df.to_excel(writer, index=False, sheet_name='Sheet1')

            print("Data saved to Excel")
        except Exception as e:
            print(f"Error saving data to Excel: {e}")
    def table(self):
        table.DataFrameViewer()
        self.root.withdraw()
    def table1(self):
        table2.DataFrameViewer()
        self.root.withdraw()


      
    def firstframe1(self):
        self.label1 = Label(self.root, text="Modbus Data", font=("Mongolian Baiti", 25), bg="white")
        self.label1.place(x=225, y=0)

        self.label2 = Label(self.root, text="Baudrate", font=("Mongolian Baiti", 15), bg="white")
        self.label2.place(x=50, y=70)
        Baudrate = [9600,9700,9800 ,19600,9300,9400]  
        self.selectedBaudrate = StringVar(value=Baudrate[0])
        self.Baudrate_dropdown2 = ttk.Combobox(self.root, textvariable=self.selectedBaudrate, values=Baudrate)
        self.Baudrate_dropdown2.place(x=150, y=70)
        
        self.label3 = Label(self.root, text="Port", font=("Mongolian Baiti", 15), bg="white")
        self.label3.place(x=50, y=120)
        Port = ['COM2','COM1', 'COM3', 'COM4']  
        self.selected_Port = StringVar(value=Port[0])
        self.Port_dropdown3 = ttk.Combobox(self.root, textvariable=self.selected_Port, values=Port)
        self.Port_dropdown3.place(x=150, y=120)
        
        self.label4 = Label(self.root, text="Method", font=("Mongolian Baiti", 15), bg="white")
        self.label4.place(x=50, y=170)
        Method = ['rtu','Asic'] 
        self.selected_Method= StringVar(value=Method[0])
        self.Method_dropdown4 = ttk.Combobox(self.root, textvariable=self.selected_Method, values=Method)
        self.Method_dropdown4.place(x=150, y=170)

        self.label5 = Label(self.root, text="Parity", font=("Mongolian Baiti", 15), bg="white")
        self.label5.place(x=50, y=230)
        Parity = ['N','1','2','3']  
        self.selected_Parity= StringVar(value=Parity[0])
        self.Parity_dropdown5 = ttk.Combobox(self.root, textvariable=self.selected_Parity, values=Parity)
        self.Parity_dropdown5.place(x=150, y=230)

        self.label5 = Label(self.root, text="Value read", font=("Mongolian Baiti", 15), bg="white")
        self.label5.place(x=50, y=280)
        value_read = ['Frequency','Phase to Neutral voltage phase 1','Phase to Neutral voltage phase 2','Phase to Neutral voltage phase 3','Neutral current','Phase 1 Current','Phase 2 Current','Phase 3 Current','Current Transformer secondary1','Current Transformer secondary','Phase to Phase Voltage: U12','Phase to Phase Voltage: U23','Phase to Phase Voltage: U31','Reactiva Power','Active Power','Active Power phase1 +/-','Active Power phase2 +/-','Active Power phase3+/-','Reactive Power phase1 +/-','Reactive Power phase2+/-','Reactive Power phase3 +/-','power factor : -: leadiing et + : lagging : PF','Power factor phase 1 -:leading and +: lagging','Power factor phase 2 -:leading and +: lagging','Power factor phase 3 -:leading and +: lagging','Voltage Transformer primary','Current Transformer secondary (1 : NO,2: YES)'] 
        self.selected_value_read= StringVar(value=value_read[0])
        self.value_read_dropdown5 = ttk.Combobox(self.root, textvariable=self.selected_value_read, values=value_read)
        self.value_read_dropdown5.place(x=150, y=280)
        
        self.label6 = Label(self.root, text="Write Value", font=("Mongolian Baiti", 15), bg="white")
        self.label6.place(x=50, y=330)
        validate_cmd = (self.root.register(self.validate_input), '%P')
        self.entry=ttk.Entry(self.root,validate='key',validatecommand=validate_cmd)
        self.entry.place(x=160,y=330)

        

        
        self.button1=ttk.Button(self.root,text="Show",style="TButton",command=lambda:Modbus2.show(self) )
        
        self.button1.place(x=40,y=420)

       
        self.button2=ttk.Button(self.root,text="Exit",style="TButton" ,command=self.exit_program)
        self.button2.place(x=120,y=420)

        self.disconnect_button = ttk.Button(self.root, text="Disconnect", style="TButton",command=self.disconnect)
        self.disconnect_button.place(x=200, y=420)

        self.disconnect_button1 = ttk.Button(self.root, text="Change reading", style="TButton",command=self.change)
        self.disconnect_button1.place(x=285, y=420)
        self.disconnect_button6 = ttk.Button(self.root, text="Table", style="TButton",command=self.table)
        self.disconnect_button6.place(x=460, y=420)
        self.disconnect_button7 = ttk.Button(self.root, text="Table1", style="TButton",command=self.table1)
        self.disconnect_button7.place(x=540, y=420)

        self.disconnect_button2 = ttk.Button(self.root, text="Write", style="TButton" ,command=self.write_button)
        self.disconnect_button2.place(x=380, y=420)
        self.disconnect_button2['state']="disable"
        self.disconnect_button1['state']="disable"
        self.root.mainloop()


if __name__ == "__main__":
    obj= Modbus2()
    obj.firstframe1()
