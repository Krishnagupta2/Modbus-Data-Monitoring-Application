from pymodbus import ModbusException
from pymodbus.client import ModbusSerialClient

from time import sleep
import pandas as pd




class logic2():
    def __init__(self,gui_instance):
        self.error_message = None
        self.gui_instance = gui_instance  
        self.Word_Count=0
        self.address=0
    def find_row_by_value(self,file_path, target_value):
        try:
            
            df = pd.read_excel(file_path)

            
            matching_rows = df[df.eq(target_value).any(axis=1)]

            if not matching_rows.empty:
                print(f"Rows containing the value '{target_value}':")
                print(matching_rows)

                
                first_matching_row_values = matching_rows.iloc[0].values

                print(f"Values of the first matching row:")
                self.address=int(first_matching_row_values[1])
                self.Word_Count=int(first_matching_row_values[2])
                print(self.Word_Count)
                self.divisior=int(first_matching_row_values[4])
                self.unit=first_matching_row_values[6]
                
                
            else:
                print(f"No rows found with the value '{target_value}'")

        except Exception as e:
            return f"Error: {e}"
            self.error_message = f"Error in find_row_by_value: {e}"


    def baudrate(self, baudrate, port, method, parity):
        self.serial = ModbusSerialClient(method=method, port=port, baudrate=baudrate, parity=parity)
        if self.serial.connect():
            
            # print('connected')
            return "Connected successfully"
        else:
            # print("not connected")
            return "Connection failed"
            # obj.
    def close_port(self):
        if self.serial and self.serial.is_socket_open():
            self.serial.close()
            print("Port closed")
        else:
            print("Port is not open")
    def write_data(self, value):
        try:
            response = self.serial.write_register(self.address, value, unit=1)
            print(f"Write response: {response}")
        except ModbusException as e:
            print(f"Modbus IO Exception: {e}")
            self.error_message = f"Modbus IO Exception: {e}"
        except Exception as e:
            print(f"Unexpected Exception: {e}")
            self.error_message = f"Unexpected Exception: {e}"


    def continuously_read_voltage(self, queue1,stop_event):
        try:
            while not stop_event.is_set() :
                if hasattr(self, 'serial') and self.serial.is_socket_open():
                    response1 = self.serial.read_holding_registers(782, 2, 5)
                    
                    voltage1 = response1.registers
                    m=voltage1[0]
                    k=voltage1[1]
                    voltage=((m<<1)+k)/100
                     
                    if self.Word_Count ==2:

                    
                        response2 = self.serial.read_holding_registers(self.address,self.Word_Count, 5)
                        Frequency1=response2.registers
                        l=Frequency1[0]
                        r=Frequency1[1]
                        Frequency=(str(((l<<1)+r)/self.divisior))+str(self.unit)
                        queue1.put((voltage,Frequency))
                    else:
                        response2=self.serial.read_holding_registers(self.address,self.Word_Count,5)
                        Frequency=response2.registers[0]





                    sleep(0.3)
                    queue1.put((voltage,Frequency))
                    

                    
                    
                   
                    
                else:
                 break
        
        except ModbusException as e:
            print(f"Modbus IO Exception: {e}")
            self.error_message = f"Modbus IO Exception: {e}"
        except AttributeError as e:
            if str(e) == "'ModbusIOException' object has no attribute 'registers'":
                print(f"AttributeError: {e}")
                self.error_message = f"AttributeError: {e}"
                # Update the error label in the GUI with the specific error message
                # self.error_label.config(text=self.error_message)
                self.gui_instance.error_label.config(text=self.error_message)
            else:
                print(f"Unexpected Attribute Error: {e}")
                self.error_message = f"Unexpected Attribute Error: {e}"

        except Exception as e:
            print(f"Unexpected Exception: {e}")
            self.error_message = f"Unexpected Exception: {e}"
        finally:
            if hasattr(self, 'serial') and self.serial.is_socket_open():
                self.serial.close()
                print("Port closed")

