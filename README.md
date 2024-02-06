# Modbus Data Monitoring Application

## Overview
This project is a Modbus data monitoring application developed using Python and the Tkinter library for the graphical user interface. The application allows users to connect to a Modbus device, read real-time data, write data, and display the data in tabular format. It also provides functionality to save the monitored data to an Excel file.

## Features
- **Connectivity**: The application supports connecting to a Modbus device with configurable parameters such as baudrate, port, method, and parity.
- **Real-time Monitoring**: Users can continuously monitor the voltage and frequency data from the connected Modbus device.
- **Write Data**: Users can write a specific value to the Modbus device.
- **Data Logging**: The application periodically saves the monitored data to an Excel file (`modbus_data.xlsx`) for historical reference.
- **Table Views**: Two table views are available:
  - `modbus.xlsx`: Displays the configuration parameters for different Modbus values.
  - `modbus_data.xlsx`: Displays the historical data saved from the real-time monitoring.

## Usage
1. Run the `Modbus2.py` file to launch the main application window.
2. Configure the Modbus connection settings (Baudrate, Port, Method, Parity).
3. Select the Modbus value to monitor from the dropdown list.
4. Click the "Show" button to establish a connection and start monitoring.
5. Real-time voltage and frequency data will be displayed on the interface.
6. Optionally, write a specific value to the Modbus device using the "Write" button.
7. Click the "Disconnect" button to stop monitoring and close the connection.
8. The application periodically saves monitored data to `modbus_data.xlsx`.
9. Table views for `modbus.xlsx` and `modbus_data.xlsx` can be accessed using the "Table" and "Table1" buttons, respectively.
10. Click the "Exit" button to close the application.

## Dependencies
- Python 3.x
- Tkinter
- Pandas
- Openpyxl
- Pymodbus

## File Structure
- `Modbus2.py`: Main application file with the Tkinter GUI.
- `logic.py`: Backend logic for Modbus communication and data handling.
- `table.py`: Table view for `modbus.xlsx`.
- `table2.py`: Table view for `modbus_data.xlsx`.

## Notes
- The application uses the Pymodbus library for Modbus communication.
- The configuration parameters for different Modbus values are stored in the `modbus.xlsx` file.
- The monitored data is saved in the `modbus_data.xlsx` file.
- The application utilizes threading for real-time data monitoring and periodic data saving.

## Author
[Krishna Gupta]

Feel free to modify and extend this project according to your needs.