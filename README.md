# attendance_recorder_fingerprint

This project is an integrated attendance and departure tracking system using fingerprint recognition, consisting of two main components:

1- Attendance and Departure Management System: The main application for managing employee attendance.

2- Fingerprint Device Simulator: A tool to simulate a fingerprint device for testing and development purposes.

---

## Screenshots

### 1. Attendance and Departure Management System

*Description: Main interface showing employee attendance records.*

![Attendance System](assets/Attendance_and_Departure.PNG)

### 2.Export Data to Excel:
The system includes a button that allows exporting attendance data to an Excel file. The information is organized based on check-in or check-out status, and the user can choose a specific time period (day, week, or month). This feature facilitates efficient analysis and tracking of employee attendance.

![Fingerprint Simulator](assets/Exel.png) ![Fingerprint Simulator](assets/Exel_pres.PNG)
---

### 3. Fingerprint Device Simulator

*Description: Simulator interface for testing fingerprint devices.*

![Fingerprint Simulator](assets/Fingerprint_Device_Simulator.PNG)



### 4.la base de donnée
*This project uses a MySQL database to manage employee attendance with two tables: **employees** and **attendance**. XAMPP is used to host the MySQL server locally. The Python app connects to the database using mysql.connector.*



CREATE TABLE **employees** (
    id INT PRIMARY KEY AUTO_INCREMENT,
    name VARCHAR(100),
    fingerprint_id INT UNIQUE
);

CREATE TABLE **attendance** (
    id INT PRIMARY KEY AUTO_INCREMENT,
    employee_id INT,
    date DATE,
    time TIME,
    status ENUM('IN', 'OUT'),
    FOREIGN KEY (employee_id) REFERENCES employees(id)
);



![Fingerprint Simulator](assets/base_de_donne.PNG)

## 5.How to Run

To run the attendance recorder and fingerprint simulator programs, use the following commands in the terminal after navigating to the project folder:

```bash
python attendance_recorder.py
python fingerprint_simulator.py
```
If you encounter any issues or errors, please feel free to contact me at **ramdanchennib@gmail.com**.


