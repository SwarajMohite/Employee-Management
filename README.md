# Employee Management System (EMS)

![Python](https://img.shields.io/badge/Python-3.6+-blue.svg)

A desktop application for managing employee records, attendance, and generating reports, built with Python and Tkinter.

## Features

### User Management
- **Role-based access control** (Employee/Manager)
- Secure login system
- User registration with password confirmation

### Employee Management (Manager Only)
- Add new employees with unique IDs
- Remove employees from the system
- View complete employee list

### Attendance Tracking
- Mark daily attendance (Present/Absent)
- View attendance records
- Filter attendance by date and employee
- Export attendance data to Excel

### Reporting (Manager Only)
- Generate attendance summary reports
- View present/absent counts per employee
- Export reports to Excel

### Data Management (Manager Only)
- Create system backups
- Restore from backups
- Export all data to Excel
- Clear all system data

## Requirements

- Python 3.6+
- Required packages:
  - pandas
  - openpyxl
  - Pillow
  - matplotlib

Install dependencies with:
```bash
pip install pandas openpyxl Pillow matplotlib
