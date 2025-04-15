import datetime
import pandas as pd
import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import shutil

# File Paths
USER_FILE = "users.txt"
EMPLOYEE_FILE = "employees.txt"
ATTENDANCE_FILE = "attendance.xlsx"

class EmployeeManagementSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("🏢 Employee Management System")
        self.root.geometry("1000x650")
        self.root.resizable(False, False)
        self.current_user = None
        self.current_role = None
        self.current_emp_id = None
        
        # Configure styles
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('.', font=('Segoe UI', 10))
        self.style.configure('TFrame', background='#f5f5f5')
        self.style.configure('TLabel', background='#f5f5f5')
        self.style.configure('Header.TLabel', font=('Segoe UI', 14, 'bold'))
        self.style.configure('TButton', padding=6)
        self.style.map('TButton', 
                      foreground=[('pressed', 'white'), ('active', 'white')],
                      background=[('pressed', '#0052cc'), ('active', '#0066ff')])
        
        # Initialize the application
        self.show_login_screen()
    
    def clear_window(self):
        """Clear all widgets from the window"""
        for widget in self.root.winfo_children():
            widget.destroy()
    
    def show_login_screen(self):
        """Display the login/registration screen"""
        self.clear_window()
        self.current_user = None
        self.current_role = None
        self.current_emp_id = None
        
        # Main container
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=40, pady=40)
        
        # Header
        header = ttk.Label(main_frame, text="Employee Management System", style='Header.TLabel')
        header.pack(pady=(0, 30))
        
        # Login Frame
        login_frame = ttk.LabelFrame(main_frame, text="Login", padding=20)
        login_frame.pack(pady=10)
        
        # Username
        ttk.Label(login_frame, text="Username:").grid(row=0, column=0, padx=10, pady=10, sticky=tk.E)
        self.username_entry = ttk.Entry(login_frame, width=25)
        self.username_entry.grid(row=0, column=1, padx=10, pady=10)
        
        # Password
        ttk.Label(login_frame, text="Password:").grid(row=1, column=0, padx=10, pady=10, sticky=tk.E)
        self.password_entry = ttk.Entry(login_frame, width=25, show="•")
        self.password_entry.grid(row=1, column=1, padx=10, pady=10)
        
        # Role Selection
        ttk.Label(login_frame, text="Role:").grid(row=2, column=0, padx=10, pady=10, sticky=tk.E)
        self.role_var = tk.StringVar(value="employee")
        ttk.Radiobutton(login_frame, text="Employee", variable=self.role_var, value="employee").grid(row=2, column=1, padx=10, pady=5, sticky=tk.W)
        ttk.Radiobutton(login_frame, text="Manager", variable=self.role_var, value="manager").grid(row=3, column=1, padx=10, pady=5, sticky=tk.W)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        login_btn = ttk.Button(button_frame, text="Login", command=self.login)
        login_btn.grid(row=0, column=0, padx=10)
        
        register_btn = ttk.Button(button_frame, text="Register", command=self.register_user_gui)
        register_btn.grid(row=0, column=1, padx=10)
        
        exit_btn = ttk.Button(button_frame, text="Exit", command=self.root.quit)
        exit_btn.grid(row=0, column=2, padx=10)
        
        # Focus on username field
        self.username_entry.focus_set()
    
    def register_user_gui(self):
        """Display user registration dialog"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Register User")
        dialog.geometry("400x300")
        dialog.resizable(False, False)
        
        ttk.Label(dialog, text="Register New User", style='Header.TLabel').pack(pady=10)
        
        frame = ttk.Frame(dialog)
        frame.pack(pady=10, padx=20)
        
        ttk.Label(frame, text="Username:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        new_user_entry = ttk.Entry(frame, width=25)
        new_user_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="Password:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
        new_pass_entry = ttk.Entry(frame, width=25, show="*")
        new_pass_entry.grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="Confirm Password:").grid(row=2, column=0, padx=5, pady=5, sticky=tk.E)
        confirm_pass_entry = ttk.Entry(frame, width=25, show="*")
        confirm_pass_entry.grid(row=2, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="Role:").grid(row=3, column=0, padx=5, pady=5, sticky=tk.E)
        role_var = tk.StringVar(value="employee")
        ttk.Radiobutton(frame, text="Employee", variable=role_var, value="employee").grid(row=3, column=1, padx=5, pady=2, sticky=tk.W)
        ttk.Radiobutton(frame, text="Manager", variable=role_var, value="manager").grid(row=4, column=1, padx=5, pady=2, sticky=tk.W)
        
        ttk.Label(frame, text="Employee ID (if employee):").grid(row=5, column=0, padx=5, pady=5, sticky=tk.E)
        emp_id_entry = ttk.Entry(frame, width=25)
        emp_id_entry.grid(row=5, column=1, padx=5, pady=5)
        
        def register():
            username = new_user_entry.get().strip()
            password = new_pass_entry.get().strip()
            confirm_pass = confirm_pass_entry.get().strip()
            role = role_var.get()
            emp_id = emp_id_entry.get().strip() if role == "employee" else None
            
            if not username or not password:
                messagebox.showerror("Error", "Username and password are required!")
                return
            
            if password != confirm_pass:
                messagebox.showerror("Error", "Passwords do not match!")
                return
            
            if role == "employee" and not emp_id:
                messagebox.showerror("Error", "Employee ID is required for employee role!")
                return
            
            if role == "employee":
                employees = self.load_employees()
                if emp_id not in employees:
                    messagebox.showerror("Error", "Employee ID not found in system!")
                    return
            
            self.register_user(username, password, role, emp_id)
            dialog.destroy()
            messagebox.showinfo("Success", "Account created successfully!")
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Register", command=register).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        new_user_entry.focus_set()
    
    def register_user(self, username, password, role, emp_id=None):
        """Register a new user"""
        with open(USER_FILE, "a") as file:
            file.write(f"{username},{password},{role},{emp_id if emp_id else ''}\n")
    
    def login(self):
        """Handle login process"""
        username = self.username_entry.get().strip()
        password = self.password_entry.get().strip()
        role = self.role_var.get()
        
        if not username or not password:
            messagebox.showerror("Error", "Username and password are required!")
            return
        
        if not os.path.exists(USER_FILE):
            messagebox.showinfo("Info", "No users found! Please register first.")
            return
        
        with open(USER_FILE, "r") as file:
            users = file.readlines()
        
        for user in users:
            parts = user.strip().split(",")
            if len(parts) >= 4:  # username,password,role,emp_id
                stored_user, stored_pass, stored_role, stored_emp_id = parts[:4]
                if (username == stored_user and password == stored_pass and 
                    role == stored_role):
                    self.current_user = username
                    self.current_role = stored_role
                    self.current_emp_id = stored_emp_id if stored_emp_id else None
                    messagebox.showinfo("Success", f"Login successful! Welcome, {username}")
                    self.show_dashboard()
                    return
        
        messagebox.showerror("Error", "Invalid credentials or role! Try again.")
    
    def show_dashboard(self):
        """Display the main dashboard based on user role"""
        self.clear_window()
        
        # Header
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(header_frame, text=f"Welcome, {self.current_user} ({self.current_role})", 
                 style='Header.TLabel').pack(side=tk.LEFT)
        
        logout_btn = ttk.Button(header_frame, text="Logout", command=self.show_login_screen)
        logout_btn.pack(side=tk.RIGHT, padx=5)
        
        # Main content
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        if self.current_role == "manager":
            # Employee Management Tab
            emp_frame = ttk.Frame(notebook)
            notebook.add(emp_frame, text="👥 Employee Management")
            self.setup_employee_tab(emp_frame)
            
            # Attendance Tab
            att_frame = ttk.Frame(notebook)
            notebook.add(att_frame, text="📅 Attendance")
            self.setup_attendance_tab(att_frame, manager_view=True)
            
            # Reports Tab
            report_frame = ttk.Frame(notebook)
            notebook.add(report_frame, text="📊 Reports")
            self.setup_reports_tab(report_frame)
            
            # Settings Tab
            settings_frame = ttk.Frame(notebook)
            notebook.add(settings_frame, text="⚙️ Settings")
            self.setup_settings_tab(settings_frame)
        else:
            # Employee can only view their own attendance
            att_frame = ttk.Frame(notebook)
            notebook.add(att_frame, text="📅 My Attendance")
            self.setup_attendance_tab(att_frame, manager_view=False)
    
    def setup_employee_tab(self, parent):
        """Setup the employee management tab (Manager only)"""
        # Add/Remove Frame
        manage_frame = ttk.LabelFrame(parent, text="Manage Employees", padding=10)
        manage_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(manage_frame, text="Add Employee", command=self.show_add_employee_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(manage_frame, text="Remove Employee", command=self.show_remove_employee_dialog).pack(side=tk.LEFT, padx=5)
        
        # Employee List
        list_frame = ttk.LabelFrame(parent, text="Employee List", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        columns = ("ID", "Name")
        self.employee_tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse")
        
        for col in columns:
            self.employee_tree.heading(col, text=col)
            self.employee_tree.column(col, width=150, anchor=tk.CENTER)
        
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.employee_tree.yview)
        self.employee_tree.configure(yscroll=scrollbar.set)
        
        self.employee_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.update_employee_list()
    
    def show_add_employee_dialog(self):
        """Show dialog to add new employee (Manager only)"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Employee")
        dialog.geometry("400x200")
        dialog.resizable(False, False)
        
        ttk.Label(dialog, text="Add New Employee", style='Header.TLabel').pack(pady=10)
        
        frame = ttk.Frame(dialog)
        frame.pack(pady=10, padx=20)
        
        ttk.Label(frame, text="Employee ID:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        emp_id_entry = ttk.Entry(frame, width=25)
        emp_id_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(frame, text="Employee Name:").grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
        emp_name_entry = ttk.Entry(frame, width=25)
        emp_name_entry.grid(row=1, column=1, padx=5, pady=5)
        
        def add_employee():
            emp_id = emp_id_entry.get().strip()
            emp_name = emp_name_entry.get().strip()
            
            if not emp_id or not emp_name:
                messagebox.showerror("Error", "Employee ID and Name are required!")
                return
            
            employees = self.load_employees()
            if emp_id in employees:
                messagebox.showerror("Error", "Employee ID already exists!")
                return
            
            with open(EMPLOYEE_FILE, "a") as file:
                file.write(f"{emp_id},{emp_name}\n")
            
            messagebox.showinfo("Success", f"Employee {emp_name} added successfully!")
            self.update_employee_list()
            dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Add", command=add_employee).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        emp_id_entry.focus_set()
    
    def show_remove_employee_dialog(self):
        """Show dialog to remove employee (Manager only)"""
        employees = self.load_employees()
        if not employees:
            messagebox.showinfo("Info", "No employees to remove.")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Remove Employee")
        dialog.geometry("400x200")
        dialog.resizable(False, False)
        
        ttk.Label(dialog, text="Remove Employee", style='Header.TLabel').pack(pady=10)
        
        frame = ttk.Frame(dialog)
        frame.pack(pady=10, padx=20)
        
        ttk.Label(frame, text="Employee ID:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        emp_id_entry = ttk.Entry(frame, width=25)
        emp_id_entry.grid(row=0, column=1, padx=5, pady=5)
        
        def remove_employee():
            emp_id = emp_id_entry.get().strip()
            
            if not emp_id:
                messagebox.showerror("Error", "Employee ID is required!")
                return
            
            if emp_id not in employees:
                messagebox.showerror("Error", "Employee ID not found!")
                return
            
            del employees[emp_id]
            with open(EMPLOYEE_FILE, "w") as file:
                for e_id, e_name in employees.items():
                    file.write(f"{e_id},{e_name}\n")
            
            messagebox.showinfo("Success", "Employee removed successfully!")
            self.update_employee_list()
            dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Remove", command=remove_employee).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        emp_id_entry.focus_set()
    
    def load_employees(self):
        """Load employees from file"""
        employees = {}
        if os.path.exists(EMPLOYEE_FILE):
            with open(EMPLOYEE_FILE, "r") as file:
                for line in file:
                    emp_id, emp_name = line.strip().split(",")
                    employees[emp_id] = emp_name
        return employees
    
    def update_employee_list(self):
        """Update the employee list in the treeview"""
        employees = self.load_employees()
        self.employee_tree.delete(*self.employee_tree.get_children())
        
        for emp_id, emp_name in employees.items():
            self.employee_tree.insert("", tk.END, values=(emp_id, emp_name))
    
    def setup_attendance_tab(self, parent, manager_view=True):
        """Setup the attendance tab"""
        if manager_view:
            # Mark Attendance Frame (Manager only)
            mark_frame = ttk.LabelFrame(parent, text="Mark Attendance", padding=10)
            mark_frame.pack(fill=tk.X, padx=10, pady=10)
            
            ttk.Button(mark_frame, text="Mark Attendance for Today", command=self.mark_attendance).pack(pady=5)
        
        # View Attendance Frame
        view_frame = ttk.LabelFrame(parent, text="Attendance Records" if manager_view else "My Attendance Records", padding=10)
        view_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        columns = ("Date", "Time", "Employee ID", "Employee Name", "Status")
        self.attendance_tree = ttk.Treeview(view_frame, columns=columns, show="headings", selectmode="extended")
        
        for col in columns:
            self.attendance_tree.heading(col, text=col)
            self.attendance_tree.column(col, width=120, anchor=tk.CENTER)
        
        scroll_y = ttk.Scrollbar(view_frame, orient=tk.VERTICAL, command=self.attendance_tree.yview)
        scroll_x = ttk.Scrollbar(view_frame, orient=tk.HORIZONTAL, command=self.attendance_tree.xview)
        self.attendance_tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.attendance_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Buttons for attendance
        btn_frame = ttk.Frame(view_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(btn_frame, text="Refresh", command=lambda: self.update_attendance_list(manager_view)).pack(side=tk.LEFT, padx=5)
        
        if manager_view:
            ttk.Button(btn_frame, text="Export to Excel", command=self.export_attendance).pack(side=tk.LEFT, padx=5)
        
        self.update_attendance_list(manager_view)
    
    def mark_attendance(self):
        """Mark attendance for all employees (Manager only)"""
        employees = self.load_employees()
        if not employees:
            messagebox.showinfo("Info", "No employees available to mark attendance.")
            return
        
        date = datetime.date.today().strftime("%d-%m-%Y")
        time = datetime.datetime.now().strftime("%H:%M:%S")
        records = []
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Mark Attendance")
        dialog.geometry("600x400")
        
        ttk.Label(dialog, text=f"Mark Attendance for {date}", style='Header.TLabel').pack(pady=10)
        
        # Create a frame for the employee list
        frame = ttk.Frame(dialog)
        frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Create a canvas and scrollbar
        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Create radio buttons for each employee
        self.attendance_vars = {}
        for emp_id, emp_name in employees.items():
            frame = ttk.Frame(scrollable_frame)
            frame.pack(fill=tk.X, padx=5, pady=2)
            
            ttk.Label(frame, text=f"{emp_name} ({emp_id})", width=30).pack(side=tk.LEFT)
            
            var = tk.StringVar(value="P")
            self.attendance_vars[emp_id] = var
            
            ttk.Radiobutton(frame, text="Present", variable=var, value="P").pack(side=tk.LEFT, padx=5)
            ttk.Radiobutton(frame, text="Absent", variable=var, value="A").pack(side=tk.LEFT, padx=5)
        
        def save_attendance():
            for emp_id, emp_name in employees.items():
                status = self.attendance_vars[emp_id].get()
                records.append({
                    "Date": date,
                    "Time": time,
                    "Employee ID": emp_id,
                    "Employee Name": emp_name,
                    "Status": status
                })
            
            self.save_attendance_to_excel(records)
            messagebox.showinfo("Success", "Attendance recorded successfully!")
            self.update_attendance_list(manager_view=True)
            dialog.destroy()
        
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Save Attendance", command=save_attendance).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def save_attendance_to_excel(self, records):
        """Save attendance records to Excel file"""
        df = pd.DataFrame(records)
        
        if os.path.exists(ATTENDANCE_FILE):
            existing_data = pd.read_excel(ATTENDANCE_FILE)
            df = pd.concat([existing_data, df], ignore_index=True)
        
        df.to_excel(ATTENDANCE_FILE, index=False)
    
    def update_attendance_list(self, manager_view=True):
        """Update the attendance records in the treeview"""
        if not os.path.exists(ATTENDANCE_FILE):
            return
        
        df = pd.read_excel(ATTENDANCE_FILE)
        self.attendance_tree.delete(*self.attendance_tree.get_children())
        
        if not manager_view and self.current_emp_id:
            # Employee view - only show their records
            df = df[df['Employee ID'] == self.current_emp_id]
        
        for _, row in df.iterrows():
            self.attendance_tree.insert("", tk.END, values=(
                row['Date'],
                row['Time'],
                row['Employee ID'],
                row['Employee Name'],
                row['Status']
            ))
    
    def export_attendance(self):
        """Export attendance data to a new Excel file (Manager only)"""
        if not os.path.exists(ATTENDANCE_FILE):
            messagebox.showinfo("Info", "No attendance data to export.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save Attendance File As"
        )
        
        if file_path:
            df = pd.read_excel(ATTENDANCE_FILE)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Attendance data exported to:\n{file_path}")
    
    def setup_reports_tab(self, parent):
        """Setup the reports tab (Manager only)"""
        # Generate Report Frame
        report_frame = ttk.LabelFrame(parent, text="Generate Reports", padding=10)
        report_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Button(report_frame, text="Generate Attendance Report", command=self.generate_attendance_report).pack(pady=5)
        
        # Report Display Frame
        display_frame = ttk.LabelFrame(parent, text="Attendance Report", padding=10)
        display_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        
        columns = ("Employee ID", "Employee Name", "Attendance Summary")
        self.report_tree = ttk.Treeview(display_frame, columns=columns, show="headings", selectmode="browse")
        
        for col in columns:
            self.report_tree.heading(col, text=col)
            self.report_tree.column(col, width=200, anchor=tk.CENTER)
        
        scrollbar = ttk.Scrollbar(display_frame, orient=tk.VERTICAL, command=self.report_tree.yview)
        self.report_tree.configure(yscrollcommand=scrollbar.set)
        
        self.report_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Export button
        btn_frame = ttk.Frame(display_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Button(btn_frame, text="Export Report", command=self.export_report).pack(side=tk.LEFT, padx=5)
    
    def generate_attendance_report(self):
        """Generate and display attendance report (Manager only)"""
        if not os.path.exists(ATTENDANCE_FILE):
            messagebox.showinfo("Info", "No attendance data available.")
            return
        
        df = pd.read_excel(ATTENDANCE_FILE)
        report = df.groupby(["Employee ID", "Employee Name"])["Status"].apply(
            lambda x: f"Present: {sum(x == 'P')}, Absent: {sum(x == 'A')}"
        ).reset_index()
        
        self.report_tree.delete(*self.report_tree.get_children())
        
        for _, row in report.iterrows():
            self.report_tree.insert("", tk.END, values=(
                row['Employee ID'],
                row['Employee Name'],
                row['Status']
            ))
    
    def export_report(self):
        """Export the generated report to Excel (Manager only)"""
        items = self.report_tree.get_children()
        if not items:
            messagebox.showinfo("Info", "No report data to export.")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="Save Report As"
        )
        
        if file_path:
            data = []
            for item in items:
                values = self.report_tree.item(item, 'values')
                data.append({
                    "Employee ID": values[0],
                    "Employee Name": values[1],
                    "Attendance Summary": values[2]
                })
            
            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", f"Report exported to:\n{file_path}")
    
    def setup_settings_tab(self, parent):
        """Setup the settings tab (Manager only)"""
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Backup section
        backup_frame = ttk.LabelFrame(frame, text="Data Backup & Restore", padding=15)
        backup_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(backup_frame, text="Create Backup", command=self.create_backup).pack(pady=5)
        ttk.Button(backup_frame, text="Restore from Backup", command=self.restore_backup).pack(pady=5)
        
        # Data management section
        data_frame = ttk.LabelFrame(frame, text="Data Management", padding=15)
        data_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(data_frame, text="Export All Data", command=self.export_all_data).pack(pady=5)
        ttk.Button(data_frame, text="Clear All Data", command=self.clear_all_data).pack(pady=5)
        
        # System section
        system_frame = ttk.LabelFrame(frame, text="System", padding=15)
        system_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(system_frame, text="About", command=self.show_about).pack(pady=5)
    
    def create_backup(self):
        """Create a backup of all data files (Manager only)"""
        backup_dir = filedialog.askdirectory(title="Select Backup Location")
        if not backup_dir:
            return
        
        try:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_folder = os.path.join(backup_dir, f"ems_backup_{timestamp}")
            os.makedirs(backup_folder)
            
            files_to_backup = [USER_FILE, EMPLOYEE_FILE, ATTENDANCE_FILE]
            
            for file in files_to_backup:
                if os.path.exists(file):
                    shutil.copy(file, backup_folder)
            
            messagebox.showinfo("Success", f"Backup created successfully at:\n{backup_folder}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create backup:\n{str(e)}")
    
    def restore_backup(self):
        """Restore data from a backup (Manager only)"""
        backup_dir = filedialog.askdirectory(title="Select Backup Location")
        if not backup_dir:
            return
        
        try:
            # Verify required files exist in backup
            required_files = [USER_FILE, EMPLOYEE_FILE, ATTENDANCE_FILE]
            missing_files = []
            
            for file in required_files:
                if not os.path.exists(os.path.join(backup_dir, os.path.basename(file))):
                    missing_files.append(file)
            
            if missing_files:
                messagebox.showerror("Error", f"Backup is missing required files:\n{', '.join(missing_files)}")
                return
            
            # Confirm with user
            if not messagebox.askyesno("Confirm", "This will overwrite all current data. Continue?"):
                return
            
            # Restore files
            for file in required_files:
                backup_file = os.path.join(backup_dir, os.path.basename(file))
                shutil.copy(backup_file, file)
            
            messagebox.showinfo("Success", "Data restored successfully!")
            self.show_login_screen()  # Restart to reload all data
        except Exception as e:
            messagebox.showerror("Error", f"Failed to restore backup:\n{str(e)}")
    
    def export_all_data(self):
        """Export all system data to Excel (Manager only)"""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Save All Data As"
        )
        
        if not file_path:
            return
        
        try:
            with pd.ExcelWriter(file_path) as writer:
                # Export users
                if os.path.exists(USER_FILE):
                    users = pd.read_csv(USER_FILE, header=None, names=["Username", "Password", "Role", "Employee ID"])
                    users.to_excel(writer, sheet_name="Users", index=False)
                
                # Export employees
                if os.path.exists(EMPLOYEE_FILE):
                    employees = pd.read_csv(EMPLOYEE_FILE, header=None, names=["Employee ID", "Employee Name"])
                    employees.to_excel(writer, sheet_name="Employees", index=False)
                
                # Export attendance
                if os.path.exists(ATTENDANCE_FILE):
                    attendance = pd.read_excel(ATTENDANCE_FILE)
                    attendance.to_excel(writer, sheet_name="Attendance", index=False)
            
            messagebox.showinfo("Success", f"All data exported to:\n{file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data:\n{str(e)}")
    
    def clear_all_data(self):
        """Clear all system data (with confirmation, Manager only)"""
        if not messagebox.askyesno("Confirm", "This will DELETE ALL DATA in the system. Are you sure?"):
            return
        
        try:
            if os.path.exists(USER_FILE):
                os.remove(USER_FILE)
            if os.path.exists(EMPLOYEE_FILE):
                os.remove(EMPLOYEE_FILE)
            if os.path.exists(ATTENDANCE_FILE):
                os.remove(ATTENDANCE_FILE)
            
            messagebox.showinfo("Success", "All data has been cleared.")
            self.show_login_screen()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to clear data:\n{str(e)}")
    
    def show_about(self):
        """Show about dialog"""
        about_text = (
            "Employee Management System\n"
            "Version 2.0\n\n"
            "Developed for MicroProject\n"
            "© 2023 All Rights Reserved"
        )
        
        messagebox.showinfo("About", about_text)

# Main application
if __name__ == "__main__":
    root = tk.Tk()
    app = EmployeeManagementSystem(root)
    root.mainloop()