import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc

def create_connection():
    try:
        conn = pyodbc.connect(
            'DRIVER={ODBC Driver 17 for SQL Server};'
            'SERVER=USHYDPNETHRAVA1;'
            'DATABASE=EmployeeDB;'
            'Trusted_Connection=yes;'
        )
        return conn
    except pyodbc.Error as e:
        messagebox.showerror("Database Connection Error", str(e))
        return None

class EmployeeManagementSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("Employee Management System")
        self.root.geometry("1250x600")
        self.root.config(bg="#f4f4f9")

        # GUI Setup
        self.setup_ui()

    def setup_ui(self):
        # Title
        title = tk.Label(self.root, text="Employee Management System", font=("Helvetica", 24), bg="#344955", fg="#f9aa33")
        title.pack(pady=20)

        # Frame for Form
        form_frame = tk.Frame(self.root, bg="#f4f4f9")
        form_frame.pack(pady=20)

        # Labels and Entries for Employee Data
        self.first_name_var = tk.StringVar()
        self.last_name_var = tk.StringVar()
        self.email_var = tk.StringVar()
        self.position_var = tk.StringVar()
        self.salary_var = tk.StringVar()  # Changed to StringVar for validation

        tk.Label(form_frame, text="First Name:", bg="#f4f4f9").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(form_frame, textvariable=self.first_name_var, width=25).grid(row=0, column=1, padx=10, pady=5)

        tk.Label(form_frame, text="Last Name:", bg="#f4f4f9").grid(row=1, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(form_frame, textvariable=self.last_name_var, width=25).grid(row=1, column=1, padx=10, pady=5)

        tk.Label(form_frame, text="Email:", bg="#f4f4f9").grid(row=2, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(form_frame, textvariable=self.email_var, width=25).grid(row=2, column=1, padx=10, pady=5)

        tk.Label(form_frame, text="Position:", bg="#f4f4f9").grid(row=3, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(form_frame, textvariable=self.position_var, width=25).grid(row=3, column=1, padx=10, pady=5)

        tk.Label(form_frame, text="Salary:", bg="#f4f4f9").grid(row=4, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(form_frame, textvariable=self.salary_var, width=25).grid(row=4, column=1, padx=10, pady=5)

        # Buttons for CRUD Operations
        tk.Button(form_frame, text="Add Employee", command=self.add_employee, bg="#344955", fg="#ffffff").grid(row=5, column=0, padx=10, pady=10)
        tk.Button(form_frame, text="Update Employee", command=self.update_employee, bg="#344955", fg="#ffffff").grid(row=5, column=1, padx=10, pady=10)
        tk.Button(form_frame, text="Delete Employee", command=self.delete_employee, bg="#344955", fg="#ffffff").grid(row=5, column=2, padx=10, pady=10)

        # Search Bar and Button
        search_frame = tk.Frame(self.root, bg="#f4f4f9")
        search_frame.pack(pady=10)

        self.search_name_var = tk.StringVar()
        self.search_position_var = tk.StringVar()

        tk.Label(search_frame, text="Search by Name:", bg="#f4f4f9").grid(row=0, column=0, padx=10, pady=5, sticky="w")
        tk.Entry(search_frame, textvariable=self.search_name_var, width=25).grid(row=0, column=1, padx=10, pady=5)

        tk.Label(search_frame, text="Search by Position:", bg="#f4f4f9").grid(row=0, column=2, padx=10, pady=5, sticky="w")
        tk.Entry(search_frame, textvariable=self.search_position_var, width=25).grid(row=0, column=3, padx=10, pady=5)

        tk.Button(search_frame, text="Search", command=self.search_employee, bg="#344955", fg="#ffffff").grid(row=0, column=4, padx=10, pady=10)

        # Treeview for Displaying Employees
        self.tree = ttk.Treeview(self.root, columns=("ID", "First Name", "Last Name", "Email", "Position", "Salary"), show="headings")
        self.tree.heading("ID", text="ID")
        self.tree.heading("First Name", text="First Name")
        self.tree.heading("Last Name", text="Last Name")
        self.tree.heading("Email", text="Email")
        self.tree.heading("Position", text="Position")
        self.tree.heading("Salary", text="Salary")
        self.tree.column("ID", width=0, stretch=tk.NO)  # Hide the ID column
        self.tree.pack(pady=20)

        self.tree.bind("<ButtonRelease-1>", self.select_employee)

        # Load employees initially
        self.view_employees()

    def add_employee(self):
        first_name = self.first_name_var.get()
        last_name = self.last_name_var.get()
        email = self.email_var.get()
        position = self.position_var.get()
        salary = self.salary_var.get()

        if first_name and last_name and email and position and self.validate_salary(salary):
            conn = create_connection()
            if conn:
                cursor = conn.cursor()
                try:
                    cursor.execute("""
                        INSERT INTO Employees (FirstName, LastName, Email, Position, Salary)
                        VALUES (?, ?, ?, ?, ?)
                    """, (first_name, last_name, email, position, salary))
                    conn.commit()
                    messagebox.showinfo("Success", "Employee added successfully!")
                    self.clear_form()
                    self.view_employees()  # Refresh employee list
                except pyodbc.Error as e:
                    messagebox.showerror("Database Error", str(e))
                finally:
                    conn.close()
        else:
            messagebox.showwarning("Input Error", "All fields are required and must be valid")

    def view_employees(self):
        conn = create_connection()
        if conn:
            cursor = conn.cursor()
            try:
                cursor.execute("SELECT * FROM Employees")
                rows = cursor.fetchall()
                self.update_treeview(rows)
            except pyodbc.Error as e:
                messagebox.showerror("Database Error", str(e))
            finally:
                conn.close()

    def update_employee(self):
        selected_item = self.tree.selection()
        if selected_item:
            selected_item = selected_item[0]
            values = self.tree.item(selected_item, "values")
            emp_id = values[0]
            first_name = self.first_name_var.get()
            last_name = self.last_name_var.get()
            email = self.email_var.get()
            position = self.position_var.get()
            salary = self.salary_var.get()

            if first_name and last_name and email and position and self.validate_salary(salary):
                conn = create_connection()
                if conn:
                    cursor = conn.cursor()
                    try:
                        cursor.execute("""
                            UPDATE Employees
                            SET FirstName = ?, LastName = ?, Email = ?, Position = ?, Salary = ?
                            WHERE EmployeeID = ?
                        """, (first_name, last_name, email, position, salary, emp_id))
                        conn.commit()
                        messagebox.showinfo("Success", "Employee updated successfully!")
                        self.clear_form()
                        self.view_employees()  # Refresh employee list
                    except pyodbc.Error as e:
                        messagebox.showerror("Database Error", str(e))
                    finally:
                        conn.close()
            else:
                messagebox.showwarning("Input Error", "All fields are required and must be valid")
        else:
            messagebox.showwarning("Selection Error", "No employee selected")

    def delete_employee(self):
        selected_item = self.tree.selection()
        if selected_item:
            selected_item = selected_item[0]
            values = self.tree.item(selected_item, "values")
            emp_id = values[0]

            conn = create_connection()
            if conn:
                cursor = conn.cursor()
                try:
                    cursor.execute("DELETE FROM Employees WHERE EmployeeID = ?", (emp_id,))
                    conn.commit()
                    messagebox.showinfo("Success", "Employee deleted successfully!")
                    self.clear_form()
                    self.view_employees()  # Refresh employee list
                except pyodbc.Error as e:
                    messagebox.showerror("Database Error", str(e))
                finally:
                    conn.close()
        else:
            messagebox.showwarning("Selection Error", "No employee selected")

    def select_employee(self, event):
        selected_item = self.tree.selection()
        if selected_item:
            selected_item = selected_item[0]
            values = self.tree.item(selected_item, "values")
            self.first_name_var.set(values[1])
            self.last_name_var.set(values[2])
            self.email_var.set(values[3])
            self.position_var.set(values[4])
            self.salary_var.set(values[5])

    def search_employee(self):
        search_name = self.search_name_var.get().lower()
        search_position = self.search_position_var.get().lower()

        conn = create_connection()
        if conn:
            cursor = conn.cursor()
            try:
                cursor.execute("""
                    SELECT * FROM Employees
                    WHERE LOWER(FirstName) LIKE ? OR LOWER(LastName) LIKE ? AND LOWER(Position) LIKE ?
                """, ('%' + search_name + '%', '%' + search_name + '%', '%' + search_position + '%'))
                rows = cursor.fetchall()
                self.update_treeview(rows)
            except pyodbc.Error as e:
                messagebox.showerror("Database Error", str(e))
            finally:
                conn.close()

    def update_treeview(self, employees):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for emp in employees:
            self.tree.insert("", "end", values=(emp.EmployeeID, emp.FirstName, emp.LastName, emp.Email, emp.Position, emp.Salary))

    def clear_form(self):
        self.first_name_var.set("")
        self.last_name_var.set("")
        self.email_var.set("")
        self.position_var.set("")
        self.salary_var.set("")

    def validate_salary(self, salary):
        try:
            float(salary)
            return True
        except ValueError:
            return False

if __name__ == "__main__":
    root = tk.Tk()
    root.resizable(0,0)
    app = EmployeeManagementSystem(root)
    root.mainloop()
