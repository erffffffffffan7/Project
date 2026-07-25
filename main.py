import customtkinter as ctk
from tkinter import messagebox, ttk, filedialog
import sqlite3
import re
import csv
from datetime import datetime, timedelta

# =========================
# System Settings
# =========================
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

DB_NAME = "gym.db"

def init_db():
    """Ensures the database and tables exist, and upgrades old databases."""
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        
        # Create or verify main members table
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS members (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            age TEXT,
            phone TEXT,
            membership TEXT
        )
        """)
        
        # Upgrade old databases seamlessly (adds date tracking if missing)
        try:
            cursor.execute("ALTER TABLE members ADD COLUMN join_date TEXT")
            cursor.execute("ALTER TABLE members ADD COLUMN expiry_date TEXT")
        except sqlite3.OperationalError:
            pass # Columns already exist, no upgrade needed

        # Create table for daily check-ins
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS checkins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            member_id INTEGER,
            date TEXT,
            time TEXT,
            FOREIGN KEY(member_id) REFERENCES members(id)
        )
        """)
        conn.commit()

def calculate_expiry(membership_plan):
    """Calculates the expiry date based on the chosen plan."""
    today = datetime.now()
    if membership_plan == "1 Month":
        expiry = today + timedelta(days=30)
    elif membership_plan == "3 Months":
        expiry = today + timedelta(days=90)
    elif membership_plan == "6 Months":
        expiry = today + timedelta(days=180)
    elif membership_plan == "1 Year":
        expiry = today + timedelta(days=365)
    elif membership_plan == "VIP Plan":
        expiry = today + timedelta(days=3650) # 10 years for VIP
    else:
        expiry = today
    return today.strftime("%Y-%m-%d"), expiry.strftime("%Y-%m-%d")

# =========================
# Main Application Class
# =========================
class GymApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Gym Management System - Pro Edition")
        self.geometry("1300x850") 
        
        self.withdraw() 
        init_db()
        self.open_login_window()

    def open_login_window(self):
        self.login_window = ctk.CTkToplevel(self)
        self.login_window.geometry("400x350")
        self.login_window.title("Admin Login")
        self.login_window.protocol("WM_DELETE_WINDOW", self.quit)

        ctk.CTkLabel(self.login_window, text="Admin Login", font=("Arial", 28, "bold")).pack(pady=(30, 20))

        self.username_entry = ctk.CTkEntry(self.login_window, width=250, placeholder_text="Username")
        self.username_entry.pack(pady=10)

        self.password_entry = ctk.CTkEntry(self.login_window, width=250, placeholder_text="Password", show="*")
        self.password_entry.pack(pady=10)

        ctk.CTkButton(self.login_window, text="Login", command=self.check_login).pack(pady=30)

    def check_login(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        if username == "admin" and password == "1234":
            self.login_window.destroy()
            self.deiconify() 
            self.build_main_ui() 
        else:
            messagebox.showerror("Login Failed", "Wrong Username or Password.")

    def build_main_ui(self):
        """Builds the Tabbed Interface."""
        self.tabs = ctk.CTkTabview(self)
        self.tabs.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_manage = self.tabs.add("Manage Members")
        self.tab_checkin = self.tabs.add("Daily Check-in")
        self.tab_stats = self.tabs.add("Dashboard & Stats")

        # Configure tabs
        self.build_manage_tab()
        self.build_checkin_tab()
        self.build_stats_tab()
        
        # Refresh data when clicking tabs
        self.tabs.configure(command=self.refresh_all_data)

    def refresh_all_data(self):
        self.load_members()
        self.load_todays_checkins()
        self.update_stats()

    # =========================
    # TAB 1: Manage Members
    # =========================
    def build_manage_tab(self):
        left_panel = ctk.CTkFrame(self.tab_manage, width=400)
        left_panel.pack(side="left", fill="y", padx=10, pady=10)

        ctk.CTkLabel(left_panel, text="Member Controls", font=("Arial", 20, "bold")).pack(pady=10)

        # Selected ID Tracker (Invisible)
        self.selected_member_id = None

        self.name_entry = ctk.CTkEntry(left_panel, width=280, placeholder_text="Full Name")
        self.name_entry.pack(pady=5)

        self.age_entry = ctk.CTkEntry(left_panel, width=280, placeholder_text="Age")
        self.age_entry.pack(pady=5)

        self.phone_entry = ctk.CTkEntry(left_panel, width=280, placeholder_text="Phone (e.g. 0912...)")
        self.phone_entry.pack(pady=5)

        self.membership_option = ctk.CTkOptionMenu(
            left_panel, width=280,
            values=["1 Month", "3 Months", "6 Months", "1 Year", "VIP Plan"]
        )
        self.membership_option.pack(pady=5)

        # Action Buttons
        btn_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        btn_frame.pack(pady=15)
        
        ctk.CTkButton(btn_frame, text="Register New", width=135, command=self.register_member).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Update Selected", width=135, fg_color="#c48a16", hover_color="#9c6d12", command=self.update_member).pack(side="left", padx=5)

        ctk.CTkFrame(left_panel, height=2, fg_color="gray").pack(fill="x", pady=15, padx=20) # Divider

        # Search & Export
        self.search_entry = ctk.CTkEntry(left_panel, width=280, placeholder_text="Search by Name or Phone")
        self.search_entry.pack(pady=5)
        ctk.CTkButton(left_panel, text="Search", width=280, command=self.search_member).pack(pady=5)
        ctk.CTkButton(left_panel, text="Clear Search", width=280, fg_color="gray", command=self.load_members).pack(pady=5)
        
        ctk.CTkButton(left_panel, text="Export to Excel (CSV)", width=280, fg_color="#1f7a43", hover_color="#14522d", command=self.export_csv).pack(pady=(20, 5))
        ctk.CTkButton(left_panel, text="Delete Selected", width=280, fg_color="#c93434", hover_color="#962626", command=self.delete_member).pack(pady=5)

        # Right Panel: Data Table
        right_panel = ctk.CTkFrame(self.tab_manage)
        right_panel.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        # Treeview Setup
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", rowheight=30, fieldbackground="#2b2b2b", borderwidth=0)
        style.configure("Treeview.Heading", background="#1f538d", foreground="white", font=('Arial', 11, 'bold'))
        style.map('Treeview', background=[('selected', '#14375e')])

        self.tree = ttk.Treeview(right_panel, columns=("ID", "Name", "Age", "Phone", "Plan", "Join Date", "Expiry Date"), show="headings")
        
        columns = {"ID": 40, "Name": 180, "Age": 50, "Phone": 120, "Plan": 100, "Join Date": 100, "Expiry Date": 100}
        for col, width in columns.items():
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width, anchor="center")

        # Color Tags for Expiry
        self.tree.tag_configure('expired', background='#5c1a1a') 
        self.tree.tag_configure('active', background='#2b2b2b')

        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_select)
        
        self.load_members()

    # =========================
    # TAB 2: Check-In System
    # =========================
    def build_checkin_tab(self):
        top_frame = ctk.CTkFrame(self.tab_checkin)
        top_frame.pack(fill="x", padx=20, pady=20)

        ctk.CTkLabel(top_frame, text="Quick Check-In", font=("Arial", 24, "bold")).pack(pady=10)
        
        self.checkin_entry = ctk.CTkEntry(top_frame, width=350, placeholder_text="Enter Member ID or Phone Number", font=("Arial", 16))
        self.checkin_entry.pack(pady=10)
        
        ctk.CTkButton(top_frame, text="Log Attendance", font=("Arial", 16, "bold"), width=200, height=40, command=self.log_checkin).pack(pady=10)

        ctk.CTkLabel(self.tab_checkin, text="Today's Attendance", font=("Arial", 18, "bold")).pack(pady=(20, 5))

        self.checkin_tree = ttk.Treeview(self.tab_checkin, columns=("Time", "ID", "Name", "Plan"), show="headings")
        for col in ("Time", "ID", "Name", "Plan"):
            self.checkin_tree.heading(col, text=col)
            self.checkin_tree.column(col, anchor="center")
        self.checkin_tree.pack(fill="both", expand=True, padx=20, pady=10)
        self.load_todays_checkins()

    # =========================
    # TAB 3: Dashboard & Stats
    # =========================
    def build_stats_tab(self):
        self.stats_frame = ctk.CTkFrame(self.tab_stats)
        self.stats_frame.pack(fill="both", expand=True, padx=40, pady=40)

        ctk.CTkLabel(self.stats_frame, text="Gym Overview", font=("Arial", 32, "bold")).pack(pady=20)

        self.lbl_total = ctk.CTkLabel(self.stats_frame, text="Total Members: 0", font=("Arial", 22))
        self.lbl_total.pack(pady=10)

        self.lbl_active = ctk.CTkLabel(self.stats_frame, text="Active Members: 0", font=("Arial", 22), text_color="#43d675")
        self.lbl_active.pack(pady=10)

        self.lbl_expired = ctk.CTkLabel(self.stats_frame, text="Expired Members: 0", font=("Arial", 22), text_color="#d64343")
        self.lbl_expired.pack(pady=10)
        
        self.lbl_today = ctk.CTkLabel(self.stats_frame, text="Check-ins Today: 0", font=("Arial", 22), text_color="#42a4f5")
        self.lbl_today.pack(pady=10)

    # =========================
    # Core Functions
    # =========================
    def load_members(self, search_query=None):
        for row in self.tree.get_children():
            self.tree.delete(row)

        today_str = datetime.now().strftime("%Y-%m-%d")

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            if search_query:
                cursor.execute("SELECT * FROM members WHERE name LIKE ? OR phone LIKE ?", (f'%{search_query}%', f'%{search_query}%'))
            else:
                cursor.execute("SELECT * FROM members")
            
            for member in cursor.fetchall():
                expiry = member[6]
                status_tag = 'expired' if (expiry and expiry < today_str) else 'active'
                self.tree.insert("", "end", values=member, tags=(status_tag,))

    def on_tree_select(self, event):
        selected = self.tree.selection()
        if selected:
            item = self.tree.item(selected[0])['values']
            self.selected_member_id = item[0]
            
            self.name_entry.delete(0, "end")
            self.name_entry.insert(0, item[1])
            
            self.age_entry.delete(0, "end")
            self.age_entry.insert(0, item[2])
            
            self.phone_entry.delete(0, "end")
            self.phone_entry.insert(0, f"0{item[3]}" if len(str(item[3])) == 10 else item[3]) 
            
            self.membership_option.set(item[4])

    def validate_inputs(self, name, age, phone):
        if not name: 
            return False, "Please enter a name."
        
        # تایید اصالت اسم: فقط حروف انگلیسی، حروف فارسی و اسپیس مجاز هستند
        name_pattern = r"^[a-zA-Z\s\u0600-\u06FF]+$"
        if not re.match(name_pattern, name):
            return False, "Name can only contain letters and spaces (English or Persian)."

        if not age.isdigit() or not (10 <= int(age) <= 120): 
            return False, "Enter a valid age (10-120)."
        
        if not re.match(r"^(0|98|\+98)?9\d{9}$", phone): 
            return False, "Enter a valid Iranian mobile number."
        
        return True, ""

    def register_member(self):
        name, age, phone = self.name_entry.get().strip(), self.age_entry.get().strip(), self.phone_entry.get().strip()
        membership = self.membership_option.get()

        valid, msg = self.validate_inputs(name, age, phone)
        if not valid:
            messagebox.showerror("Validation Error", msg)
            return

        join_date, expiry_date = calculate_expiry(membership)

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("INSERT INTO members (name, age, phone, membership, join_date, expiry_date) VALUES (?, ?, ?, ?, ?, ?)", 
                           (name, age, phone, membership, join_date, expiry_date))
            conn.commit()

        self.refresh_all_data()
        self.clear_inputs()
        messagebox.showinfo("Success", f"{name} registered!\nExpires on: {expiry_date}")

    def update_member(self):
        if not self.selected_member_id:
            messagebox.showerror("Error", "Please select a member from the table first.")
            return

        name, age, phone = self.name_entry.get().strip(), self.age_entry.get().strip(), self.phone_entry.get().strip()
        membership = self.membership_option.get()

        valid, msg = self.validate_inputs(name, age, phone)
        if not valid:
            messagebox.showerror("Validation Error", msg)
            return

        join_date, expiry_date = calculate_expiry(membership)

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE members SET name=?, age=?, phone=?, membership=?, join_date=?, expiry_date=? WHERE id=?
            """, (name, age, phone, membership, join_date, expiry_date, self.selected_member_id))
            conn.commit()

        self.refresh_all_data()
        self.clear_inputs()
        messagebox.showinfo("Success", "Member updated successfully!")

    def delete_member(self):
        if not self.selected_member_id:
            messagebox.showerror("Error", "Please select a member from the table first.")
            return

        if messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this member permanently?"):
            with sqlite3.connect(DB_NAME) as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM members WHERE id = ?", (self.selected_member_id,))
                cursor.execute("DELETE FROM checkins WHERE member_id = ?", (self.selected_member_id,)) 
                conn.commit()
            self.refresh_all_data()
            self.clear_inputs()

    def search_member(self):
        self.search_entry.get().strip()
        self.load_members(search_query=self.search_entry.get().strip())

    def clear_inputs(self):
        self.selected_member_id = None
        self.name_entry.delete(0, "end")
        self.age_entry.delete(0, "end")
        self.phone_entry.delete(0, "end")
        self.search_entry.delete(0, "end")

    def export_csv(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")], title="Save Database as Excel")
        if not filepath: return

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM members")
            rows = cursor.fetchall()

        with open(filepath, 'w', newline='', encoding='utf-8-sig') as file:
            writer = csv.writer(file)
            writer.writerow(["ID", "Name", "Age", "Phone", "Plan", "Join Date", "Expiry Date"])
            writer.writerows(rows)
            
        messagebox.showinfo("Export Successful", f"Database exported to:\n{filepath}")

    # =========================
    # Check-In Logic
    # =========================
    def log_checkin(self):
        query = self.checkin_entry.get().strip()
        if not query: return
        
        today = datetime.now().strftime("%Y-%m-%d")
        time_now = datetime.now().strftime("%H:%M:%S")

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT id, name, expiry_date FROM members WHERE id=? OR phone=?", (query, query))
            member = cursor.fetchone()

            if not member:
                messagebox.showerror("Not Found", "No member found with that ID or Phone.")
                return

            mem_id, mem_name, expiry = member

            if expiry and expiry < today:
                messagebox.showwarning("Access Denied", f"{mem_name}'s membership expired on {expiry}!\nPlease renew.")
                return

            cursor.execute("INSERT INTO checkins (member_id, date, time) VALUES (?, ?, ?)", (mem_id, today, time_now))
            conn.commit()
        
        self.checkin_entry.delete(0, "end")
        self.load_todays_checkins()
        self.update_stats()
        messagebox.showinfo("Welcome!", f"Checked in: {mem_name} at {time_now}")

    def load_todays_checkins(self):
        for row in self.checkin_tree.get_children():
            self.checkin_tree.delete(row)

        today = datetime.now().strftime("%Y-%m-%d")
        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT checkins.time, members.id, members.name, members.membership 
                FROM checkins 
                JOIN members ON checkins.member_id = members.id 
                WHERE checkins.date = ? 
                ORDER BY checkins.time DESC
            """, (today,))
            
            for row in cursor.fetchall():
                self.checkin_tree.insert("", "end", values=row)

    # =========================
    # Dashboard Logic
    # =========================
    def update_stats(self):
        today = datetime.now().strftime("%Y-%m-%d")
        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            
            cursor.execute("SELECT COUNT(*) FROM members")
            total = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM members WHERE expiry_date >= ? OR expiry_date IS NULL", (today,))
            active = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM members WHERE expiry_date < ?", (today,))
            expired = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM checkins WHERE date = ?", (today,))
            checkins = cursor.fetchone()[0]

        self.lbl_total.configure(text=f"Total Members: {total}")
        self.lbl_active.configure(text=f"Active Members: {active}")
        self.lbl_expired.configure(text=f"Expired Members: {expired}")
        self.lbl_today.configure(text=f"Check-ins Today: {checkins}")

# =========================
# Run Application
# =========================
if __name__ == "__main__":
    app = GymApp()
    app.mainloop()