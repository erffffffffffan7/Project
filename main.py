import sqlite3
from datetime import datetime, timedelta
from tkinter import messagebox, filedialog, ttk
import customtkinter as ctk
import openpyxl

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

DB_NAME = "gym_master.db"


def init_db():
    with sqlite3.connect(DB_NAME) as conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS members (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                age INTEGER,
                phone TEXT,
                plan TEXT NOT NULL,
                join_date TEXT NOT NULL,
                expiry_date TEXT NOT NULL
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS checkins (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                member_id INTEGER NOT NULL,
                date TEXT NOT NULL,
                time TEXT NOT NULL,
                FOREIGN KEY (member_id) REFERENCES members (id)
            )
        """)
        conn.commit()


class LoginWindow(ctk.CTk):
    def __init__(self, on_login_success):
        super().__init__()
        self.title("Admin Login - Gym Management")
        self.geometry("400x380")
        self.resizable(False, False)
        self.on_login_success = on_login_success

        frame = ctk.CTkFrame(self, corner_radius=15, fg_color="#212121")
        frame.pack(padx=30, pady=30, fill="both", expand=True)

        ctk.CTkLabel(frame, text="🔒 Admin Login", font=("Tahoma", 18, "bold"), text_color="white").pack(pady=(35, 20))

        self.user_entry = ctk.CTkEntry(
            frame, placeholder_text="Username", width=260, height=40,
            corner_radius=8, font=("Tahoma", 12), fg_color="#2b2b2b"
        )
        self.user_entry.pack(pady=10)

        self.pass_entry = ctk.CTkEntry(
            frame, placeholder_text="Password", show="*", width=260, height=40,
            corner_radius=8, font=("Tahoma", 12), fg_color="#2b2b2b"
        )
        self.pass_entry.pack(pady=10)
        self.pass_entry.bind("<Return>", lambda event: self.check_login())

        btn_login = ctk.CTkButton(
            frame, text="Login to System", command=self.check_login,
            width=260, height=42, corner_radius=8, font=("Tahoma", 12, "bold"),
            fg_color="#1f6aa5", hover_color="#144870"
        )
        btn_login.pack(pady=25)

    def check_login(self):
        username = self.user_entry.get().strip()
        password = self.pass_entry.get().strip()

        if username == "admin" and password == "admin":
            self.withdraw()
            self.on_login_success()
        else:
            messagebox.showerror("Error", "Invalid username or password!")


class GymManagementApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Gym Management System - Pro Edition")
        self.geometry("1150x700")
        self.resizable(True, True)

        self.selected_member_id = None

        self.tabview = ctk.CTkTabview(self, corner_radius=15)
        self.tabview.pack(fill="both", expand=True, padx=15, pady=15)

        self.tab_manage = self.tabview.add("Manage Members")
        self.tab_checkin = self.tabview.add("Daily Check-in")
        self.tab_stats = self.tabview.add("Dashboard & Stats")

        self._setup_manage_tab()
        self._setup_checkin_tab()
        self._setup_stats_tab()

        self.load_members()
        self.load_todays_checkins()
        self.update_stats()

    def _setup_manage_tab(self):
        grid_frame = ctk.CTkFrame(self.tab_manage, fg_color="transparent")
        grid_frame.pack(fill="both", expand=True, padx=5, pady=5)

        left_frame = ctk.CTkFrame(grid_frame, width=320, corner_radius=12, fg_color="#212121")
        left_frame.pack(side="left", fill="y", padx=(0, 10), pady=5)

        ctk.CTkLabel(left_frame, text="Member Controls", font=("Tahoma", 15, "bold"), text_color="white").pack(pady=(15, 10))

        self.ent_name = ctk.CTkEntry(left_frame, placeholder_text="Full Name", width=270, height=35, corner_radius=8, fg_color="#2b2b2b")
        self.ent_name.pack(pady=6)

        self.ent_age = ctk.CTkEntry(left_frame, placeholder_text="Age", width=270, height=35, corner_radius=8, fg_color="#2b2b2b")
        self.ent_age.pack(pady=6)

        self.ent_phone = ctk.CTkEntry(left_frame, placeholder_text="Phone (e.g. 0912...)", width=270, height=35, corner_radius=8, fg_color="#2b2b2b")
        self.ent_phone.pack(pady=6)

        self.cmb_plan = ctk.CTkOptionMenu(
            left_frame, values=["1 Month", "3 Months", "6 Months", "1 Year"],
            width=270, height=35, corner_radius=8, fg_color="#1f6aa5", button_color="#144870"
        )
        self.cmb_plan.pack(pady=6)

        btn_row1 = ctk.CTkFrame(left_frame, fg_color="transparent")
        btn_row1.pack(pady=8)

        btn_reg = ctk.CTkButton(btn_row1, text="Register New", command=self.register_member, width=130, height=36, corner_radius=8, fg_color="#00a8ff", hover_color="#0086cc")
        btn_reg.pack(side="left", padx=4)

        btn_upd = ctk.CTkButton(btn_row1, text="Update Selected", command=self.update_member, width=130, height=36, corner_radius=8, fg_color="#e1b12c", hover_color="#c49a23", text_color="black")
        btn_upd.pack(side="left", padx=4)

        ctk.CTkFrame(left_frame, height=2, fg_color="#333333", width=270).pack(pady=8)

        self.ent_search = ctk.CTkEntry(left_frame, placeholder_text="Search by Name, Phone, ID", width=270, height=35, corner_radius=8, fg_color="#2b2b2b")
        self.ent_search.pack(pady=6)
        self.ent_search.bind("<Return>", lambda event: self.search_member())

        btn_search = ctk.CTkButton(left_frame, text="Search", command=self.search_member, width=270, height=36, corner_radius=8, fg_color="#00d2d3", hover_color="#00a8a9", text_color="black")
        btn_search.pack(pady=4)

        btn_clear = ctk.CTkButton(left_frame, text="Clear Search", command=self.clear_search, width=270, height=36, corner_radius=8, fg_color="#8395a7", hover_color="#637282")
        btn_clear.pack(pady=4)

        btn_excel = ctk.CTkButton(left_frame, text="Export to Excel", command=self.export_members_excel, width=270, height=36, corner_radius=8, fg_color="#10ac84", hover_color="#0e8c6b")
        btn_excel.pack(pady=8)

        btn_del = ctk.CTkButton(left_frame, text="Delete Selected", command=self.delete_member, width=270, height=36, corner_radius=8, fg_color="#ee5253", hover_color="#c83839")
        btn_del.pack(pady=(0, 15))

        right_frame = ctk.CTkFrame(grid_frame, corner_radius=12, fg_color="#212121")
        right_frame.pack(side="right", fill="both", expand=True, pady=5)

        cols = ("ID", "Name", "Age", "Phone", "Plan", "Join Date", "Expiry Date")

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="#2b2b2b", foreground="white", fieldbackground="#2b2b2b", rowheight=30, font=("Tahoma", 10))
        style.configure("Treeview.Heading", background="#1a1a1a", foreground="white", font=("Tahoma", 10, "bold"))
        style.map("Treeview", background=[("selected", "#1f6aa5")], foreground=[("selected", "white")])

        self.member_tree = ttk.Treeview(right_frame, columns=cols, show="headings")
        for col in cols:
            self.member_tree.heading(col, text=col)
            self.member_tree.column(col, anchor="center", width=100)

        self.member_tree.column("ID", width=50)
        self.member_tree.column("Name", width=150)

        self.member_tree.pack(fill="both", expand=True, padx=10, pady=10)
        self.member_tree.bind("<<TreeviewSelect>>", self.on_member_select)

    def _setup_checkin_tab(self):
        frame_input = ctk.CTkFrame(self.tab_checkin, corner_radius=12, fg_color="#212121")
        frame_input.pack(fill="x", padx=15, pady=15)

        ctk.CTkLabel(frame_input, text="Member ID:", font=("Tahoma", 13, "bold"), text_color="white").pack(side="left", padx=15, pady=15)

        self.checkin_entry = ctk.CTkEntry(frame_input, font=("Tahoma", 12), width=180, height=35, corner_radius=8, fg_color="#2b2b2b")
        self.checkin_entry.pack(side="left", padx=5, pady=15)
        self.checkin_entry.bind("<Return>", lambda event: self.checkin_member())

        btn_checkin = ctk.CTkButton(
            frame_input, text="Check-in ↵", command=self.checkin_member,
            fg_color="#10ac84", hover_color="#0e8c6b", font=("Tahoma", 12, "bold"), height=35, corner_radius=8
        )
        btn_checkin.pack(side="left", padx=15, pady=15)

        frame_table = ctk.CTkFrame(self.tab_checkin, corner_radius=12, fg_color="#212121")
        frame_table.pack(fill="both", expand=True, padx=15, pady=10)

        ctk.CTkLabel(frame_table, text="📋 Today's Check-ins:", font=("Tahoma", 13, "bold"), text_color="white").pack(anchor="w", padx=15, pady=10)

        cols = ("Time", "Member ID", "Name", "Plan")
        self.checkin_tree = ttk.Treeview(frame_table, columns=cols, show="headings", height=10)
        for col in cols:
            self.checkin_tree.heading(col, text=col)
            self.checkin_tree.column(col, anchor="center")

        self.checkin_tree.pack(fill="both", expand=True, padx=15, pady=10)

    def _setup_stats_tab(self):
        frame_stats = ctk.CTkFrame(self.tab_stats, corner_radius=15, fg_color="#212121")
        frame_stats.pack(fill="both", expand=True, padx=20, pady=20)

        self.lbl_stat_total = ctk.CTkLabel(frame_stats, text="Total Members: 0", font=("Tahoma", 16, "bold"), text_color="white")
        self.lbl_stat_total.pack(pady=20)

        self.lbl_stat_active = ctk.CTkLabel(frame_stats, text="Active Members: 0", font=("Tahoma", 16, "bold"), text_color="#10ac84")
        self.lbl_stat_active.pack(pady=20)

        self.lbl_stat_expired = ctk.CTkLabel(frame_stats, text="Expired Members: 0", font=("Tahoma", 16, "bold"), text_color="#ee5253")
        self.lbl_stat_expired.pack(pady=20)

        self.lbl_stat_today = ctk.CTkLabel(frame_stats, text="Today Check-ins: 0", font=("Tahoma", 16, "bold"), text_color="#00d2d3")
        self.lbl_stat_today.pack(pady=20)

    def load_members(self):
        for row in self.member_tree.get_children():
            self.member_tree.delete(row)
        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM members ORDER BY id DESC")
            for row in cursor.fetchall():
                self.member_tree.insert("", "end", values=row)

    def register_member(self):
        name = self.ent_name.get().strip()
        age = self.ent_age.get().strip()
        phone = self.ent_phone.get().strip()
        plan = self.cmb_plan.get()

        if not name:
            messagebox.showwarning("Warning", "Full Name is required!")
            return

        days_dict = {"1 Month": 30, "3 Months": 90, "6 Months": 180, "1 Year": 365}
        days = days_dict.get(plan, 30)

        join_date = datetime.now().strftime("%Y-%m-%d")
        expiry_date = (datetime.now() + timedelta(days=days)).strftime("%Y-%m-%d")

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO members (name, age, phone, plan, join_date, expiry_date)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (name, age, phone, plan, join_date, expiry_date))
            conn.commit()

        self.clear_entries()
        self.load_members()
        self.update_stats()
        messagebox.showinfo("Success", "Member registered successfully!")

    def on_member_select(self, event):
        selected = self.member_tree.selection()
        if not selected:
            return
        item = self.member_tree.item(selected[0])
        vals = item["values"]

        self.selected_member_id = vals[0]
        self.clear_entries(keep_id=True)

        self.ent_name.insert(0, vals[1])
        if vals[2] is not None: self.ent_age.insert(0, str(vals[2]))
        if vals[3] is not None: self.ent_phone.insert(0, str(vals[3]))
        if vals[4]: self.cmb_plan.set(str(vals[4]))

    def update_member(self):
        if not self.selected_member_id:
            messagebox.showwarning("Warning", "Please select a member from the table first!")
            return

        name = self.ent_name.get().strip()
        age = self.ent_age.get().strip()
        phone = self.ent_phone.get().strip()
        plan = self.cmb_plan.get()

        if not name:
            messagebox.showwarning("Warning", "Full Name is required!")
            return

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                UPDATE members 
                SET name=?, age=?, phone=?, plan=? 
                WHERE id=?
            """, (name, age, phone, plan, self.selected_member_id))
            conn.commit()

        self.clear_entries()
        self.load_members()
        messagebox.showinfo("Success", "Member details updated successfully!")

    def delete_member(self):
        if not self.selected_member_id:
            messagebox.showwarning("Warning", "Please select a member from the table to delete!")
            return

        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete Member ID: {self.selected_member_id}?"):
            with sqlite3.connect(DB_NAME) as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM members WHERE id=?", (self.selected_member_id,))
                cursor.execute("DELETE FROM checkins WHERE member_id=?", (self.selected_member_id,))
                conn.commit()

            self.clear_entries()
            self.load_members()
            self.update_stats()
            messagebox.showinfo("Success", "Member deleted successfully!")

    def search_member(self):
        query = self.ent_search.get().strip()
        if not query:
            self.load_members()
            return

        for row in self.member_tree.get_children():
            self.member_tree.delete(row)

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT * FROM members 
                WHERE name LIKE ? OR phone LIKE ? OR id LIKE ?
            """, (f"%{query}%", f"%{query}%", f"%{query}%"))
            for row in cursor.fetchall():
                self.member_tree.insert("", "end", values=row)

    def clear_search(self):
        self.ent_search.delete(0, "end")
        self.load_members()

    def checkin_member(self):
        mem_id = self.checkin_entry.get().strip()
        if not mem_id:
            messagebox.showwarning("Warning", "Please enter Member ID!")
            return

        today = datetime.now().strftime("%Y-%m-%d")
        time_now = datetime.now().strftime("%H:%M:%S")

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name, expiry_date FROM members WHERE id = ?", (mem_id,))
            member = cursor.fetchone()

            if not member:
                messagebox.showerror("Error", "Member ID not found!")
                return

            mem_name, expiry = member
            if expiry < today:
                messagebox.showwarning("Access Denied", f"Membership for {mem_name} has expired!")
                return

            cursor.execute("INSERT INTO checkins (member_id, date, time) VALUES (?, ?, ?)", (mem_id, today, time_now))
            conn.commit()

        self.checkin_entry.delete(0, "end")
        self.load_todays_checkins()
        self.update_stats()
        messagebox.showinfo("Success", f"Check-in recorded for {mem_name}")

    def load_todays_checkins(self):
        for row in self.checkin_tree.get_children():
            self.checkin_tree.delete(row)

        today = datetime.now().strftime("%Y-%m-%d")
        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT checkins.time, members.id, members.name, members.plan 
                FROM checkins 
                JOIN members ON checkins.member_id = members.id 
                WHERE checkins.date = ? 
                ORDER BY checkins.time DESC
            """, (today,))
            for row in cursor.fetchall():
                self.checkin_tree.insert("", "end", values=row)

    def update_stats(self):
        today = datetime.now().strftime("%Y-%m-%d")
        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM members")
            total = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM members WHERE expiry_date >= ?", (today,))
            active = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM members WHERE expiry_date < ?", (today,))
            expired = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM checkins WHERE date = ?", (today,))
            checkins = cursor.fetchone()[0]

        self.lbl_stat_total.configure(text=f"Total Members: {total}")
        self.lbl_stat_active.configure(text=f"Active Members: {active}")
        self.lbl_stat_expired.configure(text=f"Expired Members: {expired}")
        self.lbl_stat_today.configure(text=f"Today Check-ins: {checkins}")

    def export_members_excel(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
        if not file_path:
            return

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Members List"

        ws.append(["ID", "Name", "Age", "Phone", "Plan", "Join Date", "Expiry Date"])

        with sqlite3.connect(DB_NAME) as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM members")
            for row in cursor.fetchall():
                ws.append(list(row))

        wb.save(file_path)
        messagebox.showinfo("Export Success", "Members list exported to Excel successfully!")

    def clear_entries(self, keep_id=False):
        if not keep_id:
            self.selected_member_id = None
        self.ent_name.delete(0, "end")
        self.ent_age.delete(0, "end")
        self.ent_phone.delete(0, "end")


def start_app():
    app = GymManagementApp()
    app.mainloop()


if __name__ == "__main__":
    init_db()
    login = LoginWindow(on_login_success=start_app)
    login.mainloop()