import os
import sqlite3
import datetime
import pandas as pd
import customtkinter as ctk
from tkinter import messagebox, filedialog, ttk
from pypinyin import lazy_pinyin
from tkcalendar import DateEntry

# 设置简约主题
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

DB_FILE = "data.db"


class SignInSystem(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("自习签到系统")
        self.geometry("950x700")

        # 数据库初始化
        self.init_db()

        # 顶部控制面板
        self.top_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.top_frame.pack(pady=10, padx=20, fill="x")

        self.btn_import = ctk.CTkButton(self.top_frame, text="导入名单", command=self.import_students, width=80)
        self.btn_import.pack(side="left", padx=5)

        self.btn_late = ctk.CTkButton(self.top_frame, text="迟到登记", command=self.open_late_window,
                                      fg_color="#CD5C5C", hover_color="#8B0000", width=80)
        self.btn_late.pack(side="left", padx=5)

        self.btn_history = ctk.CTkButton(self.top_frame, text="查看历史记录", command=self.show_history, width=100)
        self.btn_history.pack(side="left", padx=5)

        self.btn_export = ctk.CTkButton(self.top_frame, text="数据汇总", command=self.export_data, width=80)
        self.btn_export.pack(side="left", padx=5)

        # 状态显示
        self.lbl_stats = ctk.CTkLabel(self.top_frame, text="总人数: 0 | 已签到: 0 | 签到率: 0%",
                                      font=("Arial", 16, "bold"))
        self.lbl_stats.pack(side="right", padx=10)

        # 泡泡区域 (自动换行滚动区域)
        self.scroll_frame = ctk.CTkScrollableFrame(self)
        self.scroll_frame.pack(pady=10, padx=20, fill="both", expand=True)

        # 存储按钮的字典
        self.bubbles = {}
        self.load_today_data()

    def init_db(self):
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute("CREATE TABLE IF NOT EXISTS students (name TEXT PRIMARY KEY)")
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sign_ins (
                date TEXT,
                name TEXT,
                slot1_time TEXT,
                slot2_time TEXT,
                PRIMARY KEY (date, name)
            )
        """)

        # 新建带有迟到类型和备注的表
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS late_records (
                date TEXT,
                name TEXT,
                late_type TEXT,
                remark TEXT,
                PRIMARY KEY (date, name)
            )
        """)

        # 自动升级旧数据库（如果你之前运行过旧版代码，这会自动为你添加新列）
        try:
            cursor.execute("ALTER TABLE late_records ADD COLUMN late_type TEXT")
            cursor.execute("ALTER TABLE late_records ADD COLUMN remark TEXT")
        except sqlite3.OperationalError:
            pass  # 列已存在，忽略报错

        cursor.execute("CREATE TABLE IF NOT EXISTS daily_status (date TEXT PRIMARY KEY, late_saved INTEGER)")
        conn.commit()
        conn.close()

    def get_current_date(self):
        return datetime.date.today().strftime("%Y-%m-%d")

    def load_today_data(self):
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        self.bubbles.clear()

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        cursor.execute("SELECT name FROM students")
        students = [row[0] for row in cursor.fetchall()]
        students = sorted(students, key=lambda x: ''.join(lazy_pinyin(x)))

        today = self.get_current_date()
        cursor.execute("SELECT name, slot1_time, slot2_time FROM sign_ins WHERE date=?", (today,))
        today_data = {row[0]: (row[1], row[2]) for row in cursor.fetchall()}
        conn.close()

        total = len(students)
        signed_in_count = 0

        # 每行显示 6 个泡泡
        columns = 6
        for index, name in enumerate(students):
            row = index // columns
            col = index % columns

            slot1, slot2 = today_data.get(name, (None, None))

            if slot1 and slot2:
                bg_color = "#4682B4"
                text_color = "white"
                display_text = f"{name}\n① {slot1}\n② {slot2}"
                signed_in_count += 1
            elif slot1:
                bg_color = "#87CEFA"
                text_color = "black"
                display_text = f"{name}\n① {slot1}\n② --:--:--"
                signed_in_count += 1
            else:
                bg_color = "#E0E0E0"
                text_color = "black"
                display_text = f"{name}\n未签到"

            btn = ctk.CTkButton(
                self.scroll_frame,
                text=display_text,
                width=120, height=80,
                fg_color=bg_color,
                text_color=text_color,
                font=("Arial", 14),
                command=lambda n=name: self.handle_sign_in(n)
            )
            btn.grid(row=row, column=col, padx=10, pady=10)
            self.bubbles[name] = {"btn": btn}

        rate = (signed_in_count / total * 100) if total > 0 else 0
        self.lbl_stats.configure(text=f"总人数: {total} | 已签到: {signed_in_count} | 签到率: {rate:.1f}%")

    def import_students(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath:
            return

        try:
            df = pd.read_excel(filepath)
            if "姓名" not in df.columns:
                messagebox.showerror("错误", "Excel表中必须包含“姓名”列！")
                return

            names = df["姓名"].dropna().astype(str).tolist()

            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM students")
            for name in names:
                cursor.execute("INSERT OR IGNORE INTO students (name) VALUES (?)", (name.strip(),))
            conn.commit()
            conn.close()

            self.load_today_data()
            messagebox.showinfo("成功", f"成功导入 {len(names)} 名学生！")
        except Exception as e:
            messagebox.showerror("导入失败", str(e))

    def handle_sign_in(self, name):
        now = datetime.datetime.now()
        current_time = now.time()

        slot1_start, slot1_end = datetime.time(6, 0), datetime.time(6, 51)
        slot2_start, slot2_end = datetime.time(17, 20), datetime.time(17, 51)

        is_slot1 = slot1_start <= current_time <= slot1_end
        is_slot2 = slot2_start <= current_time <= slot2_end

        # 测试签到功能时，可以将下一行及后面的两行注释掉
        if not (is_slot1 or is_slot2):
            messagebox.showwarning("不在签到时间", "当前时间不允许签到！\n允许时间：6:00-6:51 和 17:20-17:51")
            return

        today = self.get_current_date()
        time_str = now.strftime("%H:%M:%S")

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        cursor.execute("SELECT slot1_time, slot2_time FROM sign_ins WHERE date=? AND name=?", (today, name))
        record = cursor.fetchone()

        if not record:
            cursor.execute("INSERT INTO sign_ins (date, name, slot1_time, slot2_time) VALUES (?, ?, ?, ?)",
                           (today, name, None, None))
            record = (None, None)

        slot1, slot2 = record

        if is_slot1:
            if slot1:
                messagebox.showinfo("提示", "该时段已签到过！")
                conn.close()
                return
            cursor.execute("UPDATE sign_ins SET slot1_time=? WHERE date=? AND name=?", (time_str, today, name))

        elif is_slot2:
            if slot2:
                messagebox.showinfo("提示", "该时段已签到过！")
                conn.close()
                return
            cursor.execute("UPDATE sign_ins SET slot2_time=? WHERE date=? AND name=?", (time_str, today, name))

        conn.commit()
        conn.close()
        self.load_today_data()

    # =========== 更新迟到登记功能 ===========
    def open_late_window(self):
        win = ctk.CTkToplevel(self)
        win.title("迟到登记")
        win.geometry("600x600")
        win.transient(self)

        today = self.get_current_date()
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        cursor.execute("SELECT late_saved FROM daily_status WHERE date=?", (today,))
        status = cursor.fetchone()
        is_saved = True if (status and status[0] == 1) else False

        cursor.execute("SELECT name FROM students")
        students = sorted([row[0] for row in cursor.fetchall()], key=lambda x: ''.join(lazy_pinyin(x)))

        cursor.execute("SELECT name, late_type, remark FROM late_records WHERE date=?", (today,))
        late_today = {row[0]: {'type': row[1], 'remark': row[2]} for row in cursor.fetchall()}
        conn.close()

        lbl_info = ctk.CTkLabel(win, text=f"日期: {today} | 请选择今日迟到情况", font=("Arial", 14, "bold"))
        lbl_info.pack(pady=10)

        scroll = ctk.CTkScrollableFrame(win)
        scroll.pack(fill="both", expand=True, padx=20, pady=5)

        self.late_widgets = {}

        # 逐行生成学生配置项
        for name in students:
            row_frame = ctk.CTkFrame(scroll)
            row_frame.pack(fill="x", pady=3, padx=5)

            lbl_name = ctk.CTkLabel(row_frame, text=name, width=60, anchor="center")
            lbl_name.pack(side="left", padx=10)

            # 迟到类型选择
            current_type = late_today.get(name, {}).get('type', "正常")
            type_var = ctk.StringVar(value=current_type)

            opt = ctk.CTkOptionMenu(row_frame, values=["正常", "早自习", "晚自习", "其他"], variable=type_var,
                                    width=100)
            opt.pack(side="left", padx=10)

            # 备注输入框
            entry_remark = ctk.CTkEntry(row_frame, width=250, placeholder_text="输入备注（仅选'其他'时可用）")
            entry_remark.pack(side="left", padx=10)

            # 恢复已保存的备注
            current_remark = late_today.get(name, {}).get('remark', "")
            if current_type == "其他":
                entry_remark.insert(0, current_remark if current_remark else "")
            else:
                entry_remark.configure(state="disabled")

            # 动态启禁用备注框的回调函数
            def on_type_change(val, e=entry_remark):
                if val == "其他":
                    e.configure(state="normal")
                else:
                    e.delete(0, 'end')
                    e.configure(state="disabled")

            opt.configure(command=on_type_change)

            if is_saved:
                opt.configure(state="disabled")
                entry_remark.configure(state="disabled")

            self.late_widgets[name] = {'type_var': type_var, 'remark_entry': entry_remark}

        if is_saved:
            ctk.CTkLabel(win, text="⚠ 今日迟到名单已保存，不可再更改", text_color="red",
                         font=("Arial", 14, "bold")).pack(pady=10)
        else:
            btn_save = ctk.CTkButton(win, text="保存登记", command=lambda: self.save_late_records(win))
            btn_save.pack(pady=15)

    def save_late_records(self, win):
        if not messagebox.askyesno("确认保存", "保存后今日的迟到记录不可再更改，确定要保存吗？"):
            return

        today = self.get_current_date()
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        cursor.execute("DELETE FROM late_records WHERE date=?", (today,))

        for name, widgets in self.late_widgets.items():
            late_type = widgets['type_var'].get()
            if late_type != "正常":
                remark = widgets['remark_entry'].get() if late_type == "其他" else ""
                cursor.execute("INSERT INTO late_records (date, name, late_type, remark) VALUES (?, ?, ?, ?)",
                               (today, name, late_type, remark))

        cursor.execute("INSERT OR REPLACE INTO daily_status (date, late_saved) VALUES (?, 1)", (today,))

        conn.commit()
        conn.close()

        messagebox.showinfo("成功", "迟到记录保存成功！")
        win.destroy()

    # =========== 历史记录与导出功能更新 ===========
    def show_history(self):
        self.query_window("历史记录查询", self.execute_history_query)

    def export_data(self):
        self.query_window("数据汇总导出", self.execute_export)

    def query_window(self, title, command_func):
        win = ctk.CTkToplevel(self)
        win.title(title)
        win.geometry("400x200")
        win.transient(self)

        ctk.CTkLabel(win, text="开始日期:").pack(pady=5)
        cal_start = DateEntry(win, width=12, background='darkblue', foreground='white', borderwidth=2,
                              date_pattern='yyyy-mm-dd')
        cal_start.pack(pady=5)

        ctk.CTkLabel(win, text="结束日期:").pack(pady=5)
        cal_end = DateEntry(win, width=12, background='darkblue', foreground='white', borderwidth=2,
                            date_pattern='yyyy-mm-dd')
        cal_end.pack(pady=5)

        ctk.CTkButton(win, text="确认", command=lambda: command_func(cal_start.get(), cal_end.get(), win)).pack(pady=15)

    def get_summary_dataframe(self, start_date, end_date):
        """核心查询逻辑：联合统计并拼接迟到详情"""
        conn = sqlite3.connect(DB_FILE)
        query = """
            SELECT 
                students.name as 姓名, 
                IFNULL(signs.sign_count, 0) as 签到总次数,
                IFNULL(lates.late_count, 0) as 迟到总次数,
                IFNULL(lates.late_details, '无') as 迟到详情
            FROM students
            LEFT JOIN (
                SELECT name, COUNT(slot1_time) + COUNT(slot2_time) as sign_count
                FROM sign_ins 
                WHERE date BETWEEN ? AND ?
                GROUP BY name
            ) signs ON students.name = signs.name
            LEFT JOIN (
                SELECT 
                    name, 
                    COUNT(*) as late_count,
                    GROUP_CONCAT(date || ' ' || late_type || CASE WHEN remark IS NOT NULL AND remark != '' THEN '('||remark||')' ELSE '' END, '；') as late_details
                FROM late_records
                WHERE date BETWEEN ? AND ?
                GROUP BY name
            ) lates ON students.name = lates.name
            ORDER BY 迟到总次数 DESC, 签到总次数 DESC, students.name
        """
        df = pd.read_sql_query(query, conn, params=(start_date, end_date, start_date, end_date))
        conn.close()
        return df

    def execute_history_query(self, start_date, end_date, win):
        win.destroy()
        df = self.get_summary_dataframe(start_date, end_date)

        res_win = ctk.CTkToplevel(self)
        res_win.title(f"{start_date} 至 {end_date} 数据查询")
        res_win.geometry("750x450")

        tree = ttk.Treeview(res_win, columns=("Name", "SignCount", "LateCount", "LateDetails"), show='headings')
        tree.heading("Name", text="姓名")
        tree.heading("SignCount", text="签到总次数")
        tree.heading("LateCount", text="迟到总次数")
        tree.heading("LateDetails", text="迟到详情")

        tree.column("Name", width=80, anchor="center")
        tree.column("SignCount", width=80, anchor="center")
        tree.column("LateCount", width=80, anchor="center")
        tree.column("LateDetails", width=350, anchor="w")

        tree.pack(fill="both", expand=True, padx=10, pady=10)

        for _, row in df.iterrows():
            tree.insert("", "end", values=(row['姓名'], row['签到总次数'], row['迟到总次数'], row['迟到详情']))

    def execute_export(self, start_date, end_date, win):
        win.destroy()
        df = self.get_summary_dataframe(start_date, end_date)

        if df.empty:
            messagebox.showinfo("提示", "所选范围内无学生名单数据。")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"考勤汇总_{start_date}至{end_date}.xlsx"
        )
        if filepath:
            df.to_excel(filepath, index=False)
            messagebox.showinfo("成功", "数据汇总导出成功！")


if __name__ == "__main__":
    app = SignInSystem()
    app.mainloop()