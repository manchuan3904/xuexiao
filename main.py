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
        self.geometry("900x700")

        # 数据库初始化
        self.init_db()

        # 顶部控制面板
        self.top_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.top_frame.pack(pady=10, padx=20, fill="x")

        self.btn_import = ctk.CTkButton(self.top_frame, text="导入名单", command=self.import_students)
        self.btn_import.pack(side="left", padx=5)

        self.btn_history = ctk.CTkButton(self.top_frame, text="查看历史记录", command=self.show_history)
        self.btn_history.pack(side="left", padx=5)

        self.btn_export = ctk.CTkButton(self.top_frame, text="数据汇总", command=self.export_data)
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
        conn.commit()
        conn.close()

    def get_current_date(self):
        return datetime.date.today().strftime("%Y-%m-%d")

    def load_today_data(self):
        # 清空当前显示
        for widget in self.scroll_frame.winfo_children():
            widget.destroy()
        self.bubbles.clear()

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        # 获取所有学生并按拼音排序
        cursor.execute("SELECT name FROM students")
        students = [row[0] for row in cursor.fetchall()]
        students = sorted(students, key=lambda x: ''.join(lazy_pinyin(x)))

        # 获取今日签到数据
        today = self.get_current_date()
        cursor.execute("SELECT name, slot1_time, slot2_time FROM sign_ins WHERE date=?", (today,))
        today_data = {row[0]: (row[1], row[2]) for row in cursor.fetchall()}
        conn.close()

        total = len(students)
        signed_in_count = 0

        # 每行显示 6 个泡泡，适配常规班级人数
        columns = 6
        for index, name in enumerate(students):
            row = index // columns
            col = index % columns

            slot1, slot2 = today_data.get(name, (None, None))

            if slot1 and slot2:
                bg_color = "#4682B4"  # 深蓝色
                text_color = "white"
                state = 2
                display_text = f"{name}\n① {slot1}\n② {slot2}"
                signed_in_count += 1
            elif slot1:
                bg_color = "#87CEFA"  # 浅蓝色
                text_color = "black"
                state = 1
                display_text = f"{name}\n① {slot1}\n② --:--:--"
                signed_in_count += 1
            else:
                bg_color = "#E0E0E0"  # 灰色
                text_color = "black"
                state = 0
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
            self.bubbles[name] = {"btn": btn, "state": state}

        # 更新统计
        rate = (signed_in_count / total * 100) if total > 0 else 0
        self.lbl_stats.configure(text=f"总人数: {total} | 已签到: {signed_in_count} | 签到率: {rate:.1f}%")

    def import_students(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not filepath:
            return

        try:
            df = pd.read_excel(filepath)
            # 假设Excel中有一列叫“姓名”
            if "姓名" not in df.columns:
                messagebox.showerror("错误", "Excel表中必须包含“姓名”列！")
                return

            names = df["姓名"].dropna().astype(str).tolist()

            conn = sqlite3.connect(DB_FILE)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM students")  # 清空旧名单
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

        if not (is_slot1 or is_slot2):
            messagebox.showwarning("不在签到时间", "当前时间不允许签到！\n允许时间：6:00-6:51 和 17:20-17:51")
            return

        today = self.get_current_date()
        time_str = now.strftime("%H:%M:%S")

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        # 检查是否已有记录
        cursor.execute("SELECT slot1_time, slot2_time FROM sign_ins WHERE date=? AND name=?", (today, name))
        record = cursor.fetchone()

        if not record:
            # 首次签到记录创建
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

        # 刷新界面状态
        self.load_today_data()

    def show_history(self):
        self.query_window("历史记录查询", self.execute_history_query)

    def export_data(self):
        self.query_window("数据汇总导出", self.execute_export)

    def query_window(self, title, command_func):
        win = ctk.CTkToplevel(self)
        win.title(title)
        win.geometry("400x200")
        win.transient(self)  # 保持在主窗口上方

        ctk.CTkLabel(win, text="开始日期:").pack(pady=5)
        cal_start = DateEntry(win, width=12, background='darkblue', foreground='white', borderwidth=2,
                              date_pattern='yyyy-mm-dd')
        cal_start.pack(pady=5)

        ctk.CTkLabel(win, text="结束日期:").pack(pady=5)
        cal_end = DateEntry(win, width=12, background='darkblue', foreground='white', borderwidth=2,
                            date_pattern='yyyy-mm-dd')
        cal_end.pack(pady=5)

        ctk.CTkButton(win, text="确认", command=lambda: command_func(cal_start.get(), cal_end.get(), win)).pack(pady=15)

    def execute_history_query(self, start_date, end_date, win):
        win.destroy()
        conn = sqlite3.connect(DB_FILE)
        query = """
            SELECT name, COUNT(slot1_time) + COUNT(slot2_time) as total_signs
            FROM sign_ins 
            WHERE date BETWEEN ? AND ?
            GROUP BY name
        """
        df = pd.read_sql_query(query, conn, params=(start_date, end_date))
        conn.close()

        res_win = ctk.CTkToplevel(self)
        res_win.title(f"{start_date} 至 {end_date} 签到记录")
        res_win.geometry("500x400")

        tree = ttk.Treeview(res_win, columns=("Name", "Count"), show='headings')
        tree.heading("Name", text="姓名")
        tree.heading("Count", text="签到总次数")
        tree.pack(fill="both", expand=True, padx=10, pady=10)

        for _, row in df.iterrows():
            tree.insert("", "end", values=(row['name'], row['total_signs']))

    def execute_export(self, start_date, end_date, win):
        win.destroy()
        conn = sqlite3.connect(DB_FILE)
        query = """
            SELECT date as 日期, name as 姓名, slot1_time as 第一次签到, slot2_time as 第二次签到
            FROM sign_ins 
            WHERE date BETWEEN ? AND ?
            ORDER BY date, name
        """
        df = pd.read_sql_query(query, conn, params=(start_date, end_date))
        conn.close()

        if df.empty:
            messagebox.showinfo("提示", "所选日期范围内没有签到数据。")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=f"签到汇总_{start_date}至{end_date}.xlsx"
        )
        if filepath:
            df.to_excel(filepath, index=False)
            messagebox.showinfo("成功", "数据导出成功！")


if __name__ == "__main__":
    app = SignInSystem()
    app.mainloop()