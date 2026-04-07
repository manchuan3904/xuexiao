import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import os
import datetime
import csv


class SignInSystem(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("自习签到系统")
        self.geometry("900x700")
        self.configure(bg="#F5F7FA")

        # 数据存储路径初始化
        self.data_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data')
        if not os.path.exists(self.data_dir):
            os.makedirs(self.data_dir)

        self.students = []
        self.bubbles = {}
        self.today_str = datetime.datetime.now().strftime("%Y-%m-%d")
        self.daily_record_file = os.path.join(self.data_dir, f"record_{self.today_str}.json")
        self.list_file = os.path.join(self.data_dir, "student_list.json")

        self.init_ui()
        self.load_students()
        self.load_today_record()
        self.check_new_day()

    def init_ui(self):
        # 顶部：标题与统计信息
        top_frame = tk.Frame(self, bg="#F5F7FA")
        top_frame.pack(side=tk.TOP, fill=tk.X, pady=10)

        title_label = tk.Label(top_frame, text="自习签到系统", font=("PingFang SC", 24, "bold"), bg="#F5F7FA",
                               fg="#2C3E50")
        title_label.pack()

        self.stats_var = tk.StringVar(value="班级总人数: 0  |  已签到人数: 0  |  签到率: 0%")
        stats_label = tk.Label(top_frame, textvariable=self.stats_var, font=("PingFang SC", 14), bg="#F5F7FA",
                               fg="#34495E")
        stats_label.pack(pady=5)

        # 中部：学生泡泡显示区 (带滚动条)
        middle_frame = tk.Frame(self, bg="#F5F7FA")
        middle_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=20)

        self.canvas = tk.Canvas(middle_frame, bg="#F5F7FA", highlightthickness=0)
        scrollbar = ttk.Scrollbar(middle_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg="#F5F7FA")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 底部：功能按键
        bottom_frame = tk.Frame(self, bg="#F5F7FA")
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=20)

        btn_import = tk.Button(bottom_frame, text="导入名单", command=self.import_list, font=("PingFang SC", 12),
                               width=10)
        btn_late = tk.Button(bottom_frame, text="迟到登记", command=self.open_late_register, font=("PingFang SC", 12),
                             width=10, fg="#E6A23C")
        btn_history = tk.Button(bottom_frame, text="查看历史记录", command=self.view_history, font=("PingFang SC", 12),
                                width=15)
        btn_summary = tk.Button(bottom_frame, text="数据汇总 (导出Excel)", command=self.export_summary,
                                font=("PingFang SC", 12), width=20)

        btn_import.pack(side=tk.LEFT, padx=15)
        btn_late.pack(side=tk.LEFT, padx=15)
        btn_history.pack(side=tk.LEFT, padx=15)
        btn_summary.pack(side=tk.RIGHT, padx=15)

    def import_list(self):
        file_path = filedialog.askopenfilename(title="选择学生名单", filetypes=[("Text Files", "*.txt")])
        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                names = [line.strip() for line in f if line.strip()]

            self.students = list(dict.fromkeys(names))  # 去重并保持原有顺序

            with open(self.list_file, 'w', encoding='utf-8') as f:
                json.dump(self.students, f, ensure_ascii=False)

            self.refresh_bubbles()
            self.save_today_record()
            messagebox.showinfo("成功", f"成功导入 {len(self.students)} 名学生！")
        except Exception as e:
            messagebox.showerror("错误", f"导入失败: {str(e)}")

    def load_students(self):
        if os.path.exists(self.list_file):
            with open(self.list_file, 'r', encoding='utf-8') as f:
                self.students = json.load(f)
            self.refresh_bubbles()

    def refresh_bubbles(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.bubbles.clear()

        columns = 6
        for index, name in enumerate(self.students):
            btn = tk.Button(self.scrollable_frame, text=name, width=12, height=3,
                            font=("PingFang SC", 12, "bold"), relief="groove")
            btn.grid(row=index // columns, column=index % columns, padx=10, pady=10)
            btn.config(command=lambda n=name: self.handle_signin(n))

            self.bubbles[name] = {
                'widget': btn,
                'sign1': None,
                'sign2': None,
                'late': []  # 新增迟到记录列表
            }

        self.update_bubble_ui()
        self.update_stats()

    def update_bubble_ui(self):
        for name, data in self.bubbles.items():
            btn = data['widget']
            sign_count = 0
            text = f"{name}"

            if data['sign1']:
                text += f"\n1: {data['sign1']}"
                sign_count += 1
            if data['sign2']:
                text += f"\n2: {data['sign2']}"
                sign_count += 1

            btn.config(text=text)

            if sign_count == 0:
                btn.config(bg="#E0E0E0", fg="#333333")
            elif sign_count == 1:
                btn.config(bg="#87CEFA", fg="#000000")
            else:
                btn.config(bg="#4682B4", fg="#FFFFFF")

    def handle_signin(self, name):
        now = datetime.datetime.now()
        current_time = now.strftime("%H:%M:%S")

        t1_start = datetime.time(6, 0, 0)
        t1_end = datetime.time(6, 51, 0)
        t2_start = datetime.time(17, 20, 0)
        t2_end = datetime.time(17, 51, 0)

        data = self.bubbles[name]

        if t1_start <= now.time() <= t1_end:
            if data['sign1']:
                messagebox.showwarning("提示", f"【{name}】在第一时间段已签到，请勿重复签到！")
            else:
                data['sign1'] = current_time
                self.update_bubble_ui()
                self.save_today_record()

        elif t2_start <= now.time() <= t2_end:
            if data['sign2']:
                messagebox.showwarning("提示", f"【{name}】在第二时间段已签到，请勿重复签到！")
            else:
                data['sign2'] = current_time
                self.update_bubble_ui()
                self.save_today_record()
        else:
            messagebox.showwarning("签到失败",
                                   "当前不在允许的签到时间段内！\n第一时间段: 6:00:00-6:51:00\n第二时间段: 17:20:00-17:51:00")
            return

        self.update_stats()

    def update_stats(self):
        total = len(self.students)
        signed_in = sum(1 for data in self.bubbles.values() if data['sign1'] or data['sign2'])
        rate = (signed_in / total * 100) if total > 0 else 0
        self.stats_var.set(f"班级总人数: {total}  |  已签到人数: {signed_in}  |  签到率: {rate:.1f}%")

    def load_today_record(self):
        if os.path.exists(self.daily_record_file):
            with open(self.daily_record_file, 'r', encoding='utf-8') as f:
                records = json.load(f)
                for name, times in records.items():
                    if name in self.bubbles:
                        self.bubbles[name]['sign1'] = times.get('sign1')
                        self.bubbles[name]['sign2'] = times.get('sign2')
                        self.bubbles[name]['late'] = times.get('late', [])
        self.update_bubble_ui()
        self.update_stats()

    def save_today_record(self):
        records = {}
        for name, data in self.bubbles.items():
            if data['sign1'] or data['sign2'] or data['late']:
                records[name] = {
                    'sign1': data['sign1'],
                    'sign2': data['sign2'],
                    'late': data['late']
                }
        with open(self.daily_record_file, 'w', encoding='utf-8') as f:
            json.dump(records, f, ensure_ascii=False, indent=4)

    def check_new_day(self):
        real_today = datetime.datetime.now().strftime("%Y-%m-%d")
        if real_today != self.today_str:
            self.today_str = real_today
            self.daily_record_file = os.path.join(self.data_dir, f"record_{self.today_str}.json")
            for data in self.bubbles.values():
                data['sign1'] = None
                data['sign2'] = None
                data['late'] = []
            self.update_bubble_ui()
            self.update_stats()
        self.after(60000, self.check_new_day)

    # ---------------- 迟到登记功能区 ----------------
    def open_late_register(self):
        if not self.students:
            messagebox.showwarning("提示", "请先导入学生名单！")
            return

        win = tk.Toplevel(self)
        win.title("迟到登记系统")
        win.geometry("750x500")
        win.configure(bg="#FFF8E7")

        tk.Label(win, text="请点击需要登记迟到的学生姓名", font=("PingFang SC", 18, "bold"), bg="#FFF8E7",
                 fg="#E6A23C").pack(pady=15)

        canvas = tk.Canvas(win, bg="#FFF8E7", highlightthickness=0)
        scrollbar = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
        frame = tk.Frame(canvas, bg="#FFF8E7")

        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=20)
        scrollbar.pack(side="right", fill="y")

        columns = 5
        for index, name in enumerate(self.students):
            btn = tk.Button(frame, text=name, width=12, height=2, font=("PingFang SC", 12), relief="groove",
                            bg="#FFFFFF")
            btn.grid(row=index // columns, column=index % columns, padx=10, pady=10)
            btn.config(command=lambda n=name: self.open_late_dialog(n))

    def open_late_dialog(self, name):
        dialog = tk.Toplevel(self)
        dialog.title(f"{name} - 迟到登记")
        dialog.geometry("300x320")

        tk.Label(dialog, text=f"为 {name} 登记迟到", font=("PingFang SC", 14, "bold")).pack(pady=10)

        var = tk.StringVar(value="早自习")

        # 选项配置
        tk.Radiobutton(dialog, text="早自习", variable=var, value="早自习", font=("Arial", 12)).pack(anchor='w',
                                                                                                     padx=30, pady=5)
        tk.Radiobutton(dialog, text="晚自习", variable=var, value="晚自习", font=("Arial", 12)).pack(anchor='w',
                                                                                                     padx=30, pady=5)
        tk.Radiobutton(dialog, text="其他", variable=var, value="其他", font=("Arial", 12)).pack(anchor='w', padx=30,
                                                                                                 pady=5)

        # 备注文本框
        tk.Label(dialog, text="备注 (选择“其他”时填写):", font=("Arial", 10)).pack(anchor='w', padx=30, pady=(10, 0))
        entry = tk.Entry(dialog, state='disabled', font=("Arial", 12))
        entry.pack(padx=30, pady=5, fill='x')

        def on_change(*args):
            if var.get() == "其他":
                entry.config(state='normal')
                entry.focus()
            else:
                entry.delete(0, tk.END)
                entry.config(state='disabled')

        var.trace_add('write', on_change)

        def save():
            late_type = var.get()
            remark = entry.get().strip() if late_type == "其他" else ""

            if late_type == "其他" and not remark:
                messagebox.showwarning("提示", "请填写其他迟到的备注原因！", parent=dialog)
                return

            current_lates = self.bubbles[name].get('late', [])

            # 检查是否已存在同类型的迟到记录（控制不可更改与重复）
            if any(l['type'] == late_type for l in current_lates):
                messagebox.showwarning("提示", f"今天已保存过【{late_type}】的迟到记录，不可重复或更改！", parent=dialog)
                return

            current_lates.append({'type': late_type, 'remark': remark})
            self.bubbles[name]['late'] = current_lates
            self.save_today_record()

            messagebox.showinfo("成功", f"已成功保存 {name} 的迟到记录！", parent=dialog)
            dialog.destroy()

        tk.Button(dialog, text="保存 (保存后不可更改)", command=save, bg="#F56C6C", fg="black",
                  font=("Arial", 11, "bold")).pack(pady=15)

    # ---------------- 历史查询与导出升级 ----------------
    def get_aggregated_data(self, start_date_str, end_date_str):
        try:
            start_dt = datetime.datetime.strptime(start_date_str, "%Y-%m-%d")
            end_dt = datetime.datetime.strptime(end_date_str, "%Y-%m-%d")
        except ValueError:
            return None, "日期格式错误，请使用 YYYY-MM-DD 格式！"

        if start_dt > end_dt:
            return None, "开始日期不能晚于结束日期！"

        # 扩充数据结构：包含签到次数、迟到次数、迟到详细备注
        totals = {name: {'sign': 0, 'late': 0, 'late_details': []} for name in self.students}

        current_dt = start_dt
        delta = datetime.timedelta(days=1)

        while current_dt <= end_dt:
            date_str = current_dt.strftime("%Y-%m-%d")
            file_path = os.path.join(self.data_dir, f"record_{date_str}.json")

            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    records = json.load(f)
                    for name in self.students:
                        # 兼容老数据与新数据的读取
                        times = records.get(name, {})

                        # 统计签到
                        sign_count = sum(1 for k in ['sign1', 'sign2'] if times.get(k))
                        totals[name]['sign'] += sign_count

                        # 统计迟到及备注
                        lates = times.get('late', [])
                        totals[name]['late'] += len(lates)
                        for l in lates:
                            detail_str = f"[{date_str}] {l['type']}"
                            if l.get('remark'):
                                detail_str += f"({l['remark']})"
                            totals[name]['late_details'].append(detail_str)

            current_dt += delta

        return totals, ""

    def view_history(self):
        win = tk.Toplevel(self)
        win.title("查看历史记录")
        win.geometry("500x550")

        frame = tk.Frame(win)
        frame.pack(pady=10)

        last_week = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime("%Y-%m-%d")

        tk.Label(frame, text="开始日期 (YYYY-MM-DD):").grid(row=0, column=0, pady=5)
        start_entry = tk.Entry(frame, width=12)
        start_entry.insert(0, last_week)
        start_entry.grid(row=0, column=1, pady=5)

        tk.Label(frame, text="结束日期 (YYYY-MM-DD):").grid(row=1, column=0, pady=5)
        end_entry = tk.Entry(frame, width=12)
        end_entry.insert(0, self.today_str)
        end_entry.grid(row=1, column=1, pady=5)

        # 表格增加一列：迟到次数
        tree = ttk.Treeview(win, columns=("Name", "SignCount", "LateCount"), show="headings", height=18)
        tree.heading("Name", text="学生姓名")
        tree.heading("SignCount", text="签到次数")
        tree.heading("LateCount", text="迟到次数")

        tree.column("Name", width=120, anchor="center")
        tree.column("SignCount", width=100, anchor="center")
        tree.column("LateCount", width=100, anchor="center")

        def do_search():
            for row in tree.get_children():
                tree.delete(row)

            totals, error_msg = self.get_aggregated_data(start_entry.get(), end_entry.get())
            if error_msg:
                messagebox.showerror("错误", error_msg, parent=win)
                return

            for name in self.students:
                tree.insert("", "end", values=(name, totals[name]['sign'], totals[name]['late']))

        tk.Button(win, text="查询区间数据", command=do_search).pack(pady=5)
        tree.pack(pady=10)

    def export_summary(self):
        win = tk.Toplevel(self)
        win.title("数据汇总导出")
        win.geometry("350x250")

        tk.Label(win, text="选择需要汇总的日期范围").pack(pady=10)

        frame = tk.Frame(win)
        frame.pack()

        last_week = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime("%Y-%m-%d")

        tk.Label(frame, text="开始日期:").grid(row=0, column=0, pady=5)
        start_entry = tk.Entry(frame, width=15)
        start_entry.insert(0, last_week)
        start_entry.grid(row=0, column=1, pady=5)

        tk.Label(frame, text="结束日期:").grid(row=1, column=0, pady=5)
        end_entry = tk.Entry(frame, width=15)
        end_entry.insert(0, self.today_str)
        end_entry.grid(row=1, column=1, pady=5)

        def do_export():
            start_str = start_entry.get()
            end_str = end_entry.get()
            totals, error_msg = self.get_aggregated_data(start_str, end_str)

            if error_msg:
                messagebox.showerror("错误", error_msg, parent=win)
                return

            file_path = filedialog.asksaveasfilename(
                title="保存汇总数据",
                defaultextension=".csv",
                filetypes=[("CSV Files", "*.csv")],
                initialfile=f"签到与迟到汇总_{start_str}至{end_str}.csv",
                parent=win
            )

            if not file_path:
                return

            try:
                # 使用 utf-8-sig 编码，确保用 Excel 打开时不乱码
                with open(file_path, 'w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.writer(f)
                    # 写入表头，增加迟到次数及明细
                    writer.writerow(["姓名", f"总签到次数", "总迟到次数", "迟到详情 (日期/类型/备注)"])

                    for name in self.students:
                        # 将迟到详情列表拼接为一段长字符串，便于在Excel的一个单元格内查看
                        details_str = " | ".join(totals[name]['late_details']) if totals[name]['late_details'] else "无"
                        writer.writerow([name, totals[name]['sign'], totals[name]['late'], details_str])

                messagebox.showinfo("成功", f"数据已成功导出至：\n{file_path}\n(请直接使用 Excel 打开该 CSV 文件)",
                                    parent=win)
                win.destroy()
            except Exception as e:
                messagebox.showerror("错误", f"导出失败: {str(e)}", parent=win)

        tk.Button(win, text="导出为 Excel (CSV格式)", command=do_export, font=("PingFang SC", 12), bg="#409EFF",
                  fg="white").pack(pady=20)


if __name__ == '__main__':
    app = SignInSystem()
    app.mainloop()


# 这是一个示例 Python 脚本。

# 按 ⌃R 执行或将其替换为您的代码。
# 按 双击 ⇧ 在所有地方搜索类、文件、工具窗口、操作和设置。


def print_hi(name):
    # 在下面的代码行中使用断点来调试脚本。
    print(f'Hi, {name}')  # 按 ⌘F8 切换断点。


# 按装订区域中的绿色按钮以运行脚本。
if __name__ == '__main__':
    print_hi('PyCharm')

# 访问 https://www.jetbrains.com/help/pycharm/ 获取 PyCharm 帮助
