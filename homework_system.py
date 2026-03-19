import sys
import os
import json
import pandas as pd
from datetime import datetime, timedelta
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout,
                             QHBoxLayout, QWidget, QLabel, QFileDialog, QStackedWidget,
                             QGridLayout, QScrollArea, QMessageBox)
from PyQt6.QtCore import Qt

# --- 自动路径识别 (解决不同电脑盘符问题) ---
if getattr(sys, 'frozen', False):
    BASE_PATH = os.path.dirname(sys.executable)
else:
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))

# 数据存储在程序同级目录下的 HomeworkData 文件夹内
BASE_DIR = os.path.join(BASE_PATH, "HomeworkData")
DATA_DIR = os.path.join(BASE_DIR, "records")
CONFIG_FILE = os.path.join(BASE_DIR, "students.json")


class HomeworkApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.subjects = ["语文", "数学", "英语", "物理", "化学", "生物", "历史", "政治", "地理"]
        self.students = []
        self.current_date = datetime.now().strftime("%Y-%m-%d")

        # 初始化目录
        if not os.path.exists(BASE_DIR): os.makedirs(BASE_DIR)
        if not os.path.exists(DATA_DIR): os.makedirs(DATA_DIR)

        self.load_students()
        self.init_ui()

    def load_students(self):
        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                self.students = json.load(f)

    def load_daily_data(self, date_str):
        path = os.path.join(DATA_DIR, f"{date_str}.json")
        if os.path.exists(path):
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {sub: {"status": "editing", "data": {}} for sub in self.subjects}

    def save_daily_data(self, date_str, data):
        path = os.path.join(DATA_DIR, f"{date_str}.json")
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

    def init_ui(self):
        self.setWindowTitle("作业登记系统")
        self.resize(1000, 750)
        self.setStyleSheet("QMainWindow { background-color: #f5f7fa; }")

        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)
        layout = QVBoxLayout(self.main_widget)

        # 1. 顶部标题
        title = QLabel("作业登记系统")
        title.setStyleSheet("font-size: 26px; font-weight: bold; color: #333; margin: 10px;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # 2. 日期与功能按钮
        top_bar = QHBoxLayout()
        self.date_lbl = QLabel(f"日期: {self.current_date}")
        top_bar.addWidget(self.date_lbl)

        btn_configs = [
            ("导入名单", "#4a90e2", self.import_names),
            ("查看历史记录", "#67c23a", self.view_history),
            ("数据导出 (Excel)", "#e6a23c", self.export_excel)
        ]
        for text, color, func in btn_configs:
            btn = QPushButton(text)
            btn.setStyleSheet(f"background-color: {color}; color: white; padding: 8px 15px; border-radius: 4px;")
            btn.clicked.connect(func)
            top_bar.addWidget(btn)
        layout.addLayout(top_bar)

        # 3. 科目列表
        grid = QGridLayout()
        for i, sub in enumerate(self.subjects):
            btn = QPushButton(sub)
            btn.setFixedSize(120, 80)
            btn.setStyleSheet(
                "background-color: white; border: 2px solid #4a90e2; border-radius: 10px; font-size: 16px;")
            btn.clicked.connect(lambda ch, s=sub: self.open_subject(s))
            grid.addWidget(btn, i // 5, i % 5)
        layout.addLayout(grid)

        # 4. 当日未交汇总框
        self.summary_box = QLabel("暂无未交记录")
        self.summary_box.setStyleSheet(
            "background-color: white; border: 1px solid #ddd; padding: 15px; border-radius: 5px;")
        self.summary_box.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.summary_box.setWordWrap(True)
        layout.addWidget(self.summary_box)
        self.refresh_summary()

    def refresh_summary(self):
        if not self.students:
            self.summary_box.setText("请先点击【导入名单】。")
            return
        data = self.load_daily_data(self.current_date)
        res = []
        for name in self.students:
            subs = [s for s in self.subjects if data.get(s, {}).get("data", {}).get(name) == "missing"]
            if subs:
                res.append(f"<b>{name}</b>: {', '.join(subs)}")
        self.summary_box.setText("<br>".join(res) if res else "今日作业全部交齐！")

    def import_names(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择学生名单", "", "文本文件 (*.txt)")
        if path:
            with open(path, 'r', encoding='utf-8') as f:
                names = [line.strip() for line in f.readlines() if line.strip()]
            self.students = names
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(names, f, ensure_ascii=False)
            QMessageBox.information(self, "成功", f"已导入 {len(names)} 名学生。")
            self.refresh_summary()

    def open_subject(self, sub):
        if not self.students: return
        self.detail = SubjectWindow(sub, self.students, self.current_date, self)
        self.detail.show()

    def view_history(self):
        QMessageBox.information(self, "查看历史", "历史记录已整合在当日汇总中，如需详细历史请点击'数据导出'查看表格。")

    def export_excel(self):
        # 导出逻辑：导出最近7天
        try:
            all_rows = []
            for i in range(7):
                d = (datetime.now() - timedelta(days=i)).strftime("%Y-%m-%d")
                d_data = self.load_daily_data(d)
                for name in self.students:
                    row = {"姓名": name, "日期": d}
                    for s in self.subjects:
                        # 补交(late)或已交(normal) 均计为0次未交；只有missing计为1次
                        status = d_data.get(s, {}).get("data", {}).get(name, "normal")
                        row[s] = 1 if status == "missing" else 0
                    all_rows.append(row)

            df = pd.DataFrame(all_rows)
            summary = df.groupby("姓名")[self.subjects].sum()
            summary['汇总'] = summary.sum(axis=1)

            out_path = os.path.join(BASE_DIR, f"作业统计_{datetime.now().strftime('%m%d')}.xlsx")
            summary.to_excel(out_path)
            os.startfile(BASE_DIR)
            QMessageBox.information(self, "成功", f"文件已保存至: {out_path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))


class SubjectWindow(QWidget):
    def __init__(self, subject, students, date_str, parent):
        super().__init__()
        self.subject = subject
        self.students = students
        self.date_str = date_str
        self.parent_win = parent
        self.full_data = parent.load_daily_data(date_str)
        self.sub_info = self.full_data[subject]
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle(f"{self.subject} 登记")
        self.resize(800, 600)
        layout = QVBoxLayout(self)

        self.info_lbl = QLabel()
        layout.addWidget(self.info_lbl)

        scroll = QScrollArea()
        content = QWidget()
        self.grid = QGridLayout(content)
        self.btns = {}
        for i, name in enumerate(self.students):
            btn = QPushButton(name)
            btn.setFixedSize(110, 50)
            status = self.sub_info["data"].get(name, "normal")
            self.style_btn(btn, status)
            btn.clicked.connect(lambda ch, n=name, b=btn: self.click_name(n, b))
            self.grid.addWidget(btn, i // 5, i % 5)
            self.btns[name] = btn

        scroll.setWidget(content)
        scroll.setWidgetResizable(True)
        layout.addWidget(scroll)

        save_btn = QPushButton("保存提交 (锁定当日状态)")
        save_btn.setStyleSheet("background-color: #4a90e2; color: white; height: 40px;")
        save_btn.clicked.connect(self.do_save)
        layout.addWidget(save_btn)
        self.update_stats()

    def style_btn(self, btn, status):
        if status == "normal":
            btn.setStyleSheet("background-color: white; color: black; border: 1px solid #ccc;")
            btn.setText(btn.text().split(' ')[0])
        elif status == "missing":
            btn.setStyleSheet("background-color: #ffeded; color: red; border: 2px solid red;")
            btn.setText(btn.text().split(' ')[0] + " ✘")
        elif status == "late":
            btn.setStyleSheet("background-color: #f5f5f5; color: #999; border: 1px solid #ddd;")
            btn.setText(btn.text().split(' ')[0] + " ✓")

    def click_name(self, name, btn):
        curr = self.sub_info["data"].get(name, "normal")
        locked = self.sub_info.get("status") == "locked"

        # 补交限时判断
        deadline = datetime.strptime(self.date_str, "%Y-%m-%d") + timedelta(days=1, hours=18, minutes=30)
        expired = datetime.now() > deadline

        if not locked:
            new_s = "missing" if curr == "normal" else "normal"
            self.sub_info["data"][name] = new_s
        else:
            if curr == "missing":
                if expired:
                    QMessageBox.warning(self, "截止", "已超次日18:30，无法补交。")
                    return
                self.sub_info["data"][name] = "late"

        self.style_btn(btn, self.sub_info["data"][name])
        self.update_stats()

    def update_stats(self):
        vals = list(self.sub_info["data"].values())
        m, l = vals.count("missing"), vals.count("late")
        self.info_lbl.setText(f"科目: {self.subject} | 未交: {m} | 补交: {l}")

    def do_save(self):
        self.sub_info["status"] = "locked"
        self.full_data[self.subject] = self.sub_info
        self.parent_win.save_daily_data(self.date_str, self.full_data)
        self.parent_win.refresh_summary()
        self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = HomeworkApp()
    window.show()
    sys.exit(app.exec())


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
