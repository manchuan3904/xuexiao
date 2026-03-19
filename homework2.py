import sys
import os
import json
import pandas as pd
from datetime import datetime, timedelta
from PyQt6.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout,
                             QHBoxLayout, QWidget, QLabel, QFileDialog, QStackedWidget,
                             QGridLayout, QScrollArea, QMessageBox, QDialog, QDateEdit,
                             QFormLayout, QDialogButtonBox)
from PyQt6.QtCore import Qt, QDate

# --- 自动路径识别 ---
if getattr(sys, 'frozen', False):
    BASE_PATH = os.path.dirname(sys.executable)
else:
    BASE_PATH = os.path.dirname(os.path.abspath(__file__))

BASE_DIR = os.path.join(BASE_PATH, "HomeworkData")
DATA_DIR = os.path.join(BASE_DIR, "records")
CONFIG_FILE = os.path.join(BASE_DIR, "students.json")


# --- 新增：日期范围选择弹窗 ---
class DateRangeDialog(QDialog):
    def __init__(self, parent=None, title="选择日期时段"):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(300, 150)
        layout = QFormLayout(self)

        # 默认开始日期为7天前，结束日期为今天
        self.start_date = QDateEdit(QDate.currentDate().addDays(-7))
        self.start_date.setCalendarPopup(True)
        self.end_date = QDateEdit(QDate.currentDate())
        self.end_date.setCalendarPopup(True)

        layout.addRow("开始日期:", self.start_date)
        layout.addRow("结束日期:", self.end_date)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_dates(self):
        # 返回 Python 的 datetime.date 对象
        return self.start_date.date().toPyDate(), self.end_date.date().toPyDate()


# --- 新增：历史记录查看窗口 ---
class HistoryResultDialog(QDialog):
    def __init__(self, start_date, end_date, app_instance, parent=None):
        super().__init__(parent)
        self.app = app_instance
        self.start_date = start_date
        self.end_date = end_date
        self.setWindowTitle(f"历史记录 ({start_date.strftime('%Y-%m-%d')} 至 {end_date.strftime('%Y-%m-%d')})")
        self.resize(700, 500)
        self.setStyleSheet("QDialog { background-color: #f5f7fa; }")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)

        # 顶部科目按钮
        btn_layout = QHBoxLayout()
        for sub in self.app.subjects:
            btn = QPushButton(sub)
            btn.setStyleSheet("background-color: white; border: 1px solid #4a90e2; padding: 5px; border-radius: 4px;")
            btn.clicked.connect(lambda ch, s=sub: self.show_subject_detail(s))
            btn_layout.addWidget(btn)
        layout.addLayout(btn_layout)

        # 下方显示区域
        self.scroll = QScrollArea()
        self.content_widget = QWidget()
        self.grid = QGridLayout(self.content_widget)
        self.scroll.setWidget(self.content_widget)
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("background-color: white; border: 1px solid #ddd;")
        layout.addWidget(self.scroll)

        # 默认提示语
        self.default_lbl = QLabel("请点击上方具体科目，查看该时段内的未交与补交名单。")
        self.default_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.default_lbl.setStyleSheet("color: #666; font-size: 14px;")
        self.grid.addWidget(self.default_lbl, 0, 0)

    def show_subject_detail(self, subject):
        # 清空当前显示的内容
        for i in reversed(range(self.grid.count())):
            widget = self.grid.itemAt(i).widget()
            if widget:
                widget.deleteLater()

        records = []
        current = self.start_date
        # 遍历所选日期范围提取数据
        while current <= self.end_date:
            d_str = current.strftime("%Y-%m-%d")
            d_data = self.app.load_daily_data(d_str)
            sub_data = d_data.get(subject, {}).get("data", {})
            for name, status in sub_data.items():
                if status in ["missing", "late"]:
                    records.append((name, d_str, status))
            current += timedelta(days=1)

        if not records:
            lbl = QLabel(f"恭喜！该时段内 {subject} 没有任何未交或补交记录。")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.grid.addWidget(lbl, 0, 0)
            return

        # 将记录按网格排列显示
        row, col = 0, 0
        for name, date, status in records:
            # 格式：名字 + 换行 + (日期)
            lbl = QLabel(f"{name}\n({date[5:]})")  # 日期只显示 月-日
            lbl.setFixedSize(100, 60)
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)

            # 未交为红色，补交为灰色
            if status == "missing":
                lbl.setStyleSheet(
                    "background-color: #ffeded; color: red; border: 1px solid red; border-radius: 5px; font-weight: bold;")
            elif status == "late":
                lbl.setStyleSheet("background-color: #f5f5f5; color: #888; border: 1px solid #ccc; border-radius: 5px;")

            self.grid.addWidget(lbl, row, col)
            col += 1
            if col > 4:  # 每行显示5个
                col = 0
                row += 1


# --- 主程序逻辑 ---
class HomeworkApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.subjects = ["语文", "数学", "英语", "物理", "化学", "生物", "历史", "政治", "地理"]
        self.students = []
        self.current_date = datetime.now().strftime("%Y-%m-%d")

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

        title = QLabel("作业登记系统")
        title.setStyleSheet("font-size: 26px; font-weight: bold; color: #333; margin: 10px;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

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

        grid = QGridLayout()
        for i, sub in enumerate(self.subjects):
            btn = QPushButton(sub)
            btn.setFixedSize(120, 80)
            btn.setStyleSheet(
                "background-color: white; border: 2px solid #4a90e2; border-radius: 10px; font-size: 16px;")
            btn.clicked.connect(lambda ch, s=sub: self.open_subject(s))
            grid.addWidget(btn, i // 5, i % 5)
        layout.addLayout(grid)

        self.summary_box = QLabel("暂无未交记录")
        self.summary_box.setStyleSheet(
            "background-color: white; border: 1px solid #ddd; padding: 15px; border-radius: 5px; line-height: 1.5;")
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
        self.summary_box.setText("<b>今日未交汇总：</b><br>" + "<br>".join(res) if res else "今日作业全部交齐！")

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
        if not self.students:
            QMessageBox.warning(self, "提示", "请先导入名单")
            return
        self.detail = SubjectWindow(sub, self.students, self.current_date, self)
        self.detail.show()

    # --- 修改：查看历史记录功能 ---
    def view_history(self):
        dialog = DateRangeDialog(self, "选择要查看的历史时段")
        if dialog.exec() == QDialog.DialogCode.Accepted:
            start_date, end_date = dialog.get_dates()
            if start_date > end_date:
                QMessageBox.warning(self, "错误", "开始日期不能晚于结束日期！")
                return

            # 打开新的历史记录展示窗口
            self.history_win = HistoryResultDialog(start_date, end_date, self)
            self.history_win.show()

    # --- 修改：数据导出功能 ---
    def export_excel(self):
        dialog = DateRangeDialog(self, "选择要导出的数据时段")
        if dialog.exec() == QDialog.DialogCode.Accepted:
            start_date, end_date = dialog.get_dates()
            if start_date > end_date:
                QMessageBox.warning(self, "错误", "开始日期不能晚于结束日期！")
                return

            try:
                all_rows = []
                current = start_date
                while current <= end_date:
                    d_str = current.strftime("%Y-%m-%d")
                    d_data = self.load_daily_data(d_str)
                    for name in self.students:
                        row = {"姓名": name, "日期": d_str}
                        for s in self.subjects:
                            status = d_data.get(s, {}).get("data", {}).get(name, "normal")
                            row[s] = 1 if status == "missing" else 0
                        all_rows.append(row)
                    current += timedelta(days=1)

                if not all_rows:
                    QMessageBox.warning(self, "提示", "该时段内没有任何记录可导出。")
                    return

                df = pd.DataFrame(all_rows)
                summary = df.groupby("姓名")[self.subjects].sum()
                summary['汇总'] = summary.sum(axis=1)

                file_name = f"作业统计_{start_date.strftime('%m%d')}-{end_date.strftime('%m%d')}.xlsx"
                out_path = os.path.join(BASE_DIR, file_name)

                with pd.ExcelWriter(out_path) as writer:
                    sheet_name = f"{start_date.strftime('%m.%d')}-{end_date.strftime('%m.%d')}汇总"
                    summary.to_excel(writer, sheet_name=sheet_name)

                os.startfile(BASE_DIR)
                QMessageBox.information(self, "导出成功", f"表格已导出至：\n{out_path}")
            except Exception as e:
                QMessageBox.critical(self, "导出失败", f"导出过程中出现错误: {str(e)}")


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
        save_btn.setStyleSheet("background-color: #4a90e2; color: white; height: 40px; font-weight: bold;")
        save_btn.clicked.connect(self.do_save)
        layout.addWidget(save_btn)
        self.update_stats()

    def style_btn(self, btn, status):
        if status == "normal":
            btn.setStyleSheet("background-color: white; color: black; border: 1px solid #ccc;")
            btn.setText(btn.text().split('\n')[0])
        elif status == "missing":
            btn.setStyleSheet("background-color: #ffeded; color: red; border: 2px solid red;")
            btn.setText(btn.text().split('\n')[0] + "\n✘")
        elif status == "late":
            btn.setStyleSheet("background-color: #f5f5f5; color: #999; border: 1px solid #ddd;")
            btn.setText(btn.text().split('\n')[0] + "\n✓")

    def click_name(self, name, btn):
        curr = self.sub_info["data"].get(name, "normal")
        locked = self.sub_info.get("status") == "locked"

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
            else:
                return  # 已交或补交的锁死后不能再点

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