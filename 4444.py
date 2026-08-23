import sys
import os
import sqlite3
from datetime import datetime, timedelta

# 显式导入 pandas 和 openpyxl，强制 PyInstaller 在打包时抓取这些依赖
import pandas as pd
import openpyxl

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget,
                             QVBoxLayout, QHBoxLayout, QGridLayout,
                             QLabel, QPushButton, QFrame, QStackedWidget,
                             QFileDialog, QDialog, QMessageBox, QTextBrowser,
                             QDateEdit, QTableWidget, QTableWidgetItem, QHeaderView,
                             QScrollArea, QTextEdit, QTimeEdit, QDoubleSpinBox)
from PyQt6.QtCore import Qt, QTimer, QDateTime, pyqtSignal, QDate, QTime
from PyQt6.QtGui import QColor

# ==========================================
# 1. 本地存储配置：极度安全的目录初始化
# ==========================================
# 尝试使用 D 盘，如果不存在则使用用户目录，防止跨平台或缺盘崩溃
if os.path.exists("D:\\"):
    WORKSPACE_DIR = r"D:\Class12_Workspace"
else:
    WORKSPACE_DIR = os.path.join(os.path.expanduser("~"), "Class12_Workspace")

# 安全创建目录，坚决不能在这里使用 print()，否则打包成 exe 后无控制台模式会直接闪退
try:
    if not os.path.exists(WORKSPACE_DIR):
        os.makedirs(WORKSPACE_DIR)
except Exception:
    # 终极备用方案：如果在C盘或D盘都没有权限，就在 exe 所在文件夹创建
    WORKSPACE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Class12_Workspace")
    if not os.path.exists(WORKSPACE_DIR):
        try:
            os.makedirs(WORKSPACE_DIR)
        except Exception:
            pass # 放弃创建，避免程序直接崩溃

DB_PATH = os.path.join(WORKSPACE_DIR, "class_data.db")


def mask_name(name):
    """脱敏函数：仅保留头尾，中间用星号代替"""
    if not name: return ""
    if len(name) <= 2: return f"{name[0]}*" if len(name) > 1 else name
    return f"{name[0]}*{name[-1]}"


# ==========================================
# 2. 数据库操作工具类
# ==========================================
class DatabaseHelper:
    @staticmethod
    def init_db():
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS students (name TEXT UNIQUE)''')
        c.execute('''CREATE TABLE IF NOT EXISTS homework 
                     (date TEXT, subject TEXT, student_name TEXT, status TEXT,
                      PRIMARY KEY (date, subject, student_name))''')
        c.execute('''CREATE TABLE IF NOT EXISTS notices 
                     (id INTEGER PRIMARY KEY AUTOINCREMENT, content TEXT, timestamp TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS deductions 
                     (id INTEGER PRIMARY KEY AUTOINCREMENT, timestamp TEXT, 
                      category TEXT, student_name TEXT, reason TEXT, points REAL)''')
        c.execute('''CREATE TABLE IF NOT EXISTS config (key TEXT PRIMARY KEY, value TEXT)''')
        c.execute("INSERT OR IGNORE INTO config (key, value) VALUES ('refresh_time', '17:00')")
        conn.commit()
        conn.close()

    @staticmethod
    def import_students(file_path):
        try:
            df = pd.read_excel(file_path)
            if '姓名' not in df.columns: return "Excel中未找到'姓名'列"
            names = df['姓名'].dropna().tolist()
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            for name in names:
                c.execute("INSERT OR IGNORE INTO students (name) VALUES (?)", (name,))
            conn.commit()
            conn.close()
            return len(names)
        except Exception as e:
            return str(e)

    @staticmethod
    def get_all_students():
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT name FROM students")
        names = [row[0] for row in c.fetchall()]
        conn.close()
        return names

    @staticmethod
    def get_refresh_time():
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT value FROM config WHERE key='refresh_time'")
        row = c.fetchone()
        conn.close()
        return row[0] if row else "17:00"

    @staticmethod
    def set_refresh_time(time_str):
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("REPLACE INTO config (key, value) VALUES ('refresh_time', ?)", (time_str,))
        conn.commit()
        conn.close()

    @staticmethod
    def get_current_period_start():
        rt_str = DatabaseHelper.get_refresh_time()
        hour, minute = map(int, rt_str.split(':'))
        now = datetime.now()
        refresh_today = now.replace(hour=hour, minute=minute, second=0, microsecond=0)

        if now >= refresh_today:
            return refresh_today.strftime("%Y-%m-%d %H:%M:%S")
        else:
            refresh_yesterday = (now - timedelta(days=1)).replace(hour=hour, minute=minute, second=0, microsecond=0)
            return refresh_yesterday.strftime("%Y-%m-%d %H:%M:%S")

    @staticmethod
    def add_deduction(category, student_name, reason, points):
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        c.execute("INSERT INTO deductions (timestamp, category, student_name, reason, points) VALUES (?, ?, ?, ?, ?)",
                  (timestamp, category, student_name, reason, float(points)))
        conn.commit()
        conn.close()

    @staticmethod
    def add_notice(content):
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("INSERT INTO notices (content, timestamp) VALUES (?, ?)",
                  (content, datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        conn.commit()
        conn.close()

    @staticmethod
    def delete_notice(notice_id):
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("DELETE FROM notices WHERE id=?", (notice_id,))
        conn.commit()
        conn.close()

    @staticmethod
    def get_notices():
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT id, content FROM notices ORDER BY id DESC")
        records = c.fetchall()
        conn.close()
        return records

    @staticmethod
    def import_schedule(file_path):
        try:
            df = pd.read_excel(file_path)
            conn = sqlite3.connect(DB_PATH)
            df.to_sql('schedule_table', conn, if_exists='replace', index=False)
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            return str(e)

    @staticmethod
    def get_schedule_df():
        try:
            conn = sqlite3.connect(DB_PATH)
            df = pd.read_sql_query("SELECT * FROM schedule_table", conn)
            conn.close()
            return df
        except:
            return pd.DataFrame()

    @staticmethod
    def clear_schedule():
        try:
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            c.execute("DROP TABLE IF EXISTS schedule_table")
            conn.commit()
            conn.close()
        except:
            pass

    @staticmethod
    def import_duty(file_path):
        try:
            df = pd.read_excel(file_path)
            conn = sqlite3.connect(DB_PATH)
            df.to_sql('duty_table', conn, if_exists='replace', index=False)
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            return str(e)

    @staticmethod
    def get_duty_df():
        try:
            conn = sqlite3.connect(DB_PATH)
            df = pd.read_sql_query("SELECT * FROM duty_table", conn)
            conn.close()
            return df
        except:
            return pd.DataFrame()

    @staticmethod
    def clear_duty():
        try:
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            c.execute("DROP TABLE IF EXISTS duty_table")
            conn.commit()
            conn.close()
        except:
            pass

    @staticmethod
    def get_today_col(df):
        weekday = datetime.now().weekday()
        keywords = ["周一", "周二", "周三", "周四", "周五", "周六", "周日"]
        search_term = keywords[weekday]
        for idx, col in enumerate(df.columns):
            if weekday >= 5:
                if "六" in col or "日" in col or "末" in col: return idx, col
            else:
                if search_term in col: return idx, col
        return -1, None


# ==========================================
# 3. 基础卡片类
# ==========================================
class ClickableFrame(QFrame):
    nav_clicked = pyqtSignal(int)

    def __init__(self, target_index, parent=None):
        super().__init__(parent)
        self.target_index = target_index
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setStyleSheet(
            "QFrame { background-color: #E8E4DA; border-radius: 6px; border: 1px solid #D5D0C5; } QFrame:hover { background-color: #DCD7CB; border: 1px solid #B5AFA1; }")

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton: event.accept()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            if self.rect().contains(event.pos()): self.nav_clicked.emit(self.target_index)
            event.accept()


class OverviewClickableCard(QFrame):
    nav_clicked = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setStyleSheet(
            "QFrame { background-color: #F8F6F1; border: 1px solid #D5D0C5; border-radius: 6px; } QFrame:hover { background-color: #E8E4DA; border: 1px solid #B5AFA1; }")
        self.layout = QVBoxLayout(self)
        self.text_label = QLabel()
        self.text_label.setStyleSheet(
            "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; line-height: 1.5; color: #000000; border: none; background: transparent;")
        self.text_label.setWordWrap(True)
        self.text_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        self.layout.addWidget(self.text_label)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton: event.accept()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            if self.rect().contains(event.pos()): self.nav_clicked.emit()
            event.accept()


# ==========================================
# 主界面的独立卡片
# ==========================================
class DeductionDashboardCard(ClickableFrame):
    def __init__(self, target_index):
        super().__init__(target_index)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)

        title_label = QLabel("扣分情况提醒")
        title_label.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 28px; font-weight: bold; color: #000000; border: none; background: transparent;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        self.content_label = QLabel()
        self.content_label.setWordWrap(True)
        self.content_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignCenter)
        self.content_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        layout.addStretch()
        layout.addWidget(title_label)
        layout.addSpacing(15)
        layout.addWidget(self.content_label)
        layout.addStretch()
        self.refresh_data()

    def refresh_data(self):
        start_time = DatabaseHelper.get_current_period_start()
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT category, student_name, reason, points FROM deductions WHERE timestamp >= ?", (start_time,))
        records = c.fetchall()
        conn.close()

        if not records:
            self.content_label.setText("当前周期暂无违纪扣分")
            self.content_label.setStyleSheet(
                "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; color: #000000; border: none; background: transparent;")
            self.content_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            return

        cat_data = {'值周班': [], '晚自习': [], '寝室': []}
        for cat, name, reason, pts in records:
            if cat in cat_data:
                masked_name = mask_name(name)
                pts_str = f"{pts:g}"
                cat_data[cat].append(f"{masked_name}：{reason}-{pts_str}")

        colors = {'值周班': ('#E6EDF2', '#2C3E50'), '晚自习': ('#F5E6E6', '#641E16'), '寝室': ('#E9F0E6', '#1D632F')}
        html_lines = []
        for cat in ['值周班', '晚自习', '寝室']:
            if cat_data[cat]:
                bg_c, txt_c = colors[cat]
                items = "<br>".join(cat_data[cat])
                block = f"<div style='background-color: {bg_c}; padding: 8px; margin-bottom: 8px; border-radius: 4px;'><span style='color: {txt_c}; font-weight: bold;'>【{cat}】</span><br><span style='color: {txt_c};'>{items}</span></div>"
                html_lines.append(block)

        self.content_label.setText("".join(html_lines))
        self.content_label.setStyleSheet(
            "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; line-height: 1.5; border: none; background: transparent;")
        self.content_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)


class NoticeDashboardCard(ClickableFrame):
    def __init__(self, target_index):
        super().__init__(target_index)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)

        layout.addStretch()
        title_label = QLabel("重要事项通知")
        title_label.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 28px; font-weight: bold; color: #000000; border: none; background: transparent;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        self.content_label = QLabel()
        self.content_label.setWordWrap(True)
        self.content_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignCenter)
        self.content_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        layout.addWidget(title_label)
        layout.addSpacing(15)
        layout.addWidget(self.content_label)
        layout.addStretch()
        self.refresh_data()

    def refresh_data(self):
        records = DatabaseHelper.get_notices()
        if not records:
            self.content_label.setText("暂无重要通知")
            self.content_label.setStyleSheet(
                "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; color: #000000; border: none; background: transparent;")
            self.content_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        else:
            lines = [f"✦ {content}" for _, content in records]
            self.content_label.setText("\n\n".join(lines))
            self.content_label.setStyleSheet(
                "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; line-height: 1.6; color: #000000; border: none; background: transparent;")
            self.content_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)


class HomeworkDashboardCard(ClickableFrame):
    def __init__(self, target_index):
        super().__init__(target_index)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)

        layout.addStretch()
        title_label = QLabel("作业上交登记")
        title_label.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 28px; font-weight: bold; color: #000000; border: none; background: transparent;")
        title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        self.content_label = QLabel()
        self.content_label.setWordWrap(True)
        self.content_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        layout.addWidget(title_label)
        layout.addSpacing(15)
        layout.addWidget(self.content_label)
        layout.addStretch()
        self.refresh_data()

    def refresh_data(self):
        today = datetime.now().strftime("%Y-%m-%d")
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT subject, student_name FROM homework WHERE date=? AND status='未交'", (today,))
        records = c.fetchall()
        conn.close()
        if not records:
            self.content_label.setText("今日暂无未交作业")
            self.content_label.setStyleSheet(
                "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; color: #000000; border: none; background: transparent;")
            self.content_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        else:
            summary = {}
            for sub, name in records:
                if sub not in summary: summary[sub] = []
                summary[sub].append(name)
            lines = [f"<b>{sub}</b>：{'、'.join(names)}" for sub, names in summary.items()]
            self.content_label.setText("\n".join(lines))
            self.content_label.setStyleSheet(
                "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; line-height: 1.5; color: #000000; border: none; background: transparent;")
            self.content_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)


class ScheduleCard(ClickableFrame):
    def __init__(self, target_index):
        super().__init__(target_index)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 20, 15, 20)
        layout.setSpacing(8)

        layout.addStretch()
        title = QLabel("课 表")
        title.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 28px; font-weight: bold; color: #000000; border: none; background: transparent;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        layout.addWidget(title)
        layout.addSpacing(15)

        self.subject_labels = []

        def add_period_row(num_str):
            row = QHBoxLayout()
            num_label = QLabel(num_str)
            num_label.setStyleSheet(
                "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 24px; font-weight: bold; color: #000000; border: none; background: transparent; padding-right: 5px;")
            num_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            num_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

            subject_label = QLabel("")
            subject_label.setStyleSheet(
                "background-color: #F4F1EA; border-radius: 6px; border: 1px solid #DFDBD1; font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; color: #000000;")
            subject_label.setMinimumHeight(40)
            subject_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            subject_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

            row.addWidget(num_label, 1)
            row.addWidget(subject_label, 3)
            layout.addLayout(row)
            self.subject_labels.append(subject_label)

        for i in ["1", "2", "3", "4", "5"]: add_period_row(i)

        layout.addSpacing(8)
        hline = QFrame()
        hline.setFrameShape(QFrame.Shape.HLine)
        hline.setStyleSheet("border: none; border-top: 2px solid #000000; background: transparent;")
        layout.addWidget(hline)
        layout.addSpacing(8)

        for i in ["6", "7", "8", "9"]: add_period_row(i)

        layout.addStretch()
        self.refresh_data()

    def refresh_data(self):
        df = DatabaseHelper.get_schedule_df()
        for label in self.subject_labels: label.setText("")
        if df.empty:
            if self.subject_labels: self.subject_labels[0].setText("暂未导入")
            return
        target_col_idx, _ = DatabaseHelper.get_today_col(df)
        if target_col_idx != -1:
            row_idx = 0
            for r in range(len(df)):
                if row_idx >= len(self.subject_labels): break
                val = str(df.iloc[r, target_col_idx])
                if val in ('nan', 'None'): val = ""
                self.subject_labels[row_idx].setText(val)
                row_idx += 1


class DutyCard(ClickableFrame):
    def __init__(self, target_index):
        super().__init__(target_index)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 20, 15, 20)

        layout.addStretch()
        title = QLabel("值日安排")
        title.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 28px; font-weight: bold; color: #000000; border: none; background: transparent;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        self.content_label = QLabel()
        self.content_label.setStyleSheet(
            "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; color: #000000; border: none; background: transparent; line-height: 1.8;")
        self.content_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.content_label.setWordWrap(True)
        self.content_label.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

        layout.addWidget(title)
        layout.addSpacing(15)
        layout.addWidget(self.content_label)
        layout.addStretch()
        self.refresh_data()

    def refresh_data(self):
        df = DatabaseHelper.get_duty_df()
        if df.empty:
            self.content_label.setText("今日值日数据\n暂未导入")
            self.content_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            return
        target_col_idx, target_col_name = DatabaseHelper.get_today_col(df)
        if target_col_name and len(df.columns) > 0:
            task_col = df.columns[0]
            lines = []
            for _, row in df.iterrows():
                task = str(row[task_col])
                student = str(row[target_col_name])
                if pd.notna(student) and student.strip() and student not in ('nan', 'None'):
                    lines.append(f"<b>{task}</b>：{student}")
            if lines:
                self.content_label.setText("<br>".join(lines))
                self.content_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            else:
                self.content_label.setText("今日无值日安排")
                self.content_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        else:
            self.content_label.setText("未匹配到今日安排列")
            self.content_label.setAlignment(Qt.AlignmentFlag.AlignCenter)


# ==========================================
# 4. 扣分情况提醒系统
# ==========================================
class RefreshTimeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置刷新时间")
        self.resize(400, 200)
        self.setStyleSheet("background-color: #F2EFE9;")
        layout = QVBoxLayout(self)
        lbl = QLabel("清空展示的刷新时间：")
        lbl.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 24px; font-weight: bold; color: #000000;")
        layout.addWidget(lbl)
        self.time_edit = QTimeEdit()
        current_rt = DatabaseHelper.get_refresh_time()
        self.time_edit.setTime(QTime.fromString(current_rt, "HH:mm"))
        self.time_edit.setStyleSheet(
            "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; padding: 5px; background: white; color: #000000;")
        layout.addWidget(self.time_edit)
        btn_save = QPushButton("保存")
        btn_save.setFixedHeight(50)
        btn_save.setStyleSheet(
            "background-color: #000000; color: white; border-radius: 4px; font-size: 24px; font-weight: bold;")
        btn_save.clicked.connect(self.save)
        layout.addWidget(btn_save)

    def save(self):
        new_time = self.time_edit.time().toString("HH:mm")
        DatabaseHelper.set_refresh_time(new_time)
        QMessageBox.information(self, "成功", f"刷新时间已更新为每天 {new_time}")
        self.accept()


class StudentMultiSelectDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("选择违纪学生 (可多选)")
        self.resize(800, 600)
        self.setStyleSheet("background-color: #F2EFE9;")
        self.selected_students = []
        layout = QVBoxLayout(self)
        grid = QGridLayout()
        grid.setSpacing(15)
        self.btns = []
        students = DatabaseHelper.get_all_students()
        for i, name in enumerate(students):
            btn = QPushButton(name)
            btn.setFixedSize(120, 50)
            btn.setCheckable(True)
            btn.setStyleSheet("""
                QPushButton { font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; background-color: #E8E4DA; border: 1px solid #D5D0C5; border-radius: 4px; color: #000000;}
                QPushButton:checked { background-color: #D32F2F; color: white; font-weight: bold; border: none; }
            """)
            self.btns.append(btn)
            grid.addWidget(btn, i // 6, i % 6)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none;")
        inner_w = QWidget()
        inner_w.setLayout(grid)
        scroll.setWidget(inner_w)
        layout.addWidget(scroll)
        btn_ok = QPushButton("✅ 确认选择")
        btn_ok.setFixedHeight(50)
        btn_ok.setStyleSheet(
            "background-color: #000000; color: white; font-size: 24px; font-weight: bold; border-radius: 4px;")
        btn_ok.clicked.connect(self.confirm)
        layout.addWidget(btn_ok)

    def confirm(self):
        self.selected_students = [btn.text() for btn in self.btns if btn.isChecked()]
        if not self.selected_students:
            QMessageBox.warning(self, "提示", "请至少选择一名学生")
            return
        self.accept()


class DeductionInputDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("录入扣分详情")
        self.resize(500, 350)
        self.setStyleSheet("background-color: #F2EFE9;")
        self.reason = ""
        self.points = 0.0
        layout = QVBoxLayout(self)
        lbl1 = QLabel("📝 扣分原因：")
        lbl1.setStyleSheet(
            "color: #000000; font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 24px; font-weight: bold;")
        layout.addWidget(lbl1)
        self.reason_edit = QTextEdit()
        self.reason_edit.setFixedHeight(100)
        self.reason_edit.setStyleSheet(
            "background: white; font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; color: #000000;")
        layout.addWidget(self.reason_edit)

        lbl2 = QLabel("🔢 扣分分数 (最小单位0.5)：")
        lbl2.setStyleSheet(
            "color: #000000; font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 24px; font-weight: bold;")
        layout.addWidget(lbl2)

        self.points_edit = QDoubleSpinBox()
        self.points_edit.setRange(0.5, 100.0)
        self.points_edit.setSingleStep(0.5)
        self.points_edit.setDecimals(1)
        self.points_edit.setStyleSheet(
            "background: white; font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; padding: 5px; color: #000000;")
        layout.addWidget(self.points_edit)

        btn_save = QPushButton("💾 保存记录")
        btn_save.setFixedHeight(50)
        btn_save.setStyleSheet(
            "background-color: #000000; color: white; font-size: 24px; font-weight: bold; border-radius: 4px;")
        btn_save.clicked.connect(self.save)
        layout.addWidget(btn_save)

    def save(self):
        self.reason = self.reason_edit.toPlainText().strip()
        if not self.reason:
            QMessageBox.warning(self, "提示", "请输入扣分原因！")
            return
        self.points = self.points_edit.value()
        self.accept()


class CategoryDeductionDialog(QDialog):
    def __init__(self, category, color_bg, color_txt, parent=None):
        super().__init__(parent)
        self.category = category
        self.color_bg = color_bg
        self.color_txt = color_txt
        self.setWindowTitle(f"【{category}】扣分管理")
        self.resize(700, 600)
        self.setStyleSheet(f"background-color: {color_bg};")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        top_l = QHBoxLayout()
        title = QLabel(f"{self.category} - 本周期扣分明细")
        title.setStyleSheet(
            f"font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 28px; font-weight: bold; color: {self.color_txt};")
        btn_add = QPushButton("➕ 增加记录")
        btn_add.setFixedSize(160, 50)
        btn_add.setStyleSheet(
            f"background-color: {self.color_txt}; color: white; border-radius: 6px; font-size: 24px; font-weight: bold;")
        btn_add.clicked.connect(self.add_record)
        top_l.addWidget(title)
        top_l.addStretch()
        top_l.addWidget(btn_add)
        layout.addLayout(top_l)
        layout.addSpacing(15)

        self.browser = QTextBrowser()
        self.browser.setStyleSheet(
            f"background-color: rgba(255,255,255,0.6); border: 1px solid {self.color_txt}; border-radius: 6px; font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; padding: 15px; line-height: 1.6; color: {self.color_txt};")
        layout.addWidget(self.browser)
        self.load_data()

    def load_data(self):
        start = DatabaseHelper.get_current_period_start()
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute(
            "SELECT student_name, reason, points FROM deductions WHERE category=? AND timestamp>=? ORDER BY id DESC",
            (self.category, start))
        records = c.fetchall()
        conn.close()

        if not records:
            self.browser.setText("当前周期暂无扣分记录。")
            return
        lines = []
        for name, reason, pts in records:
            pts_str = f"{pts:g}"
            lines.append(f"✦ {mask_name(name)}：{reason} -{pts_str}")
        self.browser.setText("\n\n".join(lines))

    def add_record(self):
        sel_dialog = StudentMultiSelectDialog(self)
        if sel_dialog.exec() == QDialog.DialogCode.Accepted:
            inp_dialog = DeductionInputDialog(self)
            if inp_dialog.exec() == QDialog.DialogCode.Accepted:
                for stu in sel_dialog.selected_students:
                    DatabaseHelper.add_deduction(self.category, stu, inp_dialog.reason, inp_dialog.points)
                self.load_data()


class DeductionQueryDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("🔍 数据查询 (扣分统计)")
        self.resize(700, 500)
        self.setStyleSheet("background-color: #F2EFE9;")
        layout = QVBoxLayout(self)

        dl = QHBoxLayout()
        self.d1 = QDateEdit()
        self.d2 = QDateEdit()
        for d in [self.d1, self.d2]:
            d.setCalendarPopup(True)
            d.setDate(QDate.currentDate())
            d.setStyleSheet("font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; padding: 5px; color: #000000;")

        btn = QPushButton("查询汇总")
        btn.setStyleSheet(
            "background-color: #000000; color: white; border-radius: 4px; font-size: 24px; font-weight: bold; padding: 5px 15px;")
        btn.clicked.connect(self.query)

        lbl_from = QLabel("从：")
        lbl_from.setStyleSheet("color: #000000; font-size: 24px;")
        lbl_to = QLabel("到：")
        lbl_to.setStyleSheet("color: #000000; font-size: 24px;")
        dl.addWidget(lbl_from)
        dl.addWidget(self.d1)
        dl.addWidget(lbl_to)
        dl.addWidget(self.d2)
        dl.addWidget(btn)
        layout.addLayout(dl)

        self.browser = QTextBrowser()
        self.browser.setStyleSheet(
            "background: white; font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; color: #000000;")
        layout.addWidget(self.browser)

    def query(self):
        sd = self.d1.date().toString("yyyy-MM-dd")
        ed = self.d2.date().addDays(1).toString("yyyy-MM-dd")

        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute(
            "SELECT student_name, category, SUM(points) FROM deductions WHERE timestamp >= ? AND timestamp < ? GROUP BY student_name, category",
            (sd, ed))
        records = c.fetchall()
        conn.close()

        if not records:
            self.browser.setText("该时段内无任何扣分记录。")
            return
        summary = {}
        for name, cat, pts in records:
            if name not in summary: summary[name] = []
            pts_str = f"{pts:g}"
            summary[name].append(f"{cat}扣{pts_str}分")

        lines = [f"<b>{name}</b>：{'，'.join(details)}" for name, details in summary.items()]
        self.browser.setHtml("<br><br>".join(lines))


class DeductionExportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📊 导出扣分记录")
        self.resize(500, 300)
        self.setStyleSheet("background-color: #F2EFE9;")
        layout = QVBoxLayout(self)

        dl = QGridLayout()
        self.d1 = QDateEdit()
        self.d2 = QDateEdit()
        for d in [self.d1, self.d2]:
            d.setCalendarPopup(True)
            d.setDate(QDate.currentDate())
            d.setStyleSheet("font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; padding: 5px; color: #000000;")

        lbl_s = QLabel("起始日期：")
        lbl_s.setStyleSheet("color: #000000; font-size: 24px; font-weight: bold;")
        lbl_e = QLabel("结束日期：")
        lbl_e.setStyleSheet("color: #000000; font-size: 24px; font-weight: bold;")
        dl.addWidget(lbl_s, 0, 0)
        dl.addWidget(self.d1, 0, 1)
        dl.addWidget(lbl_e, 1, 0)
        dl.addWidget(self.d2, 1, 1)
        layout.addLayout(dl)

        btn = QPushButton("📥 一键导出 Excel")
        btn.setFixedHeight(50)
        btn.setStyleSheet(
            "background-color: #000000; color: white; border-radius: 4px; font-size: 24px; font-weight: bold;")
        btn.clicked.connect(self.export)
        layout.addWidget(btn)

    def export(self):
        sd = self.d1.date().toString("yyyy-MM-dd")
        ed = self.d2.date().toString("yyyy-MM-dd")
        ed_query = self.d2.date().addDays(1).toString("yyyy-MM-dd")

        conn = sqlite3.connect(DB_PATH)
        df = pd.read_sql_query(
            f"SELECT student_name, category, SUM(points) as pts FROM deductions WHERE timestamp >= '{sd}' AND timestamp < '{ed_query}' GROUP BY student_name, category",
            conn)
        c = conn.cursor()
        c.execute("SELECT name FROM students")
        all_students = [row[0] for row in c.fetchall()]
        conn.close()

        if df.empty:
            QMessageBox.warning(self, "空", "该时段无扣分记录。")
            return

        pivot_df = df.pivot(index='student_name', columns='category', values='pts').fillna(0)
        missing = set(all_students) - set(pivot_df.index)
        for stu in missing: pivot_df.loc[stu] = 0

        pivot_df['总扣分'] = pivot_df.sum(axis=1)
        pivot_df = pivot_df.reset_index().rename(columns={'student_name': '姓名'})

        export_path = os.path.join(WORKSPACE_DIR, f"扣分汇总_{sd}_至_{ed}.xlsx")
        try:
            pivot_df.to_excel(export_path, index=False)
            QMessageBox.information(self, "成功", f"导出成功：\n{export_path}")
            self.accept()
        except Exception as e:
            QMessageBox.warning(self, "失败", str(e))


class DeductionPage(QWidget):
    def __init__(self, back_callback, parent=None):
        super().__init__(parent)
        self.back_callback = back_callback
        self.setStyleSheet("background-color: #F2EFE9;")
        self.cat_content_labels = {}
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none; background-color: transparent;")

        inner_w = QWidget()
        inner_w.setStyleSheet("background-color: transparent;")
        inner_layout = QVBoxLayout(inner_w)
        inner_layout.setContentsMargins(30, 30, 30, 30)

        top_layout = QHBoxLayout()
        btn_back = QPushButton("◀ 返回平台")
        btn_back.setFixedSize(160, 50)
        btn_back.setStyleSheet(
            "QPushButton { font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; background-color: #E8E4DA; border: 1px solid #D5D0C5; border-radius: 4px; font-size: 24px; font-weight: bold; color: #000000;} QPushButton:hover { background-color: #DCD7CB; }")
        btn_back.clicked.connect(self.back_callback)

        title = QLabel("扣分情况提醒")
        title.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 28px; font-weight: bold; color: #000000;")

        btn_import = QPushButton("导入名单")
        btn_query = QPushButton("数据查询")
        btn_export = QPushButton("数据导出")
        btn_time = QPushButton("刷新时间")
        for btn in [btn_import, btn_query, btn_export, btn_time]:
            btn.setFixedSize(140, 50)
            btn.setStyleSheet(btn_back.styleSheet())

        btn_import.clicked.connect(self.import_excel)
        btn_query.clicked.connect(lambda: DeductionQueryDialog(self).exec())
        btn_export.clicked.connect(lambda: DeductionExportDialog(self).exec())
        btn_time.clicked.connect(lambda: RefreshTimeDialog(self).exec())

        top_layout.addWidget(btn_back)
        top_layout.addSpacing(20)
        top_layout.addWidget(title)
        top_layout.addStretch()
        top_layout.addWidget(btn_import)
        top_layout.addWidget(btn_query)
        top_layout.addWidget(btn_export)
        top_layout.addWidget(btn_time)
        inner_layout.addLayout(top_layout)
        inner_layout.addSpacing(20)

        cats = [("值周班", "#E6EDF2", "#2C3E50"), ("晚自习", "#F5E6E6", "#641E16"), ("寝室", "#E9F0E6", "#1D632F")]
        panels_layout = QHBoxLayout()

        for cat_name, bg_c, txt_c in cats:
            frame = ClickableFrame(-1)
            frame.setStyleSheet(
                f"QFrame {{ background-color: {bg_c}; border-radius: 8px; border: 1px solid {txt_c}; }} QFrame:hover {{ border: 2px solid {txt_c}; }}")
            fl = QVBoxLayout(frame)
            lbl = QLabel(cat_name)
            lbl.setStyleSheet(
                f"font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 28px; font-weight: bold; color: {txt_c}; border: none;")
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)

            content_lbl = QLabel()
            content_lbl.setStyleSheet(
                f"font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; color: {txt_c}; border: none; line-height: 1.6;")
            content_lbl.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignCenter)
            content_lbl.setWordWrap(True)
            content_lbl.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
            self.cat_content_labels[cat_name] = content_lbl

            fl.addWidget(lbl)
            fl.addSpacing(20)
            fl.addWidget(content_lbl)
            fl.addStretch()

            frame.mouseReleaseEvent = lambda e, c=cat_name, b=bg_c, t=txt_c: self.open_category(c, b,
                                                                                                t) if e.button() == Qt.MouseButton.LeftButton else None
            panels_layout.addWidget(frame)

        inner_layout.addLayout(panels_layout)
        scroll.setWidget(inner_w)
        main_layout.addWidget(scroll)
        self.refresh_data()

    def refresh_data(self):
        start_time = DatabaseHelper.get_current_period_start()
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT category, student_name, reason, points FROM deductions WHERE timestamp >= ?", (start_time,))
        records = c.fetchall()
        conn.close()

        cat_data = {'值周班': [], '晚自习': [], '寝室': []}
        for cat, name, reason, pts in records:
            if cat in cat_data:
                masked_name = mask_name(name)
                pts_str = f"{pts:g}"
                cat_data[cat].append(f"{masked_name}：{reason} -{pts_str}")

        for cat_name, lbl in self.cat_content_labels.items():
            if not cat_data[cat_name]:
                lbl.setText("今日暂无扣分")
            else:
                lines = [f"{item}" for item in cat_data[cat_name]]
                lbl.setText("<br>".join(lines))

    def import_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择学生名单 Excel", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            res = DatabaseHelper.import_students(file_path)
            if isinstance(res, int): QMessageBox.information(self, "成功", f"导入 {res} 人")

    def open_category(self, cat, bg_c, txt_c):
        CategoryDeductionDialog(cat, bg_c, txt_c, self).exec()
        self.refresh_data()


# ==========================================
# 5. 重要通知及作业管理系统
# ==========================================
class AddNoticeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("添加通知")
        self.resize(600, 400)
        l = QVBoxLayout(self)
        self.te = QTextEdit()
        self.te.setStyleSheet("font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; color: #000000;")
        l.addWidget(self.te)
        btn = QPushButton("保存")
        btn.setFixedHeight(50)
        btn.setStyleSheet(
            "background-color: #000000; color: white; font-size: 24px; font-weight: bold; border-radius: 4px;")
        btn.clicked.connect(self.accept)
        l.addWidget(btn)

    def get_text(self):
        return self.te.toPlainText().strip()


class NoticePage(QWidget):
    def __init__(self, back_callback, parent=None):
        super().__init__(parent)
        self.back_callback = back_callback
        self.setStyleSheet("background-color: #F2EFE9;")
        self.init_ui()
        self.load_data()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none;")
        inner_w = QWidget()
        inner_w.setStyleSheet("background-color: transparent;")
        inner_layout = QVBoxLayout(inner_w)
        inner_layout.setContentsMargins(30, 30, 30, 30)

        top_layout = QHBoxLayout()
        btn_back = QPushButton("◀ 返回平台")
        btn_back.setFixedSize(160, 50)
        btn_back.setStyleSheet(
            "QPushButton { font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; background-color: #E8E4DA; border: 1px solid #D5D0C5; border-radius: 4px; font-size: 24px; font-weight: bold; color: #000000;} QPushButton:hover { background-color: #DCD7CB; }")
        btn_back.clicked.connect(self.back_callback)
        title = QLabel("重要事项通知")
        title.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 28px; font-weight: bold; color: #000000;")

        btn_add = QPushButton("➕ 添加通知")
        btn_add.setFixedSize(160, 50)
        btn_add.setStyleSheet(
            "QPushButton { font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; background-color: #000000; color: white; border-radius: 4px; font-size: 24px; font-weight: bold;} QPushButton:hover { background-color: #333333; }")
        btn_add.clicked.connect(self.add_notice)

        top_layout.addWidget(btn_back)
        top_layout.addSpacing(20)
        top_layout.addWidget(title)
        top_layout.addStretch()
        top_layout.addWidget(btn_add)
        inner_layout.addLayout(top_layout)
        inner_layout.addSpacing(20)

        self.list_widget = QWidget()
        self.list_widget.setStyleSheet("background-color: transparent;")
        self.list_layout = QVBoxLayout(self.list_widget)
        self.list_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        inner_layout.addWidget(self.list_widget)

        scroll.setWidget(inner_w)
        main_layout.addWidget(scroll)

    def load_data(self):
        while self.list_layout.count():
            c = self.list_layout.takeAt(0)
            if c.widget(): c.widget().deleteLater()
        records = DatabaseHelper.get_notices()
        if not records: return
        for nid, content in records:
            frame = QFrame()
            frame.setStyleSheet("QFrame { background-color: #F8F6F1; border: 1px solid #D5D0C5; border-radius: 6px; }")
            f_layout = QHBoxLayout(frame)
            lbl = QLabel(f"✦ {content}")
            lbl.setWordWrap(True)
            lbl.setStyleSheet(
                "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; color: #000000; border: none;")
            btn_del = QPushButton("×")
            btn_del.setFixedSize(50, 50)
            btn_del.setCursor(Qt.CursorShape.PointingHandCursor)
            btn_del.setStyleSheet(
                "QPushButton { background-color: transparent; color: #A6A6A6; font-size: 32px; border: none; } QPushButton:hover { color: #E74C3C; }")
            btn_del.clicked.connect(lambda checked, i=nid: self.del_notice(i))
            f_layout.addWidget(lbl, 1)
            f_layout.addWidget(btn_del)
            self.list_layout.addWidget(frame)

    def add_notice(self):
        d = AddNoticeDialog(self)
        if d.exec() == QDialog.DialogCode.Accepted:
            text = d.get_text()
            if text:
                DatabaseHelper.add_notice(text)
                self.load_data()

    def del_notice(self, nid):
        if QMessageBox.question(self, '确认', '删除该通知？') == QMessageBox.StandardButton.Yes:
            DatabaseHelper.delete_notice(nid)
            self.load_data()


class HomeworkHistoryDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📜 作业历史记录")
        self.resize(700, 600)
        self.setStyleSheet("background-color: #F2EFE9;")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        top_layout = QHBoxLayout()
        date_label = QLabel("选择查询日期：")
        date_label.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 24px; font-weight: bold; color: #000000;")
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setStyleSheet(
            "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; padding: 5px; background: white; color: #000000;")
        btn_search = QPushButton("🔍 查询")
        btn_search.setStyleSheet(
            "background-color: #000000; color: white; font-size: 24px; font-weight: bold; border-radius: 4px; padding: 5px 15px;")
        btn_search.clicked.connect(self.search_history)

        top_layout.addWidget(date_label)
        top_layout.addWidget(self.date_edit)
        top_layout.addWidget(btn_search)
        top_layout.addStretch()
        layout.addLayout(top_layout)
        layout.addSpacing(15)

        self.result_browser = QTextBrowser()
        self.result_browser.setStyleSheet(
            "background-color: #F8F6F1; border: 1px solid #D5D0C5; font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; padding: 15px; line-height: 1.6; color: #000000;")
        layout.addWidget(self.result_browser)
        self.search_history()

    def search_history(self):
        target_date = self.date_edit.date().toString("yyyy-MM-dd")
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT subject, student_name, status FROM homework WHERE date=? AND status IN ('未交', '已补交')",
                  (target_date,))
        records = c.fetchall()
        conn.close()

        if not records:
            self.result_browser.setHtml(f"<h3 style='color: #000000;'>{target_date} 表现完美，无未交作业记录。</h3>")
            return

        data_map = {}
        for sub, name, status in records:
            if sub not in data_map: data_map[sub] = []
            if status == "未交":
                data_map[sub].append(f"<span style='color: #D32F2F; font-weight: bold;'>{name}(未交)</span>")
            else:
                data_map[sub].append(f"<span style='color: #7A7469;'>{name}(已补交)</span>")

        html_lines = [
            f"<h3 style='border-bottom: 1px solid #ccc; padding-bottom: 5px; color: #000000;'>{target_date} 记录明细：</h3>"]
        for sub, names in data_map.items():
            html_lines.append(f"<p style='color: #000000;'><b>{sub}</b>：{'、'.join(names)}</p>")
        self.result_browser.setHtml("\n".join(html_lines))


class HomeworkExportDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📊 作业数据导出")
        self.resize(500, 300)
        self.setStyleSheet("background-color: #F2EFE9;")
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        info = QLabel("请选择需要导出的时间段：")
        info.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 24px; font-weight: bold; color: #000000;")
        layout.addWidget(info)
        layout.addSpacing(15)

        date_layout = QGridLayout()
        self.start_date = QDateEdit()
        self.start_date.setCalendarPopup(True)
        current_date = QDate.currentDate()
        self.start_date.setDate(current_date.addDays(-current_date.dayOfWeek() + 1))
        self.end_date = QDateEdit()
        self.end_date.setCalendarPopup(True)
        self.end_date.setDate(current_date)

        for w in [self.start_date, self.end_date]:
            w.setStyleSheet(
                "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; padding: 5px; background: white; color: #000000;")

        lbl_s = QLabel("起始日期：")
        lbl_s.setStyleSheet("color: #000000; font-size: 24px; font-weight: bold;")
        lbl_e = QLabel("结束日期：")
        lbl_e.setStyleSheet("color: #000000; font-size: 24px; font-weight: bold;")
        date_layout.addWidget(lbl_s, 0, 0)
        date_layout.addWidget(self.start_date, 0, 1)
        date_layout.addWidget(lbl_e, 1, 0)
        date_layout.addWidget(self.end_date, 1, 1)
        layout.addLayout(date_layout)
        layout.addStretch()

        btn_export = QPushButton("📥 一键导出 Excel")
        btn_export.setFixedHeight(50)
        btn_export.setStyleSheet(
            "background-color: #000000; color: white; border-radius: 4px; font-size: 24px; font-weight: bold;")
        btn_export.clicked.connect(self.do_export)
        layout.addWidget(btn_export)

    def do_export(self):
        start_str = self.start_date.date().toString("yyyy-MM-dd")
        end_str = self.end_date.date().toString("yyyy-MM-dd")
        if self.start_date.date() > self.end_date.date():
            QMessageBox.warning(self, "错误", "起始日期不能晚于结束日期！")
            return

        conn = sqlite3.connect(DB_PATH)
        query = f"SELECT student_name, subject, COUNT(*) as count FROM homework WHERE date BETWEEN '{start_str}' AND '{end_str}' AND status = '未交' GROUP BY student_name, subject"
        df = pd.read_sql_query(query, conn)
        c = conn.cursor()
        c.execute("SELECT name FROM students")
        all_students = [row[0] for row in c.fetchall()]
        conn.close()

        if df.empty and not all_students:
            QMessageBox.warning(self, "提示", "数据库中暂无记录。")
            return

        subjects = ["语文", "数学", "英语", "物理", "化学", "生物", "政治", "历史", "地理", "技术"]
        if df.empty:
            pivot_df = pd.DataFrame(0, index=all_students, columns=subjects)
        else:
            pivot_df = df.pivot(index='student_name', columns='subject', values='count').fillna(0)
            for sub in subjects:
                if sub not in pivot_df.columns: pivot_df[sub] = 0
            missing_students = set(all_students) - set(pivot_df.index)
            for stu in missing_students: pivot_df.loc[stu] = 0

        pivot_df = pivot_df[subjects]
        pivot_df['汇总'] = pivot_df.sum(axis=1)
        pivot_df = pivot_df.reset_index().rename(columns={'student_name': '姓名'})

        export_path = os.path.join(WORKSPACE_DIR, f"未交作业统计_{start_str}_至_{end_str}.xlsx")
        try:
            with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                pivot_df.to_excel(writer, index=False, startrow=1, sheet_name='未交统计')
                worksheet = writer.sheets['未交统计']
                end_col_letter = chr(65 + len(pivot_df.columns) - 1)
                worksheet.merge_cells(f'A1:{end_col_letter}1')
                worksheet['A1'].value = f"{start_str} 到 {end_str} 未交作业汇总名单"
            QMessageBox.information(self, "导出成功", f"文件已成功保存至：\n{export_path}")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"请检查文件是否被占用。\n{str(e)}")


class SubjectDialog(QDialog):
    def __init__(self, date_str, subject, parent=None):
        super().__init__(parent)
        self.date_str = date_str
        self.subject = subject
        self.setWindowTitle(f"{date_str} - {subject} 初始作业登记")
        self.resize(800, 600)
        self.setStyleSheet("background-color: #F2EFE9;")
        self.students = DatabaseHelper.get_all_students()
        self.buttons, self.status_map = {}, {}
        self.is_locked = False
        self.init_ui()
        self.load_data()

    def init_ui(self):
        layout = QVBoxLayout(self)
        info_label = QLabel(f"科目：{self.subject}   |   操作说明：点击姓名标记未交，保存后不可修改")
        info_label.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 24px; font-weight: bold; color: #000000;")
        layout.addWidget(info_label)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none; background: transparent;")
        inner_w = QWidget()
        grid = QGridLayout(inner_w)
        grid.setSpacing(15)
        for i, name in enumerate(self.students):
            btn = QPushButton(name)
            btn.setFixedSize(120, 60)
            self.set_btn_style(btn, "已交")
            btn.clicked.connect(lambda checked, n=name: self.toggle_status(n))
            self.buttons[name] = btn
            self.status_map[name] = "已交"
            grid.addWidget(btn, i // 6, i % 6)
        scroll.setWidget(inner_w)
        layout.addWidget(scroll)

        self.btn_save = QPushButton("💾 保存初始登记")
        self.btn_save.setFixedHeight(55)
        self.btn_save.setStyleSheet(
            "background-color: #000000; color: white; font-size: 24px; font-weight: bold; border-radius: 5px;")
        self.btn_save.clicked.connect(self.save_data)
        layout.addWidget(self.btn_save)

    def set_btn_style(self, btn, status):
        font = "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; border-radius: 4px;"
        if status == "已交":
            btn.setStyleSheet(f"{font} background-color: #E8E4DA; color: #000000; border: 1px solid #D5D0C5;")
            btn.setText(btn.text().split(' ')[0])
        elif status == "未交":
            btn.setStyleSheet(f"{font} background-color: #F8D7DA; color: #000000; border: 1px solid #F5C6CB;")
            btn.setText(f"{btn.text().split(' ')[0]} ❌")
        else:
            btn.setStyleSheet(f"{font} background-color: #E2E3E5; color: #000000; border: 1px solid #D6D8DB;")
            btn.setText(f"{btn.text().split(' ')[0]} ✓")

    def toggle_status(self, name):
        if self.is_locked: return
        current = self.status_map[name]
        btn = self.buttons[name]
        if current == "已交":
            self.status_map[name] = "未交"
            self.set_btn_style(btn, "未交")
        elif current == "未交":
            self.status_map[name] = "已交"
            self.set_btn_style(btn, "已交")

    def load_data(self):
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT student_name, status FROM homework WHERE date=? AND subject=?", (self.date_str, self.subject))
        records = c.fetchall()
        if records:
            self.is_locked = True
            self.btn_save.setEnabled(False)
            self.btn_save.setText("🔒 本科目初始登记已完成")
            self.btn_save.setStyleSheet(
                "background-color: #C4BEB1; color: #000000; font-size: 24px; font-weight: bold; border-radius: 5px; border: none;")
            for name, status in records:
                if name in self.buttons:
                    self.status_map[name] = status
                    self.set_btn_style(self.buttons[name], status)
        conn.close()

    def save_data(self):
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        for name, status in self.status_map.items():
            if status == "未交":
                c.execute("REPLACE INTO homework (date, subject, student_name, status) VALUES (?, ?, ?, ?)",
                          (self.date_str, self.subject, name, status))
        conn.commit()
        conn.close()
        self.accept()


class LateSubmissionDialog(QDialog):
    def __init__(self, date_str, parent=None):
        super().__init__(parent)
        self.date_str = date_str
        self.setWindowTitle(f"{date_str} - 补交登记界面")
        self.resize(800, 600)
        self.setStyleSheet("background-color: #F2EFE9;")
        self.buttons, self.status_map = {}, {}
        self.locked_uids = set()

        target_date = datetime.strptime(self.date_str, "%Y-%m-%d")
        deadline = target_date + timedelta(days=1)
        deadline = deadline.replace(hour=18, minute=30, second=0)
        self.is_expired = datetime.now() > deadline
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout(self)
        header_text = "⚠️ 已超过次日18:30，补交通道已关闭" if self.is_expired else "📝 请点击标红的学生姓名完成补交登记"
        info_label = QLabel(header_text)
        info_label.setStyleSheet(
            f"font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 24px; font-weight: bold; color: #000000;")
        layout.addWidget(info_label)
        layout.addSpacing(10)

        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT subject, student_name, status FROM homework WHERE date=? AND status IN ('未交', '已补交')",
                  (self.date_str,))
        records = c.fetchall()
        conn.close()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none; background: transparent;")
        inner_w = QWidget()
        inner_layout = QVBoxLayout(inner_w)

        if not records:
            empty = QLabel("🎉 迟交/未交作业记录为空！")
            empty.setStyleSheet(
                "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; color: #27AE60;")
            inner_layout.addWidget(empty)
        else:
            data_map = {}
            for sub, name, status in records:
                if sub not in data_map: data_map[sub] = []
                data_map[sub].append((name, status))

            for sub, students in data_map.items():
                sub_layout = QHBoxLayout()
                sub_label = QLabel(f"{sub}：")
                sub_label.setStyleSheet(
                    "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 24px; font-weight: bold; color: #000000;")
                sub_label.setFixedWidth(100)
                sub_layout.addWidget(sub_label)
                for name, status in students:
                    btn = QPushButton(name)
                    btn.setFixedSize(110, 50)
                    self.set_btn_style(btn, status)
                    uid = f"{sub}_{name}"
                    self.buttons[uid] = (btn, sub, name)
                    self.status_map[uid] = status
                    if status == "已补交":
                        self.locked_uids.add(uid)
                    btn.clicked.connect(lambda checked, u=uid: self.toggle_late_status(u))
                    sub_layout.addWidget(btn)
                sub_layout.addStretch()
                inner_layout.addLayout(sub_layout)

        inner_layout.addStretch()
        scroll.setWidget(inner_w)
        layout.addWidget(scroll)

        self.btn_save = QPushButton("💾 保存补交状态")
        self.btn_save.setFixedHeight(55)
        self.btn_save.setStyleSheet(
            "background-color: #000000; color: white; font-size: 24px; font-weight: bold; border-radius: 5px;")
        self.btn_save.clicked.connect(self.save_data)
        if self.is_expired or not records:
            self.btn_save.setEnabled(False)
            self.btn_save.setStyleSheet("background-color: #C4BEB1; color: #000000; border: none;")
        layout.addWidget(self.btn_save)

    def set_btn_style(self, btn, status):
        font = "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; border-radius: 4px;"
        if status == "未交":
            btn.setStyleSheet(f"{font} background-color: #F8D7DA; color: #000000; border: 1px solid #F5C6CB;")
            btn.setText(f"{btn.text().split(' ')[0]} ❌")
        elif status == "已补交":
            btn.setStyleSheet(f"{font} background-color: #E2E3E5; color: #000000; border: 1px solid #D6D8DB;")
            btn.setText(f"{btn.text().split(' ')[0]} ✓")

    def toggle_late_status(self, uid):
        if self.is_expired or uid in self.locked_uids: return
        current = self.status_map[uid]
        btn, sub, name = self.buttons[uid]
        if current == "未交":
            self.status_map[uid] = "已补交"
            self.set_btn_style(btn, "已补交")
        elif current == "已补交":
            self.status_map[uid] = "未交"
            self.set_btn_style(btn, "未交")

    def save_data(self):
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        for uid, status in self.status_map.items():
            _, sub, name = self.buttons[uid]
            c.execute("UPDATE homework SET status=? WHERE date=? AND subject=? AND student_name=?",
                      (status, self.date_str, sub, name))
        conn.commit()
        conn.close()
        self.accept()


class HomeworkPage(QWidget):
    def __init__(self, back_callback, parent=None):
        super().__init__(parent)
        self.back_callback = back_callback
        self.setStyleSheet("background-color: #F2EFE9;")
        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none;")
        inner_w = QWidget()
        inner_w.setStyleSheet("background-color: transparent;")
        inner_layout = QVBoxLayout(inner_w)
        inner_layout.setContentsMargins(30, 30, 30, 30)

        top_layout = QHBoxLayout()
        btn_back = QPushButton("◀ 返回平台")
        btn_back.setFixedSize(160, 50)
        btn_back.setStyleSheet(
            "QPushButton { font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; background-color: #E8E4DA; border: 1px solid #D5D0C5; border-radius: 4px; font-size: 24px; font-weight: bold; color: #000000;} QPushButton:hover { background-color: #DCD7CB; }")
        btn_back.clicked.connect(self.back_callback)
        title = QLabel("作业登记系统")
        title.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 28px; font-weight: bold; color: #000000;")

        btn_import = QPushButton("导入名单")
        btn_history = QPushButton("查看历史记录")
        btn_export = QPushButton("数据导出")

        for btn in [btn_import, btn_history, btn_export]:
            btn.setFixedSize(160, 50)
            btn.setStyleSheet(btn_back.styleSheet())

        btn_import.clicked.connect(self.import_excel)
        btn_history.clicked.connect(self.open_history)
        btn_export.clicked.connect(self.open_export)

        top_layout.addWidget(btn_back)
        top_layout.addSpacing(20)
        top_layout.addWidget(title)
        top_layout.addStretch()
        top_layout.addWidget(btn_import)
        top_layout.addWidget(btn_history)
        top_layout.addWidget(btn_export)
        inner_layout.addLayout(top_layout)
        inner_layout.addSpacing(30)

        grid = QGridLayout()
        grid.setSpacing(20)
        for i, sub in enumerate(["语文", "数学", "英语", "物理", "化学", "生物", "政治", "历史", "地理", "技术"]):
            btn = QPushButton(sub)
            btn.setFixedHeight(80)
            btn.setStyleSheet(
                "QPushButton { background-color: #F8F6F1; color: #000000; border: 1px solid #D5D0C5; border-radius: 6px; font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 28px; font-weight: bold;} QPushButton:hover { background-color: #E8E4DA; }")
            btn.clicked.connect(lambda checked, s=sub: self.open_subject(s))
            grid.addWidget(btn, i // 5, i % 5)
        inner_layout.addLayout(grid)
        inner_layout.addSpacing(30)

        overview_title_layout = QHBoxLayout()
        overview_label = QLabel(f"📌 {datetime.now().strftime('%Y-%m-%d')} 今日作业上交情况：")
        overview_label.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 24px; font-weight: bold; color: #000000;")
        hint_label = QLabel("（点击下方卡片进行补交登记）")
        hint_label.setStyleSheet("font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 20px; color: #000000;")
        overview_title_layout.addWidget(overview_label)
        overview_title_layout.addWidget(hint_label)
        overview_title_layout.addStretch()

        self.overview_card = OverviewClickableCard()
        self.overview_card.nav_clicked.connect(self.open_late_submission)
        inner_layout.addLayout(overview_title_layout)
        inner_layout.addWidget(self.overview_card)
        inner_layout.addStretch()

        scroll.setWidget(inner_w)
        main_layout.addWidget(scroll)
        self.refresh_today_overview()

    def import_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择学生名单 Excel", "", "Excel Files (*.xlsx *.xls)")
        if file_path:
            res = DatabaseHelper.import_students(file_path)
            if isinstance(res, int):
                QMessageBox.information(self, "成功", f"成功导入 {res} 名学生！")
            else:
                QMessageBox.warning(self, "错误", f"导入失败: {res}")

    def open_subject(self, subject):
        today = datetime.now().strftime("%Y-%m-%d")
        SubjectDialog(today, subject, self).exec()
        self.refresh_today_overview()

    def open_late_submission(self):
        today = datetime.now().strftime("%Y-%m-%d")
        LateSubmissionDialog(today, self).exec()
        self.refresh_today_overview()

    def open_history(self):
        HomeworkHistoryDialog(self).exec()

    def open_export(self):
        HomeworkExportDialog(self).exec()

    def refresh_today_overview(self):
        today = datetime.now().strftime("%Y-%m-%d")
        conn = sqlite3.connect(DB_PATH)
        c = conn.cursor()
        c.execute("SELECT subject, student_name FROM homework WHERE date=? AND status='未交'", (today,))
        records = c.fetchall()
        conn.close()
        if not records:
            self.overview_card.text_label.setText("今日表现完美，暂无未交作业记录！")
            return
        summary = {}
        for sub, name in records:
            if sub not in summary: summary[sub] = []
            summary[sub].append(name)
        display_lines = [f"<b>{sub}</b>：{'、'.join(names)}" for sub, names in summary.items()]
        self.overview_card.text_label.setText("\n".join(display_lines))


# ==========================================
# 7. 课表与值日安排页面模块
# ==========================================
class SchedulePage(QWidget):
    def __init__(self, back_callback, parent=None):
        super().__init__(parent)
        self.back_callback = back_callback
        self.setStyleSheet("background-color: #F2EFE9;")
        self.init_ui()
        self.load_data()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none;")
        inner_w = QWidget()
        inner_w.setStyleSheet("background-color: transparent;")
        inner_layout = QVBoxLayout(inner_w)
        inner_layout.setContentsMargins(30, 30, 30, 30)

        top_layout = QHBoxLayout()
        btn_back = QPushButton("◀ 返回平台")
        btn_back.setFixedSize(160, 50)
        btn_back.setStyleSheet(
            "QPushButton { font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; background-color: #E8E4DA; border: 1px solid #D5D0C5; border-radius: 4px; font-size: 24px; font-weight: bold; color: #000000;} QPushButton:hover { background-color: #DCD7CB; }")
        btn_back.clicked.connect(self.back_callback)
        title = QLabel("课表全览")
        title.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 28px; font-weight: bold; color: #000000;")
        btn_imp = QPushButton("导入Excel")
        btn_imp.clicked.connect(self.imp)
        btn_res = QPushButton("重置数据")
        btn_res.clicked.connect(self.res)
        for b in [btn_imp, btn_res]:
            b.setFixedSize(160, 50)
            b.setStyleSheet(btn_back.styleSheet())
        top_layout.addWidget(btn_back)
        top_layout.addSpacing(20)
        top_layout.addWidget(title)
        top_layout.addStretch()
        top_layout.addWidget(btn_imp)
        top_layout.addWidget(btn_res)
        inner_layout.addLayout(top_layout)
        inner_layout.addSpacing(20)

        self.table = QTableWidget()
        self.table.setStyleSheet(
            "QTableWidget { font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; background-color: white; color: #000000;} QHeaderView::section { font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 24px; font-weight: bold; color: #000000; }")
        self.table.verticalHeader().setDefaultSectionSize(60)
        inner_layout.addWidget(self.table)

        scroll.setWidget(inner_w)
        main_layout.addWidget(scroll)

    def imp(self):
        f, _ = QFileDialog.getOpenFileName(self, "选择", "", "Excel (*.xlsx *.xls)")
        if f: DatabaseHelper.import_schedule(f); self.load_data()

    def res(self):
        DatabaseHelper.clear_schedule(); self.load_data()

    def load_data(self):
        df = DatabaseHelper.get_schedule_df()
        self.table.clear()
        self.table.setRowCount(0)
        if df.empty: return
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels([str(c) for c in df.columns])
        t_idx, _ = DatabaseHelper.get_today_col(df)
        inserted = False
        for r in range(len(df)):
            p = str(df.iloc[r, 0])
            if not inserted and ("6" in p or "六" in p):
                rc = self.table.rowCount()
                self.table.insertRow(rc)
                self.table.setRowHeight(rc, 8)
                inserted = True
            rc = self.table.rowCount()
            self.table.insertRow(rc)
            for c in range(len(df.columns)):
                item = QTableWidgetItem(str(df.iloc[r, c]))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if c == t_idx: item.setBackground(QColor("#DCD7CB"))
                self.table.setItem(rc, c, item)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)


class DutyPage(QWidget):
    def __init__(self, back_callback, parent=None):
        super().__init__(parent)
        self.back_callback = back_callback
        self.setStyleSheet("background-color: #F2EFE9;")
        self.init_ui()
        self.load_data()

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: none;")
        inner_w = QWidget()
        inner_w.setStyleSheet("background-color: transparent;")
        inner_layout = QVBoxLayout(inner_w)
        inner_layout.setContentsMargins(30, 30, 30, 30)

        top_layout = QHBoxLayout()
        btn_back = QPushButton("◀ 返回平台")
        btn_back.setFixedSize(160, 50)
        btn_back.setStyleSheet(
            "QPushButton { font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; background-color: #E8E4DA; border: 1px solid #D5D0C5; border-radius: 4px; font-size: 24px; font-weight: bold; color: #000000;} QPushButton:hover { background-color: #DCD7CB; }")
        btn_back.clicked.connect(self.back_callback)
        title = QLabel("值日安排全览")
        title.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 28px; font-weight: bold; color: #000000;")
        btn_imp = QPushButton("导入Excel")
        btn_imp.clicked.connect(self.imp)
        btn_res = QPushButton("重置数据")
        btn_res.clicked.connect(self.res)
        for b in [btn_imp, btn_res]:
            b.setFixedSize(160, 50)
            b.setStyleSheet(btn_back.styleSheet())
        top_layout.addWidget(btn_back)
        top_layout.addSpacing(20)
        top_layout.addWidget(title)
        top_layout.addStretch()
        top_layout.addWidget(btn_imp)
        top_layout.addWidget(btn_res)
        inner_layout.addLayout(top_layout)
        inner_layout.addSpacing(20)

        self.table = QTableWidget()
        self.table.setStyleSheet(
            "QTableWidget { font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; background-color: white; color: #000000;} QHeaderView::section { font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 24px; font-weight: bold; color: #000000; }")
        self.table.verticalHeader().setDefaultSectionSize(60)
        inner_layout.addWidget(self.table)

        scroll.setWidget(inner_w)
        main_layout.addWidget(scroll)

    def imp(self):
        f, _ = QFileDialog.getOpenFileName(self, "选择", "", "Excel (*.xlsx *.xls)")
        if f: DatabaseHelper.import_duty(f); self.load_data()

    def res(self):
        DatabaseHelper.clear_duty(); self.load_data()

    def load_data(self):
        df = DatabaseHelper.get_duty_df()
        self.table.clear()
        self.table.setRowCount(0)
        if df.empty: return
        self.table.setRowCount(len(df))
        self.table.setColumnCount(len(df.columns))
        self.table.setHorizontalHeaderLabels([str(c) for c in df.columns])
        t_idx, _ = DatabaseHelper.get_today_col(df)
        for r in range(len(df)):
            for c in range(len(df.columns)):
                item = QTableWidgetItem(str(df.iloc[r, c]))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                if c == t_idx: item.setBackground(QColor("#DCD7CB"))
                self.table.setItem(r, c, item)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)


# ==========================================
# 8. 主窗口及路由逻辑 (完全居中自适应布局)
# ==========================================
class ClassDashboard(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("班级事务管理平台")
        self.resize(1150, 800)
        self.setStyleSheet("background-color: #F2EFE9;")
        DatabaseHelper.init_db()
        self.stacked_widget = QStackedWidget()
        self.setCentralWidget(self.stacked_widget)
        self.init_home_page()
        self.init_sub_pages()

    def init_home_page(self):
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setStyleSheet("QScrollArea { border: none; background-color: #F2EFE9; }")

        home_widget = QWidget()
        home_widget.setStyleSheet("background-color: #F2EFE9;")
        main_layout = QVBoxLayout(home_widget)
        main_layout.setContentsMargins(30, 25, 30, 30)

        header_layout = QHBoxLayout()
        title_label = QLabel("班级事务管理平台")
        title_label.setStyleSheet(
            "font-family: 'SimSun', '宋体', 'Songti SC', 'STSong', serif; font-size: 32px; font-weight: bold; color: #000000; letter-spacing: 2px;")
        self.time_label = QLabel()
        self.time_label.setStyleSheet(
            "font-family: 'KaiTi', '楷体', 'Kaiti SC', 'STKaiti', serif; font-size: 24px; font-weight: bold; color: #000000;")
        self.time_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignBottom)
        header_layout.addWidget(title_label)
        header_layout.addWidget(self.time_label)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_time)
        self.timer.start(1000)
        self.update_time()

        main_layout.addLayout(header_layout)
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        main_layout.addWidget(line)

        content_layout = QHBoxLayout()
        left_layout = QVBoxLayout()
        self.schedule_card = ScheduleCard(target_index=4)
        self.duty_card = DutyCard(target_index=5)
        self.schedule_card.nav_clicked.connect(self.switch_page)
        self.duty_card.nav_clicked.connect(self.switch_page)
        left_layout.addWidget(self.schedule_card)
        left_layout.addWidget(self.duty_card)

        right_layout = QVBoxLayout()
        self.notice_card = NoticeDashboardCard(target_index=1)
        self.deduction_card = DeductionDashboardCard(target_index=2)
        self.homework_card = HomeworkDashboardCard(target_index=3)
        self.notice_card.nav_clicked.connect(self.switch_page)
        self.deduction_card.nav_clicked.connect(self.switch_page)
        self.homework_card.nav_clicked.connect(self.switch_page)
        right_layout.addWidget(self.notice_card)
        right_layout.addWidget(self.deduction_card)
        right_layout.addWidget(self.homework_card)

        content_layout.addLayout(left_layout, 35)
        content_layout.addLayout(right_layout, 65)
        main_layout.addLayout(content_layout)

        scroll_area.setWidget(home_widget)
        self.stacked_widget.addWidget(scroll_area)

    def init_sub_pages(self):
        for index in range(1, 6):
            if index == 1:
                self.notice_page = NoticePage(back_callback=lambda: self.switch_page(0))
                self.stacked_widget.addWidget(self.notice_page)
            elif index == 2:
                self.deduction_page = DeductionPage(back_callback=lambda: self.switch_page(0))
                self.stacked_widget.addWidget(self.deduction_page)
            elif index == 3:
                self.homework_page = HomeworkPage(back_callback=lambda: self.switch_page(0))
                self.stacked_widget.addWidget(self.homework_page)
            elif index == 4:
                self.schedule_page = SchedulePage(back_callback=lambda: self.switch_page(0))
                self.stacked_widget.addWidget(self.schedule_page)
            elif index == 5:
                self.duty_page = DutyPage(back_callback=lambda: self.switch_page(0))
                self.stacked_widget.addWidget(self.duty_page)

    def update_time(self):
        self.time_label.setText(QDateTime.currentDateTime().toString("yyyy年MM月dd日   HH:mm   dddd"))

    def switch_page(self, index):
        self.stacked_widget.setCurrentIndex(index)
        if index == 0:
            self.homework_card.refresh_data()
            self.duty_card.refresh_data()
            self.schedule_card.refresh_data()
            self.notice_card.refresh_data()
            self.deduction_card.refresh_data()
        elif index == 2:
            self.deduction_page.refresh_data()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ClassDashboard()
    window.show()
    sys.exit(app.exec())