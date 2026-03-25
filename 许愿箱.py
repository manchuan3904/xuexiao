import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
from datetime import datetime
import openpyxl

# 本地数据存储文件
DATA_FILE = "wishes_data.json"


class WishingBoxApp:
    def __init__(self, root):
        self.root = root
        self.root.title("许愿箱")
        self.root.geometry("700x500")  # 稍微加宽一点以容纳更多按钮

        # 低饱和度配色方案
        self.colors = {
            "bg": "#F4F4F2",  # 背景：米白
            "primary": "#A3B18A",  # 主色：低饱和灰绿
            "secondary": "#588157",  # 强调色：深灰绿
            "text": "#333333",  # 文本：深灰
            "list_bg": "#E8E8E4",  # 列表背景：浅灰
            "danger": "#D4A373"  # 废除按钮色：低饱和土橙色
        }
        self.root.configure(bg=self.colors["bg"])
        self.data = self.load_data()

        # 兼容旧数据，增加 completed 和 trashed 状态
        for wish in self.data:
            if "completed" not in wish:
                wish["completed"] = False
            if "trashed" not in wish:
                wish["trashed"] = False
        self.save_data()

        self.setup_ui()

    def load_data(self):
        """调用储存在本地电脑的数据"""
        if os.path.exists(DATA_FILE):
            try:
                with open(DATA_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                return []
        return []

    def save_data(self):
        """每日/实时自动储存数据到程序所在文件夹"""
        with open(DATA_FILE, 'w', encoding='utf-8') as f:
            json.dump(self.data, f, ensure_ascii=False, indent=4)

    def setup_ui(self):
        # 样式设置
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TButton", padding=6, relief="flat", background=self.colors["primary"], foreground="white")
        style.map("TButton", background=[('active', self.colors["secondary"])])

        # 针对废除按钮的特殊样式
        style.configure("Danger.TButton", padding=6, relief="flat", background=self.colors["danger"],
                        foreground="white")
        style.map("Danger.TButton", background=[('active', "#CC8B50")])

        style.configure("Treeview", background=self.colors["list_bg"], fieldbackground=self.colors["list_bg"],
                        foreground=self.colors["text"], rowheight=30)
        style.configure("Treeview.Heading", background=self.colors["primary"], foreground="white",
                        font=('Arial', 10, 'bold'))

        # 顶部标题与按钮区
        top_frame = tk.Frame(self.root, bg=self.colors["bg"])
        top_frame.pack(fill=tk.X, padx=20, pady=15)

        title_label = tk.Label(top_frame, text="✨ 许愿箱 ✨", font=("Arial", 20, "bold"), bg=self.colors["bg"],
                               fg=self.colors["secondary"])
        title_label.pack(side=tk.LEFT)

        btn_frame = tk.Frame(top_frame, bg=self.colors["bg"])
        btn_frame.pack(side=tk.RIGHT)

        ttk.Button(btn_frame, text="许愿", command=self.open_wish_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="已实现愿望", command=self.open_completed_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="回收站", command=self.open_recycle_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="数据汇总(导出)", command=self.export_data).pack(side=tk.LEFT, padx=5)

        # 主界面愿望列表
        list_frame = tk.Frame(self.root, bg=self.colors["bg"])
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        columns = ("title", "time", "heat")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse")
        self.tree.heading("title", text="愿望标题")
        self.tree.heading("time", text="发布时间")
        self.tree.heading("heat", text="🔥 热度 (评论数)")

        self.tree.column("title", width=300, anchor=tk.W)
        self.tree.column("time", width=150, anchor=tk.CENTER)
        self.tree.column("heat", width=100, anchor=tk.CENTER)

        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # 双击查看详情 (传入主树状图对象及列表类型)
        self.tree.bind("<Double-1>", lambda event: self.open_detail_dialog(self.tree, "main"))

        self.refresh_tree(self.tree, "main")

    def refresh_tree(self, tree_widget, list_type):
        """通用刷新列表方法，根据不同的列表类型过滤数据"""
        for item in tree_widget.get_children():
            tree_widget.delete(item)

        for index, wish in enumerate(self.data):
            is_completed = wish.get("completed", False)
            is_trashed = wish.get("trashed", False)

            show_item = False
            if list_type == "main" and not is_completed and not is_trashed:
                show_item = True
            elif list_type == "completed" and is_completed and not is_trashed:
                show_item = True
            elif list_type == "trashed" and is_trashed:
                show_item = True

            if show_item:
                heat = len(wish.get("comments", []))
                tree_widget.insert("", tk.END, iid=index, values=(wish["title"], wish["time"], f"热度: {heat}"))

    def open_wish_dialog(self):
        """许愿界面"""
        dialog = tk.Toplevel(self.root)
        dialog.title("写下你的愿望")
        dialog.geometry("400x350")
        dialog.configure(bg=self.colors["bg"])

        tk.Label(dialog, text="愿望标题 (必填):", bg=self.colors["bg"], font=("Arial", 10)).pack(anchor=tk.W, padx=20,
                                                                                                 pady=(15, 5))
        title_entry = ttk.Entry(dialog, width=45)
        title_entry.pack(padx=20)

        tk.Label(dialog, text="愿望内容 (必填):", bg=self.colors["bg"], font=("Arial", 10)).pack(anchor=tk.W, padx=20,
                                                                                                 pady=(15, 5))
        content_text = tk.Text(dialog, width=45, height=8, font=("Arial", 10))
        content_text.pack(padx=20)

        def save_wish():
            title = title_entry.get().strip()
            content = content_text.get("1.0", tk.END).strip()
            if not title or not content:
                messagebox.showwarning("提示", "标题和内容均为必填项哦！", parent=dialog)
                return

            new_wish = {
                "title": title,
                "content": content,
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "comments": [],
                "completed": False,
                "trashed": False
            }
            self.data.append(new_wish)
            self.save_data()
            self.refresh_tree(self.tree, "main")
            dialog.destroy()

        ttk.Button(dialog, text="保存愿望", command=save_wish).pack(pady=20)

    def open_detail_dialog(self, tree_widget, list_type):
        """详情与评论区界面"""
        selected_item = tree_widget.selection()
        if not selected_item:
            return

        wish_index = int(selected_item[0])
        wish = self.data[wish_index]

        dialog = tk.Toplevel(self.root)
        dialog.title(f"愿望详情: {wish['title']}")
        dialog.geometry("500x650")  # 加高一点容纳新按钮
        dialog.configure(bg=self.colors["bg"])

        # 愿望内容展示
        tk.Label(dialog, text=wish["title"], font=("Arial", 14, "bold"), bg=self.colors["bg"],
                 fg=self.colors["secondary"]).pack(pady=10)

        content_frame = tk.Frame(dialog, bg="white", bd=1, relief="solid")
        content_frame.pack(fill=tk.X, padx=20, pady=5)
        tk.Label(content_frame, text=wish["content"], font=("Arial", 11), bg="white", justify=tk.LEFT,
                 wraplength=450).pack(padx=10, pady=10, anchor=tk.W)

        # 状态操作区
        action_frame = tk.Frame(dialog, bg=self.colors["bg"])
        action_frame.pack(pady=10)

        is_trashed = wish.get("trashed", False)
        is_completed = wish.get("completed", False)

        if is_trashed:
            # 如果在回收站中，显示复活按钮
            tk.Label(action_frame, text="⚠️ 该愿望已被废除", font=("Arial", 10, "bold"), fg=self.colors["danger"],
                     bg=self.colors["bg"]).pack(pady=5)

            def revive_wish():
                if messagebox.askyesno("确认", "确定要复活这个愿望吗？\n它将重新回到主界面的未实现列表中。",
                                       parent=dialog):
                    wish["trashed"] = False
                    wish["completed"] = False  # 强制回到主界面
                    self.save_data()
                    self.refresh_tree(self.tree, "main")  # 刷新主界面
                    self.refresh_tree(tree_widget, list_type)  # 刷新弹窗背后的回收站列表
                    dialog.destroy()

            ttk.Button(action_frame, text="♻️ 复活并返回主界面", command=revive_wish).pack()

        else:
            # 如果不在回收站中，显示完成和废除按钮
            if not is_completed:
                def mark_completed():
                    if messagebox.askyesno("确认", "确定将这个愿望标记为已实现吗？\n它将移至“已实现愿望”列表中。",
                                           parent=dialog):
                        wish["completed"] = True
                        self.save_data()
                        self.refresh_tree(self.tree, "main")
                        self.refresh_tree(tree_widget, list_type)
                        dialog.destroy()

                ttk.Button(action_frame, text="🎉 标记为已实现", command=mark_completed).pack(side=tk.LEFT, padx=10)
            else:
                tk.Label(action_frame, text="✅ 该愿望已实现", font=("Arial", 10, "bold"), fg=self.colors["secondary"],
                         bg=self.colors["bg"]).pack(side=tk.LEFT, padx=10)

            # 废除按钮
            def trash_wish():
                if messagebox.askyesno("确认", "确定将这个愿望废除吗？\n它将移入回收站。", parent=dialog):
                    wish["trashed"] = True
                    self.save_data()
                    self.refresh_tree(self.tree, "main")
                    self.refresh_tree(tree_widget, list_type)
                    dialog.destroy()

            ttk.Button(action_frame, text="🗑️ 废除 (移入回收站)", style="Danger.TButton", command=trash_wish).pack(
                side=tk.LEFT, padx=10)

        # 评论区展示
        tk.Label(dialog, text="👇 评论区", bg=self.colors["bg"], font=("Arial", 10, "bold")).pack(anchor=tk.W, padx=20,
                                                                                                 pady=(10, 5))

        comment_listbox = tk.Listbox(dialog, height=8, bg=self.colors["list_bg"], relief="flat", font=("Arial", 10))
        comment_listbox.pack(fill=tk.BOTH, expand=True, padx=20)

        for c in wish.get("comments", []):
            comment_listbox.insert(tk.END, f"• {c}")

        # 发表评论
        input_frame = tk.Frame(dialog, bg=self.colors["bg"])
        input_frame.pack(fill=tk.X, padx=20, pady=15)

        comment_entry = ttk.Entry(input_frame, width=40)
        comment_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 10))

        def add_comment():
            new_comment = comment_entry.get().strip()
            if new_comment:
                wish.setdefault("comments", []).append(new_comment)
                self.save_data()
                self.refresh_tree(self.tree, "main")
                self.refresh_tree(tree_widget, list_type)  # 保持列表热度刷新
                comment_listbox.insert(tk.END, f"• {new_comment}")
                comment_entry.delete(0, tk.END)

        ttk.Button(input_frame, text="发表评论", command=add_comment).pack(side=tk.RIGHT)

    def open_completed_dialog(self):
        """已实现的愿望列表面板"""
        self._create_sub_window("🎉 已实现愿望列表 🎉", "completed")

    def open_recycle_dialog(self):
        """回收站面板"""
        self._create_sub_window("🗑️ 回收站 (已废除的愿望) 🗑️", "trashed")

    def _create_sub_window(self, title_text, list_type):
        """用于生成次级列表窗口的内部辅助方法"""
        win = tk.Toplevel(self.root)
        win.title(title_text)
        win.geometry("550x400")
        win.configure(bg=self.colors["bg"])

        tk.Label(win, text=title_text, font=("Arial", 16, "bold"), bg=self.colors["bg"],
                 fg=self.colors["secondary"]).pack(pady=15)

        list_frame = tk.Frame(win, bg=self.colors["bg"])
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        columns = ("title", "time", "heat")
        sub_tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="browse")
        sub_tree.heading("title", text="愿望标题")
        sub_tree.heading("time", text="发布时间")
        sub_tree.heading("heat", text="🔥 热度 (评论数)")

        sub_tree.column("title", width=250, anchor=tk.W)
        sub_tree.column("time", width=150, anchor=tk.CENTER)
        sub_tree.column("heat", width=100, anchor=tk.CENTER)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=sub_tree.yview)
        sub_tree.configure(yscroll=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        sub_tree.pack(fill=tk.BOTH, expand=True)

        self.refresh_tree(sub_tree, list_type)
        sub_tree.bind("<Double-1>", lambda event: self.open_detail_dialog(sub_tree, list_type))

    def export_data(self):
        """导出/数据汇总功能"""
        if not self.data:
            messagebox.showinfo("提示", "当前没有可以导出的数据。")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        export_filename = f"愿望数据汇总_{timestamp}.xlsx"

        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "愿望汇总"

            headers = ["状态", "愿望标题", "发布时间", "热度(评论数)", "愿望内容", "评论详情"]
            ws.append(headers)

            for wish in self.data:
                # 判定当前状态
                if wish.get("trashed", False):
                    status_text = "已废除"
                elif wish.get("completed", False):
                    status_text = "已实现"
                else:
                    status_text = "未实现"

                comments = " | ".join(wish.get("comments", []))
                heat = len(wish.get("comments", []))
                ws.append([status_text, wish["title"], wish["time"], heat, wish["content"], comments])

            wb.save(export_filename)
            messagebox.showinfo("导出成功", f"数据已成功导出为Excel表格：\n{export_filename}\n保存在程序所在文件夹内。")
        except Exception as e:
            messagebox.showerror("导出失败", f"发生错误：{e}\n请确保已安装 openpyxl 库。")


if __name__ == "__main__":
    root = tk.Tk()
    app = WishingBoxApp(root)
    root.mainloop()