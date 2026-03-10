import tkinter as tk
from tkinter import ttk, messagebox
import asyncio
from datetime import datetime
from outlook_client import OutlookClient
from api_client import APIClient

class EmailWindow:
    def __init__(self, parent):
        self.parent = parent
        self.window = tk.Toplevel(parent)
        self.window.title("邮件提取")
        self.window.geometry("1000x700")

        # 设置为子窗口，不抢焦点但保持置顶
        self.window.transient(parent)
        self.window.grab_set()

        self.outlook_client = None
        self.api_client = APIClient()

        self.folders = []
        self.emails = []
        self.selected_emails = []
        self.task_ids = []

        self._setup_ui()
        self._init_outlook()

    def _setup_ui(self):
        # 筛选条件区域
        filter_frame = tk.LabelFrame(self.window, text="筛选条件", padx=10, pady=10)
        filter_frame.pack(fill="x", padx=10, pady=5)

        # 日期范围
        tk.Label(filter_frame, text="起始日期:").grid(row=0, column=0, sticky="w")
        self.start_date_var = tk.StringVar(value="2021-01-01")
        tk.Entry(filter_frame, textvariable=self.start_date_var, width=15).grid(row=0, column=1, padx=5)

        tk.Label(filter_frame, text="截止日期:").grid(row=0, column=2, sticky="w")
        self.end_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        tk.Entry(filter_frame, textvariable=self.end_date_var, width=15).grid(row=0, column=3, padx=5)

        # 文件夹选择
        tk.Label(filter_frame, text="文件夹:").grid(row=1, column=0, sticky="w", pady=5)
        self.folder_var = tk.StringVar()
        self.folder_combo = ttk.Combobox(filter_frame, textvariable=self.folder_var, width=30)
        self.folder_combo.grid(row=1, column=1, columnspan=2, padx=5, pady=5)

        # 主题搜索
        tk.Label(filter_frame, text="主题搜索:").grid(row=2, column=0, sticky="w")
        self.subject_var = tk.StringVar()
        tk.Entry(filter_frame, textvariable=self.subject_var, width=30).grid(row=2, column=1, columnspan=2, padx=5)

        # 按钮
        tk.Button(filter_frame, text="刷新列表", command=self._load_emails).grid(row=3, column=0, pady=10)
        tk.Button(filter_frame, text="全选", command=self._select_all).grid(row=3, column=1, pady=10, padx=5)
        tk.Button(filter_frame, text="取消全选", command=self._deselect_all).grid(row=3, column=2, pady=10)
        tk.Button(filter_frame, text="提取选中", command=self._extract_selected).grid(row=3, column=3, pady=10, padx=5)

        # 邮件列表
        list_frame = tk.Frame(self.window)
        list_frame.pack(fill="both", expand=True, padx=10, pady=5)

        columns = ("select", "subject", "sender", "received_time", "conversation_topic")
        self.tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=15, selectmode="extended")

        self.tree.heading("select", text="选择")
        self.tree.heading("subject", text="主题")
        self.tree.heading("sender", text="发件人")
        self.tree.heading("received_time", text="收件时间")
        self.tree.heading("conversation_topic", text="会话主题")

        self.tree.column("select", width=50)
        self.tree.column("subject", width=300)
        self.tree.column("sender", width=150)
        self.tree.column("received_time", width=150)
        self.tree.column("conversation_topic", width=200)

        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 绑定点击任意列都能选中整行
        self.tree.bind('<Button-1>', self._on_tree_click)

        # 绑定选择事件
        self.tree.bind('<<TreeviewSelect>>', self._on_tree_select)

        # 任务状态区域
        status_frame = tk.LabelFrame(self.window, text="任务状态", padx=10, pady=10)
        status_frame.pack(fill="x", padx=10, pady=5)

        self.task_label = tk.Label(status_frame, text="无任务", justify="left", anchor="w")
        self.task_label.pack(side="left")

        tk.Button(status_frame, text="刷新状态", command=self._refresh_status).pack(side="right")

    def _init_outlook(self):
        """初始化Outlook连接"""
        try:
            self.outlook_client = OutlookClient()
            self._load_folders()
        except Exception as e:
            messagebox.showerror("错误", f"无法连接Outlook: {e}\n请确保Outlook已启动")

    def _load_folders(self):
        """加载文件夹列表"""
        try:
            self.folders = self.outlook_client.get_folders()
            self.folder_combo['values'] = [f["name"] for f in self.folders]
            if self.folders:
                self.folder_var.set(self.folders[0]["name"])  # 默认第一个
        except Exception as e:
            messagebox.showerror("错误", f"加载文件夹失败: {e}")

    def _load_emails(self):
        """加载邮件列表"""
        if not self.outlook_client:
            messagebox.showerror("错误", "Outlook未连接")
            return

        try:
            # 解析日期
            start_date = datetime.strptime(self.start_date_var.get(), "%Y-%m-%d")
            end_date = datetime.strptime(self.end_date_var.get(), "%Y-%m-%d")
            end_date = end_date.replace(hour=23, minute=59, second=59)

            # 获取文件夹（默认收件箱）
            folder = self.outlook_client.get_default_folder(6)

            # 获取邮件
            self.emails = self.outlook_client.get_emails(
                folder,
                start_date=start_date,
                end_date=end_date,
                subject_filter=self.subject_var.get()
            )

            # 去重
            self.emails = self.outlook_client.deduplicate_by_conversation(self.emails)

            # 显示
            self._display_emails()

            # 先显示消息，再恢复窗口焦点
            self.window.after(100, lambda: self.window.lift())
            self.window.after(200, lambda: self.window.focus_force())

            messagebox.showinfo("成功", f"加载了 {len(self.emails)} 封邮件（去重后）")

            # 消息框关闭后再确保窗口在最前
            self.window.after(300, lambda: self.window.lift())

        except Exception as e:
            messagebox.showerror("错误", f"加载邮件失败: {e}")

    def _display_emails(self):
        """显示邮件列表"""
        # 清空
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 插入数据
        for email in self.emails:
            received = ""
            if email["received_time"]:
                received = email["received_time"].strftime("%Y-%m-%d %H:%M")

            subject = email["subject"]
            if len(subject) > 50:
                subject = subject[:50] + "..."

            topic = email["conversation_topic"]
            if topic and len(topic) > 30:
                topic = topic[:30] + "..."

            self.tree.insert("", "end", values=(
                "☐",
                subject,
                email["sender"],
                received,
                topic or ""
            ), tags=(email["entry_id"],))

    def _on_tree_click(self, event):
        """点击任意列都选中整行"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell" or region == "row":
            item = self.tree.identify_row(event.y)
            if item:
                current = self.tree.selection()
                if item in current:
                    self.tree.selection_remove(item)
                else:
                    self.tree.selection_add(item)

    def _on_tree_select(self, event):
        """处理选择事件，更新复选框显示"""
        # 更新选择复选框显示
        for item in self.tree.get_children():
            if item in self.tree.selection():
                self.tree.set(item, "select", "☑")
            else:
                self.tree.set(item, "select", "☐")

    def _select_all(self):
        """全选"""
        for item in self.tree.get_children():
            self.tree.selection_add(item)
            self.tree.set(item, "select", "☑")

    def _deselect_all(self):
        """取消全选"""
        for item in self.tree.get_children():
            self.tree.selection_remove(item)
            self.tree.set(item, "select", "☐")

    def _extract_selected(self):
        """提取选中的邮件"""
        # 获取选中项
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择邮件")
            return

        # 获取选中的邮件数据
        selected_emails = []
        for item in selected_items:
            entry_id = self.tree.item(item)["tags"][0]
            for email in self.emails:
                if email["entry_id"] == entry_id:
                    received_str = ""
                    if email["received_time"]:
                        received_str = email["received_time"].strftime("%Y-%m-%d %H:%M:%S")

                    selected_emails.append({
                        "subject": email["subject"],
                        "sender": email["sender"],
                        "received_time": received_str,
                        "conversation_topic": email["conversation_topic"],
                        "html_content": email["html_content"],
                        "images": email["images"]
                    })
                    break

        if not selected_emails:
            messagebox.showwarning("警告", "未找到选中邮件的数据")
            return

        # 构建请求数据
        json_data = {
            "emails": selected_emails,
            "options": {"desensitize": True}
        }

        try:
            # 调用API
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            result = loop.run_until_complete(self.api_client.extract_emails(json_data))

            self.task_ids = result.get("task_ids", [])
            self.task_label.config(text=f"任务已提交: {', '.join(self.task_ids)}")
            messagebox.showinfo("成功", f"已提交 {len(self.task_ids)} 个任务")

        except Exception as e:
            messagebox.showerror("错误", f"提交失败: {e}")

    def _refresh_status(self):
        """刷新任务状态"""
        if not self.task_ids:
            messagebox.showwarning("警告", "没有任务")
            return

        try:
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)

            status_text = []
            for task_id in self.task_ids:
                result = loop.run_until_complete(self.api_client.get_task_status(task_id))
                status = result.get("status", "unknown")
                progress = result.get("progress", 0)
                message = result.get("message", "")
                status_text.append(f"{task_id}: {status} ({progress}%) - {message}")

            self.task_label.config(text="\n".join(status_text))

        except Exception as e:
            messagebox.showerror("错误", f"查询失败: {e}")
