import tkinter as tk
from email_window import EmailWindow

class MainWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("邮件提取工具")
        self.root.geometry("400x300")

        self._setup_ui()

    def _setup_ui(self):
        # 居中框架
        frame = tk.Frame(self.root)
        frame.place(relx=0.5, rely=0.5, anchor="center")

        # 邮件提取按钮
        btn_extract = tk.Button(
            frame,
            text="邮件提取",
            width=15,
            height=2,
            command=self._open_email_window
        )
        btn_extract.pack(pady=10)

        # 待规划按钮（禁用）
        btn_plan = tk.Button(
            frame,
            text="待规划",
            width=15,
            height=2,
            state="disabled"  # 禁用
        )
        btn_plan.pack(pady=10)

    def _open_email_window(self):
        EmailWindow(self.root)

    def run(self):
        self.root.mainloop()

if __name__ == "__main__":
    app = MainWindow()
    app.run()
