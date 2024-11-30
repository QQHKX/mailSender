import os
import resend
import sqlite3
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

DEFAULT_DB_FILE = "email_history.db"
DEFAULT_EMAIL_FILE = "email_info.txt"
PROGRAM_VERSION = "1.0.0"


class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("邮件发送程序")

        self.api_key = tk.StringVar(value="re_L42m8BBS_AVmXaifM8HzTR8psWfHDNhre")
        self.db_file = tk.StringVar(value=DEFAULT_DB_FILE)
        self.email_file = DEFAULT_EMAIL_FILE
        self.details_window_open = False

        self.setup_ui()

    def setup_ui(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(expand=1, fill="both")

        # 创建邮件发送选项卡
        send_frame = ttk.Frame(notebook)
        notebook.add(send_frame, text="发送邮件")

        # 发件邮箱
        ttk.Label(send_frame, text="发件邮箱:").grid(row=0, column=0, padx=10, pady=5)
        self.from_entry = ttk.Entry(send_frame, width=50)
        self.from_entry.grid(row=0, column=1, padx=10, pady=5)
        self.from_entry.insert(0, self.load_email_info())

        # 收件邮箱
        ttk.Label(send_frame, text="收件邮箱:").grid(row=1, column=0, padx=10, pady=5)
        self.to_entry = ttk.Entry(send_frame, width=50)
        self.to_entry.grid(row=1, column=1, padx=10, pady=5)

        # 邮件主题
        ttk.Label(send_frame, text="邮件主题:").grid(row=2, column=0, padx=10, pady=5)
        self.subject_entry = ttk.Entry(send_frame, width=50)
        self.subject_entry.grid(row=2, column=1, padx=10, pady=5)

        # 邮件内容
        ttk.Label(send_frame, text="邮件内容:").grid(row=3, column=0, padx=10, pady=5)
        self.content_text = tk.Text(send_frame, width=50, height=10)
        self.content_text.grid(row=3, column=1, padx=10, pady=5)

        # 附件
        ttk.Label(send_frame, text="附件:").grid(row=4, column=0, padx=10, pady=5)
        self.attachment_var = tk.StringVar()
        self.attachment_entry = ttk.Entry(send_frame, textvariable=self.attachment_var, width=50)
        self.attachment_entry.grid(row=4, column=1, padx=10, pady=5)
        ttk.Button(send_frame, text="选择文件", command=self.select_attachment).grid(row=4, column=2, padx=10, pady=5)

        # 发送按钮
        ttk.Button(send_frame, text="发送邮件", command=self.send_email).grid(row=5, column=1, padx=10, pady=20)

        # 历史记录
        history_frame = ttk.Frame(send_frame)
        history_frame.grid(row=6, column=0, columnspan=3, padx=10, pady=10)
        ttk.Label(history_frame, text="发件历史记录:").pack(anchor="w")

        columns = ("from_email", "to_email", "subject", "content", "attachment", "send_time")
        self.history_tree = ttk.Treeview(history_frame, columns=columns, show="headings")
        for col in columns:
            self.history_tree.heading(col, text=col)
            self.history_tree.column(col, width=100)

        self.history_tree.pack(fill=tk.BOTH, expand=True)
        self.history_tree.bind("<Double-1>", self.show_email_details)
        self.history_tree.bind("<Button-3>", self.show_context_menu)

        self.init_db()
        self.update_history_list()

        # 创建设置选项卡
        settings_frame = ttk.Frame(notebook)
        notebook.add(settings_frame, text="设置")

        # API Key
        ttk.Label(settings_frame, text="API Key:").grid(row=0, column=0, padx=10, pady=5)
        self.api_key_entry = ttk.Entry(settings_frame, textvariable=self.api_key, width=50)
        self.api_key_entry.grid(row=0, column=1, padx=10, pady=5)

        # 数据库存储位置
        ttk.Label(settings_frame, text="数据库存储位置:").grid(row=1, column=0, padx=10, pady=5)
        self.db_file_entry = ttk.Entry(settings_frame, textvariable=self.db_file, width=50)
        self.db_file_entry.grid(row=1, column=1, padx=10, pady=5)
        ttk.Button(settings_frame, text="选择文件", command=self.select_db_file).grid(row=1, column=2, padx=10, pady=5)

        # 保存设置按钮
        ttk.Button(settings_frame, text="保存设置", command=self.save_settings).grid(row=2, column=1, padx=10, pady=20)

        # 创建关于选项卡
        about_frame = ttk.Frame(notebook)
        notebook.add(about_frame, text="关于")

        about_text = f"邮件发送程序\n版本: {PROGRAM_VERSION}\n作者: Your Name"
        ttk.Label(about_frame, text=about_text).pack(padx=10, pady=10)

    def init_db(self):
        conn = sqlite3.connect(self.db_file.get())
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS email_history
                          (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                           from_email TEXT, 
                           to_email TEXT, 
                           subject TEXT, 
                           content TEXT, 
                           attachment TEXT,
                           send_time TEXT)''')
        conn.commit()
        # 检查并添加 send_time 列
        cursor.execute("PRAGMA table_info(email_history)")
        columns = [info[1] for info in cursor.fetchall()]
        if "send_time" not in columns:
            cursor.execute("ALTER TABLE email_history ADD COLUMN send_time TEXT")
        conn.commit()
        conn.close()

    def save_email_info(self, from_email):
        with open(self.email_file, "w") as file:
            file.write(from_email)

    def load_email_info(self):
        if os.path.exists(self.email_file):
            with open(self.email_file, "r") as file:
                return file.read().strip()
        return ""

    def save_email_history(self, from_email, to_email, subject, content, attachment):
        send_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn = sqlite3.connect(self.db_file.get())
        cursor = conn.cursor()
        cursor.execute('''INSERT INTO email_history 
                          (from_email, to_email, subject, content, attachment, send_time) 
                          VALUES (?, ?, ?, ?, ?, ?)''',
                       (from_email, to_email, subject, content, attachment, send_time))
        conn.commit()
        conn.close()

    def load_email_history(self):
        conn = sqlite3.connect(self.db_file.get())
        cursor = conn.cursor()
        cursor.execute('SELECT id, from_email, to_email, subject, content, attachment, send_time FROM email_history')
        rows = cursor.fetchall()
        conn.close()
        return rows

    def send_email(self):
        resend.api_key = self.api_key.get()

        from_email = self.from_entry.get()
        to_email = self.to_entry.get()
        subject = self.subject_entry.get()
        content = self.content_text.get("1.0", tk.END)
        attachment_path = self.attachment_var.get()

        if not from_email or not to_email or not subject or not content:
            messagebox.showwarning("输入错误", "请填写所有字段")
            return

        if attachment_path:
            with open(attachment_path, "rb") as f:
                file_content = f.read()
            attachment = {"filename": os.path.basename(attachment_path), "content": list(file_content)}
            attachments = [attachment]
        else:
            attachments = []

        params = {
            "from": from_email,
            "to": [to_email],
            "subject": subject,
            "html": content,
            "headers": {"X-Entity-Ref-ID": "123456789"},
            "attachments": attachments,
        }

        try:
            email = resend.Emails.send(params)
            messagebox.showinfo("发送成功", "邮件已成功发送")
            self.save_email_history(from_email, to_email, subject, content, attachment_path)
            self.update_history_list()
        except Exception as e:
            messagebox.showerror("发送失败", f"邮件发送失败: {e}")

        self.save_email_info(from_email)

    def select_attachment(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.attachment_var.set(file_path)

    def select_db_file(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".db", filetypes=[("SQLite Database", "*.db")])
        if file_path:
            self.db_file.set(file_path)

    def save_settings(self):
        self.api_key.get()
        self.db_file.get()
        self.init_db()
        self.update_history_list()
        messagebox.showinfo("设置已保存", "API Key 和数据库存储位置已保存")

    def update_history_list(self):
        for i in self.history_tree.get_children():
            self.history_tree.delete(i)
        for row in self.load_email_history():
            self.history_tree.insert("", "end", values=row[1:])

    def show_email_details(self, event):
        if self.details_window_open:
            return

        item = self.history_tree.selection()
        if item:
            record = self.history_tree.item(item, "values")
            detail_window = tk.Toplevel(self.root)
            detail_window.title("邮件详情")
            detail_window.geometry("600x400")
            self.details_window_open = True

            details_text = f"发件邮箱: {record[0]}\n收件邮箱: {record[1]}\n主题: {record[2]}\n内容: {record[3]}\n附件: {record[4]}\n发送时间: {record[5]}"
            details_box = tk.Text(detail_window, wrap=tk.WORD)
            details_box.insert(tk.END, details_text)
            details_box.config(state=tk.DISABLED)
            details_box.pack(expand=True, fill=tk.BOTH)

            def on_close():
                self.details_window_open = False
                detail_window.destroy()

            detail_window.protocol("WM_DELETE_WINDOW", on_close)

    def show_context_menu(self, event):
        item = self.history_tree.selection()
        if item:
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="再次发送给TA", command=lambda: self.resend_email(item))
            menu.add_command(label="删除此记录", command=lambda: self.delete_record(item))
            menu.post(event.x_root, event.y_root)

    def resend_email(self, item):
        record = self.history_tree.item(item, "values")
        self.from_entry.delete(0, tk.END)
        self.from_entry.insert(0, record[0])
        self.to_entry.delete(0, tk.END)
        self.to_entry.insert(0, record[1])
        messagebox.showinfo("信息已填写", "发件邮箱和收件邮箱已填写")

    def delete_record(self, item):
        record = self.history_tree.item(item, "values")
        conn = sqlite3.connect(self.db_file.get())
        cursor = conn.cursor()
        cursor.execute(
            "DELETE FROM email_history WHERE from_email = ? AND to_email = ? AND subject = ? AND send_time = ?",
            (record[0], record[1], record[2], record[5]))
        conn.commit()
        conn.close()
        self.history_tree.delete(item)
        messagebox.showinfo("记录已删除", "邮件记录已成功删除")


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()
