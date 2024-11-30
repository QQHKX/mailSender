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

        # 创建历史记录选项卡
        history_frame = ttk.Frame(notebook)
        notebook.add(history_frame, text="发送记录")

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

        # 创建日志选项卡
        log_frame = ttk.Frame(notebook)
        notebook.add(log_frame, text="日志记录")

        self.log_text = tk.Text(log_frame, wrap=tk.WORD)
        self.log_text.pack(expand=True, fill=tk.BOTH)

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
            self.log("邮件发送失败：缺少必要字段")
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
            messagebox.showinfo("成功", "邮件发送成功")
            self.save_email_info(from_email)
            self.save_email_history(from_email, to_email, subject, content, attachment_path)
            self.update_history_list()
            self.log(f"邮件发送成功：{subject}")
        except Exception as e:
            messagebox.showerror("错误", f"发送邮件失败: {e}")
            self.log(f"邮件发送失败：{subject}，错误信息：{e}")

    def select_attachment(self):
        filepath = filedialog.askopenfilename()
        if filepath:
            self.attachment_var.set(filepath)

    def select_db_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("数据库文件", "*.db"), ("所有文件", "*.*")])
        if filepath:
            self.db_file.set(filepath)
            self.init_db()
            self.update_history_list()

    def save_settings(self):
        self.init_db()
        messagebox.showinfo("设置", "设置已保存")
        self.log("设置已保存")

    def update_history_list(self):
        for row in self.history_tree.get_children():
            self.history_tree.delete(row)

        for row in self.load_email_history():
            self.history_tree.insert("", "end", values=row)

    def show_email_details(self, event):
        item = self.history_tree.selection()
        if item:
            email_id = self.history_tree.item(item, "values")[0]
            details = self.get_email_details(email_id)
            if details and not self.details_window_open:
                self.details_window_open = True
                details_window = tk.Toplevel(self.root)
                details_window.title("邮件详情")
                details_window.geometry("400x300")
                details_window.protocol("WM_DELETE_WINDOW", lambda: self.close_details_window(details_window))

                details_text = tk.Text(details_window, wrap=tk.WORD)
                details_text.pack(expand=True, fill=tk.BOTH)
                details_text.insert(tk.END, details)

    def close_details_window(self, window):
        window.destroy()
        self.details_window_open = False

    def get_email_details(self, email_id):
        conn = sqlite3.connect(self.db_file.get())
        cursor = conn.cursor()
        cursor.execute('SELECT * FROM email_history WHERE id=?', (email_id,))
        email = cursor.fetchone()
        conn.close()
        if email:
            details = (f"发件邮箱: {email[1]}\n"
                       f"收件邮箱: {email[2]}\n"
                       f"邮件主题: {email[3]}\n"
                       f"邮件内容: {email[4]}\n"
                       f"附件: {email[5]}\n"
                       f"发送时间: {email[6]}\n")
            return details
        return ""

    def show_context_menu(self, event):
        item = self.history_tree.selection()
        if item:
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="删除记录", command=self.delete_record)
            menu.add_command(label="重新发送", command=self.resend_email)
            menu.post(event.x_root, event.y_root)

    def delete_record(self):
        item = self.history_tree.selection()
        if item:
            email_id = self.history_tree.item(item, "values")[0]
            conn = sqlite3.connect(self.db_file.get())
            cursor = conn.cursor()
            cursor.execute('DELETE FROM email_history WHERE id=?', (email_id,))
            conn.commit()
            conn.close()
            self.history_tree.delete(item)
            self.log(f"删除邮件记录：{email_id}")

    def resend_email(self):
        item = self.history_tree.selection()
        if item:
            email_id = self.history_tree.item(item, "values")[0]
            conn = sqlite3.connect(self.db_file.get())
            cursor = conn.cursor()
            cursor.execute('SELECT from_email, to_email, subject, content, attachment FROM email_history WHERE id=?', (email_id,))
            email = cursor.fetchone()
            conn.close()
            if email:
                from_email, to_email, subject, content, attachment_path = email
                try:
                    resend.api_key = self.api_key.get()

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

                    resend.Emails.send(params)
                    messagebox.showinfo("成功", "邮件重新发送成功")
                    self.log(f"重新发送邮件成功：{subject}")
                except Exception as e:
                    messagebox.showerror("错误", f"重新发送邮件失败: {e}")
                    self.log(f"重新发送邮件失败：{subject}，错误信息：{e}")

    def log(self, message):
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)


if __name__ == "__main__":
    root = tk.Tk()
    app = EmailSenderApp(root)
    root.mainloop()
