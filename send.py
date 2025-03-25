# -*- coding: utf-8 -*-
import os
import configparser
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
from tkinterdnd2 import TkinterDnD, DND_FILES
import threading
from tencentcloud.common import credential
from tencentcloud.common.exception.tencent_cloud_sdk_exception import TencentCloudSDKException
from tencentcloud.sms.v20210111 import sms_client, models
from tencentcloud.common.profile.client_profile import ClientProfile
from tencentcloud.common.profile.http_profile import HttpProfile
import ctypes
from ctypes import wintypes
class SMSApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Code By Orange~")
        
        # 设置任务栏图标
        self.setup_taskbar_icon()
        
        # 创建自定义标题栏
        title_bar = tk.Frame(self, bg='black', height=30)
        title_bar.pack(fill=tk.X)
    
        # 自定义标题文字
        title_label = tk.Label(title_bar, text="AAA招新短信发送助手", font=('二字元濑户淘气体-闪', 15), fg='white', bg='black')
        title_label.pack(side=tk.LEFT, padx=10)
        
        self.geometry("1000x800")
        self.resizable(False, False)
        self.config = configparser.ConfigParser()
        self.config.read('config.ini', encoding='utf-8')
        self.excel_path = ""
        self.running = False
        self.setup_style()
        self.create_widgets()
        self.setup_dnd()
        self.configure(background='black')
        
        # 初始化任务栏进度条
        self.taskbar_progress = None

    def setup_taskbar_icon(self):
        # 设置任务栏图标
        self.iconbitmap('icon.ico')               
        # 设置应用ID
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("AAA.Recruitment.SMS")

    def set_progress(self, value):
        """设置任务栏进度条（0-100）"""
        try:
            if os.name == 'nt' and self.taskbar_progress:
                hwnd = ctypes.windll.user32.GetParent(self.winfo_id())
                ctypes.windll.shell32.SetProgressValue(
                    hwnd,
                    ctypes.c_ulong(value),
                    ctypes.c_ulong(100)
                )
        except Exception as e:
            self.log(f"任务栏进度条设置失败: {str(e)}")

    def flash_window(self, count=3):
        """任务栏图标闪烁提示"""
        try:
            if os.name == 'nt':
                FLASHWINFO = ctypes.Structure
                class FLASHWINFO(ctypes.Structure):
                    _fields_ = [
                        ("cbSize", wintypes.UINT),
                        ("hwnd", wintypes.HWND),
                        ("dwFlags", wintypes.DWORD),
                        ("uCount", wintypes.UINT),
                        ("dwTimeout", wintypes.DWORD)
                    ]
                
                hwnd = self.winfo_id()
                flash_info = FLASHWINFO(
                    hwnd=hwnd,
                    dwFlags=0x3 | 0xC, 
                    uCount=count,
                    dwTimeout=0
                )
                flash_info.cbSize = ctypes.sizeof(flash_info)
                ctypes.windll.user32.FlashWindowEx(ctypes.byref(flash_info))
        except Exception as e:
            self.log(f"窗口闪烁设置失败: {str(e)}")

    def setup_style(self):
        style = ttk.Style()
        style.theme_use('alt')
        
        style.configure('.', background='black', foreground='white', font=('二字元濑户淘气体-闪', 12))
        style.map('Treeview', background=[('selected', '#666666')])
        
        style.configure('TFrame', background='black')
        style.configure('TButton', font=('二字元濑户淘气体-闪', 12), padding=6, background='#333333', foreground='white')
        style.map('TButton', background=[('active', '#444444'), ('disabled', '#222222')], foreground=[('active', 'white')])
        
        style.configure('TLabel', background='black', foreground='white', font=('二字元濑户淘气体-闪', 12))
        
        style.configure('Treeview', background='#222222', foreground='white', fieldbackground='#222222', rowheight=25)
        style.configure('Treeview.Heading', background='#333333', foreground='white', font=('二字元濑户淘气体-闪', 15))
        
        style.configure('TLabelFrame', background='black', foreground='white', font=('二字元濑户淘气体-闪', 15, 'bold'))

    def create_widgets(self):
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 文件选择区域
        file_frame = ttk.LabelFrame(main_frame, text="文件选择")
        file_frame.pack(fill=tk.X, pady=5)

        btn_frame = ttk.Frame(file_frame)
        btn_frame.pack(fill=tk.X, pady=5)

        self.btn_choose = ttk.Button(
            btn_frame, 
            text="选择Excel文件", 
            command=self.choose_file
        )
        self.btn_choose.pack(side=tk.LEFT, padx=5)

        self.lbl_path = ttk.Label(btn_frame, text="未选择文件")
        self.lbl_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # 数据预览区域
        preview_frame = ttk.LabelFrame(main_frame, text="数据预览")
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.tree = ttk.Treeview(preview_frame, show="headings", selectmode='browse')
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scroll_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scroll_y.set)

        scroll_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.configure(xscrollcommand=scroll_x.set)

        # 控制按钮
        btn_control_frame = ttk.Frame(main_frame)
        btn_control_frame.pack(pady=10)

        self.btn_start = ttk.Button(
            btn_control_frame,
            text="开始发送",
            command=self.start_sending
        )
        self.btn_start.pack(side=tk.LEFT, padx=5)

        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="发送日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_area = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            state='disabled',
            font=('二字元濑户淘气体-闪', 12),
            bg='#222222',
            fg='white',
            insertbackground='white'
        )
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

    def setup_dnd(self):
        self.drop_target_register(DND_FILES)
        self.dnd_bind('<<Drop>>', self.handle_drop)

    def handle_drop(self, event):
        files = event.data.split()
        if files:
            path = files[0].strip("{}")
            if path.lower().endswith(('.xls', '.xlsx')):
                self.set_excel_path(path)

    def choose_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xls *.xlsx")]
        )
        if path:
            self.set_excel_path(path)

    def set_excel_path(self, path):
        self.excel_path = path
        self.lbl_path.config(text=path)
        self.log("已选择文件: " + path)
        self.update_preview()

    def update_preview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.tree["columns"] = []

        try:
            df = pd.read_excel(self.excel_path)
            required_columns = ['名字', '电话', '日期', '面试时间', '面试地点']
            if not all(col in df.columns for col in required_columns):
                self.log("错误：Excel文件缺少必要列！")
                return

            columns = list(df.columns)
            self.tree["columns"] = columns

            col_widths = {'名字': 100, '电话': 120, '日期': 120, '面试时间': 100, '面试地点': 200}
            for col in columns:
                self.tree.heading(col, text=col, anchor=tk.CENTER)
                self.tree.column(col, width=col_widths.get(col, 100), anchor=tk.CENTER)

            for _, row in df.iterrows():
                self.tree.insert("", tk.END, values=list(row))

        except Exception as e:
            self.log(f"预览数据失败：{str(e)}")

    def start_sending(self):
        if not self.excel_path:
            self.log("请先选择Excel文件！")
            return

        if self.running:
            self.log("已有任务正在运行...")
            return

        self.running = True
        self.btn_start.config(state=tk.DISABLED)
        threading.Thread(target=self.send_sms, daemon=True).start()

    def send_sms(self):
        try:
            self.log("正在读取配置...")
            secret_id = self.config.get('TencentCloud', 'SecretId')
            secret_key = self.config.get('TencentCloud', 'SecretKey')
            region = self.config.get('TencentCloud', 'Region')
            sms_sdk_app_id = self.config.get('SMS', 'SmsSdkAppId')
            sign_name = self.config.get('SMS', 'SignName')
            template_id = self.config.get('SMS', 'TemplateId')

            cred = credential.Credential(secret_id, secret_key)
            
            httpProfile = HttpProfile()
            httpProfile.endpoint = "sms.tencentcloudapi.com"
            clientProfile = ClientProfile()
            clientProfile.httpProfile = httpProfile
            client = sms_client.SmsClient(cred, region, clientProfile)

            try:
                df = pd.read_excel(self.excel_path)
            except Exception as e:
                self.log(f"读取Excel失败: {str(e)}")
                return

            total = len(df)
            success = 0
            fail = 0

            self.log(f"开始发送短信，共 {total} 条记录...") 
            
            for index, row in df.iterrows():
                if not self.running:
                    break

                try:
                    # 更新进度条
                    progress = int((index + 1) / total * 100)
                    self.after(100, lambda: self.set_progress(progress))
                    
                    name = str(row["名字"])
                    raw_phone = str(row["电话"]).strip()
                    date = str(row["日期"])
                    time = str(row["面试时间"])
                    place = str(row["面试地点"])

                    phone = "+86" + raw_phone if not raw_phone.startswith("+") else raw_phone
                    template_params = [name, date, time, place]

                    req = models.SendSmsRequest()
                    req.SmsSdkAppId = sms_sdk_app_id
                    req.SignName = sign_name
                    req.TemplateId = template_id
                    req.TemplateParamSet = template_params
                    req.PhoneNumberSet = [phone]

                    resp = client.SendSms(req)
                    success += 1
                    self.log(f"[成功] {name} | 状态: {resp.SendStatusSet[0].Code}")
                except Exception as e:
                    fail += 1
                    self.log(f"[失败] 第{index+1}行 | 错误: {str(e)}")
                    
                    # 失败时闪烁窗口提醒
                    self.flash_window()

            self.log(f"发送完成！成功: {success} 条，失败: {fail} 条，总计: {total} 条")
            self.set_progress(0)  # 重置进度条

        except TencentCloudSDKException as err:
            self.log(f"腾讯云API错误: {err}")
        except Exception as e:
            self.log(f"程序运行错误: {str(e)}")
        finally:
            self.running = False
            self.after(100, lambda: self.btn_start.config(state=tk.NORMAL))

    def log(self, message):
        self.log_area.config(state=tk.NORMAL)
        self.log_area.insert(tk.END, message + "\n")
        self.log_area.see(tk.END)
        self.log_area.config(state=tk.DISABLED)

    def on_closing(self):
        if self.running:
            self.running = False
            self.log("\n正在停止发送任务...")
        self.destroy()

if __name__ == "__main__":
    app = SMSApp()
    app.iconbitmap("icon.ico")
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()