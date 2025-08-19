import os
import sys
import subprocess
import smtplib
import re
import time
import logging
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import tkinter as tk
from tkinter import scrolledtext, messagebox, simpledialog, filedialog
from tkinter import ttk
import threading
import configparser

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("opera_monitor.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger("OperaMonitor")

class ConfigManager:
    def get_app_dir(self):
        """获取应用程序根目录"""
        if getattr(sys, 'frozen', False):
            # 如果是打包后的可执行文件
            return os.path.dirname(sys.executable)
        else:
            # 如果是开发环境
            return os.path.dirname(os.path.abspath(__file__))
            
    def __init__(self, config_file="opera_monitor.ini"):
        # 如果配置文件路径不是绝对路径，则使用应用程序目录
        if not os.path.isabs(config_file):
            app_dir = self.get_app_dir()
            self.config_file = os.path.join(app_dir, config_file)
        else:
            self.config_file = config_file
        self.config = configparser.ConfigParser()
        self.load_config()
    
    def load_config(self):
        # 获取应用程序根目录
        app_dir = self.get_app_dir()
            
        # 默认配置
        default_config = {
            'Email': {
                'smtp_server': 'smtp.example.com',
                'smtp_port': '587',
                'sender_email': 'your_email@example.com',
                'sender_password': '',
                'recipient_emails': 'recipient1@example.com,recipient2@example.com',
                'use_tls': 'True'
            },
            'Paths': {
                'check_standby_bat': os.path.join(app_dir, 'check_standby.bat'),
                'daily_report_bat': os.path.join(app_dir, 'daily_report.bat'),
                'report_path': os.path.join(app_dir, 'logs', 'daily_report.html')
            },
            'Settings': {
                'auto_run_interval': '86400',  # 24小时，单位：秒
                'check_errors': 'True',
                'error_patterns': 'error,warning,danger,failed,ORA-,TNS-',
            }
        }
        
        # 检查配置文件是否存在
        if os.path.exists(self.config_file):
            try:
                self.config.read(self.config_file, encoding='utf-8')
                logger.info(f"配置文件已加载: {self.config_file}")
            except Exception as e:
                logger.error(f"加载配置文件时出错: {e}")
                self.create_default_config(default_config)
        else:
            logger.info(f"配置文件不存在，创建默认配置: {self.config_file}")
            self.create_default_config(default_config)
    
    def create_default_config(self, default_config):
        for section, options in default_config.items():
            if not self.config.has_section(section):
                self.config.add_section(section)
            for option, value in options.items():
                if not self.config.has_option(section, option):
                    self.config.set(section, option, value)
        
        # 确保logs目录存在
        app_dir = self.get_app_dir()
        logs_dir = os.path.join(app_dir, 'logs')
        if not os.path.exists(logs_dir):
            os.makedirs(logs_dir)
            
        # 保存配置
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)
    
    def get(self, section, option, fallback=None):
        return self.config.get(section, option, fallback=fallback)
    
    def getboolean(self, section, option, fallback=None):
        return self.config.getboolean(section, option, fallback=fallback)
    
    def getint(self, section, option, fallback=None):
        return self.config.getint(section, option, fallback=fallback)
    
    def set(self, section, option, value):
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, option, value)
        with open(self.config_file, 'w', encoding='utf-8') as f:
            self.config.write(f)

class OperaMonitor:
    def __init__(self, root):
        self.root = root
        self.root.title("Opera数据库监控工具")
        self.root.geometry("900x700")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 加载配置
        self.config_manager = ConfigManager()
        
        # 创建UI组件
        self.create_widgets()
        
        # 初始化变量
        self.is_running = False
        self.auto_run_thread = None
        self.auto_run_active = False
        
        # 检查路径是否存在
        self.check_paths()
    
    def check_paths(self):
        check_standby_bat = self.config_manager.get('Paths', 'check_standby_bat')
        daily_report_bat = self.config_manager.get('Paths', 'daily_report_bat')
        report_path = self.config_manager.get('Paths', 'report_path')
        
        missing_paths = []
        if not os.path.exists(check_standby_bat):
            missing_paths.append(f"检查备用数据库脚本: {check_standby_bat}")
        if not os.path.exists(daily_report_bat):
            missing_paths.append(f"每日报告脚本: {daily_report_bat}")
        
        if missing_paths:
            message = "以下文件路径不存在，请在设置中更新:\n" + "\n".join(missing_paths)
            messagebox.showwarning("路径错误", message)
    
    def create_widgets(self):
        # 创建菜单栏
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="文件", menu=file_menu)
        file_menu.add_command(label="保存日志", command=self.save_log)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.on_closing)
        
        # 设置菜单
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="设置", menu=settings_menu)
        settings_menu.add_command(label="邮件设置", command=self.open_email_settings)
        settings_menu.add_command(label="路径设置", command=self.open_path_settings)
        settings_menu.add_command(label="监控设置", command=self.open_monitor_settings)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="关于", command=self.show_about)
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=5)
        
        # 添加按钮
        self.run_button = ttk.Button(button_frame, text="运行监控", command=self.run_monitor)
        self.run_button.pack(side=tk.LEFT, padx=5)
        
        self.auto_run_button = ttk.Button(button_frame, text="启动自动监控", command=self.toggle_auto_run)
        self.auto_run_button.pack(side=tk.LEFT, padx=5)
        
        self.send_email_button = ttk.Button(button_frame, text="发送邮件报告", command=self.send_email_report)
        self.send_email_button.pack(side=tk.LEFT, padx=5)
        
        self.view_report_button = ttk.Button(button_frame, text="查看HTML报告", command=self.view_html_report)
        self.view_report_button.pack(side=tk.LEFT, padx=5)
        
        self.clear_button = ttk.Button(button_frame, text="清除日志", command=self.clear_log)
        self.clear_button.pack(side=tk.LEFT, padx=5)
        
        # 创建状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 创建选项卡控件
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 日志选项卡
        log_frame = ttk.Frame(notebook)
        notebook.add(log_frame, text="执行日志")
        
        # 日志文本框
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # 分析选项卡
        analysis_frame = ttk.Frame(notebook)
        notebook.add(analysis_frame, text="分析结果")
        
        # 分析文本框
        self.analysis_text = scrolledtext.ScrolledText(analysis_frame, wrap=tk.WORD)
        self.analysis_text.pack(fill=tk.BOTH, expand=True)
        
        # 自动滚动日志
        self.log_text.see(tk.END)
    
    def run_monitor(self):
        if self.is_running:
            messagebox.showinfo("提示", "监控任务正在运行中，请等待完成")
            return
        
        # 在新线程中运行监控任务
        threading.Thread(target=self._run_monitor_thread, daemon=True).start()
    
    def _run_monitor_thread(self):
        self.is_running = True
        self.status_var.set("正在运行监控...")
        self.run_button.config(state=tk.DISABLED)
        
        try:
            # 清除分析结果
            self.analysis_text.delete(1.0, tk.END)
            
            # 获取批处理文件路径
            check_standby_bat = self.config_manager.get('Paths', 'check_standby_bat')
            daily_report_bat = self.config_manager.get('Paths', 'daily_report_bat')
            
            # 检查文件是否存在
            if not os.path.exists(check_standby_bat):
                self.log_message(f"错误: 文件不存在 - {check_standby_bat}")
                return
            if not os.path.exists(daily_report_bat):
                self.log_message(f"错误: 文件不存在 - {daily_report_bat}")
                return
            
            # 运行check_standby.bat
            self.log_message("开始执行 check_standby.bat...")
            check_standby_output = self.run_batch_file(check_standby_bat)
            self.log_message("check_standby.bat 执行完成")
            self.log_message("输出:\n" + check_standby_output)
            
            # 运行daily_report.bat
            self.log_message("开始执行 daily_report.bat...")
            daily_report_output = self.run_batch_file(daily_report_bat)
            self.log_message("daily_report.bat 执行完成")
            self.log_message("输出:\n" + daily_report_output)
            
            # 分析结果
            self.analyze_results(check_standby_output, daily_report_output)
            
            # 更新状态
            self.status_var.set("监控完成")
            self.log_message("监控任务完成")
            
            # 如果设置了自动发送邮件，则发送
            if self.config_manager.getboolean('Settings', 'auto_send_email', fallback=False):
                self.send_email_report()
        
        except Exception as e:
            self.log_message(f"执行监控时出错: {str(e)}")
            self.status_var.set("监控出错")
            logger.error(f"执行监控时出错: {str(e)}", exc_info=True)
        
        finally:
            self.is_running = False
            self.run_button.config(state=tk.NORMAL)
    
    def run_batch_file(self, batch_file):
        try:
            # 使用subprocess运行批处理文件并捕获输出
            process = subprocess.Popen(
                batch_file, 
                stdout=subprocess.PIPE, 
                stderr=subprocess.PIPE,
                shell=True,
                text=True,
                encoding='utf-8',
                errors='replace'
            )
            
            # 实时获取输出
            output = ""
            while True:
                line = process.stdout.readline()
                if not line and process.poll() is not None:
                    break
                if line:
                    output += line
                    self.log_message(line.strip())
                    self.root.update_idletasks()
            
            # 获取剩余输出和错误
            stdout, stderr = process.communicate()
            output += stdout
            
            if stderr:
                self.log_message("错误输出:\n" + stderr)
                output += "\n错误输出:\n" + stderr
            
            return output
        
        except Exception as e:
            error_msg = f"运行批处理文件时出错: {str(e)}"
            self.log_message(error_msg)
            logger.error(error_msg, exc_info=True)
            return error_msg
    
    def analyze_results(self, check_standby_output, daily_report_output):
        self.analysis_text.delete(1.0, tk.END)
        self.analysis_text.insert(tk.END, "===== 分析结果 =====\n\n")
        
        # 检查是否启用错误检查
        if not self.config_manager.getboolean('Settings', 'check_errors', fallback=True):
            self.analysis_text.insert(tk.END, "错误检查已禁用，跳过分析。\n")
            return
        
        # 获取错误模式
        error_patterns_str = self.config_manager.get('Settings', 'error_patterns', 
                                                 fallback='error,warning,danger,failed,ORA-,TNS-')
        error_patterns = [pattern.strip().lower() for pattern in error_patterns_str.split(',')]
        
        # 分析check_standby输出
        self.analysis_text.insert(tk.END, "1. Check Standby 分析:\n")
        standby_issues = self.check_for_issues(check_standby_output, error_patterns)
        
        if standby_issues:
            self.analysis_text.insert(tk.END, "   发现以下问题:\n")
            for issue in standby_issues:
                self.analysis_text.insert(tk.END, f"   - {issue}\n")
        else:
            self.analysis_text.insert(tk.END, "   未发现问题\n")
        
        # 提取最后10条SEQUENCE# APPLIED记录
        sequence_applied_pattern = r"\s*(\d+)\s+(YES|NO)\s*"
        sequence_applied_matches = re.findall(sequence_applied_pattern, check_standby_output)
        
        if sequence_applied_matches:
            self.analysis_text.insert(tk.END, "\n   最后10条 SEQUENCE# APPLIED:\n")
            # 取最后10条记录
            last_ten = sequence_applied_matches[-10:] if len(sequence_applied_matches) >= 10 else sequence_applied_matches
            for seq, applied in last_ten:
                self.analysis_text.insert(tk.END, f"   SEQUENCE#: {seq}, APPLIED: {applied}\n")
        
        # 分析HTML报告
        self.analysis_text.insert(tk.END, "\n2. HTML报告分析:\n")
        
        # 获取HTML报告路径
        report_path = self.config_manager.get('Paths', 'report_path')
        
        if os.path.exists(report_path):
            try:
                with open(report_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                
                # 检查HTML报告中的问题
                html_issues = self.check_for_issues(html_content, error_patterns)
                
                if html_issues:
                    self.analysis_text.insert(tk.END, "   HTML报告中发现以下问题:\n")
                    for issue in html_issues:
                        self.analysis_text.insert(tk.END, f"   - {issue}\n")
                else:
                    self.analysis_text.insert(tk.END, "   HTML报告中未发现问题\n")
                
                # 检查特定的数据库状态
                self.check_database_status(html_content)
                
            except Exception as e:
                self.analysis_text.insert(tk.END, f"   读取HTML报告时出错: {str(e)}\n")
        else:
            self.analysis_text.insert(tk.END, f"   HTML报告文件不存在: {report_path}\n")
        
        # 总结
        self.analysis_text.insert(tk.END, "\n3. 总结:\n")
        if standby_issues or (os.path.exists(report_path) and html_issues):
            self.analysis_text.insert(tk.END, "   监控发现异常情况，建议检查系统状态\n")
        else:
            self.analysis_text.insert(tk.END, "   所有检查正常\n")
    
    def check_for_issues(self, text, error_patterns):
        issues = []
        lines = text.lower().split('\n')
        
        for i, line in enumerate(lines):
            for pattern in error_patterns:
                if pattern in line:
                    # 获取上下文（前后各1行）
                    start = max(0, i - 1)
                    end = min(len(lines), i + 2)
                    context = '\n'.join(lines[start:end])
                    issues.append(f"发现 '{pattern}': {context}")
                    break
        
        return issues
    
    def check_database_status(self, html_content):
        # 检查数据库角色
        if "PRIMARY" in html_content and "PHYSICAL STANDBY" in html_content:
            self.analysis_text.insert(tk.END, "   数据库角色检查: ✅ 正常 (主库和备库都存在)\n")
        else:
            self.analysis_text.insert(tk.END, "   数据库角色检查: ❌ 异常 (可能缺少主库或备库)\n")
        
        # 检查归档日志间隙
        if "GAPS" in html_content and re.search(r'"GAPS"[^0-9]*0', html_content):
            self.analysis_text.insert(tk.END, "   归档日志间隙检查: ✅ 正常 (无间隙)\n")
        elif "GAPS" in html_content:
            self.analysis_text.insert(tk.END, "   归档日志间隙检查: ❌ 异常 (存在间隙)\n")
        
        # 检查未应用的日志
        if "NOT APPLIED" in html_content and re.search(r'"NOT APPLIED"[^0-9]*0', html_content):
            self.analysis_text.insert(tk.END, "   未应用日志检查: ✅ 正常 (无未应用日志)\n")
        elif "NOT APPLIED" in html_content:
            self.analysis_text.insert(tk.END, "   未应用日志检查: ❌ 异常 (存在未应用日志)\n")
        
        # 检查表空间使用情况
        if "DANGER" in html_content:
            self.analysis_text.insert(tk.END, "   表空间使用检查: 🔴 危险 (有表空间使用率超过90%)\n")
        elif "WARNING" in html_content:
            self.analysis_text.insert(tk.END, "   表空间使用检查: 🟡 警告 (有表空间使用率超过80%)\n")
        else:
            self.analysis_text.insert(tk.END, "   表空间使用检查: ✅ 正常\n")
            
        # 服务器运行时间分析
        uptime_pattern = r'START TIME[^\n]*([\d-]+\s+[\d:]+)'
        uptime_match = re.search(uptime_pattern, html_content)
        
        if uptime_match:
            start_time_str = uptime_match.group(1).strip()
            try:
                start_time = datetime.datetime.strptime(start_time_str, '%d-%b-%Y %H:%M')
                current_time = datetime.datetime.now()
                uptime_days = (current_time - start_time).days
                
                if uptime_days > 90:
                    self.analysis_text.insert(tk.END, f"   服务器运行时间分析: 🔴 警告 (已运行{uptime_days}天，建议重启)\n")
                elif uptime_days > 60:
                    self.analysis_text.insert(tk.END, f"   服务器运行时间分析: 🟡 注意 (已运行{uptime_days}天)\n")
                else:
                    self.analysis_text.insert(tk.END, f"   服务器运行时间分析: ✅ 正常 (已运行{uptime_days}天)\n")
            except Exception as e:
                self.analysis_text.insert(tk.END, f"   服务器运行时间分析: ❓ 无法解析 ({start_time_str})\n")
    
    def send_email_report(self):
        try:
            # 获取邮件设置
            smtp_server = self.config_manager.get('Email', 'smtp_server')
            smtp_port = self.config_manager.getint('Email', 'smtp_port')
            sender_email = self.config_manager.get('Email', 'sender_email')
            sender_password = self.config_manager.get('Email', 'sender_password')
            recipient_emails_str = self.config_manager.get('Email', 'recipient_emails')
            use_tls = self.config_manager.getboolean('Email', 'use_tls')
            
            # 检查必要的设置
            if not smtp_server or not sender_email or not recipient_emails_str:
                messagebox.showerror("邮件设置错误", "请先完成邮件设置")
                return
            
            # 解析收件人列表
            recipient_emails = [email.strip() for email in recipient_emails_str.split(',')]
            
            # 获取报告路径
            report_path = self.config_manager.get('Paths', 'report_path')
            if not os.path.exists(report_path):
                messagebox.showerror("文件错误", f"HTML报告文件不存在: {report_path}")
                return
            
            # 创建邮件
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = ", ".join(recipient_emails)
            msg['Subject'] = f"Opera数据库监控报告 - {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}"
            
            # 添加邮件正文
            body = "这是自动生成的Opera数据库监控报告，请查看附件。\n\n"
            
            # 添加分析结果
            analysis_text = self.analysis_text.get(1.0, tk.END)
            body += "分析结果:\n" + analysis_text
            
            # 创建HTML格式的邮件正文
            html_body = "<html><body>"
            html_body += "<p>这是自动生成的Opera数据库监控报告，请查看附件。</p>"
            html_body += "<h3>分析结果:</h3>"
            html_body += "<pre>"
            
            # 将分析文本中的表情符号转换为HTML格式
            analysis_lines = analysis_text.split('\n')
            for line in analysis_lines:
                if "服务器运行时间分析" in line:
                    if "🔴 警告" in line:
                        line = line.replace("🔴 警告", "<span style='color: red; font-weight: bold;'>⚠️ 警告</span>")
                    elif "🟡 注意" in line:
                        line = line.replace("🟡 注意", "<span style='color: orange; font-weight: bold;'>⚠️ 注意</span>")
                    elif "✅ 正常" in line:
                        line = line.replace("✅ 正常", "<span style='color: green; font-weight: bold;'>✓ 正常</span>")
                    elif "❓ 无法解析" in line:
                        line = line.replace("❓ 无法解析", "<span style='color: gray; font-weight: bold;'>❓ 无法解析</span>")
                elif "数据库角色检查" in line or "归档日志间隙检查" in line or "未应用日志检查" in line or "表空间使用检查" in line:
                    if "✅ 正常" in line:
                        line = line.replace("✅ 正常", "<span style='color: green; font-weight: bold;'>✓ 正常</span>")
                    elif "❌ 异常" in line:
                        line = line.replace("❌ 异常", "<span style='color: red; font-weight: bold;'>⚠️ 异常</span>")
                    elif "🔴 危险" in line:
                        line = line.replace("🔴 危险", "<span style='color: red; font-weight: bold;'>⚠️ 危险</span>")
                    elif "🟡 警告" in line:
                        line = line.replace("🟡 警告", "<span style='color: orange; font-weight: bold;'>⚠️ 警告</span>")
                
                html_body += line + "<br>"
            
            html_body += "</pre>"
            html_body += "</body></html>"
            
            # 添加纯文本和HTML格式的邮件正文
            msg.attach(MIMEText(body, 'plain'))
            msg.attach(MIMEText(html_body, 'html'))
            
            # 添加HTML报告附件
            with open(report_path, 'rb') as f:
                attachment = MIMEApplication(f.read(), _subtype='html')
                attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(report_path))
                msg.attach(attachment)
            
            # 添加日志附件
            log_content = self.log_text.get(1.0, tk.END)
            log_attachment = MIMEText(log_content, 'plain')
            log_attachment.add_header('Content-Disposition', 'attachment', filename='execution_log.txt')
            msg.attach(log_attachment)
            
            # 连接到SMTP服务器并发送邮件
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                if use_tls:
                    server.starttls()
                if sender_password:  # 只有在提供密码时才尝试登录
                    server.login(sender_email, sender_password)
                server.send_message(msg)
            
            messagebox.showinfo("成功", "邮件已成功发送")
            self.log_message("邮件已成功发送")
        
        except Exception as e:
            error_msg = f"发送邮件时出错: {str(e)}"
            messagebox.showerror("错误", error_msg)
            self.log_message(error_msg)
            logger.error(error_msg, exc_info=True)
    
    def view_html_report(self):
        report_path = self.config_manager.get('Paths', 'report_path')
        if os.path.exists(report_path):
            # 在默认浏览器中打开HTML报告
            if sys.platform == 'darwin':  # macOS
                subprocess.call(['open', report_path])
            elif sys.platform == 'win32':  # Windows
                os.startfile(report_path)
            else:  # Linux
                subprocess.call(['xdg-open', report_path])
        else:
            messagebox.showerror("文件错误", f"HTML报告文件不存在: {report_path}")
    
    def toggle_auto_run(self):
        if self.auto_run_active:
            # 停止自动运行
            self.auto_run_active = False
            self.auto_run_button.config(text="启动自动监控")
            self.status_var.set("自动监控已停止")
            self.log_message("自动监控已停止")
        else:
            # 启动自动运行
            self.auto_run_active = True
            self.auto_run_button.config(text="停止自动监控")
            self.status_var.set("自动监控已启动")
            self.log_message("自动监控已启动")
            
            # 在新线程中运行自动监控
            if self.auto_run_thread is None or not self.auto_run_thread.is_alive():
                self.auto_run_thread = threading.Thread(target=self._auto_run_thread, daemon=True)
                self.auto_run_thread.start()
    
    def _auto_run_thread(self):
        while self.auto_run_active:
            if not self.is_running:
                self._run_monitor_thread()
            
            # 获取自动运行间隔（秒）
            interval = self.config_manager.getint('Settings', 'auto_run_interval', fallback=86400)
            
            # 等待指定的时间，但每秒检查一次是否应该停止
            for _ in range(interval):
                if not self.auto_run_active:
                    break
                time.sleep(1)
    
    def log_message(self, message):
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        # 在UI中显示日志
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        
        # 同时记录到日志文件
        logger.info(message)
    
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
    
    def save_log(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            title="保存日志文件"
        )
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.get(1.0, tk.END))
                messagebox.showinfo("成功", f"日志已保存到: {file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"保存日志时出错: {str(e)}")
    
    def open_email_settings(self):
        settings_window = tk.Toplevel(self.root)
        settings_window.title("邮件设置")
        settings_window.geometry("500x350")
        settings_window.transient(self.root)
        settings_window.grab_set()
        
        frame = ttk.Frame(settings_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # SMTP服务器
        ttk.Label(frame, text="SMTP服务器:").grid(row=0, column=0, sticky=tk.W, pady=5)
        smtp_server_var = tk.StringVar(value=self.config_manager.get('Email', 'smtp_server'))
        ttk.Entry(frame, textvariable=smtp_server_var, width=40).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # SMTP端口
        ttk.Label(frame, text="SMTP端口:").grid(row=1, column=0, sticky=tk.W, pady=5)
        smtp_port_var = tk.StringVar(value=self.config_manager.get('Email', 'smtp_port'))
        ttk.Entry(frame, textvariable=smtp_port_var, width=10).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # 发件人邮箱
        ttk.Label(frame, text="发件人邮箱:").grid(row=2, column=0, sticky=tk.W, pady=5)
        sender_email_var = tk.StringVar(value=self.config_manager.get('Email', 'sender_email'))
        ttk.Entry(frame, textvariable=sender_email_var, width=40).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # 发件人密码
        ttk.Label(frame, text="发件人密码:").grid(row=3, column=0, sticky=tk.W, pady=5)
        sender_password_var = tk.StringVar(value=self.config_manager.get('Email', 'sender_password'))
        ttk.Entry(frame, textvariable=sender_password_var, width=40, show="*").grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # 收件人邮箱
        ttk.Label(frame, text="收件人邮箱:").grid(row=4, column=0, sticky=tk.W, pady=5)
        ttk.Label(frame, text="(多个邮箱用逗号分隔)").grid(row=5, column=0, sticky=tk.W)
        recipient_emails_var = tk.StringVar(value=self.config_manager.get('Email', 'recipient_emails'))
        ttk.Entry(frame, textvariable=recipient_emails_var, width=40).grid(row=4, column=1, rowspan=2, sticky=tk.W, pady=5)
        
        # 使用TLS
        use_tls_var = tk.BooleanVar(value=self.config_manager.getboolean('Email', 'use_tls'))
        ttk.Checkbutton(frame, text="使用TLS加密", variable=use_tls_var).grid(row=6, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # 自动发送邮件
        auto_send_var = tk.BooleanVar(value=self.config_manager.getboolean('Settings', 'auto_send_email', fallback=False))
        ttk.Checkbutton(frame, text="监控完成后自动发送邮件", variable=auto_send_var).grid(row=7, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # 测试按钮
        ttk.Button(frame, text="测试邮件设置", command=lambda: self.test_email_settings(
            smtp_server_var.get(),
            smtp_port_var.get(),
            sender_email_var.get(),
            sender_password_var.get(),
            recipient_emails_var.get(),
            use_tls_var.get()
        )).grid(row=8, column=0, pady=10)
        
        # 保存按钮
        ttk.Button(frame, text="保存设置", command=lambda: self.save_email_settings(
            smtp_server_var.get(),
            smtp_port_var.get(),
            sender_email_var.get(),
            sender_password_var.get(),
            recipient_emails_var.get(),
            use_tls_var.get(),
            auto_send_var.get(),
            settings_window
        )).grid(row=8, column=1, pady=10)
    
    def test_email_settings(self, smtp_server, smtp_port, sender_email, sender_password, recipient_emails, use_tls):
        try:
            # 验证输入
            if not smtp_server or not smtp_port or not sender_email or not recipient_emails:
                messagebox.showerror("输入错误", "请填写所有必要的字段")
                return
            
            # 解析收件人列表
            recipient_list = [email.strip() for email in recipient_emails.split(',')]
            if not recipient_list:
                messagebox.showerror("输入错误", "请提供至少一个有效的收件人邮箱")
                return
            
            # 创建测试邮件
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = ", ".join(recipient_list)
            msg['Subject'] = "Opera监控工具 - 邮件设置测试"
            msg.attach(MIMEText("这是一封测试邮件，用于验证Opera监控工具的邮件设置。", 'plain'))
            
            # 连接到SMTP服务器并发送邮件
            with smtplib.SMTP(smtp_server, int(smtp_port)) as server:
                if use_tls:
                    server.starttls()
                if sender_password:  # 只有在提供密码时才尝试登录
                    server.login(sender_email, sender_password)
                server.send_message(msg)
            
            messagebox.showinfo("成功", "测试邮件已成功发送")
        
        except Exception as e:
            messagebox.showerror("错误", f"发送测试邮件时出错: {str(e)}")
    
    def save_email_settings(self, smtp_server, smtp_port, sender_email, sender_password, recipient_emails, use_tls, auto_send, window):
        try:
            # 保存设置
            self.config_manager.set('Email', 'smtp_server', smtp_server)
            self.config_manager.set('Email', 'smtp_port', smtp_port)
            self.config_manager.set('Email', 'sender_email', sender_email)
            self.config_manager.set('Email', 'sender_password', sender_password)
            self.config_manager.set('Email', 'recipient_emails', recipient_emails)
            self.config_manager.set('Email', 'use_tls', str(use_tls))
            self.config_manager.set('Settings', 'auto_send_email', str(auto_send))
            
            messagebox.showinfo("成功", "邮件设置已保存")
            window.destroy()
        
        except Exception as e:
            messagebox.showerror("错误", f"保存设置时出错: {str(e)}")
    
    def open_path_settings(self):
        settings_window = tk.Toplevel(self.root)
        settings_window.title("路径设置")
        settings_window.geometry("600x250")
        settings_window.transient(self.root)
        settings_window.grab_set()
        
        frame = ttk.Frame(settings_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # check_standby.bat路径
        ttk.Label(frame, text="检查备用数据库脚本:").grid(row=0, column=0, sticky=tk.W, pady=5)
        check_standby_var = tk.StringVar(value=self.config_manager.get('Paths', 'check_standby_bat'))
        check_standby_entry = ttk.Entry(frame, textvariable=check_standby_var, width=50)
        check_standby_entry.grid(row=0, column=1, sticky=tk.W, pady=5)
        ttk.Button(frame, text="浏览", command=lambda: self.browse_file(check_standby_var)).grid(row=0, column=2, padx=5)
        
        # daily_report.bat路径
        ttk.Label(frame, text="每日报告脚本:").grid(row=1, column=0, sticky=tk.W, pady=5)
        daily_report_var = tk.StringVar(value=self.config_manager.get('Paths', 'daily_report_bat'))
        daily_report_entry = ttk.Entry(frame, textvariable=daily_report_var, width=50)
        daily_report_entry.grid(row=1, column=1, sticky=tk.W, pady=5)
        ttk.Button(frame, text="浏览", command=lambda: self.browse_file(daily_report_var)).grid(row=1, column=2, padx=5)
        
        # HTML报告路径
        ttk.Label(frame, text="HTML报告路径:").grid(row=2, column=0, sticky=tk.W, pady=5)
        report_path_var = tk.StringVar(value=self.config_manager.get('Paths', 'report_path'))
        report_path_entry = ttk.Entry(frame, textvariable=report_path_var, width=50)
        report_path_entry.grid(row=2, column=1, sticky=tk.W, pady=5)
        ttk.Button(frame, text="浏览", command=lambda: self.browse_file(report_path_var)).grid(row=2, column=2, padx=5)
        
        # 保存按钮
        ttk.Button(frame, text="保存设置", command=lambda: self.save_path_settings(
            check_standby_var.get(),
            daily_report_var.get(),
            report_path_var.get(),
            settings_window
        )).grid(row=3, column=1, pady=10)
    
    def browse_file(self, var):
        file_path = filedialog.askopenfilename()
        if file_path:
            var.set(file_path)
    
    def save_path_settings(self, check_standby_bat, daily_report_bat, report_path, window):
        try:
            # 保存设置
            self.config_manager.set('Paths', 'check_standby_bat', check_standby_bat)
            self.config_manager.set('Paths', 'daily_report_bat', daily_report_bat)
            self.config_manager.set('Paths', 'report_path', report_path)
            
            messagebox.showinfo("成功", "路径设置已保存")
            window.destroy()
            
            # 检查路径是否存在
            self.check_paths()
        
        except Exception as e:
            messagebox.showerror("错误", f"保存设置时出错: {str(e)}")
    
    def open_monitor_settings(self):
        settings_window = tk.Toplevel(self.root)
        settings_window.title("监控设置")
        settings_window.geometry("500x300")
        settings_window.transient(self.root)
        settings_window.grab_set()
        
        frame = ttk.Frame(settings_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 自动运行间隔
        ttk.Label(frame, text="自动运行间隔 (小时):").grid(row=0, column=0, sticky=tk.W, pady=5)
        interval_hours = self.config_manager.getint('Settings', 'auto_run_interval', fallback=86400) / 3600
        interval_var = tk.DoubleVar(value=interval_hours)
        ttk.Spinbox(frame, from_=1, to=48, increment=0.5, textvariable=interval_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # 启用错误检查
        check_errors_var = tk.BooleanVar(value=self.config_manager.getboolean('Settings', 'check_errors', fallback=True))
        ttk.Checkbutton(frame, text="启用错误检查", variable=check_errors_var).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # 错误模式
        ttk.Label(frame, text="错误模式:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Label(frame, text="(用逗号分隔)").grid(row=3, column=0, sticky=tk.W)
        error_patterns_var = tk.StringVar(value=self.config_manager.get('Settings', 'error_patterns', fallback='error,warning,danger,failed,ORA-,TNS-'))
        ttk.Entry(frame, textvariable=error_patterns_var, width=40).grid(row=2, column=1, rowspan=2, sticky=tk.W, pady=5)
        
        # 保存按钮
        ttk.Button(frame, text="保存设置", command=lambda: self.save_monitor_settings(
            interval_var.get(),
            check_errors_var.get(),
            error_patterns_var.get(),
            settings_window
        )).grid(row=4, column=0, columnspan=2, pady=10)
    
    def save_monitor_settings(self, interval_hours, check_errors, error_patterns, window):
        try:
            # 将小时转换为秒
            interval_seconds = int(interval_hours * 3600)
            
            # 保存设置
            self.config_manager.set('Settings', 'auto_run_interval', str(interval_seconds))
            self.config_manager.set('Settings', 'check_errors', str(check_errors))
            self.config_manager.set('Settings', 'error_patterns', error_patterns)
            
            messagebox.showinfo("成功", "监控设置已保存")
            window.destroy()
        
        except Exception as e:
            messagebox.showerror("错误", f"保存设置时出错: {str(e)}")
    
    def show_about(self):
        about_text = """Opera数据库监控工具

版本: 1.0

这是一个用于监控Opera数据库系统的工具，可以自动执行检查脚本、生成报告并通过邮件通知管理员。

© 2023 All Rights Reserved"""
        messagebox.showinfo("关于", about_text)
    
    def on_closing(self):
        if self.auto_run_active:
            if messagebox.askyesno("确认", "自动监控正在运行中，确定要退出吗？"):
                self.auto_run_active = False
                self.root.destroy()
        else:
            self.root.destroy()

def main():
    root = tk.Tk()
    app = OperaMonitor(root)
    root.mainloop()

if __name__ == "__main__":
    main()