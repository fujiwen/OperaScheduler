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
        self.config = configparser.ConfigParser(interpolation=None)
        self.load_config()
    
    def load_config(self):
        # 获取应用程序根目录
        app_dir = self.get_app_dir()
            
        # 默认配置
        default_config = {
            'Email': {
                'smtp_server': 'smtp-sin02.aa.accor.net',
                'smtp_port': '587',
                'smtp_username': '',
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
                'auto_run_interval': '86400',  # 24小时，单位：秒（GUI模式使用）
                'run_hour': '8',  # 服务模式运行时间：小时
                'run_minute': '0',  # 服务模式运行时间：分钟
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
    def __init__(self, root=None):
        self.root = root
        self.service_mode = root is None
        
        if not self.service_mode:
            self.root.title("Opera DataGuard状态监测工具")
            self.root.geometry("900x700")
            self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # 加载配置
        self.config_manager = ConfigManager()
        
        # 创建UI组件（仅在GUI模式下）
        if not self.service_mode:
            self.create_widgets()
        
        # 初始化变量
        self.is_running = False
        self.auto_run_thread = None
        self.auto_run_active = False
        self.stop_requested = False
        
        # 存储执行输出
        self.check_standby_output = ""
        self.daily_report_output = ""
        
        # 检查路径是否存在
        self.check_paths()
        
        # 服务模式下的最后执行时间
        self.last_run_time = None
    
    def should_run_now(self):
        """检查是否应该运行监控任务（服务模式）"""
        if not self.service_mode:
            return False
            
        # 获取配置的运行时间（默认8:00）
        run_hour = self.config_manager.getint('Settings', 'run_hour', fallback=8)
        run_minute = self.config_manager.getint('Settings', 'run_minute', fallback=0)
        
        # 获取当前时间
        now = datetime.datetime.now()
        current_date = now.date()
        
        # 构造今天的目标运行时间
        target_time = datetime.datetime.combine(current_date, datetime.time(run_hour, run_minute))
        
        # 如果从未运行过
        if self.last_run_time is None:
            # 如果当前时间已经过了今天的运行时间，则立即运行
            if now >= target_time:
                return True
            else:
                return False
        
        # 将last_run_time转换为datetime对象
        last_run_datetime = datetime.datetime.fromtimestamp(self.last_run_time)
        last_run_date = last_run_datetime.date()
        
        # 如果今天还没有运行过，且当前时间已经到了或超过了运行时间
        if last_run_date < current_date and now >= target_time:
            return True
            
        return False
    
    def run_monitoring_task(self):
        """运行监控任务（服务模式）"""
        if self.is_running:
            logger.info("监控任务正在运行中，跳过此次执行")
            return
            
        self.is_running = True
        self.last_run_time = time.time()
        
        try:
            logger.info("开始执行监控任务...")
            
            # 获取批处理文件路径
            check_standby_bat = self.config_manager.get('Paths', 'check_standby_bat')
            daily_report_bat = self.config_manager.get('Paths', 'daily_report_bat')
            
            # 检查文件是否存在
            if not os.path.exists(check_standby_bat):
                logger.error(f"文件不存在: {check_standby_bat}")
                return
            if not os.path.exists(daily_report_bat):
                logger.error(f"文件不存在: {daily_report_bat}")
                return
            
            # 运行check_standby.bat
            logger.info("开始执行 check_standby.bat...")
            self.check_standby_output = self.run_batch_file(check_standby_bat)
            logger.info("check_standby.bat 执行完成")
            
            # 运行daily_report.bat
            logger.info("开始执行 daily_report.bat...")
            self.daily_report_output = self.run_batch_file(daily_report_bat)
            logger.info("daily_report.bat 执行完成")
            
            # 分析结果
            analysis_result = self.analyze_results_service_mode(self.check_standby_output, self.daily_report_output)
            logger.info(f"分析结果: {analysis_result}")
            
            # 如果设置了自动发送邮件，则发送
            if self.config_manager.getboolean('Settings', 'auto_send_email', fallback=False):
                self.send_email_report_service_mode(analysis_result)
                logger.info("邮件报告已发送")
            
            logger.info("监控任务完成")
            
        except Exception as e:
            logger.error(f"执行监控任务时出错: {str(e)}", exc_info=True)
        
        finally:
            self.is_running = False
    
    def stop_monitoring(self):
        """停止监控（服务模式）"""
        self.stop_requested = True
        self.auto_run_active = False
        logger.info("监控停止请求已发送")
    
    def analyze_results_service_mode(self, check_standby_output, daily_report_output):
        """分析结果（服务模式，不更新GUI）- 增强版"""
        analysis_result = []
        
        # 检查是否启用错误检查
        if not self.config_manager.getboolean('Settings', 'check_errors', fallback=True):
            analysis_result.append("错误检查已禁用")
            return "\n".join(analysis_result)
        
        # 获取错误模式
        error_patterns = self.config_manager.get('Settings', 'error_patterns', fallback='error,warning,danger,failed,ORA-,TNS-')
        patterns = [p.strip().lower() for p in error_patterns.split(',') if p.strip()]
        
        # 基础分析
        check_standby_errors = self._analyze_basic_errors(check_standby_output, patterns)
        daily_report_errors = self._analyze_basic_errors(daily_report_output, patterns)
        html_analysis = self.analyze_html_report_service_mode()
        
        # 1. Check Standby 分析
        analysis_result.append("1. Check Standby 分析:")
        if check_standby_errors:
            analysis_result.append(f"   发现 {len(check_standby_errors)} 个潜在问题:")
            analysis_result.extend([f"   - {error}" for error in check_standby_errors[:5]])
            if len(check_standby_errors) > 5:
                analysis_result.append(f"   ... 还有 {len(check_standby_errors) - 5} 个问题")
        else:
            analysis_result.append("   未发现问题")
        
        # 2. HTML报告分析
        analysis_result.append("\n2. HTML报告分析:")
        if html_analysis:
            analysis_result.append(f"   {html_analysis}")
        else:
            analysis_result.append("   HTML报告中未发现问题")
        
        # 3. 性能监控分析
        performance_analysis = self._analyze_performance(check_standby_output, daily_report_output)
        analysis_result.append("\n3. 性能监控分析:")
        analysis_result.extend([f"   {item}" for item in performance_analysis])
        
        # 4. 容量和空间分析
        capacity_analysis = self._analyze_capacity()
        analysis_result.append("\n4. 容量和空间分析:")
        analysis_result.extend([f"   {item}" for item in capacity_analysis])
        
        # 5. 风险预警分析
        risk_analysis = self._analyze_risks(check_standby_output, daily_report_output)
        analysis_result.append("\n5. 风险预警分析:")
        analysis_result.extend([f"   {item}" for item in risk_analysis])
        
        # 6. 智能告警分级
        alert_analysis = self._analyze_alert_levels(check_standby_errors, daily_report_errors, html_analysis)
        analysis_result.append("\n6. 智能告警分级:")
        analysis_result.extend([f"   {item}" for item in alert_analysis])
        
        # 7. 综合总结
        analysis_result.append("\n7. 综合总结:")
        summary = self._generate_comprehensive_summary(check_standby_errors, daily_report_errors, html_analysis, performance_analysis, capacity_analysis, risk_analysis)
        analysis_result.extend([f"   {item}" for item in summary])
        
        return "\n".join(analysis_result)
    
    def _analyze_basic_errors(self, output, patterns):
        """分析基础错误"""
        errors = []
        for line in output.split('\n'):
            line_lower = line.lower()
            for pattern in patterns:
                if pattern in line_lower:
                    errors.append(line.strip())
                    break
        return errors
    
    def _analyze_performance(self, check_standby_output, daily_report_output):
        """性能监控分析"""
        performance_issues = []
        
        # 分析同步延迟
        if "last applied" in daily_report_output.lower() and "last received" in daily_report_output.lower():
            performance_issues.append("同步状态: 正在监控主备库同步延迟")
        else:
            performance_issues.append("同步状态: 无法获取同步时间信息")
        
        # 分析进程状态
        if "mrp" in check_standby_output.lower():
            if "wait_for_log" in check_standby_output.lower():
                performance_issues.append("MRP进程: 正常运行，等待日志")
            elif "applying_log" in check_standby_output.lower():
                performance_issues.append("MRP进程: 正在应用日志")
            else:
                performance_issues.append("MRP进程: 状态需要关注")
        else:
            performance_issues.append("MRP进程: 未检测到进程信息")
        
        # 分析网络传输
        if "rfs" in check_standby_output.lower():
            performance_issues.append("网络传输: RFS进程正常运行")
        else:
            performance_issues.append("网络传输: 需要检查RFS进程状态")
        
        return performance_issues if performance_issues else ["性能监控数据不足"]
    
    def _analyze_capacity(self):
        """容量和空间分析"""
        capacity_issues = []
        
        try:
            # 分析HTML报告中的表空间使用情况
            report_path = self.config_manager.get('Paths', 'report_path')
            if os.path.exists(report_path):
                try:
                    with open(report_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                    
                    # 分析表空间使用情况
                    tablespace_stats = self._analyze_tablespace_usage(content)
                    if tablespace_stats:
                        capacity_issues.extend(tablespace_stats)
                    
                except Exception as e:
                    capacity_issues.append(f"分析表空间信息时出错: {str(e)}")
            
            # 检查日志目录空间
            if getattr(sys, 'frozen', False):
                app_dir = os.path.dirname(sys.executable)
            else:
                app_dir = os.path.dirname(os.path.abspath(__file__))
            
            logs_dir = os.path.join(app_dir, 'logs')
            if os.path.exists(logs_dir):
                # 计算日志目录大小
                total_size = 0
                file_count = 0
                for root, dirs, files in os.walk(logs_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        if os.path.exists(file_path):
                            total_size += os.path.getsize(file_path)
                            file_count += 1
                
                size_mb = total_size / (1024 * 1024)
                capacity_issues.append(f"日志目录: {file_count} 个文件，总大小 {size_mb:.2f} MB")
                
                if size_mb > 100:
                    capacity_issues.append("建议: 日志文件较大，建议定期清理")
                else:
                    capacity_issues.append("日志空间: 使用正常")
            else:
                capacity_issues.append("日志目录: 不存在，可能需要创建")
            
            # 检查磁盘空间（Windows）
            import shutil
            try:
                total, used, free = shutil.disk_usage(app_dir)
                free_gb = free / (1024**3)
                total_gb = total / (1024**3)
                usage_percent = (used / total) * 100
                
                capacity_issues.append(f"磁盘空间: 总计 {total_gb:.1f} GB，剩余 {free_gb:.1f} GB，使用率 {usage_percent:.1f}%")
                
                if usage_percent > 90:
                    capacity_issues.append("警告: 磁盘空间不足，使用率超过90%")
                elif usage_percent > 80:
                    capacity_issues.append("注意: 磁盘空间使用率超过80%")
                else:
                    capacity_issues.append("磁盘空间: 充足")
            except:
                capacity_issues.append("磁盘空间: 无法获取磁盘使用信息")
                
        except Exception as e:
            capacity_issues.append(f"容量分析失败: {str(e)}")
        
        return capacity_issues if capacity_issues else ["容量分析数据不足"]
    
    def _analyze_tablespace_usage(self, content):
        """分析表空间使用情况，统计SIZE (M)和MAX SIZE (M)"""
        tablespace_stats = []
        
        if "Tablespace usage:" not in content:
            return tablespace_stats
        
        try:
            import re
            
            # 查找表空间使用表格
            lines = content.split('\n')
            in_tablespace_section = False
            total_size = 0
            total_max_size = 0
            high_usage_count = 0
            tablespace_count = 0
            
            i = 0
            while i < len(lines):
                line = lines[i].strip()
                
                if "Tablespace usage:" in line:
                    in_tablespace_section = True
                    i += 1
                    continue
                
                if in_tablespace_section:
                    # 检查是否到了下一个section
                    if "</table>" in line:
                        break
                    
                    # 查找表空间名称行（不包含表头，不包含align="right"）
                    if line == "<td>" and i + 1 < len(lines):
                        # 下一行应该是表空间名称
                        next_line = lines[i + 1].strip()
                        if next_line and not next_line.startswith("<") and next_line not in ["TABLESPACE", "SIZE (M)", "MAX SIZE (M)", "USED %", "TYPE", "STATUS"]:
                            tablespace_name = next_line
                            
                            # 查找后续的SIZE (M)值
                            size_value = None
                            max_size_value = None
                            used_pct = None
                            
                            # 从当前位置开始查找数值
                            j = i + 2
                            numeric_values = []
                            
                            while j < len(lines) and len(numeric_values) < 3:
                                current_line = lines[j].strip()
                                
                                # 如果遇到下一个表空间或表格结束，停止
                                if current_line == "<td>" or "</table>" in current_line or "</tr>" in current_line:
                                    if "</tr>" in current_line:
                                        break
                                    j += 1
                                    continue
                                
                                # 尝试提取数值
                                if re.match(r'^[\d,]+\.\d+$', current_line):
                                    numeric_values.append(current_line.replace(',', ''))
                                elif re.match(r'^[\d.]+$', current_line) and '.' in current_line:
                                    numeric_values.append(current_line)
                                
                                j += 1
                            
                            # 如果找到了3个数值，分别是SIZE, MAX_SIZE, USED%
                            if len(numeric_values) >= 3:
                                try:
                                    size_value = float(numeric_values[0])
                                    max_size_value = float(numeric_values[1])
                                    used_pct = float(numeric_values[2])
                                    
                                    total_size += size_value
                                    total_max_size += max_size_value
                                    tablespace_count += 1
                                    
                                    if used_pct > 80:
                                        high_usage_count += 1
                                        
                                except ValueError:
                                    pass
                
                i += 1
            
            # 生成统计信息
            if tablespace_count > 0:
                # 转换为GB单位
                total_size_gb = total_size / 1024
                total_max_size_gb = total_max_size / 1024
                
                tablespace_stats.append(f"表空间统计: 共 {tablespace_count} 个表空间")
                tablespace_stats.append(f"总SIZE (G): {total_size_gb:,.2f} GB")
                tablespace_stats.append(f"总MAX SIZE (G): {total_max_size_gb:,.2f} GB")
                
                if total_max_size > 0:
                    overall_usage = (total_size / total_max_size) * 100
                    tablespace_stats.append(f"整体使用率: {overall_usage:.2f}%")
                    
                    if overall_usage > 80:
                        tablespace_stats.append("⚠️ 整体使用率较高，建议关注")
                    elif overall_usage > 60:
                        tablespace_stats.append("📊 整体使用率正常，需要监控")
                    else:
                        tablespace_stats.append("✅ 整体使用率良好")
                
                if high_usage_count > 0:
                    tablespace_stats.append(f"🔴 高使用率表空间数量: {high_usage_count} 个 (>80%)")
                else:
                    tablespace_stats.append("✅ 所有表空间使用率正常")
            
        except Exception as e:
            tablespace_stats.append(f"解析表空间数据时出错: {str(e)}")
        
        return tablespace_stats
    
    def _analyze_risks(self, check_standby_output, daily_report_output):
        """风险预警分析"""
        risk_issues = []
        
        # 检查连接状态
        if "ora-" in check_standby_output.lower() or "tns-" in check_standby_output.lower():
            risk_issues.append("连接风险: 检测到Oracle连接错误")
        elif "sqlplus" in check_standby_output.lower() and "not recognized" in check_standby_output.lower():
            risk_issues.append("环境风险: Oracle客户端未正确安装或配置")
        else:
            risk_issues.append("连接状态: 基础检查正常")
        
        # 检查进程健康
        process_count = 0
        if "mrp" in check_standby_output.lower():
            process_count += 1
        if "rfs" in check_standby_output.lower():
            process_count += 1
        if "lgwr" in check_standby_output.lower():
            process_count += 1
        
        if process_count >= 2:
            risk_issues.append(f"进程健康: 检测到 {process_count} 个关键进程运行")
        elif process_count == 1:
            risk_issues.append("进程健康: 部分关键进程运行，需要关注")
        else:
            risk_issues.append("进程健康: 未检测到关键进程，可能存在风险")
        
        # 检查配置一致性
        if "protection_mode" in daily_report_output.lower():
            risk_issues.append("配置检查: 保护模式信息可用")
        else:
            risk_issues.append("配置检查: 无法获取保护模式信息")
        
        # 检查备份完整性
        if "database_role" in check_standby_output.lower():
            if "standby" in check_standby_output.lower():
                risk_issues.append("角色检查: 数据库角色为备库，正常")
            else:
                risk_issues.append("角色检查: 数据库角色异常，需要确认")
        else:
            risk_issues.append("角色检查: 无法确认数据库角色")
        
        return risk_issues if risk_issues else ["风险评估数据不足"]
    
    def _analyze_alert_levels(self, check_standby_errors, daily_report_errors, html_analysis):
        """智能告警分级"""
        alerts = []
        alert_level = "normal"  # 默认正常级别
        critical_issues = []
        warning_issues = []
        
        # 1. 检查数据库运行时间
        db_runtime_alert = self._check_database_runtime_alert(html_analysis)
        if db_runtime_alert:
            if "alert-critical" in db_runtime_alert:
                alert_level = "alert-critical"
                critical_issues.append(db_runtime_alert)
            elif "alert-warning" in db_runtime_alert:
                if alert_level != "alert-critical":
                    alert_level = "alert-warning"
                warning_issues.append(db_runtime_alert)
        
        # 2. 检查HTML报告中的STATUS
        status_alert = self._check_status_alert(html_analysis)
        if status_alert:
            alert_level = "alert-critical"
            critical_issues.append(status_alert)
        
        # 3. 检查RMAN备份时间差
        backup_alert = self._check_backup_time_alert(html_analysis)
        if backup_alert:
            alert_level = "alert-critical"
            critical_issues.append(backup_alert)
        
        # 4. 检查容量使用率
        capacity_alert = self._check_capacity_alert(html_analysis)
        if capacity_alert:
            if "alert-critical" in capacity_alert:
                alert_level = "alert-critical"
                critical_issues.append(capacity_alert)
            elif "alert-warning" in capacity_alert:
                if alert_level != "alert-critical":
                    alert_level = "alert-warning"
                warning_issues.append(capacity_alert)
        
        # 5. 检查其他传统错误
        total_errors = len(check_standby_errors) + len(daily_report_errors)
        if html_analysis:
            if "danger" in html_analysis.lower() or "error" in html_analysis.lower() or "异常" in html_analysis:
                if alert_level == "normal":
                    alert_level = "alert-warning"
                    warning_issues.append("HTML报告中发现异常信息")
        
        # 生成告警级别描述
        if alert_level == "alert-critical":
            alerts.append("告警级别: 🔴 紧急 - 发现严重问题，需要立即处理")
            if critical_issues:
                alerts.extend(critical_issues)
        elif alert_level == "alert-warning":
            alerts.append("告警级别: 🟡 重要 - 发现重要问题，建议尽快处理")
            if warning_issues:
                alerts.extend(warning_issues)
            if critical_issues:
                alerts.extend(critical_issues)
        else:
            alerts.append("告警级别: ✅ 正常 - 系统运行正常")
        
        # 添加传统错误信息
        if total_errors > 0:
            alerts.append(f"检测到其他错误数: {total_errors}")
        
        # 根因分析
        if "sqlplus" in str(check_standby_errors).lower() and "not recognized" in str(check_standby_errors).lower():
            alerts.append("根因分析: Oracle客户端环境配置问题")
        elif "ora-" in str(check_standby_errors).lower():
            alerts.append("根因分析: Oracle数据库连接或配置问题")
        elif html_analysis and "间隙" in html_analysis:
            alerts.append("根因分析: 归档日志同步问题")
        
        return alerts if alerts else ["告警分析数据不足"]
    
    def _check_database_runtime_alert(self, html_analysis):
        """检查数据库运行时间告警"""
        if not html_analysis:
            return None
        
        try:
            # 从html_analysis中提取数据库运行时间信息
            lines = html_analysis.split('\n')
            alerts = []
            
            for line in lines:
                # 检查Standby数据库运行时间
                if "Standby数据库运行时间:" in line and "天" in line:
                    import re
                    match = re.search(r'(\d+)天', line)
                    if match:
                        days = int(match.group(1))
                        if days >= 63:
                            alerts.append("Standby数据库运行时间: {}天 - alert-critical (数据库启用已超过2个月，建议重启服务器以释放资源)".format(days))
                        elif days >= 32:
                            alerts.append("Standby数据库运行时间: {}天 - alert-warning (32-62天，需要关注)".format(days))
                
                # 检查Production数据库运行时间
                elif "Production数据库运行时间:" in line and "天" in line:
                    import re
                    match = re.search(r'(\d+)天', line)
                    if match:
                        days = int(match.group(1))
                        if days >= 63:
                            alerts.append("Production数据库运行时间: {}天 - alert-critical (数据库启用已超过2个月，建议重启服务器以释放资源)".format(days))
                        elif days >= 32:
                            alerts.append("Production数据库运行时间: {}天 - alert-warning (32-62天，需要关注)".format(days))
            
            # 返回最高级别的告警
            if alerts:
                # 优先返回critical级别的告警
                for alert in alerts:
                    if "alert-critical" in alert:
                        return alert
                # 如果没有critical，返回warning级别的告警
                for alert in alerts:
                    if "alert-warning" in alert:
                        return alert
            
            return None
        except Exception as e:
            logger.error(f"检查数据库运行时间告警时出错: {e}")
            return None
    
    def _check_status_alert(self, html_analysis):
        """检查HTML报告中的STATUS告警"""
        if not html_analysis:
            return None
        
        try:
            # 查找STATUS信息
            lines = html_analysis.split('\n')
            alerts = []
            
            for line in lines:
                # 检查Standby Database的STATUS
                if "Standby Database" in line and "STATUS" in line.upper():
                    if "MOUNTED" not in line.upper():
                        alerts.append("Standby数据库STATUS异常 - alert-critical (STATUS应为MOUNTED状态)")
                
                # 检查Production Database的STATUS
                elif "Production Database" in line and "STATUS" in line.upper():
                    if "OPEN" not in line.upper():
                        alerts.append("Production数据库STATUS异常 - alert-critical (STATUS应为OPEN状态)")
            
            # 返回第一个发现的STATUS异常
            if alerts:
                return alerts[0]
            
            return None
        except Exception as e:
            logger.error(f"检查STATUS告警时出错: {e}")
            return None
    
    def _check_backup_time_alert(self, html_analysis):
        """检查RMAN备份时间告警"""
        if not html_analysis:
            return None
        
        try:
            # 从html_analysis中提取备份END_TIME和SYSTEM DATE
            lines = html_analysis.split('\n')
            end_time = None
            system_date = None
            
            for line in lines:
                if "END_TIME:" in line:
                    # 提取END_TIME
                    import re
                    match = re.search(r'END_TIME:\s*([^\n]+)', line)
                    if match:
                        end_time = match.group(1).strip()
                elif "SYSTEM DATE:" in line:
                    # 提取SYSTEM DATE
                    match = re.search(r'SYSTEM DATE:\s*([^\n]+)', line)
                    if match:
                        system_date = match.group(1).strip()
            
            if end_time and system_date:
                # 计算时间差
                try:
                    import datetime
                    # 解析时间格式，例如: "22-JUL-2025 16:36"
                    end_dt = datetime.datetime.strptime(end_time, "%d-%b-%Y %H:%M")
                    system_dt = datetime.datetime.strptime(system_date, "%d-%b-%Y %H:%M")
                    
                    # 计算小时差
                    time_diff = system_dt - end_dt
                    hours_diff = time_diff.total_seconds() / 3600
                    
                    if hours_diff > 48:
                        return "RMAN备份故障，请立即检查 - alert-critical (备份时间超过48小时)"
                except Exception as e:
                    logger.error(f"解析备份时间时出错: {e}")
            
            return None
        except Exception as e:
            logger.error(f"检查备份时间告警时出错: {e}")
            return None
    
    def _check_capacity_alert(self, html_analysis):
        """检查容量使用率告警"""
        if not html_analysis:
            return None
        
        try:
            # 从html_analysis中提取整体使用率信息
            lines = html_analysis.split('\n')
            for line in lines:
                if "整体使用率:" in line and "%" in line:
                    # 提取使用率百分比
                    import re
                    match = re.search(r'整体使用率:\s*([\d.]+)%', line)
                    if match:
                        usage_percent = float(match.group(1))
                        if usage_percent > 80:
                            return "容量使用率: {:.2f}% - alert-critical (超过80%)".format(usage_percent)
                        elif usage_percent > 70:
                            return "容量使用率: {:.2f}% - alert-warning (超过70%)".format(usage_percent)
                        # 70%以下为正常，不返回告警
            return None
        except Exception as e:
            logger.error(f"检查容量使用率告警时出错: {e}")
            return None

    def _generate_comprehensive_summary(self, check_standby_errors, daily_report_errors, html_analysis, performance_analysis, capacity_analysis, risk_analysis):
        """生成综合总结"""
        summary = []
        
        # 获取告警级别分析结果
        alert_analysis = self._analyze_alert_levels(check_standby_errors, daily_report_errors, html_analysis)
        
        # 统计alert-warning和alert-critical的数量
        warning_count = 0
        critical_count = 0
        
        for alert in alert_analysis:
            if "alert-warning" in str(alert) or "🟡 重要" in str(alert):
                warning_count += 1
            elif "alert-critical" in str(alert) or "🔴 紧急" in str(alert):
                critical_count += 1
        
        # 计算总问题数（包括传统错误和新的告警）
        traditional_errors = len(check_standby_errors) + len(daily_report_errors)
        total_issues = warning_count + critical_count + traditional_errors
        
        # 移除问题总数显示
        
        # 性能状态
        if any("正常" in item for item in performance_analysis):
            summary.append("性能状态: 基本正常")
        else:
            summary.append("性能状态: 需要关注")
        
        # 容量状态
        if any("充足" in item or "正常" in item for item in capacity_analysis):
            summary.append("容量状态: 正常")
        else:
            summary.append("容量状态: 需要关注")
        
        # 风险评估
        if any("风险" in item for item in risk_analysis):
            summary.append("风险评估: 存在潜在风险")
        else:
            summary.append("风险评估: 风险可控")
        
        # 下次检查建议
        if total_issues > 0:
            summary.append("建议: 1小时后重新检查")
        else:
            summary.append("建议: 按计划进行下次检查")
        
        # 联系信息
        if total_issues > 3:
            summary.append("紧急情况: 建议联系数据库管理员")
        
        return summary
    
    def _analyze_trends_and_history(self):
        """趋势和历史分析"""
        trends = []
        
        try:
            # 检查历史日志文件
            if getattr(sys, 'frozen', False):
                app_dir = os.path.dirname(sys.executable)
            else:
                app_dir = os.path.dirname(os.path.abspath(__file__))
            
            log_file = os.path.join(app_dir, 'logs', 'opera_task.log')
            
            if os.path.exists(log_file):
                # 分析最近的日志记录
                with open(log_file, 'r', encoding='utf-8', errors='ignore') as f:
                    lines = f.readlines()
                
                # 统计最近的错误模式
                recent_errors = 0
                sqlplus_errors = 0
                connection_errors = 0
                
                # 只分析最后100行
                for line in lines[-100:]:
                    line_lower = line.lower()
                    if any(pattern in line_lower for pattern in ['error', 'failed', 'exception']):
                        recent_errors += 1
                    if 'sqlplus' in line_lower and 'not recognized' in line_lower:
                        sqlplus_errors += 1
                    if 'ora-' in line_lower or 'tns-' in line_lower:
                        connection_errors += 1
                
                trends.append(f"历史错误统计: 最近检测到 {recent_errors} 个错误")
                
                if sqlplus_errors > 0:
                    trends.append(f"环境问题: 检测到 {sqlplus_errors} 次Oracle客户端问题")
                    trends.append("趋势分析: 建议优先解决Oracle客户端配置问题")
                
                if connection_errors > 0:
                    trends.append(f"连接问题: 检测到 {connection_errors} 次数据库连接问题")
                    trends.append("趋势分析: 建议检查网络连接稳定性")
                
                if recent_errors == 0:
                    trends.append("趋势分析: 系统运行稳定，无明显异常趋势")
                elif recent_errors > 10:
                    trends.append("趋势分析: 错误频率较高，建议深入排查")
                else:
                    trends.append("趋势分析: 偶发性问题，建议持续监控")
                
                # 预测建议
                if sqlplus_errors > connection_errors:
                    trends.append("预测建议: 主要问题为环境配置，解决后系统稳定性将显著提升")
                elif connection_errors > 0:
                    trends.append("预测建议: 存在网络或数据库连接问题，需要网络团队协助")
                else:
                    trends.append("预测建议: 系统整体稳定，建议保持当前监控频率")
                    
            else:
                trends.append("历史数据: 日志文件不存在，无法进行趋势分析")
                
        except Exception as e:
            trends.append(f"趋势分析失败: {str(e)}")
        
        return trends if trends else ["趋势分析数据不足"]
    

    

    
    def analyze_html_report_service_mode(self):
        """分析HTML报告（服务模式）"""
        try:
            # 自动检测HTML报告文件路径
            if getattr(sys, 'frozen', False):
                # 如果是打包的可执行文件
                app_dir = os.path.dirname(sys.executable)
            else:
                # 如果是Python脚本
                app_dir = os.path.dirname(os.path.abspath(__file__))
            
            report_path = os.path.join(app_dir, 'logs', 'daily_report.html')
            
            if not os.path.exists(report_path):
                return f"HTML报告文件不存在: {report_path}"
            
            with open(report_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 提取详细的数据库信息
            analysis_details = self._extract_database_details(content)
            
            # 检查特定的异常情况
            issues = []
            if "归档日志间隙检查: 异常" in content:
                issues.append("发现归档日志间隙")
            if "未应用日志检查: 异常" in content:
                issues.append("发现未应用日志")
            if "DANGER" in content:
                issues.append("发现危险状态")
            if "ERROR" in content:
                issues.append("发现错误")
            if "Exception" in content:
                issues.append("发现异常")
            
            # 合并问题和详细信息
            result = analysis_details
            if issues:
                result += "\n\n异常情况: " + "; ".join(issues)
            
            return result if result.strip() else None
            
        except Exception as e:
            logger.error(f"分析HTML报告时出错: {e}")
            return f"HTML报告分析失败: {str(e)}"
    
    def _extract_database_details(self, content):
        """从HTML报告中提取详细的数据库信息"""
        details = []
        
        try:
            # 1. 提取Standby Database信息
            standby_info = self._extract_standby_database_info(content)
            if standby_info:
                details.append("Standby Database 备库信息:")
                details.append(standby_info)
            
            # 2. 提取Production Database信息
            production_info = self._extract_production_database_info(content)
            if production_info:
                details.append("\n Production Database 主库信息:")
                details.append(production_info)
            
            # 3. 分析START TIME和SYSTEM DATE的时间差
            time_analysis = self._analyze_time_difference(content)
            if time_analysis:
                details.append("\n 服务器运行时间分析:")
                details.append(time_analysis)
            
            # 4. 提取Opera版本信息
            opera_version = self._extract_opera_version_info(content)
            if opera_version:
                details.append("\n Opera Version Information:")
                details.append(opera_version)
            
            # 5. 提取Oracle版本信息
            oracle_version = self._extract_oracle_version_info(content)
            if oracle_version:
                details.append("\n Oracle Version Information:")
                details.append(oracle_version)
            
            # 6. 提取最近3天备份的最后一条数据
            backup_info = self._extract_last_backup_info(content)
            if backup_info:
                details.append("\n List of last 3 days backups (最后一条数据):")
                details.append(backup_info)
            
            return "\n".join(details)
            
        except Exception as e:
            logger.error(f"提取数据库详细信息时出错: {e}")
            return f"提取详细信息失败: {str(e)}"
    
    def _analyze_time_difference(self, content):
        """分析START TIME和SYSTEM DATE的时间差"""
        try:
            # 提取Standby和Production的START TIME和SYSTEM DATE
            standby_times = self._extract_times_from_section(content, "Standby Database")
            production_times = self._extract_times_from_section(content, "Production Database")
            
            results = []
            
            # 分析Standby数据库时间差
            if standby_times:
                start_time, system_date = standby_times
                if start_time and system_date:
                    days_diff = self._calculate_days_difference(start_time, system_date)
                    color_status = self._get_time_status_with_color(days_diff)
                    results.append(f"   Standby数据库运行时间: {days_diff}天 {color_status}")
            
            # 分析Production数据库时间差
            if production_times:
                start_time, system_date = production_times
                if start_time and system_date:
                    days_diff = self._calculate_days_difference(start_time, system_date)
                    color_status = self._get_time_status_with_color(days_diff)
                    results.append(f"   Production数据库运行时间: {days_diff}天 {color_status}")
            
            return "\n".join(results) if results else None
            
        except Exception as e:
            logger.error(f"分析时间差时出错: {e}")
            return None
    
    def _extract_times_from_section(self, content, section_name):
        """从指定数据库部分提取START TIME和SYSTEM DATE"""
        try:
            # 查找对应数据库部分的General Database Information表格
            if section_name == "Standby Database":
                pattern = r'<h2>Standby Database</h2>.*?<h3>General Database Information:</h3>.*?<table[^>]*>(.*?)</table>'
            else:
                pattern = r'<h2>Production Database</h2>.*?<h3>General Database Information:</h3>.*?<table[^>]*>(.*?)</table>'
            
            table_match = re.search(pattern, content, re.DOTALL | re.IGNORECASE)
            
            if not table_match:
                return None
            
            table_content = table_match.group(1)
            
            # 查找数据行（跳过表头）
            rows = re.findall(r'<tr[^>]*>(.*?)</tr>', table_content, re.DOTALL)
            
            for row in rows:
                # 跳过表头行
                if re.search(r'<th[^>]*>', row, re.IGNORECASE):
                    continue
                    
                cells = re.findall(r'<td[^>]*>(.*?)</td>', row, re.DOTALL)
                if len(cells) >= 7:
                    # 清理HTML标签
                    clean_cells = [re.sub(r'<[^>]+>', '', cell).strip() for cell in cells]
                    
                    # 根据实际HTML结构确定列索引
                    if section_name == "Standby Database" and len(clean_cells) >= 9:
                        # Standby: INST_ID, DATABASE NAME, INSTANCE NAME, STATUS, HOST NAME, DATABASE ROLE, PROTECTION MODE, START TIME, SYSTEM DATE
                        start_time = clean_cells[7] if len(clean_cells) > 7 else None  # START TIME在第8列
                        system_date = clean_cells[8] if len(clean_cells) > 8 else None  # SYSTEM DATE在第9列
                    elif section_name == "Production Database" and len(clean_cells) >= 8:
                        # Production: INST_ID, DATABASE NAME, INSTANCE, STATUS, HOST NAME, DATABASE ROLE, START TIME, SYSTEM DATE
                        start_time = clean_cells[6] if len(clean_cells) > 6 else None  # START TIME在第7列
                        system_date = clean_cells[7] if len(clean_cells) > 7 else None  # SYSTEM DATE在第8列
                    
                    # 如果找到了有效的时间信息就返回
                    if start_time and system_date and start_time.strip() != '' and system_date.strip() != '':
                        return (start_time.strip(), system_date.strip())
            
            return None
            
        except Exception as e:
            logger.error(f"提取{section_name}时间信息时出错: {e}")
            return None
    
    def _calculate_days_difference(self, start_time_str, system_date_str):
        """计算两个时间字符串之间的天数差"""
        try:
            # 解析时间格式，例如: "22-JUL-2025 16:36" 和 "07-AUG-2025 10:21"
            start_time = datetime.datetime.strptime(start_time_str, "%d-%b-%Y %H:%M")
            system_date = datetime.datetime.strptime(system_date_str, "%d-%b-%Y %H:%M")
            
            # 计算天数差
            time_diff = system_date - start_time
            return time_diff.days
            
        except Exception as e:
            logger.error(f"计算时间差时出错: {e}")
            return 0
    
    def _get_time_status_with_color(self, days):
        """根据天数返回带颜色的状态信息"""
        if days <= 31:
            return "<span style='color: green; font-weight: bold;'>状态正常</span>"
        elif days <= 62:
            return "<span style='color: orange; font-weight: bold;'>需要关注</span>"
        else:
            return "<span style='color: red; font-weight: bold;'>立即关注，建议立即重启服务器以释放资源</span>"
    
    def _extract_standby_database_info(self, content):
        """提取Standby Database信息"""
        try:
            # 查找Standby Database部分
            standby_section = re.search(r'<h2>Standby Database</h2>.*?<h3>General Database Information:</h3>(.*?)(?=<h3>|<h2>|</body>|$)', content, re.DOTALL | re.IGNORECASE)
            if not standby_section:
                return None
            
            section_content = standby_section.group(1)
            
            # 提取表格数据 - 查找表格行
            table_rows = re.findall(r'<tr[^>]*>(.*?)</tr>', section_content, re.DOTALL)
            
            info = {}
            
            # 查找数据行（非表头行）
            for row in table_rows:
                if '<th' not in row:  # 跳过表头行
                    cells = re.findall(r'<td[^>]*>(.*?)</td>', row, re.DOTALL)
                    if len(cells) >= 9:  # 确保有足够的列
                        # 根据表格列顺序提取数据
                        info['INST_ID'] = re.sub(r'<[^>]+>', '', cells[0]).strip()
                        info['DATABASE NAME'] = re.sub(r'<[^>]+>', '', cells[1]).strip()
                        info['INSTANCE NAME'] = re.sub(r'<[^>]+>', '', cells[2]).strip()
                        info['STATUS'] = re.sub(r'<[^>]+>', '', cells[3]).strip()
                        info['HOST NAME'] = re.sub(r'<[^>]+>', '', cells[4]).strip()
                        info['DATABASE ROLE'] = re.sub(r'<[^>]+>', '', cells[5]).strip()
                        info['PROTECTION MODE'] = re.sub(r'<[^>]+>', '', cells[6]).strip()
                        info['START TIME'] = re.sub(r'<[^>]+>', '', cells[7]).strip()
                        info['SYSTEM DATE'] = re.sub(r'<[^>]+>', '', cells[8]).strip()
                        break
            
            if info:
                result = []
                for key in ['STATUS', 'HOST NAME', 'DATABASE ROLE', 'PROTECTION MODE', 'START TIME', 'SYSTEM DATE']:
                    if key in info:
                        result.append(f"   {key}: {info[key]}")
                return "\n".join(result)
            
            return None
            
        except Exception as e:
            logger.error(f"提取Standby数据库信息时出错: {e}")
            return None
    
    def _extract_production_database_info(self, content):
        """提取Production Database信息"""
        try:
            # 查找Production Database部分
            production_section = re.search(r'<h2>Production Database</h2>.*?<h3>General Database Information:</h3>(.*?)(?=<h3>|<h2>|</body>|$)', content, re.DOTALL | re.IGNORECASE)
            if not production_section:
                return None
            
            section_content = production_section.group(1)
            
            # 提取表格数据 - 查找表格行
            table_rows = re.findall(r'<tr[^>]*>(.*?)</tr>', section_content, re.DOTALL)
            
            info = {}
            
            # 查找数据行（非表头行）
            for row in table_rows:
                if '<th' not in row:  # 跳过表头行
                    cells = re.findall(r'<td[^>]*>(.*?)</td>', row, re.DOTALL)
                    if len(cells) >= 8:  # Production Database表格有8列
                        # 根据表格列顺序提取数据（Production Database没有PROTECTION MODE列）
                        info['INST_ID'] = re.sub(r'<[^>]+>', '', cells[0]).strip()
                        info['DATABASE NAME'] = re.sub(r'<[^>]+>', '', cells[1]).strip()
                        info['INSTANCE'] = re.sub(r'<[^>]+>', '', cells[2]).strip()
                        info['STATUS'] = re.sub(r'<[^>]+>', '', cells[3]).strip()
                        info['HOST NAME'] = re.sub(r'<[^>]+>', '', cells[4]).strip()
                        info['DATABASE ROLE'] = re.sub(r'<[^>]+>', '', cells[5]).strip()
                        info['START TIME'] = re.sub(r'<[^>]+>', '', cells[6]).strip()
                        info['SYSTEM DATE'] = re.sub(r'<[^>]+>', '', cells[7]).strip()
                        break
            
            if info:
                result = []
                # Production Database没有PROTECTION MODE字段
                for key in ['STATUS', 'HOST NAME', 'DATABASE ROLE', 'START TIME', 'SYSTEM DATE']:
                    if key in info:
                        result.append(f"   {key}: {info[key]}")
                return "\n".join(result)
            
            return None
            
        except Exception as e:
            logger.error(f"提取Production数据库信息时出错: {e}")
            return None
    
    def _extract_opera_version_info(self, content):
        """提取Opera版本信息"""
        try:
            # 查找Opera Version Information部分
            opera_match = re.search(r'<h3>Opera Version Information:</h3>(.*?)(?=<h3>|</body>|$)', content, re.DOTALL | re.IGNORECASE)
            if not opera_match:
                return None
            
            section_content = opera_match.group(1)
            
            # 提取版本信息
            version_match = re.search(r'<td[^>]*>(.*?)</td>', section_content, re.DOTALL)
            if version_match:
                version = version_match.group(1).strip()
                # 清理HTML标签
                version = re.sub(r'<[^>]+>', '', version).strip()
                return f"   Opera Version: {version}"
            
            return None
            
        except Exception as e:
            logger.error(f"提取Opera版本信息时出错: {e}")
            return None
    
    def _extract_oracle_version_info(self, content):
        """提取Oracle版本信息"""
        try:
            # 查找Oracle version Information部分
            oracle_match = re.search(r'<h3>Oracle version Information:</h3>(.*?)(?=<h3>|<h2>|</body>|$)', content, re.DOTALL | re.IGNORECASE)
            if not oracle_match:
                return None
            
            section_content = oracle_match.group(1)
            
            # 提取版本信息和平台信息
            version_info = []
            
            # 提取VERSION表格
            version_match = re.search(r'<th[^>]*>\s*VERSION\s*</th>.*?<td[^>]*>(.*?)</td>', section_content, re.DOTALL)
            if version_match:
                version_text = re.sub(r'<[^>]+>', '', version_match.group(1)).strip()
                version_info.append(f"   Oracle Version: {version_text}")
            
            # 提取PLATFORM_NAME表格
            platform_match = re.search(r'<th[^>]*>\s*PLATFORM_NAME\s*</th>.*?<td[^>]*>(.*?)</td>', section_content, re.DOTALL)
            if platform_match:
                platform_text = re.sub(r'<[^>]+>', '', platform_match.group(1)).strip()
                version_info.append(f"   Platform: {platform_text}")
            
            if version_info:
                return "\n".join(version_info)
            
            return None
            
        except Exception as e:
            logger.error(f"提取Oracle版本信息时出错: {e}")
            return None
    
    def _extract_last_backup_info(self, content):
        """提取最近3天备份的最后一条数据"""
        try:
            # 查找List of last 3 days backups部分
            backup_match = re.search(r'<h3>List of last 3 days backups:</h3>(.*?)(?=<h3>|<h2>|</body>|spool off)', content, re.DOTALL | re.IGNORECASE)
            if not backup_match:
                return None
            
            section_content = backup_match.group(1)
            
            # 查找表格中的最后一行数据
            # 提取所有表格行
            rows = re.findall(r'<tr[^>]*>(.*?)</tr>', section_content, re.DOTALL)
            
            if not rows:
                # 如果没有找到表格行，查找是否有"No backup information found"消息
                if "No backup information found" in section_content:
                    return "   未找到备份信息，请检查备份日志"
                return None
            
            # 获取最后一行（排除表头）
            last_row = None
            for row in reversed(rows):
                # 跳过表头行
                if not re.search(r'<th[^>]*>', row, re.IGNORECASE):
                    last_row = row
                    break
            
            if last_row:
                # 提取单元格数据
                cells = re.findall(r'<td[^>]*>(.*?)</td>', last_row, re.DOTALL)
                if cells:
                    # 清理HTML标签并格式化
                    clean_cells = []
                    for cell in cells:
                        clean_cell = re.sub(r'<[^>]+>', '', cell).strip()
                        clean_cells.append(clean_cell)
                    
                    # 根据生产环境表格的列顺序格式化输出
                    if len(clean_cells) >= 9:
                        result = []
                        result.append(f"   SESSION_RECID: {clean_cells[0]}")
                        result.append(f"   START_TIME: {clean_cells[1]}")
                        result.append(f"   END_TIME: {clean_cells[2]}")
                        result.append(f"   OUTPUT_MBYTES: {clean_cells[3]}")
                        result.append(f"   STATUS: {clean_cells[4]}")
                        result.append(f"   INPUT_TYPE: {clean_cells[5]}")
                        result.append(f"   DAY: {clean_cells[6]}")
                        result.append(f"   TIME_TAKEN: {clean_cells[7]}")
                        result.append(f"   OUTPUT_INSTANCE: {clean_cells[8]}")
                        
                        return "\n".join(result)
            
            return None
            
        except Exception as e:
            logger.error(f"提取备份信息时出错: {e}")
            return None
    
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
            if not self.service_mode:
                messagebox.showwarning("路径错误", message)
            else:
                logger.warning(message)
    
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
        settings_menu.add_separator()
        settings_menu.add_command(label="服务管理", command=self.open_service_management)
        
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
            self.check_standby_output = self.run_batch_file(check_standby_bat)
            self.log_message("check_standby.bat 执行完成")
            self.log_message("输出:\n" + self.check_standby_output)
            
            # 运行daily_report.bat
            self.log_message("开始执行 daily_report.bat...")
            self.daily_report_output = self.run_batch_file(daily_report_bat)
            self.log_message("daily_report.bat 执行完成")
            self.log_message("输出:\n" + self.daily_report_output)
            
            # 分析结果
            self.analyze_results(self.check_standby_output, self.daily_report_output)
            
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
            # 获取应用程序目录作为工作目录
            app_dir = self.config_manager.get_app_dir()
            
            # 使用subprocess运行批处理文件并捕获输出
            process = subprocess.Popen(
                batch_file, 
                stdout=subprocess.PIPE, 
                stderr=subprocess.PIPE,
                shell=True,
                text=True,
                encoding='utf-8',
                errors='replace',
                cwd=app_dir  # 设置工作目录为应用程序目录
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
                    # 只在GUI模式下更新界面
                    if not self.service_mode and self.root:
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
            self.analysis_text.insert(tk.END, "   数据库角色检查: 正常 (主库和备库都存在)\n")
        else:
            self.analysis_text.insert(tk.END, "   数据库角色检查: 异常 (可能缺少主库或备库)\n")
        
        # 检查归档日志间隙 - 提取详细信息
        if "归档日志间隙检查: 异常" in html_content:
            self.analysis_text.insert(tk.END, "   归档日志间隙检查: 异常 (存在间隙)\n")
            
            # 提取间隙数量
            gap_count_match = re.search(r'间隙数量: (\d+) 个日志文件', html_content)
            if gap_count_match:
                gap_count = gap_count_match.group(1)
                self.analysis_text.insert(tk.END, f"     - 间隙数量: {gap_count} 个日志文件\n")
            
            # 提取间隙范围
            gap_range_match = re.search(r'间隙范围: 序列号 (\d+) 到 (\d+)', html_content)
            if gap_range_match:
                low_seq = gap_range_match.group(1)
                high_seq = gap_range_match.group(2)
                self.analysis_text.insert(tk.END, f"     - 缺失序列号范围: {low_seq} 到 {high_seq}\n")
                self.analysis_text.insert(tk.END, "     - 建议: 检查网络连接，手动复制缺失的归档日志文件\n")
        elif "归档日志间隙检查: 正常" in html_content or ("GAPS" in html_content and re.search(r'"GAPS"[^0-9]*0', html_content)):
            self.analysis_text.insert(tk.END, "   归档日志间隙检查: 正常 (无间隙)\n")
        elif "GAPS" in html_content:
            self.analysis_text.insert(tk.END, "   归档日志间隙检查: 异常 (存在间隙)\n")
        
        # 检查未应用的日志 - 提取详细信息
        if "未应用日志检查: 异常" in html_content:
            self.analysis_text.insert(tk.END, "   未应用日志检查: 异常 (存在未应用日志)\n")
            
            # 提取未应用日志数量
            unapplied_count_match = re.search(r'未应用日志数量: (\d+) 个日志文件', html_content)
            if unapplied_count_match:
                unapplied_count = unapplied_count_match.group(1)
                self.analysis_text.insert(tk.END, f"     - 未应用日志数量: {unapplied_count} 个日志文件\n")
                self.analysis_text.insert(tk.END, "     - 建议: 检查MRP进程状态，重启应用进程或运行start_standby.bat\n")
        elif "未应用日志检查: 正常" in html_content or ("NOT APPLIED" in html_content and re.search(r'"NOT APPLIED"[^0-9]*0', html_content)):
            self.analysis_text.insert(tk.END, "   未应用日志检查: 正常 (无未应用日志)\n")
        elif "NOT APPLIED" in html_content:
            self.analysis_text.insert(tk.END, "   未应用日志检查: 异常 (存在未应用日志)\n")
        
        # 检查表空间使用情况
        if "DANGER" in html_content:
            self.analysis_text.insert(tk.END, "   表空间使用检查: 危险 (有表空间使用率超过90%)\n")
        elif "WARNING" in html_content:
            self.analysis_text.insert(tk.END, "   表空间使用检查: 警告 (有表空间使用率超过80%)\n")
        else:
            self.analysis_text.insert(tk.END, "   表空间使用检查: 正常\n")
    
    def send_email_report_service_mode(self, analysis_result):
        """服务模式下发送邮件报告（不依赖GUI组件）"""
        try:
            # 获取邮件设置
            smtp_server = self.config_manager.get('Email', 'smtp_server')
            smtp_port = self.config_manager.getint('Email', 'smtp_port')
            smtp_username = self.config_manager.get('Email', 'smtp_username', fallback='')
            sender_email = self.config_manager.get('Email', 'sender_email')
            sender_password = self.config_manager.get('Email', 'sender_password')
            recipient_emails_str = self.config_manager.get('Email', 'recipient_emails')
            use_tls = self.config_manager.getboolean('Email', 'use_tls')
            
            # 检查必要的设置
            if not smtp_server or not sender_email or not recipient_emails_str:
                logger.error("邮件设置不完整，无法发送邮件")
                return
            
            # 解析收件人列表
            recipient_emails = [email.strip() for email in recipient_emails_str.split(',')]
            
            # 获取报告路径
            report_path = self.config_manager.get('Paths', 'report_path')
            html_report_exists = os.path.exists(report_path)
            if not html_report_exists:
                logger.warning(f"HTML报告文件不存在: {report_path}，将继续发送邮件但不包含HTML报告附件")
            
            # 检查HTML报告中是否有异常
            has_exception = False
            if html_report_exists:
                try:
                    with open(report_path, 'r', encoding='utf-8') as f:
                        html_content = f.read()
                        # 更精确地检查异常状态
                        if ("归档日志间隙检查: 异常" in html_content or 
                            "未应用日志检查: 异常" in html_content or
                            "DANGER" in html_content or
                            "ERROR" in html_content or
                            "Exception" in html_content):
                            has_exception = True
                except Exception as e:
                    logger.error(f"检查HTML报告异常状态时出错: {str(e)}")
            
            # 创建邮件
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = ", ".join(recipient_emails)
            
            # 从分析结果中提取告警级别
            alert_level = "✅ 正常"
            if "告警级别:" in analysis_result:
                # 提取告警级别信息
                lines = analysis_result.split('\n')
                for line in lines:
                    if "告警级别:" in line:
                        # 提取告警级别部分，例如："告警级别: 🔴 紧急 - 发现严重问题，需要立即处理"
                        alert_part = line.split("告警级别:")[1].strip()
                        if alert_part:
                            # 只取告警级别的前半部分，例如："🔴 紧急"
                            alert_level = alert_part.split(" - ")[0].strip()
                        break
            
            # 根据告警级别和异常状态设置邮件标题
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
            base_subject = f"[{alert_level}] {current_time} - Opera DataGuard状态监测报告"
            if has_exception:
                msg['Subject'] = base_subject
            else:
                msg['Subject'] = base_subject
            
            # 添加邮件正文（HTML格式）
            html_body = """
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="UTF-8">
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <title>Opera DataGuard状态监测报告</title>
                <style>
                    * {{
                        margin: 0;
                        padding: 0;
                        box-sizing: border-box;
                    }}
                    body {{
                        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                        line-height: 1.6;
                        color: #333;
                        background-color: #f5f5f5;
                        padding: 20px;
                    }}
                    .container {{
                        max-width: 1200px;
                        margin: 0 auto;
                        background: white;
                        border-radius: 10px;
                        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
                        overflow: hidden;
                    }}
                    .header {{
                        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                        color: white;
                        padding: 30px;
                        text-align: center;
                    }}
                    .header h1 {{
                        font-size: 28px;
                        margin-bottom: 10px;
                        font-weight: 300;
                    }}
                    .header .subtitle {{
                        font-size: 16px;
                        opacity: 0.9;
                    }}
                    .content {{
                        padding: 30px;
                    }}
                    .alert-badge {{
                        display: inline-block;
                        padding: 8px 16px;
                        border-radius: 20px;
                        font-weight: bold;
                        font-size: 14px;
                        margin-bottom: 20px;
                    }}
                    .alert-normal {{
                        background-color: #d4edda;
                        color: #155724;
                        border: 1px solid #c3e6cb;
                    }}
                    .alert-warning {{
                        background-color: #fff3cd;
                        color: #856404;
                        border: 1px solid #ffeaa7;
                    }}
                    .alert-critical {{
                        background-color: #f8d7da;
                        color: #721c24;
                        border: 1px solid #f5c6cb;
                    }}
                    .report-section {{
                        background: #f8f9fa;
                        border-left: 4px solid #667eea;
                        padding: 20px;
                        margin: 20px 0;
                        border-radius: 0 8px 8px 0;
                    }}
                    .report-content {{
                        font-family: 'Courier New', monospace;
                        font-size: 13px;
                        line-height: 1.5;
                        white-space: pre-wrap;
                        background: white;
                        padding: 15px;
                        border-radius: 5px;
                        border: 1px solid #e9ecef;
                        column-count: 2;
                        column-gap: 30px;
                        column-rule: 1px solid #e9ecef;
                    }}
                    .report-content-item {{
                        break-inside: avoid;
                        margin-bottom: 15px;
                        padding: 10px;
                        background: #f8f9fa;
                        border-radius: 4px;
                        border-left: 3px solid #667eea;
                    }}
                    .footer {{
                        background: #f8f9fa;
                        padding: 20px;
                        text-align: center;
                        border-top: 1px solid #e9ecef;
                        color: #6c757d;
                        font-size: 14px;
                    }}
                    .timestamp {{
                        color: #6c757d;
                        font-size: 14px;
                        margin-top: 10px;
                    }}
                    .highlight {{
                        background-color: #fff3cd;
                        padding: 2px 4px;
                        border-radius: 3px;
                    }}
                    @media (max-width: 600px) {{
                        .container {{
                            margin: 10px;
                            border-radius: 5px;
                        }}
                        .header, .content {{
                            padding: 20px;
                        }}
                        .header h1 {{
                            font-size: 24px;
                        }}
                        .report-content {{
                            column-count: 1;
                            column-gap: 0;
                            column-rule: none;
                        }}
                    }}
                </style>
            </head>
            <body>
                <div class="container">
                    <div class="header">
                        <h1>🛡️ Opera DataGuard</h1>
                        <div class="subtitle">数据库状态监测报告</div>
                        <div class="timestamp">{}</div>
                    </div>
                    <div class="content">
                        <div class="alert-badge {}">
                            {}
                        </div>
                        <div class="report-section">
                            <h3 style="margin-bottom: 15px; color: #495057;">📊 详细分析结果</h3>
                            <div class="report-content">{}</div>
                        </div>
                        <div style="margin-top: 20px; padding: 15px; background: #e3f2fd; border-radius: 5px; border-left: 4px solid #2196f3;">
                            <strong>📎 附件说明：</strong><br>
                            • HTML详细报告：包含完整的数据库状态信息<br>
                            • 检查输出文件：原始监控数据<br>
                            • 日报输出文件：每日统计数据
                        </div>
                    </div>
                    <div class="footer">
                        <p>此邮件由 Opera DataGuard 自动监控系统生成</p>
                        <p>如有问题，请联系数据库管理员</p>
                    </div>
                </div>
            </body>
            </html>
            """.format(
                datetime.datetime.now().strftime('%Y年%m月%d日 %H:%M'),
                'alert-normal' if '✅' in alert_level else ('alert-warning' if '🟡' in alert_level or '🟢' in alert_level else 'alert-critical'),
                alert_level,
                analysis_result.replace('\n', '\n').replace('<', '&lt;').replace('>', '&gt;')
            )
            
            msg.attach(MIMEText(html_body, 'html'))
            
            # 添加HTML报告附件（如果存在）
            if html_report_exists:
                with open(report_path, 'rb') as f:
                    attachment = MIMEApplication(f.read(), _subtype='html')
                    attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(report_path))
                    msg.attach(attachment)
            

            
            # 添加check_standby输出附件
            if hasattr(self, 'check_standby_output') and self.check_standby_output:
                check_standby_attachment = MIMEText(self.check_standby_output, 'plain')
                check_standby_attachment.add_header('Content-Disposition', 'attachment', filename='check_standby_output.txt')
                msg.attach(check_standby_attachment)
            
            # 添加daily_report输出附件
            if hasattr(self, 'daily_report_output') and self.daily_report_output:
                daily_report_attachment = MIMEText(self.daily_report_output, 'plain')
                daily_report_attachment.add_header('Content-Disposition', 'attachment', filename='daily_report_output.txt')
                msg.attach(daily_report_attachment)
            
            # 连接到SMTP服务器并发送邮件
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                if use_tls:
                    server.starttls()
                if sender_password:  # 只有在提供密码时才尝试登录
                    # 使用SMTP用户名进行身份验证，如果没有设置则使用发件人邮箱
                    login_username = smtp_username if smtp_username else sender_email
                    server.login(login_username, sender_password)
                server.send_message(msg)
            
            logger.info("邮件已成功发送")
        
        except Exception as e:
            error_msg = f"发送邮件时出错: {str(e)}"
            logger.error(error_msg, exc_info=True)
    
    def send_email_report(self):
        try:
            # 获取邮件设置
            smtp_server = self.config_manager.get('Email', 'smtp_server')
            smtp_port = self.config_manager.getint('Email', 'smtp_port')
            smtp_username = self.config_manager.get('Email', 'smtp_username', fallback='')
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
            
            # 检查HTML报告中是否有异常
            has_exception = False
            try:
                with open(report_path, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                    # 更精确地检查异常状态
                    if ("归档日志间隙检查: 异常" in html_content or 
                        "未应用日志检查: 异常" in html_content or
                        "DANGER" in html_content or
                        "ERROR" in html_content or
                        "Exception" in html_content):
                        has_exception = True
            except Exception as e:
                self.log_message(f"检查HTML报告异常状态时出错: {str(e)}")
            
            # 创建邮件
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = ", ".join(recipient_emails)
            
            # 从分析结果中提取告警级别
            alert_level = "✅ 正常"
            analysis_content = self.analysis_text.get(1.0, tk.END)
            if "告警级别:" in analysis_content:
                # 提取告警级别信息
                lines = analysis_content.split('\n')
                for line in lines:
                    if "告警级别:" in line:
                        # 提取告警级别部分，例如："告警级别: 🔴 紧急 - 发现严重问题，需要立即处理"
                        alert_part = line.split("告警级别:")[1].strip()
                        if alert_part:
                            # 只取告警级别的前半部分，例如："🔴 紧急"
                            alert_level = alert_part.split(" - ")[0].strip()
                        break
            
            # 根据告警级别和异常状态设置邮件标题
            current_time = datetime.datetime.now().strftime('%Y-%m-%d %H:%M')
            base_subject = f"[{alert_level}] {current_time} - Opera DataGuard状态监测报告"
            if has_exception:
                msg['Subject'] = base_subject
            else:
                msg['Subject'] = base_subject
            
            # 添加邮件正文
            body = "这是自动生成的Opera DataGuard状态监测报告，请查看附件。\n\n"
            
            # 添加分析结果
            body += "分析结果:\n" + self.analysis_text.get(1.0, tk.END)
            
            msg.attach(MIMEText(body, 'plain'))
            
            # 添加HTML报告附件
            with open(report_path, 'rb') as f:
                attachment = MIMEApplication(f.read(), _subtype='html')
                attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(report_path))
                msg.attach(attachment)
            

            
            # 添加check_standby输出附件
            if self.check_standby_output:
                check_standby_attachment = MIMEText(self.check_standby_output, 'plain')
                check_standby_attachment.add_header('Content-Disposition', 'attachment', filename='check_standby_output.txt')
                msg.attach(check_standby_attachment)
            
            # 添加daily_report输出附件
            if self.daily_report_output:
                daily_report_attachment = MIMEText(self.daily_report_output, 'plain')
                daily_report_attachment.add_header('Content-Disposition', 'attachment', filename='daily_report_output.txt')
                msg.attach(daily_report_attachment)
            
            # 连接到SMTP服务器并发送邮件
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                if use_tls:
                    server.starttls()
                if sender_password:  # 只有在提供密码时才尝试登录
                    # 使用SMTP用户名进行身份验证，如果没有设置则使用发件人邮箱
                    login_username = smtp_username if smtp_username else sender_email
                    server.login(login_username, sender_password)
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
        
        # 在UI中显示日志（仅在GUI模式下）
        if not self.service_mode and hasattr(self, 'log_text'):
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
        settings_window.geometry("500x400")
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
        
        # SMTP用户名
        ttk.Label(frame, text="SMTP用户名:").grid(row=2, column=0, sticky=tk.W, pady=5)
        smtp_username_var = tk.StringVar(value=self.config_manager.get('Email', 'smtp_username', fallback=''))
        ttk.Entry(frame, textvariable=smtp_username_var, width=40).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # 发件人邮箱
        ttk.Label(frame, text="发件人邮箱:").grid(row=3, column=0, sticky=tk.W, pady=5)
        sender_email_var = tk.StringVar(value=self.config_manager.get('Email', 'sender_email'))
        ttk.Entry(frame, textvariable=sender_email_var, width=40).grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # 发件人密码
        ttk.Label(frame, text="发件人密码:").grid(row=4, column=0, sticky=tk.W, pady=5)
        sender_password_var = tk.StringVar(value=self.config_manager.get('Email', 'sender_password'))
        ttk.Entry(frame, textvariable=sender_password_var, width=40, show="*").grid(row=4, column=1, sticky=tk.W, pady=5)
        
        # 收件人邮箱
        ttk.Label(frame, text="收件人邮箱:").grid(row=5, column=0, sticky=tk.W, pady=5)
        ttk.Label(frame, text="(多个邮箱用逗号分隔)").grid(row=6, column=0, sticky=tk.W)
        recipient_emails_var = tk.StringVar(value=self.config_manager.get('Email', 'recipient_emails'))
        ttk.Entry(frame, textvariable=recipient_emails_var, width=40).grid(row=5, column=1, rowspan=2, sticky=tk.W, pady=5)
        
        # 使用TLS
        use_tls_var = tk.BooleanVar(value=self.config_manager.getboolean('Email', 'use_tls'))
        ttk.Checkbutton(frame, text="使用TLS加密", variable=use_tls_var).grid(row=7, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # 自动发送邮件
        auto_send_var = tk.BooleanVar(value=self.config_manager.getboolean('Settings', 'auto_send_email', fallback=False))
        ttk.Checkbutton(frame, text="监控完成后自动发送邮件", variable=auto_send_var).grid(row=8, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # 测试按钮
        ttk.Button(frame, text="测试邮件设置", command=lambda: self.test_email_settings(
            smtp_server_var.get(),
            smtp_port_var.get(),
            smtp_username_var.get(),
            sender_email_var.get(),
            sender_password_var.get(),
            recipient_emails_var.get(),
            use_tls_var.get()
        )).grid(row=9, column=0, pady=10)
        
        # 保存按钮
        ttk.Button(frame, text="保存设置", command=lambda: self.save_email_settings(
            smtp_server_var.get(),
            smtp_port_var.get(),
            smtp_username_var.get(),
            sender_email_var.get(),
            sender_password_var.get(),
            recipient_emails_var.get(),
            use_tls_var.get(),
            auto_send_var.get(),
            settings_window
        )).grid(row=9, column=1, pady=10)
    
    def test_email_settings(self, smtp_server, smtp_port, smtp_username, sender_email, sender_password, recipient_emails, use_tls):
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
                    # 使用SMTP用户名进行身份验证，如果没有设置则使用发件人邮箱
                    login_username = smtp_username if smtp_username else sender_email
                    server.login(login_username, sender_password)
                server.send_message(msg)
            
            messagebox.showinfo("成功", "测试邮件已成功发送")
        
        except Exception as e:
            messagebox.showerror("错误", f"发送测试邮件时出错: {str(e)}")
    
    def save_email_settings(self, smtp_server, smtp_port, smtp_username, sender_email, sender_password, recipient_emails, use_tls, auto_send, window):
        try:
            # 保存设置
            self.config_manager.set('Email', 'smtp_server', smtp_server)
            self.config_manager.set('Email', 'smtp_port', smtp_port)
            self.config_manager.set('Email', 'smtp_username', smtp_username)
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
        settings_window.geometry("650x300")
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
        
        # 自动检测按钮
        ttk.Button(frame, text="自动检测路径", command=lambda: self.auto_detect_paths(
            check_standby_var, daily_report_var, report_path_var
        )).grid(row=3, column=0, columnspan=3, pady=10)
        
        # 按钮框架
        button_frame = ttk.Frame(frame)
        button_frame.grid(row=4, column=0, columnspan=3, pady=10)
        
        # 保存按钮
        ttk.Button(button_frame, text="保存设置", command=lambda: self.save_path_settings(
            check_standby_var.get(),
            daily_report_var.get(),
            report_path_var.get(),
            settings_window
        )).pack(side=tk.LEFT, padx=5)
        
        # 取消按钮
        ttk.Button(button_frame, text="取消", command=settings_window.destroy).pack(side=tk.LEFT, padx=5)
    
    def browse_file(self, var):
        file_path = filedialog.askopenfilename()
        if file_path:
            var.set(file_path)
    
    def auto_detect_paths(self, check_standby_var, daily_report_var, report_path_var):
        """自动检测当前目录下的相关文件路径"""
        try:
            # 获取当前脚本所在目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            
            # 检测check_standby.bat
            check_standby_path = os.path.join(current_dir, 'check_standby.bat')
            if os.path.exists(check_standby_path):
                check_standby_var.set(check_standby_path)
            
            # 检测daily_report.bat
            daily_report_path = os.path.join(current_dir, 'daily_report.bat')
            if os.path.exists(daily_report_path):
                daily_report_var.set(daily_report_path)
            
            # 检测HTML报告路径（优先检测logs目录下的文件）
            logs_dir = os.path.join(current_dir, 'logs')
            html_report_path = os.path.join(logs_dir, 'daily_report.html')
            
            # 如果logs目录不存在，创建它
            if not os.path.exists(logs_dir):
                os.makedirs(logs_dir)
            
            # 设置HTML报告路径
            report_path_var.set(html_report_path)
            
            # 显示检测结果
            detected_files = []
            if os.path.exists(check_standby_path):
                detected_files.append("✓ check_standby.bat")
            else:
                detected_files.append("✗ check_standby.bat (未找到)")
            
            if os.path.exists(daily_report_path):
                detected_files.append("✓ daily_report.bat")
            else:
                detected_files.append("✗ daily_report.bat (未找到)")
            
            detected_files.append(f"✓ HTML报告路径: {html_report_path}")
            
            message = "自动检测结果:\n\n" + "\n".join(detected_files)
            messagebox.showinfo("自动检测完成", message)
            
        except Exception as e:
            messagebox.showerror("错误", f"自动检测路径时出错: {str(e)}")
    
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
        settings_window.geometry("500x400")
        settings_window.transient(self.root)
        settings_window.grab_set()
        
        frame = ttk.Frame(settings_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # GUI模式设置
        gui_frame = ttk.LabelFrame(frame, text="GUI模式设置", padding="10")
        gui_frame.grid(row=0, column=0, columnspan=2, sticky=tk.EW, pady=5)
        
        # 自动运行间隔
        ttk.Label(gui_frame, text="自动运行间隔 (小时):").grid(row=0, column=0, sticky=tk.W, pady=5)
        interval_hours = self.config_manager.getint('Settings', 'auto_run_interval', fallback=86400) / 3600
        interval_var = tk.DoubleVar(value=interval_hours)
        ttk.Spinbox(gui_frame, from_=1, to=48, increment=0.5, textvariable=interval_var, width=10).grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # 服务模式设置
        service_frame = ttk.LabelFrame(frame, text="服务模式设置", padding="10")
        service_frame.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=5)
        
        # 运行时间设置
        ttk.Label(service_frame, text="每日运行时间:").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        time_frame = ttk.Frame(service_frame)
        time_frame.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # 小时
        run_hour = self.config_manager.getint('Settings', 'run_hour', fallback=8)
        hour_var = tk.IntVar(value=run_hour)
        ttk.Spinbox(time_frame, from_=0, to=23, textvariable=hour_var, width=5).pack(side=tk.LEFT)
        ttk.Label(time_frame, text="时").pack(side=tk.LEFT, padx=2)
        
        # 分钟
        run_minute = self.config_manager.getint('Settings', 'run_minute', fallback=0)
        minute_var = tk.IntVar(value=run_minute)
        ttk.Spinbox(time_frame, from_=0, to=59, textvariable=minute_var, width=5).pack(side=tk.LEFT)
        ttk.Label(time_frame, text="分").pack(side=tk.LEFT, padx=2)
        
        # 通用设置
        common_frame = ttk.LabelFrame(frame, text="通用设置", padding="10")
        common_frame.grid(row=2, column=0, columnspan=2, sticky=tk.EW, pady=5)
        
        # 启用错误检查
        check_errors_var = tk.BooleanVar(value=self.config_manager.getboolean('Settings', 'check_errors', fallback=True))
        ttk.Checkbutton(common_frame, text="启用错误检查", variable=check_errors_var).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # 错误模式
        ttk.Label(common_frame, text="错误模式:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Label(common_frame, text="(用逗号分隔)").grid(row=2, column=0, sticky=tk.W)
        error_patterns_var = tk.StringVar(value=self.config_manager.get('Settings', 'error_patterns', fallback='error,warning,danger,failed,ORA-,TNS-'))
        ttk.Entry(common_frame, textvariable=error_patterns_var, width=40).grid(row=1, column=1, rowspan=2, sticky=tk.W, pady=5)
        
        # 保存按钮
        ttk.Button(frame, text="保存设置", command=lambda: self.save_monitor_settings(
            interval_var.get(),
            hour_var.get(),
            minute_var.get(),
            check_errors_var.get(),
            error_patterns_var.get(),
            settings_window
        )).grid(row=3, column=0, columnspan=2, pady=10)
    
    def save_monitor_settings(self, interval_hours, run_hour, run_minute, check_errors, error_patterns, window):
        try:
            # 将小时转换为秒
            interval_seconds = int(interval_hours * 3600)
            
            # 保存设置
            self.config_manager.set('Settings', 'auto_run_interval', str(interval_seconds))
            self.config_manager.set('Settings', 'run_hour', str(run_hour))
            self.config_manager.set('Settings', 'run_minute', str(run_minute))
            self.config_manager.set('Settings', 'check_errors', str(check_errors))
            self.config_manager.set('Settings', 'error_patterns', error_patterns)
            
            messagebox.showinfo("成功", "监控设置已保存")
            window.destroy()
        
        except Exception as e:
            messagebox.showerror("错误", f"保存设置时出错: {str(e)}")
    
    def show_about(self):
        about_text = """Opera DataGuard状态监测工具

版本: 1.0

这是一个用于监控Opera DataGuard环境的工具，可以自动执行检查脚本、生成报告并通过邮件通知管理员。

© 2023 All Rights Reserved"""
        messagebox.showinfo("关于", about_text)
    
    def open_service_management(self):
        """打开服务管理窗口"""
        service_window = tk.Toplevel(self.root)
        service_window.title("Windows服务管理")
        service_window.geometry("600x500")
        service_window.transient(self.root)
        service_window.grab_set()
        
        frame = ttk.Frame(service_window, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # 服务状态显示
        status_frame = ttk.LabelFrame(frame, text="服务状态", padding="10")
        status_frame.pack(fill=tk.X, pady=5)
        
        self.service_status_var = tk.StringVar(value="检查中...")
        ttk.Label(status_frame, textvariable=self.service_status_var, font=("Arial", 10, "bold")).pack()
        
        # 服务信息
        info_frame = ttk.LabelFrame(frame, text="服务信息", padding="10")
        info_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(info_frame, text="服务名称: OperaDataGuardService").pack(anchor=tk.W)
        ttk.Label(info_frame, text="显示名称: Opera DataGuard监控服务").pack(anchor=tk.W)
        ttk.Label(info_frame, text="描述: Oracle数据库状态监控服务").pack(anchor=tk.W)
        
        # 按钮框架
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=10)
        
        # 服务管理按钮
        self.install_service_btn = ttk.Button(button_frame, text="安装服务", command=self.install_service)
        self.install_service_btn.pack(side=tk.LEFT, padx=5)
        
        self.uninstall_service_btn = ttk.Button(button_frame, text="卸载服务", command=self.uninstall_service)
        self.uninstall_service_btn.pack(side=tk.LEFT, padx=5)
        
        self.start_service_btn = ttk.Button(button_frame, text="启动服务", command=self.start_service)
        self.start_service_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_service_btn = ttk.Button(button_frame, text="停止服务", command=self.stop_service)
        self.stop_service_btn.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="刷新状态", command=self.refresh_service_status).pack(side=tk.LEFT, padx=5)
        
        # 日志显示
        log_frame = ttk.LabelFrame(frame, text="操作日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        self.service_log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15)
        self.service_log_text.pack(fill=tk.BOTH, expand=True)
        
        # 说明文本
        help_frame = ttk.LabelFrame(frame, text="使用说明", padding="10")
        help_frame.pack(fill=tk.X, pady=5)
        
        help_text = """1. 安装服务：将Opera DataGuard注册为Windows服务
2. 启动服务：启动后台监控服务，无需用户登录即可运行
3. 停止服务：停止后台监控服务
4. 卸载服务：从系统中移除服务注册

注意：服务操作需要管理员权限"""
        ttk.Label(help_frame, text=help_text, justify=tk.LEFT).pack(anchor=tk.W)
        
        # 初始化服务状态
        self.refresh_service_status()
    
    def log_service_message(self, message):
        """在服务管理窗口中记录日志"""
        if hasattr(self, 'service_log_text'):
            timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.service_log_text.insert(tk.END, f"[{timestamp}] {message}\n")
            self.service_log_text.see(tk.END)
            self.service_log_text.update_idletasks()
    
    def get_service_status(self):
        """获取服务状态"""
        try:
            app_dir = self.config_manager.get_app_dir()
            service_exe = os.path.join(app_dir, "OperaDataGuardService.exe")
            
            if not os.path.exists(service_exe):
                return "服务文件不存在", "error"
            
            # 使用sc命令查询服务状态
            result = subprocess.run(
                ["sc", "query", "OperaDataGuardService"],
                capture_output=True,
                text=True,
                encoding='gbk'
            )
            
            if result.returncode == 0:
                output = result.stdout
                if "RUNNING" in output:
                    return "服务正在运行", "running"
                elif "STOPPED" in output:
                    return "服务已停止", "stopped"
                elif "START_PENDING" in output:
                    return "服务正在启动", "starting"
                elif "STOP_PENDING" in output:
                    return "服务正在停止", "stopping"
                else:
                    return "服务状态未知", "unknown"
            else:
                return "服务未安装", "not_installed"
                
        except Exception as e:
            return f"检查服务状态失败: {str(e)}", "error"
    
    def refresh_service_status(self):
        """刷新服务状态显示"""
        status_text, status_code = self.get_service_status()
        self.service_status_var.set(status_text)
        
        # 根据状态启用/禁用按钮
        if status_code == "not_installed":
            self.install_service_btn.config(state=tk.NORMAL)
            self.uninstall_service_btn.config(state=tk.DISABLED)
            self.start_service_btn.config(state=tk.DISABLED)
            self.stop_service_btn.config(state=tk.DISABLED)
        elif status_code == "stopped":
            self.install_service_btn.config(state=tk.DISABLED)
            self.uninstall_service_btn.config(state=tk.NORMAL)
            self.start_service_btn.config(state=tk.NORMAL)
            self.stop_service_btn.config(state=tk.DISABLED)
        elif status_code == "running":
            self.install_service_btn.config(state=tk.DISABLED)
            self.uninstall_service_btn.config(state=tk.DISABLED)
            self.start_service_btn.config(state=tk.DISABLED)
            self.stop_service_btn.config(state=tk.NORMAL)
        elif status_code == "error":
            self.install_service_btn.config(state=tk.NORMAL)
            self.uninstall_service_btn.config(state=tk.NORMAL)
            self.start_service_btn.config(state=tk.DISABLED)
            self.stop_service_btn.config(state=tk.DISABLED)
        else:
            # starting, stopping, unknown状态
            self.install_service_btn.config(state=tk.DISABLED)
            self.uninstall_service_btn.config(state=tk.DISABLED)
            self.start_service_btn.config(state=tk.DISABLED)
            self.stop_service_btn.config(state=tk.DISABLED)
        
        self.log_service_message(f"服务状态: {status_text}")
    
    def install_service(self):
        """安装Windows服务"""
        try:
            self.log_service_message("开始安装服务...")
            
            app_dir = self.config_manager.get_app_dir()
            service_exe = os.path.join(app_dir, "OperaDataGuardService.exe")
            
            if not os.path.exists(service_exe):
                self.log_service_message(f"错误: 服务文件不存在 - {service_exe}")
                messagebox.showerror("错误", f"服务文件不存在:\n{service_exe}\n\n请确保已正确构建服务版本")
                return
            
            # 使用服务可执行文件安装服务
            result = subprocess.run(
                [service_exe, "install"],
                capture_output=True,
                text=True,
                cwd=app_dir
            )
            
            if result.returncode == 0:
                self.log_service_message("服务安装成功")
                messagebox.showinfo("成功", "服务安装成功")
            else:
                error_msg = result.stderr or result.stdout or "未知错误"
                self.log_service_message(f"服务安装失败: {error_msg}")
                messagebox.showerror("错误", f"服务安装失败:\n{error_msg}")
            
            self.refresh_service_status()
            
        except Exception as e:
            error_msg = f"安装服务时出错: {str(e)}"
            self.log_service_message(error_msg)
            messagebox.showerror("错误", error_msg)
    
    def uninstall_service(self):
        """卸载Windows服务"""
        try:
            if not messagebox.askyesno("确认", "确定要卸载Opera DataGuard服务吗？"):
                return
            
            self.log_service_message("开始卸载服务...")
            
            app_dir = self.config_manager.get_app_dir()
            service_exe = os.path.join(app_dir, "OperaDataGuardService.exe")
            
            if os.path.exists(service_exe):
                # 使用服务可执行文件卸载服务
                result = subprocess.run(
                    [service_exe, "remove"],
                    capture_output=True,
                    text=True,
                    cwd=app_dir
                )
            else:
                # 如果服务文件不存在，尝试使用sc命令删除
                result = subprocess.run(
                    ["sc", "delete", "OperaDataGuardService"],
                    capture_output=True,
                    text=True
                )
            
            if result.returncode == 0:
                self.log_service_message("服务卸载成功")
                messagebox.showinfo("成功", "服务卸载成功")
            else:
                error_msg = result.stderr or result.stdout or "未知错误"
                self.log_service_message(f"服务卸载失败: {error_msg}")
                messagebox.showerror("错误", f"服务卸载失败:\n{error_msg}")
            
            self.refresh_service_status()
            
        except Exception as e:
            error_msg = f"卸载服务时出错: {str(e)}"
            self.log_service_message(error_msg)
            messagebox.showerror("错误", error_msg)
    
    def start_service(self):
        """启动Windows服务"""
        try:
            self.log_service_message("开始启动服务...")
            
            result = subprocess.run(
                ["sc", "start", "OperaDataGuardService"],
                capture_output=True,
                text=True,
                encoding='gbk'
            )
            
            if result.returncode == 0:
                self.log_service_message("服务启动成功")
                messagebox.showinfo("成功", "服务启动成功")
            else:
                error_msg = result.stderr or result.stdout or "未知错误"
                self.log_service_message(f"服务启动失败: {error_msg}")
                messagebox.showerror("错误", f"服务启动失败:\n{error_msg}")
            
            # 等待一下再刷新状态
            self.root.after(2000, self.refresh_service_status)
            
        except Exception as e:
            error_msg = f"启动服务时出错: {str(e)}"
            self.log_service_message(error_msg)
            messagebox.showerror("错误", error_msg)
    
    def stop_service(self):
        """停止Windows服务"""
        try:
            if not messagebox.askyesno("确认", "确定要停止Opera DataGuard服务吗？"):
                return
            
            self.log_service_message("开始停止服务...")
            
            result = subprocess.run(
                ["sc", "stop", "OperaDataGuardService"],
                capture_output=True,
                text=True,
                encoding='gbk'
            )
            
            if result.returncode == 0:
                self.log_service_message("服务停止成功")
                messagebox.showinfo("成功", "服务停止成功")
            else:
                error_msg = result.stderr or result.stdout or "未知错误"
                self.log_service_message(f"服务停止失败: {error_msg}")
                messagebox.showerror("错误", f"服务停止失败:\n{error_msg}")
            
            # 等待一下再刷新状态
            self.root.after(2000, self.refresh_service_status)
            
        except Exception as e:
            error_msg = f"停止服务时出错: {str(e)}"
            self.log_service_message(error_msg)
            messagebox.showerror("错误", error_msg)
    
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