#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Opera DataGuard 计划任务版本

该脚本用于Windows计划任务，直接执行监控检查、分析和邮件发送功能。
不包含GUI界面和服务管理功能，适合通过计划任务定时执行。
"""

import os
import sys
import logging
import time
from pathlib import Path

# 添加当前目录到Python路径
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

# 配置日志 - 必须在导入opera_monitor之前配置
def setup_logging():
    """设置日志配置"""
    try:
        # 获取应用程序目录
        if getattr(sys, 'frozen', False):
            app_dir = os.path.dirname(sys.executable)
        else:
            app_dir = os.path.dirname(os.path.abspath(__file__))
        
        log_dir = os.path.join(app_dir, 'logs')
        os.makedirs(log_dir, exist_ok=True)
        
        log_file = os.path.join(log_dir, 'opera_task.log')
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler()
            ]
        )
        
        logger = logging.getLogger('OperaTask')
        logger.info(f"日志初始化完成，日志文件: {log_file}")
        return logger
        
    except Exception as e:
        # 如果日志设置失败，使用基本配置
        logging.basicConfig(level=logging.INFO)
        logger = logging.getLogger('OperaTask')
        logger.error(f"日志设置失败: {e}")
        return logger

def main():
    """主函数 - 执行完整的监控任务"""
    logger = setup_logging()
    
    try:
        logger.info("="*50)
        logger.info("开始执行Opera DataGuard监控任务")
        logger.info("="*50)
        
        # 导入opera_monitor（在日志配置之后）
        from opera_monitor import OperaMonitor
        
        # 创建监控器实例（服务模式，无GUI）
        logger.info("初始化监控器...")
        monitor = OperaMonitor(root=None)  # root=None 表示服务模式
        
        # 检查配置
        logger.info("检查配置信息...")
        
        # 获取应用程序目录
        if getattr(sys, 'frozen', False):
            app_dir = os.path.dirname(sys.executable)
        else:
            app_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 自动检测批处理文件路径
        check_standby_bat = os.path.join(app_dir, 'check_standby.bat')
        daily_report_bat = os.path.join(app_dir, 'daily_report.bat')
        
        if not os.path.exists(check_standby_bat):
            logger.error(f"批处理文件不存在: {check_standby_bat}")
            return 1
        
        if not os.path.exists(daily_report_bat):
            logger.error(f"批处理文件不存在: {daily_report_bat}")
            return 1
        
        logger.info(f"check_standby.bat: {check_standby_bat}")
        logger.info(f"daily_report.bat: {daily_report_bat}")
        
        # 检查邮件配置
        smtp_server = monitor.config_manager.get('Email', 'smtp_server', fallback='')
        smtp_username = monitor.config_manager.get('Email', 'smtp_username', fallback='')
        sender_email = monitor.config_manager.get('Email', 'sender_email', fallback='')
        recipient_emails = monitor.config_manager.get('Email', 'recipient_emails', fallback='')
        
        if not smtp_server or smtp_server == 'smtp.example.com':
            logger.warning("邮件配置未设置，将跳过邮件发送")
            send_email = False
        else:
            logger.info(f"邮件服务器: {smtp_server}")
            if smtp_username:
                logger.info(f"SMTP用户名: {smtp_username}")
            logger.info(f"发送者: {sender_email}")
            logger.info(f"接收者: {recipient_emails}")
            send_email = True
        
        # 执行监控任务
        logger.info("开始执行监控检查...")
        
        # 运行check_standby.bat
        logger.info("执行 check_standby.bat...")
        check_standby_output = monitor.run_batch_file(check_standby_bat)
        monitor.check_standby_output = check_standby_output  # 保存到monitor对象属性
        logger.info("check_standby.bat 执行完成")
        
        # 运行daily_report.bat
        logger.info("执行 daily_report.bat...")
        daily_report_output = monitor.run_batch_file(daily_report_bat)
        monitor.daily_report_output = daily_report_output  # 保存到monitor对象属性
        logger.info("daily_report.bat 执行完成")
        
        # 分析结果
        logger.info("分析执行结果...")
        analysis_result = monitor.analyze_results_service_mode(check_standby_output, daily_report_output)
        logger.info("结果分析完成")
        
        # 输出分析结果到日志
        if analysis_result:
            logger.info("分析结果:")
            for line in analysis_result.split('\n'):
                if line.strip():
                    logger.info(f"  {line}")
        
        # 发送邮件报告
        if send_email:
            logger.info("发送邮件报告...")
            try:
                # 使用服务模式的邮件发送方法
                monitor.send_email_report_service_mode(analysis_result)
                logger.info("邮件报告发送成功")
            except Exception as e:
                logger.error(f"邮件发送失败: {e}")
                return 1
        else:
            logger.info("跳过邮件发送（配置未设置）")
        
        logger.info("="*50)
        logger.info("Opera DataGuard监控任务执行完成")
        logger.info("="*50)
        
        return 0
        
    except Exception as e:
        logger.error(f"执行监控任务时出错: {e}")
        logger.exception("详细错误信息:")
        return 1

if __name__ == '__main__':
    exit_code = main()
    sys.exit(exit_code)