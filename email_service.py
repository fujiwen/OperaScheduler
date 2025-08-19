import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os
from datetime import datetime
import logging
import socket
from typing import List, Optional, Callable, Dict, Union
from pathlib import Path

class EmailService:
    """邮件服务类，用于处理供应商对账确认函的邮件发送。
    
    属性:
        config (dict): 邮件配置信息
        skipped_vendors (list): 跳过的供应商列表
        ui_callback (callable): UI回调函数
        logger (logging.Logger): 日志记录器
    """
    
    # 默认SMTP服务器配置
    DEFAULT_SMTP_HOST = "smtp-sin02.aa.accor.net"
    DEFAULT_SMTP_PORT = 587
    
    # 配置文件相关
    CONFIG_FILE = 'email.ini'
    REQUIRED_CONFIGS = ['smtp_username', 'smtp_password']
    
    # 文件命名相关
    CONFIRMATION_KEYWORD = '确认函'
    DATE_SEPARATOR = '-'
    def __init__(self, ui_callback: Optional[Callable[[str, str], None]] = None):
        # 配置日志记录器
        self.logger = logging.getLogger(__name__)
        self.logger.setLevel(logging.INFO)
        
        # 添加控制台处理器
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # 设置日志格式
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        console_handler.setFormatter(formatter)
        
        # 添加处理器到日志记录器
        if not self.logger.handlers:
            self.logger.addHandler(console_handler)
        
        self.config = self.load_config()
        self.skipped_vendors = []
        self.ui_callback = ui_callback  # UI回调函数，用于在界面显示消息
        
        # 设置SMTP配置（优先使用ini文件配置，否则使用默认值）
        self.smtp_host = self.config.get('smtp_host', self.DEFAULT_SMTP_HOST)
        self.smtp_port = int(self.config.get('smtp_port', self.DEFAULT_SMTP_PORT))
        
        # 智能判断加密方式：优先使用配置，否则根据端口号自动判断
        if 'smtp_encryption' in self.config:
            self.smtp_encryption = self.config['smtp_encryption'].lower()
        else:
            self.smtp_encryption = self._auto_detect_encryption(self.smtp_port)
        
        # 提示使用的SMTP服务器来源
        if 'smtp_host' in self.config and 'smtp_port' in self.config:
            self.log_message("✓ 使用配置文件中的SMTP服务器", "info")
        else:
            self.log_message("✓ 使用内置默认SMTP服务器", "info")
        
    def load_config(self) -> Dict[str, str]:
        """从配置文件加载邮件服务配置。

        Returns:
            Dict[str, str]: 配置键值对

        Raises:
            FileNotFoundError: 配置文件不存在
            ValueError: 配置格式错误或缺少必需配置项
        """
        config_path = Path(self.CONFIG_FILE)
        if not config_path.exists():
            raise FileNotFoundError(f"配置文件不存在: {config_path}")

        config: Dict[str, str] = {}
        try:
            with config_path.open('r', encoding='utf-8') as f:
                for line_num, line in enumerate(f, 1):
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue
                        
                    if ':' not in line:
                        self.logger.warning(f"第{line_num}行格式不正确: {line}")
                        continue
                        
                    key, value = map(str.strip, line.split(':', 1))
                    if not key or not value:
                        self.logger.warning(f"第{line_num}行键值对不完整: {line}")
                        continue
                        
                    config[key] = value

            # 验证必需配置项
            missing_configs = [cfg for cfg in self.REQUIRED_CONFIGS if cfg not in config]
            if missing_configs:
                raise ValueError(f"缺少必需的配置项: {', '.join(missing_configs)}")

            return config

        except Exception as e:
            if isinstance(e, ValueError):
                raise
            raise ValueError(f"加载配置文件失败: {e}")
    
    def _auto_detect_encryption(self, port: int) -> str:
        """根据端口号自动检测加密方式。

        Args:
            port: SMTP端口号

        Returns:
            str: 加密方式 ('ssl', 'starttls', 'none')
        """
        # 处理空端口号或无效端口号的情况
        if not port or port <= 0:
            return 'starttls'
        
        # 常见端口号对应的加密方式
        port_encryption_map = {
            25: 'none',      # 标准SMTP端口，通常无加密
            587: 'starttls', # 提交端口，通常使用STARTTLS
            465: 'ssl',      # SMTPS端口，使用SSL/TLS加密
            2525: 'starttls' # 备用端口，通常使用STARTTLS
        }
        
        return port_encryption_map.get(port, 'starttls')  # 默认使用STARTTLS
    
    def _create_smtp_connection(self, timeout: int = 30):
        """创建SMTP连接。

        Args:
            timeout: 连接超时时间（秒）

        Returns:
            SMTP连接对象
        """
        if self.smtp_encryption == 'ssl':
            return smtplib.SMTP_SSL(self.smtp_host, self.smtp_port, timeout=timeout)
        else:
            smtp = smtplib.SMTP(self.smtp_host, self.smtp_port, timeout=timeout)
            if self.smtp_encryption == 'starttls':
                smtp.starttls()
            return smtp
    
    def _get_vendor_email(self, vendor_name: str) -> str:
        """获取供应商邮箱地址。

        Args:
            vendor_name: 供应商名称

        Returns:
            str: 供应商邮箱地址

        Raises:
            ValueError: 未找到供应商邮箱配置
        """
        # 标准化供应商名称（移除下划线和空格）
        normalized_name = vendor_name.replace('_', '').strip()
        
        # 在配置中查找匹配的供应商
        for key, value in self.config.items():
            if key.replace('_', '').strip() == normalized_name:
                self.logger.debug(f"找到供应商 {vendor_name} 的邮箱配置")
                return value
        
        error_msg = f"未找到供应商 {vendor_name} 的邮箱配置"
        self.logger.error(error_msg)
        raise ValueError(error_msg)
    
    def test_smtp_connection(self) -> bool:
        """测试SMTP服务器连接。

        Returns:
            bool: 连接测试是否成功
        """
        self.log_message("正在测试SMTP服务器连接...", "info")
        self.logger.info("开始SMTP连接测试")
        
        try:
            # 创建SMTP连接
            smtp = self._create_smtp_connection(timeout=30)
            self.log_message("✓ 成功连接到SMTP服务器", "success")
            
            # 验证用户凭据
            self.log_message("正在验证用户凭据...", "info")
            smtp.login(self.config['smtp_username'], self.config['smtp_password'])
            self.log_message("✓ 用户凭据验证成功", "success")
            self.logger.info("用户凭据验证成功")
            
            # 关闭连接
            smtp.quit()
            self.log_message("✓ SMTP连接测试完成，所有检查通过", "success")
            self.logger.info("SMTP连接测试完成")
            
            return True
            
        except smtplib.SMTPAuthenticationError as e:
            error_msg = f"用户凭据验证失败: {e}"
            self.log_message(f"✗ {error_msg}", "error")
            self.log_message("请检查配置文件中的smtp_username和smtp_password", "error")
            self.logger.error(error_msg)
            return False
            
        except smtplib.SMTPConnectError as e:
            error_msg = f"无法连接到SMTP服务器: {e}"
            self.log_message(f"✗ {error_msg}", "error")
            self.log_message("请检查网络连接或SMTP服务器配置", "error")
            self.logger.error(error_msg)
            return False
            
        except socket.timeout as e:
            error_msg = f"连接超时: 无法在30秒内连接到SMTP服务器"
            self.log_message(f"✗ {error_msg}", "error")
            self.log_message("请检查网络连接和防火墙设置", "error")
            self.logger.error(error_msg)
            return False
            
        except socket.gaierror as e:
            error_msg = f"DNS解析失败: 无法解析服务器地址"
            self.log_message(f"✗ {error_msg}", "error")
            self.log_message("请检查服务器地址是否正确", "error")
            self.logger.error(error_msg)
            return False
            
        except ConnectionRefusedError as e:
            error_msg = f"连接被拒绝: 服务器拒绝连接"
            self.log_message(f"✗ {error_msg}", "error")
            self.log_message("请检查端口号是否正确或服务器是否可用", "error")
            self.logger.error(error_msg)
            return False
            
        except smtplib.SMTPException as e:
            error_msg = f"SMTP错误: {e}"
            self.log_message(f"✗ {error_msg}", "error")
            self.logger.error(error_msg)
            return False
            
        except Exception as e:
            error_msg = f"连接测试失败: {e}"
            self.log_message(f"✗ {error_msg}", "error")
            self.logger.error(f"意外错误: {e}", exc_info=True)
            return False

    def log_message(self, message: str, tag: str = "info") -> None:
        """记录日志信息并通过UI回调显示。

        Args:
            message: 日志消息
            tag: 消息标签，用于UI显示样式
        """
        # 根据tag选择合适的日志级别
        log_levels = {
            "info": self.logger.info,
            "error": self.logger.error,
            "success": self.logger.info,
            "header": self.logger.info,
            "warning": self.logger.warning
        }
        log_levels.get(tag, self.logger.info)(message)
        
        # 如果有UI回调函数，发送到UI
        if self.ui_callback:
            self.ui_callback(message, tag)

    def _extract_vendor_name(self, filename: str) -> Optional[str]:
        """从文件名中提取供应商名称。

        Args:
            filename: 文件名

        Returns:
            Optional[str]: 供应商名称，如果无法提取则返回None
        """
        if self.CONFIRMATION_KEYWORD not in filename:
            return None

        parts = filename.split(self.DATE_SEPARATOR)
        if len(parts) < 3:
            return None

        vendor_name = parts[2].strip()
        if '%' in vendor_name:
            vendor_name = vendor_name.split('%')[0].strip()
        return vendor_name

    def _print_summary(self, total_vendors: int, total_files: int) -> None:
        """打印处理完成的汇总信息。

        Args:
            total_vendors: 总供应商数
            total_files: 总文件数
        """
        self.log_message('\n' + '='*50, "header")
        self.log_message('处理完成汇总：', "header")
        self.log_message(f'- 总供应商数：{total_vendors}', "info")
        self.log_message(f'- 总文件数：{total_files}', "info")
        self.log_message(f'- 成功处理供应商数：{total_vendors - len(self.skipped_vendors)}', "success")
        self.log_message(f'- 跳过的供应商数：{len(self.skipped_vendors)}', "error")
        if self.skipped_vendors:
            self.log_message(f'- 跳过的供应商列表：{", ".join(self.skipped_vendors)}', "error")
        self.log_message('='*50, "header")
    def process_folder(self, folder_path: str, progress_callback: Optional[Callable[[int, int, str], None]] = None) -> None:
        """处理文件夹内的所有确认函文件。

        Args:
            folder_path: 文件夹路径
            progress_callback: 进度回调函数

        Raises:
            FileNotFoundError: 文件夹不存在
            ValueError: 处理过程中的错误
        """
        folder = Path(folder_path)
        if not folder.exists():
            raise FileNotFoundError(f"文件夹不存在：{folder}")

        # 获取所有Excel文件
        excel_files = list(folder.glob("*.xlsx"))
        confirmation_files = [f for f in excel_files if self.CONFIRMATION_KEYWORD in f.name]
        total_files = len(confirmation_files)
        self.log_message(f"找到{total_files}个确认函文件", "info")

        # 按供应商分组文件
        vendor_files: Dict[str, List[Path]] = {}
        for file_path in confirmation_files:
            vendor_name = self._extract_vendor_name(file_path.name)
            if vendor_name:
                if vendor_name not in vendor_files:
                    vendor_files[vendor_name] = []
                vendor_files[vendor_name].append(file_path)
            else:
                self.log_message(f"无法从文件名解析供应商信息：{file_path.name}", "error")

        # 处理每个供应商的所有文件
        total_vendors = len(vendor_files)
        self.log_message(f'开始处理{total_vendors}个供应商的邮件发送', "header")
        
        for i, (vendor_name, file_paths) in enumerate(vendor_files.items(), 1):
            try:
                self.log_message(f'正在处理供应商：{vendor_name}（{len(file_paths)}个文件）', "info")
                try:
                    self.send_reconciliation_email([str(p) for p in file_paths], vendor_name)
                    self.log_message(f"✓ 成功发送邮件给供应商：{vendor_name}", "success")
                except ValueError as e:
                    if "未找到供应商" in str(e):
                        self.skipped_vendors.append(vendor_name)
                        self.log_message(f"✗ 跳过供应商（未找到邮箱配置）：{vendor_name}", "error")
                    else:
                        self.log_message(f"✗ 发送失败（{vendor_name}）：{e}", "error")
                except Exception as e:
                    self.log_message(f"✗ 发送失败（{vendor_name}）：{e}", "error")
                    self.logger.error(f"处理供应商 {vendor_name} 时出错", exc_info=True)

            except Exception as e:
                self.log_message(f"处理供应商时出错（{vendor_name}）：{e}", "error")
                self.logger.error(f"处理供应商 {vendor_name} 时发生意外错误", exc_info=True)

            # 更新进度
            if progress_callback:
                progress_callback(i, total_vendors, f"正在处理第 {i}/{total_vendors} 个供应商")

        # 处理完成后的汇总信息
        self._print_summary(total_vendors, total_files)

    def _extract_year_month(self, file_path: str) -> Optional[str]:
        """从文件路径中提取年月信息。

        支持两种格式：
        1. 路径格式：export/2025-07/Confirmed/xxx
        2. 文件名格式：年-月-供应商名称[-税率%]-确认函.xlsx

        Args:
            file_path: 文件路径

        Returns:
            Optional[str]: 年月信息（格式：YYYY-MM），如果无法提取则返回None
        """
        # 尝试从路径中提取
        path_parts = file_path.split(os.sep)
        for part in path_parts:
            if len(part) == 7 and part[4] == '-':
                return part

        # 尝试从文件名中提取
        filename = os.path.basename(file_path)
        if self.CONFIRMATION_KEYWORD not in filename:
            return None

        parts = filename.split(self.DATE_SEPARATOR)
        if len(parts) >= 3:
            year = parts[0].strip()
            month = parts[1].strip()
            if year.isdigit() and month.isdigit() and len(year) == 4 and len(month) == 2:
                return f"{year}-{month}"

        return None

    def send_reconciliation_email(self, file_paths: Union[str, List[str]], vendor_name: str) -> bool:
        """发送对账确认函邮件。

        Args:
            file_paths: 单个文件路径或文件路径列表
            vendor_name: 供应商名称

        Returns:
            bool: 发送是否成功

        Raises:
            ValueError: 供应商邮箱未配置
            FileNotFoundError: 附件文件不存在
            Exception: 其他邮件发送错误
        """
        file_paths = [file_paths] if isinstance(file_paths, str) else file_paths
        try:
            # 从文件路径中提取年月
            year_month = self._extract_year_month(file_paths[0])
            if not year_month:
                self.logger.warning(f"无法从文件路径提取年月信息：{file_paths[0]}")

            # 验证所有文件是否存在
            for file_path in file_paths:
                if not os.path.exists(file_path):
                    raise FileNotFoundError(f"附件文件不存在：{file_path}")

            # 获取供应商邮箱
            recipient_email = self._get_vendor_email(vendor_name)
            
            # 构建邮件
            msg = MIMEMultipart()
            
            # 设置发件人（优先使用配置的sender_email，否则使用smtp_username）
            sender_email = self.config.get('sender_email', '').strip() or self.config['smtp_username']
            msg['From'] = sender_email
            msg['To'] = recipient_email
            
            # 设置主题
            subject = self.config.get('email_subject', '对账确认函')
            if year_month:
                subject = f'{year_month}月{subject}'
            msg['Subject'] = subject
            
            # 添加正文（处理换行符）
            body = self.config.get('email_body', '').replace('\\n', '\n')
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
            
            # 添加附件
            for file_path in file_paths:
                with open(file_path, 'rb') as f:
                    part = MIMEApplication(f.read())
                    part.add_header('Content-Disposition', 'attachment',
                                 filename=os.path.basename(file_path))
                    msg.attach(part)
            
            # 发送邮件（设置30秒超时）
            with self._create_smtp_connection(timeout=30) as smtp:
                smtp.login(self.config['smtp_username'], self.config['smtp_password'])
                smtp.send_message(msg)
            
            return True
            
        except FileNotFoundError:
            raise
        except ValueError:
            raise
        except Exception as e:
            self.logger.error(f"发送邮件失败", exc_info=True)
            raise Exception(f"发送邮件失败：{e}")