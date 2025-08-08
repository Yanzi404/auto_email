"""
邮件发送工具类
"""
import imaplib
from datetime import datetime, timezone
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from typing import Optional
from .logger import logger


class EmailSender:
    """邮件发送类，负责创建和发送邮件"""
    
    def __init__(self, config: dict):
        """
        初始化邮件发送器
        
        Args:
            config: IMAP配置字典
        """
        self.smtp_server = config["smtp_server"]
        self.username = config["username"]
        self.password = config["password"]
        self.drafts_folder = config["drafts"]
        self.imap_server = config["server"]
        
        # 验证必要的配置
        if not all([self.smtp_server, self.username, self.password]):
            raise ValueError("SMTP配置不完整，请检查环境变量设置")
    
    def create_message(self, subject: str, content: str, to: Optional[str] = None,
                       cc: Optional[str] = None) -> MIMEMultipart:
        """
        创建邮件对象
        
        Args:
            subject: 邮件主题
            content: 邮件内容
            to: 收件人
            cc: 抄送人，多个邮箱用逗号分隔
        
        Returns:
            邮件对象
        """
        msg = MIMEMultipart()
        msg['From'] = self.username
        if to:
            msg['To'] = to
        if cc:
            msg['Cc'] = cc
        msg['Subject'] = subject
        
        msg.attach(MIMEText(content, 'html'))
        
        # 添加嵌入图片
        try:
            with open('template/mengxiang.PNG', 'rb') as f:
                img = MIMEImage(f.read())
                img.add_header('Content-ID', '<image1>')  # 这里的image1对应HTML中的cid:image1
                msg.attach(img)
        except FileNotFoundError:
            logger.warning("未找到图片文件 template/mengxiang.PNG，跳过图片附件")
        except Exception as e:
            logger.error(f"添加图片附件时出错: {str(e)}")
        
        return msg
    
    def save_to_drafts(self, subject: str, content: str, to: Optional[str] = None, 
                       cc: Optional[str] = None) -> bool:
        """
        创建邮件并保存到草稿箱
        
        Args:
            subject: 邮件主题
            content: 邮件内容
            to: 收件人
            cc: 抄送人，多个邮箱用逗号分隔
        
        Returns:
            是否成功
        """
        try:
            # 创建邮件
            msg = self.create_message(subject, content, to, cc)
            # 保存到草稿箱
            self._save_to_drafts_folder(msg)
            logger.info("邮件已成功保存到草稿箱")
            return True
        except Exception as e:
            logger.error(f"保存到草稿箱失败: {str(e)}")
            return False
    
    def _save_to_drafts_folder(self, msg: MIMEMultipart) -> None:
        """
        保存邮件到草稿箱
        
        Args:
            msg: 邮件对象
        
        Raises:
            Exception: 当保存失败时
        """
        try:
            with imaplib.IMAP4_SSL(self.imap_server) as imap:
                imap.login(self.username, self.password)
                # 使用带时区的datetime
                current_time = datetime.now(timezone.utc)
                imap.append(
                    f'"{self.drafts_folder}"',
                    '',
                    imaplib.Time2Internaldate(current_time),
                    msg.as_bytes()
                )
                logger.info("邮件已保存到草稿箱")
        except Exception as e:
            logger.error(f"保存到草稿箱失败: {str(e)}")
            raise