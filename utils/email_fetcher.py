"""
邮件获取工具类
"""
import imaplib
import email
from datetime import datetime, timedelta
from email.header import decode_header
from typing import List, Tuple
from .logger import logger


class EmailFetcher:
    """邮件获取类，负责从邮箱获取邮件内容"""
    
    def __init__(self, config: dict):
        """
        初始化邮件获取器
        
        Args:
            config: IMAP配置字典
        """
        self.imap_server = config["server"]
        self.username = config["username"]
        self.password = config["password"]
        self.sent = config["sent"]
        self.drafts = config["drafts"]
        self.imap = None
        
        # 验证必要的配置
        if not all([self.imap_server, self.username, self.password]):
            raise ValueError("IMAP配置不完整，请检查环境变量设置")
    
    def __enter__(self):
        """连接IMAP服务器"""
        try:
            self.imap = imaplib.IMAP4_SSL(self.imap_server)
            self.imap.login(self.username, self.password)
            self.select_folder(self.sent)  # 默认选择已发送文件夹
            return self
        except Exception as e:
            logger.error(f"连接IMAP服务器失败: {str(e)}")
            raise
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """关闭IMAP连接"""
        if self.imap:
            try:
                self.imap.close()
                self.imap.logout()
            except Exception as e:
                logger.warning(f"关闭IMAP连接时出错: {str(e)}")
    
    def select_folder(self, folder_name: str) -> bool:
        """
        选择邮件文件夹
        
        Args:
            folder_name: 文件夹名称
        
        Returns:
            是否成功
        """
        try:
            status, data = self.imap.select(f'"{folder_name}"')
            if status != 'OK':
                logger.warning(f"选择文件夹 {folder_name} 失败: {data}")
                return False
            return True
        except Exception as e:
            logger.error(f"选择文件夹时出错: {str(e)}")
            return False
    
    def fetch_weekly_reports(self, keyword: str = "日报") -> List[Tuple[str, str]]:
        """
        获取本周的日报邮件内容
        
        Args:
            keyword: 邮件主题中的关键词
        
        Returns:
            邮件内容列表，每项为(日期, 内容)元组
        """
        try:
            # 计算本周一的日期
            today = datetime.now().date()
            this_monday = today - timedelta(days=today.weekday())
            since_date = this_monday.strftime("%d-%b-%Y")
            
            # 搜索邮件
            status, messages = self.imap.search(None, f'SINCE "{since_date}"')
            if status != 'OK':
                logger.warning(f"搜索邮件失败: {messages}")
                return []
            
            mail_ids = messages[0].split()
            logger.info(f"找到 {len(mail_ids)} 封邮件")
            
            results = []
            for mail_id in mail_ids:
                try:
                    status, msg_data = self.imap.fetch(mail_id, '(RFC822)')
                    if status != 'OK':
                        logger.warning(f"获取邮件 {mail_id} 失败: {msg_data}")
                        continue
                    
                    msg = email.message_from_bytes(msg_data[0][1])
                    subject = self._decode_header(msg["Subject"])
                    
                    if keyword in subject:
                        date = msg.get("Date")
                        body = self._extract_email_body(msg)
                        if body:
                            results.append((date, body))
                            logger.debug(f"处理邮件: {subject}, 日期: {date}")
                except Exception as e:
                    logger.error(f"处理邮件 {mail_id} 时出错: {str(e)}")
                    continue
            
            return results
        except Exception as e:
            logger.error(f"获取周报邮件时出错: {str(e)}")
            return []
    
    def fetch_drafts(self, keyword: str = "日报") -> List[Tuple[str, str]]:
        """
        获取草稿箱中的日报
        
        Args:
            keyword: 邮件主题中的关键词
        
        Returns:
            邮件内容列表
        """
        if self.select_folder(self.drafts):
            return self.fetch_weekly_reports(keyword)
        return []
    
    def _decode_header(self, header: str) -> str:
        """
        解码邮件头
        
        Args:
            header: 邮件头
        
        Returns:
            解码后的字符串
        """
        if not header:
            return ""
        
        try:
            decoded = decode_header(header)[0][0]
            if isinstance(decoded, bytes):
                return decoded.decode('utf-8', errors='ignore')
            return str(decoded)
        except Exception as e:
            logger.error(f"解码邮件头时出错: {str(e)}")
            return ""
    
    def _extract_email_body(self, msg) -> str:
        """
        提取邮件正文内容
        
        Args:
            msg: 邮件对象
        
        Returns:
            邮件正文
        """
        try:
            if msg.is_multipart():
                for part in msg.walk():
                    content_type = part.get_content_type()
                    content_disposition = str(part.get("Content-Disposition") or "")
                    
                    if "attachment" not in content_disposition and content_type == "text/plain":
                        payload = part.get_payload(decode=True)
                        if payload:
                            return payload.decode('utf-8', errors='ignore')
            else:
                payload = msg.get_payload(decode=True)
                if payload:
                    return payload.decode('utf-8', errors='ignore')
            return ""
        except Exception as e:
            logger.error(f"提取邮件正文时出错: {str(e)}")
            return ""