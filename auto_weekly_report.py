import imaplib
import email
import smtplib
import requests
from datetime import datetime, timedelta, timezone
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from typing import List, Tuple, Optional

import os
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()

# 配置参数
CONFIG = {
    "imap": {
        "server": os.getenv("IMAP_SERVER"),
        "smtp_server": os.getenv("SMTP_SERVER"),
        "username": os.getenv("EMAIL_USERNAME"),  # 邮箱地址
        "password": os.getenv("EMAIL_PASSWORD"),  # 邮箱密码
        "sent": os.getenv("SENT_FOLDER"),
        "drafts": os.getenv("DRAFTS_FOLDER")
    },
    "ai": {
        "api_key": os.getenv("AI_API_KEY"),
        "api_url": os.getenv("AI_API_URL"),
        "model": os.getenv("AI_MODEL"),
        "system_prompt": os.getenv("AI_SYSTEM_PROMPT")
    },
    "report": {
        "title_prefix": os.getenv("REPORT_TITLE_PREFIX"),  # 周报标题前缀
        "title_date_format": os.getenv("REPORT_TITLE_DATE_FORMAT"),  # 周报标题日期格式
        "default_to": os.getenv("REPORT_DEFAULT_TO"),  # 默认收件人
        "default_cc": os.getenv("REPORT_DEFAULT_CC")  # 默认抄送人，多个邮箱用逗号分隔
    }
}

contents = []

import logging

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("email_automation.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class EmailSender:
    def __init__(self, config: dict):
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
        :param subject: 邮件主题
        :param content: 邮件内容
        :param to: 收件人
        :param cc: 抄送人，多个邮箱用逗号分隔
        :return: 邮件对象
        """
        msg = MIMEMultipart()
        msg['From'] = self.username
        if to:
            msg['To'] = to
        if cc:
            msg['Cc'] = cc
        msg['Subject'] = subject
        msg.attach(MIMEText(content, 'plain', 'utf-8'))
        return msg

    def save_to_drafts(self, subject: str, content: str, to: Optional[str] = None, cc: Optional[str] = None) -> bool:
        """
        创建邮件并保存到草稿箱
        :param subject: 邮件主题
        :param content: 邮件内容
        :param to: 收件人
        :param cc: 抄送人，多个邮箱用逗号分隔
        :return: 是否成功
        """
        try:
            # 创建邮件
            msg = self.create_message(subject, content, to, cc)
            # 保存到草稿箱
            self._save_to_drafts_folder(msg)
            return True
        except Exception as e:
            logger.error(f"保存到草稿箱失败: {str(e)}")
            return False

    def _save_to_drafts_folder(self, msg: MIMEMultipart) -> None:
        """
        保存邮件到草稿箱
        :param msg: 邮件对象
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


class EmailFetcher:
    def __init__(self, config: dict):
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
        """选择邮件文件夹
        :param folder_name: 文件夹名称
        :return: 是否成功
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
        """获取本周的日报邮件内容
        :param keyword: 邮件主题中的关键词
        :return: 邮件内容列表，每项为(日期, 内容)元组
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
        """获取草稿箱中的日报
        :param keyword: 邮件主题中的关键词
        :return: 邮件内容列表
        """
        if self.select_folder(self.drafts):
            return self.fetch_weekly_reports(keyword)
        return []

    def _decode_header(self, header: str) -> str:
        """解码邮件头
        :param header: 邮件头
        :return: 解码后的字符串
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
        """提取邮件正文内容
        :param msg: 邮件对象
        :return: 邮件正文
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


class AIAssistant:
    def __init__(self, config: dict):
        self.api_key = config["api_key"]
        self.api_url = config["api_url"]
        self.model = config["model"]
        self.system_prompt = config["system_prompt"]

        # 验证必要的配置
        if not self.api_key:
            raise ValueError("AI API密钥未设置，请检查环境变量")

    def generate_weekly_summary(self, daily_reports: List[Tuple[str, str]]) -> str:
        """根据日报生成周报
        :param daily_reports: 日报列表，每项为(日期, 内容)元组
        :return: 生成的周报内容
        """
        if not daily_reports:
            logger.warning("没有找到日报内容，无法生成周报")
            return "本周没有找到日报内容，无法生成周报。"

        try:
            # 格式化日报内容，使其更易于AI处理
            formatted_reports = self._format_reports(daily_reports)

            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}"
            }

            data = {
                "model": self.model,
                "messages": [
                    {"role": "system", "content": self.system_prompt},
                    {"role": "user", "content": formatted_reports}
                ],
                "stream": False
            }

            logger.info(f"正在请求AI生成周报，使用模型: {self.model}")
            response = requests.post(self.api_url, headers=headers, json=data, timeout=30)

            if response.status_code == 200:
                content = response.json()['choices'][0]['message']['content']
                logger.info("AI成功生成周报")
                return content
            else:
                error_msg = f"AI请求失败: HTTP {response.status_code}, {response.text}"
                logger.error(error_msg)
                raise Exception(error_msg)
        except requests.RequestException as e:
            logger.error(f"网络请求错误: {str(e)}")
            raise
        except Exception as e:
            logger.error(f"生成周报时出错: {str(e)}")
            raise

    def _format_reports(self, daily_reports: List[Tuple[str, str]]) -> str:
        """格式化日报内容，使其更易于AI处理
        :param daily_reports: 日报列表
        :return: 格式化后的日报内容
        """
        try:
            # 按日期排序
            sorted_reports = sorted(daily_reports,
                                    key=lambda x: email.utils.parsedate_to_datetime(x[0]) if x[0] else datetime.now())

            formatted = []
            for i, (date, content) in enumerate(sorted_reports):
                try:
                    # 解析日期
                    if date:
                        parsed_date = email.utils.parsedate_to_datetime(date)
                        date_str = parsed_date.strftime("%Y-%m-%d %A")
                    else:
                        date_str = "未知日期"

                    formatted.append(f"===== 日报 {i + 1}: {date_str} =====\n{content}\n")
                except Exception as e:
                    logger.warning(f"格式化日报 {i + 1} 时出错: {str(e)}")
                    formatted.append(f"===== 日报 {i + 1} =====\n{content}\n")

            return "\n\n".join(formatted)
        except Exception as e:
            logger.error(f"格式化日报内容时出错: {str(e)}")
            # 返回原始格式，确保程序不会崩溃
            return str(daily_reports)


def main():
    try:
        logger.info("开始执行自动生成周报程序")

        # 验证配置
        if not CONFIG["imap"]["username"] or not CONFIG["imap"]["password"]:
            logger.error("邮箱配置不完整，请检查环境变量设置")
            return

        if not CONFIG["ai"]["api_key"]:
            logger.error("AI API密钥未设置，请检查环境变量")
            return

        # 获取本周日报
        daily_reports = []
        logger.info("开始获取本周日报")
        with EmailFetcher(CONFIG["imap"]) as fetcher:
            # 从已发送邮件中获取日报
            sent_reports = fetcher.fetch_weekly_reports()
            logger.info(f"从已发送邮件中获取到 {len(sent_reports)} 条日报")
            daily_reports.extend(sent_reports)

            # 从草稿箱中获取日报
            draft_reports = fetcher.fetch_drafts()
            logger.info(f"从草稿箱中获取到 {len(draft_reports)} 条日报")
            daily_reports.extend(draft_reports)

        if not daily_reports:
            logger.warning("没有找到任何日报，程序结束")
            return

        # 生成周报
        logger.info("开始生成周报")
        assistant = AIAssistant(CONFIG["ai"])
        weekly_summary = assistant.generate_weekly_summary(daily_reports)

        # 打印周报内容
        print("\n" + "=" * 50)
        print("生成的周报内容:")
        print("=" * 50)
        print(weekly_summary)
        print("=" * 50 + "\n")

        # 保存周报到草稿箱
        save_to_drafts = input("是否保存周报到草稿箱？(y/n): ").strip().lower() == 'y'
        if save_to_drafts:
            sender = EmailSender(CONFIG["imap"])
            # 使用配置的标题前缀和日期格式
            title_prefix = CONFIG["report"]["title_prefix"]
            date_format = CONFIG["report"]["title_date_format"]
            subject = f"{title_prefix}{datetime.now().strftime(date_format)}"

            # 使用配置的默认收件人和抄送人
            default_to = CONFIG["report"]["default_to"]
            default_cc = CONFIG["report"]["default_cc"]

            if sender.save_to_drafts(subject, weekly_summary, default_to, default_cc):
                if default_to:
                    logger.info(f"默认收件人: {default_to}")
                if default_cc:
                    logger.info(f"默认抄送人: {default_cc}")
            else:
                logger.error("保存周报到草稿箱失败")

        logger.info("程序执行完成")
    except Exception as e:
        logger.error(f"程序运行出错: {str(e)}")
        print(f"程序运行出错: {str(e)}")


if __name__ == "__main__":
    main()
