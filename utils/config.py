"""
配置管理模块
"""
import os
from dotenv import load_dotenv

# 加载环境变量
load_dotenv()


class Config:
    """配置管理类"""

    def __init__(self):
        self.imap_config = {
            "server": os.getenv("IMAP_SERVER"),
            "smtp_server": os.getenv("SMTP_SERVER"),
            "username": os.getenv("EMAIL_USERNAME"),
            "password": os.getenv("EMAIL_PASSWORD"),
            "sent": os.getenv("SENT_FOLDER"),
            "drafts": os.getenv("DRAFTS_FOLDER")
        }

        self.ai_config = {
            "api_key": os.getenv("AI_API_KEY"),
            "api_url": os.getenv("AI_API_URL"),
            "model": os.getenv("AI_MODEL"),
            "system_prompt": os.getenv("AI_SYSTEM_PROMPT")
        }

        self.report_config = {
            "title_prefix": os.getenv("REPORT_TITLE_PREFIX"),
            "title_date_format": os.getenv("REPORT_TITLE_DATE_FORMAT"),
            "default_to": os.getenv("REPORT_DEFAULT_TO"),
            "default_cc": os.getenv("REPORT_DEFAULT_CC")
        }

    def validate_config(self) -> bool:
        """验证配置是否完整"""
        # 验证邮箱配置
        if not all([
            self.imap_config["server"],
            self.imap_config["smtp_server"],
            self.imap_config["username"],
            self.imap_config["password"]
        ]):
            return False

        # 验证AI配置
        if not self.ai_config["api_key"]:
            return False

        return True

    def get_missing_config(self) -> list:
        """获取缺失的配置项"""
        missing = []

        # 检查邮箱配置
        required_imap = ["server", "smtp_server", "username", "password"]
        for key in required_imap:
            if not self.imap_config[key]:
                missing.append(f"IMAP_{key.upper()}")

        # 检查AI配置
        if not self.ai_config["api_key"]:
            missing.append("AI_API_KEY")

        return missing


# 全局配置实例
config = Config()
