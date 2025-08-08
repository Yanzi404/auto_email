"""
AI助手工具类
"""
import requests
import email
from datetime import datetime
from typing import List, Tuple
from .logger import logger


class AIAssistant:
    """AI助手类，负责调用AI API生成周报"""

    def __init__(self, config: dict):
        """
        初始化AI助手
        
        Args:
            config: AI配置字典
        """
        self.api_key = config["api_key"]
        self.api_url = config["api_url"]
        self.model = config["model"]
        self.system_prompt = config["system_prompt"]

        # 验证必要的配置
        if not self.api_key:
            raise ValueError("AI API密钥未设置，请检查环境变量")

    def generate_weekly_summary(self, daily_reports: List[Tuple[str, str]]) -> str:
        """
        根据日报生成周报
        
        Args:
            daily_reports: 日报列表，每项为(日期, 内容)元组
        
        Returns:
            生成的周报内容
        
        Raises:
            Exception: 当AI请求失败时
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
                ai_content = response.json()['choices'][0]['message']['content']
                logger.info("AI成功生成周报")
                return ai_content
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
        """
        格式化日报内容，使其更易于AI处理
        
        Args:
            daily_reports: 日报列表
        
        Returns:
            格式化后的日报内容
        """
        try:
            # 按日期排序
            sorted_reports = sorted(
                daily_reports,
                key=lambda x: email.utils.parsedate_to_datetime(x[0]) if x[0] else datetime.now()
            )

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
