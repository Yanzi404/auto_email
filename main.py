"""
自动周报生成主程序
"""
from datetime import datetime
from utils.config import config
from utils.logger import logger
from utils.ai_assistant import AIAssistant
from utils.email_fetcher import EmailFetcher
from utils.email_sender import EmailSender


def load_template(template_path: str = 'template/template.html') -> str:
    """
    加载HTML模板
    
    Args:
        template_path: 模板文件路径
    
    Returns:
        模板内容
    """
    try:
        with open(template_path, 'r', encoding='utf-8') as file:
            return file.read()
    except FileNotFoundError:
        logger.error(f"模板文件 {template_path} 不存在")
        return "{weekly_summary}"  # 返回简单的占位符
    except Exception as e:
        logger.error(f"读取模板文件时出错: {str(e)}")
        return "{weekly_summary}"


def generate_weekly_report():
    """生成周报的主要逻辑"""
    try:
        logger.info("开始执行自动生成周报程序")

        # 验证配置
        if not config.validate_config():
            missing_config = config.get_missing_config()
            logger.error(f"配置不完整，缺少以下配置项: {', '.join(missing_config)}")
            return False

        # 获取本周日报
        daily_reports = []
        logger.info("开始获取本周日报")

        with EmailFetcher(config.imap_config) as fetcher:
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
            return False

        # 生成周报
        logger.info("开始生成周报")
        assistant = AIAssistant(config.ai_config)
        ai_content = assistant.generate_weekly_summary(daily_reports)

        # 打印周报内容
        print("\n" + "=" * 50)
        print("生成的周报内容:")
        print("=" * 50)
        print(ai_content)
        print("=" * 50 + "\n")

        # 使用模板生成周报
        html_template = load_template()
        weekly_summary = html_template.format(weekly_summary=ai_content)

        # 询问是否保存到草稿箱
        save_to_drafts = input("是否保存周报到草稿箱？(y/n): ").strip().lower() == 'y'
        if save_to_drafts:
            sender = EmailSender(config.imap_config)

            # 生成邮件主题
            title_prefix = config.report_config["title_prefix"] or "周报-"
            date_format = config.report_config["title_date_format"] or "%Y年第%U周"
            subject = f"{title_prefix}{datetime.now().strftime(date_format)}"

            # 获取默认收件人和抄送人
            default_to = config.report_config["default_to"]
            default_cc = config.report_config["default_cc"]

            if sender.save_to_drafts(subject, weekly_summary, default_to, default_cc):
                if default_to:
                    logger.info(f"默认收件人: {default_to}")
                if default_cc:
                    logger.info(f"默认抄送人: {default_cc}")
                logger.info("周报已成功保存到草稿箱")
            else:
                logger.error("保存周报到草稿箱失败")
                return False

        logger.info("程序执行完成")
        return True

    except Exception as e:
        logger.error(f"程序运行出错: {str(e)}")
        print(f"程序运行出错: {str(e)}")
        return False


def main():
    """主入口函数"""
    try:
        success = generate_weekly_report()
        if not success:
            exit(1)
    except KeyboardInterrupt:
        logger.info("程序被用户中断")
        print("\n程序已停止")
    except Exception as e:
        logger.error(f"程序异常退出: {str(e)}")
        print(f"程序异常退出: {str(e)}")
        exit(1)


if __name__ == "__main__":
    main()
