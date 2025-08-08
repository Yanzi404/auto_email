"""
日志配置模块
"""
import logging
import os


def setup_logger(name: str = __name__, log_file: str = "email_automation.log") -> logging.Logger:
    """
    设置日志配置
    
    Args:
        name: 日志器名称
        log_file: 日志文件名
    
    Returns:
        配置好的日志器
    """
    logger = logging.getLogger(name)

    # 避免重复添加处理器
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)

    # 创建格式器
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )

    # 文件处理器
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)

    # 控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    # 添加处理器
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger


# 全局日志器
logger = setup_logger("auto_email")
