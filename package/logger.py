import logging


def clear_log_file(base_path):
    path = base_path + r"\app.log"
    print(path)
    with open(path, "w") as f:
        pass  # 清空文件内容
def logger_init(base_path):
    clear_log_file(base_path)
    logging.basicConfig(level=logging.DEBUG, 
                        format='%(asctime)s - %(levelname)s - %(message)s',
                        handlers=[
                            logging.FileHandler("app.log"),
                            logging.StreamHandler()
                        ])


# 日志级别映射
LOG_LEVELS = {
    "DEBUG": logging.DEBUG,
    "INFO": logging.INFO,
    "WARNING": logging.WARNING,
    "ERROR": logging.ERROR,
    "CRITICAL": logging.CRITICAL
}
def log_message(message, level="INFO"):
    """记录日志信息，支持字符串形式的日志级别"""
    log_level = LOG_LEVELS.get(level.upper(), logging.INFO)  # 默认为 INFO
    logging.log(log_level, message)