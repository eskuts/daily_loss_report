import logging
import sys
from pathlib import Path
from typing import Optional


def setup_logging(
        log_file: Optional[str] = "app.log",
        console_level: int = logging.INFO,
        file_level: int = logging.DEBUG,
        module: Optional[str] = None
) -> logging.Logger:
    """Настройка и получение логгера для модуля

    Args:
        log_file: Путь к файлу логов. Если None - запись в файл не ведется
        console_level: Уровень логирования для консоли
        file_level: Уровень логирования для файла
        module: Имя модуля для логгера. Если None - используется корневой логгер

    Returns:
        Настроенный логгер
    """
    logger = logging.getLogger(module)
    logger.setLevel(logging.DEBUG)  # Самый низкий уровень для обработки всех сообщений

    # Форматтер
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    # Обработчик для консоли
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(console_level)
    console_handler.setFormatter(formatter)

    # Обработчик для файла (если указан)
    if log_file:
        # Создаем директорию для логов, если ее нет
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)

        file_handler = logging.FileHandler(log_file)
        file_handler.setLevel(file_level)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    logger.addHandler(console_handler)

    return logger