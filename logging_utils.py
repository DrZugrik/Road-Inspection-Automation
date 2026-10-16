"""Настройка единого журналирования запуска."""

import logging
from datetime import datetime
from pathlib import Path


def configure_logging(output_dir: Path, level_name: str) -> logging.Logger:
    output_dir.mkdir(parents=True, exist_ok=True)
    log_dir = output_dir / 'logs'
    log_dir.mkdir(parents=True, exist_ok=True)

    logger = logging.getLogger()
    logger.setLevel(getattr(logging, level_name, logging.INFO))
    logger.handlers.clear()
    formatter = logging.Formatter('%(asctime)s | %(levelname)-8s | %(name)s | %(message)s')

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    log_path = log_dir / f"run_{datetime.now():%Y%m%d_%H%M%S}.log"
    file_handler = logging.FileHandler(log_path, encoding='utf-8')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    return logging.getLogger(__name__)