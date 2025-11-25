import logging
import sys
from functools import lru_cache

LOG_FORMAT = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'


@lru_cache(maxsize=4)
def get_logger(name: str = 'robo_backend') -> logging.Logger:
    logger = logging.getLogger(name)
    if not logger.handlers:
        handler = logging.StreamHandler(sys.stdout)
        formatter = logging.Formatter(LOG_FORMAT)
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
        logger.propagate = False
    return logger
