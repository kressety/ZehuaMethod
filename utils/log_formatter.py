import logging


class LogColors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


class ColoredFormatter(logging.Formatter):
    LEVEL_COLORS = {
        logging.DEBUG: LogColors.OKBLUE + '%s' + LogColors.ENDC,
        logging.INFO: LogColors.OKGREEN + '%s' + LogColors.ENDC,
        logging.WARNING: LogColors.WARNING + '%s' + LogColors.ENDC,
        logging.ERROR: LogColors.FAIL + '%s' + LogColors.ENDC,
        logging.CRITICAL: LogColors.FAIL + LogColors.BOLD + '%s' + LogColors.ENDC,
    }

    def format(self, record):
        log_fmt = self.LEVEL_COLORS.get(record.levelno) % self._fmt
        formatter = logging.Formatter(log_fmt)
        return formatter.format(record)


def setup_logger(name: str) -> logging.Logger:
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    if not logger.handlers:
        ch = logging.StreamHandler()
        ch.setLevel(logging.DEBUG)
        formatter = ColoredFormatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        ch.setFormatter(formatter)
        logger.addHandler(ch)

    return logger
