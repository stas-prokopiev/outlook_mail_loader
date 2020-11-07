# Standard library imports
import sys
import logging

# Third party imports
from logging_nice_handlers import JupyterStreamHandler

# Local imports
from .class_mail_dumper import MailFolderDumper

LOGGER = logging.getLogger("outlook_mail_loader")
LOGGER.setLevel(level=10)  # Or any level you see suitable now
LOGGER.propagate = False


# Add stdout handler
# stdout_format = logging.Formatter("[%(levelname)s]: %(message)s")
# stdout_handler = logging.StreamHandler(sys.stdout)
# stdout_handler.setFormatter(stdout_format)
# LOGGER.addHandler(stdout_handler)


LOGGER.addHandler(JupyterStreamHandler(20, 30))


__all__ = ["MailFolderDumper"]


