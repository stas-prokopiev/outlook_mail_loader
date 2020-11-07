"""
Module with class to listen to some outlook folder and dump all letters
to the local directory
"""
# Standard library imports
import logging
import datetime

# Third party imports
import win32com.client
from char import char
from local_simple_database import LocalSimpleDatabase

# Local imports
from .class_mail_dumper import MailFolderDumper

LOGGER = logging.getLogger("outlook_mail_loader")


class MailFolderListener(MailFolderDumper):
    """"""
    @char
    def __init__(
            self,
            str_folder_to_get="root",
            str_path_dir_where_to_save="mails",
    ):
        pass

    # self.l_datetimes_when_letter_saved = []
    #     self.l_datetimes_when_letter_saved.append(datetime.datetime.now())