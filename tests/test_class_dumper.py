# -*- coding: utf-8 -*-
from __future__ import unicode_literals
from outlook_mail_loader import MailFolderDumper



def test_main_functionality():
    """"""
    # 1
    mail_loader_obj = MailFolderDumper("Входящие", "./tests/test_mails")
    mail_loader_obj.dump_new(1)
    mail_loader_obj.print_stats_about_initialized_folders()
    mail_loader_obj.print_full_folders_hierarchy_from_root()
    #####
    # 2
    mail_loader_obj2 = MailFolderDumper("inbox", "./tests/test_mails")
    mail_loader_obj2.dump_new(1)
