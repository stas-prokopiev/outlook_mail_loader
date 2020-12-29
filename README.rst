===================
outlook_mail_loader
===================

.. image:: https://img.shields.io/github/last-commit/stas-prokopiev/outlook_mail_loader
   :target: https://img.shields.io/github/last-commit/stas-prokopiev/outlook_mail_loader
   :alt: GitHub last commit

.. image:: https://img.shields.io/github/license/stas-prokopiev/outlook_mail_loader
    :target: https://github.com/stas-prokopiev/outlook_mail_loader/blob/master/LICENSE.txt
    :alt: GitHub license<space><space>

.. image:: https://readthedocs.org/projects/outlook_mail_loader/badge/?version=latest
    :target: https://outlook_mail_loader.readthedocs.io/en/latest/?badge=latest
    :alt: Documentation Status

.. image:: https://img.shields.io/badge/pre--commit-enabled-brightgreen?logo=pre-commit&logoColor=white
   :target: https://github.com/pre-commit/pre-commit
   :alt: pre-commit

.. image:: https://img.shields.io/pypi/v/outlook_mail_loader
   :target: https://img.shields.io/pypi/v/outlook_mail_loader
   :alt: PyPI

.. image:: https://img.shields.io/pypi/pyversions/outlook_mail_loader
   :target: https://img.shields.io/pypi/pyversions/outlook_mail_loader
   :alt: PyPI - Python Version


.. contents:: **Table of Contents**

Short Overview.
=========================
outlook_mail_loader is a python package (**py>=3.6**) which helps in handling outlook letters

This library helps to dump letters from outlook to some local folder in human readable format.

Installation via pip:
======================

.. code-block:: bash

    pip install outlook_mail_loader

Typical usages
============================

1) Dump letters from outlook to local directory
-----------------------------------------------

| If you want to dump all new letters from outlook folder to windows local folder.
| Firstly, you should define dumper object from class **MailFolderDumper**
| And then, whenever you want you can call method dump_new(...) to dump the new letters

Simplest example
*********************

.. code-block:: python

    from outlook_mail_loader import MailFolderDumper

    mail_loader_obj = MailFolderDumper(
        str_outlook_folder_name="inbox",
        str_path_dir_where_to_save="mails",
    )

    # To dump the new letters from the asked outlook folder
    mail_loader_obj.dump_new()

After the dump you will get file-folder structure like that:

| **str_path_dir_where_to_save**
| --> **<str_outlook_folder_name>**
| ----> **LETTER_1**
| ----> **LETTER_2**
| ----> ...
| ----> **LETTER_N**
| ------> *letter.txt*
| ------> *dict_metainfo.json*
| ------> **ATTACHMENTS**
| --------> *file_1*
| --------> *file_N*

Full signature of **mail_loader_obj.dump_new** method
***************************************************************

.. code-block:: python

    # Full signature of the dump_new method
    mail_loader_obj.dump_new(
        int_max_last_letters_to_dump=10,
        is_to_mark_messages_as_read=False,
        is_to_remove_attachments=False,
        is_to_preserve_msg_obj=False,
    )

Attributes and methods of **mail_loader_obj**
***************************************************************

Attributes:

* **.str_outlook_folder_name** (str): Folder name which to dump
* **.str_path_dir_where_to_save** (str): Path where to dump

Methods:

* **.dump_new(...)** - Dump new letters to set local directory
* **.print_stats_about_initialized_folders()** - Print hierarchy for initialized outlook mail folder
* **.print_full_folders_hierarchy_from_root()** - Print full hierarchy from root outlook mail folder
* **.get_list_names_of_all_outlook_folders()** - Get list names of all outlook folders available

2) Process dumped letters
---------------------------------------------

| After the letters are dumped to local folder
| You most probably want to do some action with them
| Here are some handlers for doing it

Example
*********************

.. code-block:: python

    from outlook_mail_loader import DumpedMails

    dumped_mails_obj = DumpedMails(str_path_dir_with_mails="mails",)

    # Get dictionary with last letter
    dict_last_letter = dumped_mails_obj.get_last_letter()
    print(get_last_letter)

    # Get dictionary with last N letter
    list_dict_last_5_letter = dumped_mails_obj.get_last_n_letters(5)

    # Print statistics about all dumped letters
    dumped_mails_obj.print_stats_about_dumped_mails()

Format of the dictitonary with dumped letter
***************************************************************

* **dict_one_letter["letter"]** - Text of the letter
* **dict_one_letter["dict_metainfo"]** - All metainfo about the letter
* **dict_one_letter["list_attachments"]** - List pathes to files with letter's attachments

Attributes and methods of **dumped_mails_obj**
***************************************************************

Attributes:

* **.str_path_dir_with_mails** (str): Path to dir with dumped letters
* **.int_last_dumped_id** (str): Id of the last dumped letter

Methods:

* **.get_last_letter()** - Get dictionary with last letter
* **.get_last_n_letters(int_last_letters_to_get)** - Get list of dicts of last N letters
* **.print_stats_about_dumped_mails()** - Print statistics about all dumped letters
* **.clear_dumped_mails()** - Clear from cache dumped mails

3) Listen to some outlook folder to dump all letters continuously
-------------------------------------------------------------------

| In case if you want to run some process only once and
| then be sure that all new letters are dumped into the local folder
| Then you can start mail folder listener

Example
*********************

.. code-block:: python

    from outlook_mail_loader import listen_outlook_mail_folder

    # To dump the new letters from the asked outlook folder
    listen_outlook_mail_folder(
        str_outlook_folder_name="inbox",
        str_path_dir_where_to_save="mails",
        int_seconds_step_in_dump=60,
    )

Full signature of **listen_outlook_mail_folder**
***************************************************************

.. code-block:: python

    from outlook_mail_loader import listen_outlook_mail_folder

    # Full signature of the dump_new method
    listen_outlook_mail_folder(
        str_outlook_folder_name="inbox",
        str_path_dir_where_to_save="mails",
        int_seconds_step_in_dump=60,
        is_to_mark_messages_as_read=False,
        is_to_remove_attachments=False,
        is_to_preserve_msg_obj=False,
    )

Arguments description:

* **str_outlook_folder_name** (str, optional): Which outlook folder to listen
* **str_path_dir_where_to_save** (str, optional): Path to dir. where to save letters.
* **int_seconds_step_in_dump** (int, optional): Seconds to wait between dumping new letters.
* **is_to_mark_messages_as_read** (bool, optional): Flag if to mark as read saved letters. Default is False.
* **is_to_remove_attachments** (bool, optional): Flag if to remove attachments to save disk space. Default is False.
* **is_to_preserve_msg_obj** (bool, optional): Flag if to preserve outlook .msg object. Default is False.

Links
=====

    * `PYPI <https://pypi.org/project/outlook_mail_loader/>`_
    * `readthedocs <https://outlook_mail_loader.readthedocs.io/en/latest/>`_
    * `GitHub <https://github.com/stas-prokopiev/outlook_mail_loader>`_

Project local Links
===================

    * `CHANGELOG <https://github.com/stas-prokopiev/outlook_mail_loader/blob/master/CHANGELOG.rst>`_.
    * `CONTRIBUTING <https://github.com/stas-prokopiev/outlook_mail_loader/blob/master/CONTRIBUTING.rst>`_.

Contacts
========

    * Email: stas.prokopiev@gmail.com
    * `vk.com <https://vk.com/stas.prokopyev>`_
    * `Facebook <https://www.facebook.com/profile.php?id=100009380530321>`_

License
=======

This project is licensed under the MIT License.
