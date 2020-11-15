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

.. image:: https://travis-ci.org/stas-prokopiev/outlook_mail_loader.svg?branch=master
    :target: https://travis-ci.org/stas-prokopiev/outlook_mail_loader

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

Usage
============================

Dump letters from outlook to local directory
---------------------------------------------

.. code-block:: python

    from outlook_mail_loader import MailFolderDumper

    mail_loader_obj = MailFolderDumper(
        str_folder_to_get="inbox",
        str_path_dir_where_to_save="mails",
    )

    mail_loader_obj.dump_new(<int_max_last_letters_to_dump>)


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

