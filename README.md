CiviCRM export to Excel
=======================

This extension adds the possibility to export directly into the MS Excel
format from CiviReports and Search results, instead of CSV (less fiddling,
easier to use).

This extension uses the PhpSpreadsheet library. See the "License" section below
for more information (LGPL).

To download the latest version of this module:  
https://lab.civicrm.org/extensions/civiexportexcel/

This extension is maintained by [Coop SymbioTIC](https://www.symbiotic.coop/en).

Sponsors:

* [Canadian Credit Union Association](https://ccua.com)
* [Projet Montréal](http://projetmontreal.org)
* [Development and Peace](https://www.devp.org)

Warnings
========

* This extension does not run the buildACLClause() function, meaning that you may have deleted contacts show up in some reports. If you are using ACLs in general, this can also cause important issues.

Requirements
============

- CiviCRM >= 5.0 (previous versions untested)

Installation
============

1- Download this extension and unpack it in your 'extensions' directory.
   You may need to create it if it does not already exist, and configure
   the correct path in CiviCRM -> Administer -> System -> Directories.

2- If installing from git, run `composer install`.

3- Apply this CiviCRM core patch civiexportexcel-core.patch (optional but recommended).

4- Enable the extension from CiviCRM -> Administer -> System -> Extensions.

5- If you wish to send emails with the report as an Excel attachment,
   you must apply the patch in civiexportexcel-core-mail.patch.

Report mails
============

To send report e-mails in Excel2007 format, use: "format=excel2007" in
the "Scheduled Jobs" settings.

Support
=======

Please post bug reports in the issue tracker of this project on CiviCRM's Gitlab:  
https://lab.civicrm.org/extensions/civiexportexcel/issues

For general questions and support, please post on Stack Exchange:  
https://civicrm.stackexchange.com/

This is a community contributed extension written thanks to the financial
support of organisations using it, as well as the very helpful and collaborative
CiviCRM community.

If you appreciate this extension, please consider financially supporting the
CiviCRM project by becoming a member, a partner or a one-time donation:  
https://civicrm.org/support-us

While we do our best to provide volunteer support for this extension, please
consider financially contributing to support or development of this extension
if you can.

Commercial support via Coop SymbioTIC:  
https://www.symbiotic.coop/en

Todo
====

* Propose a new hook to CiviCRM for a cleaner postProcess implementation (incl. mail).
* Add OpenDocument (LibreOffice) support.
* Add admin settings form to enable excel/opendocument formats?

License
=======

(C) 2014-2019 Mathieu Lutfy <mathieu@symbiotic.coop>  
(C) 2018-2019 Coop SymbioTIC <info@symbiotic.coop>

Distributed under the terms of the GNU Affero General public license (AGPL).
See LICENSE.txt for details.

This extension relies on phpSpreadsheet by PHPOffice:  
https://github.com/PHPOffice/PhpSpreadsheet

See composer.json for more information about dependencies.
