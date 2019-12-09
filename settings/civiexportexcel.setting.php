<?php

use CRM_CiviExportExcel_ExtensionUtil as E;

return [
  'civiexportexcel_show_title' => [
    'group_name' => 'domain',
    'group' => 'civiexportexcel',
    'name' => 'civiexportexcel_show_title',
    'type' => 'Boolean',
    'default' => 0,
    'add' => '1.0',
    'is_domain' => 1,
    'is_contact' => 0,
    'title' => E::ts('Show the report title'),
    'description' => E::ts("Prints the report title as the first row of the report."),
    'quick_form_type' => 'YesNo',
  ],
  'civiexportexcel_show_export_date' => [
    'group_name' => 'domain',
    'group' => 'civiexportexcel',
    'name' => 'civiexportexcel_show_export_date',
    'type' => 'Boolean',
    'default' => 0,
    'add' => '1.0',
    'is_domain' => 1,
    'is_contact' => 0,
    'title' => E::ts('Show the report export date'),
    'description' => E::ts("Prints the export date at the top of the report."),
    'quick_form_type' => 'YesNo',
  ],
];
