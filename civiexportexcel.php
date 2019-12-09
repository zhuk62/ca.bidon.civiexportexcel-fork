<?php

require_once 'civiexportexcel.civix.php';
use CRM_CiviExportExcel_ExtensionUtil as E;

require_once(__DIR__ . '/vendor/autoload.php');

/**
 * Implementation of hook_civicrm_config
 */
function civiexportexcel_civicrm_config(&$config) {
  _civiexportexcel_civix_civicrm_config($config);
}

/**
 * Implementation of hook_civicrm_xmlMenu
 *
 * @param $files array(string)
 */
function civiexportexcel_civicrm_xmlMenu(&$files) {
  _civiexportexcel_civix_civicrm_xmlMenu($files);
}

/**
 * Implementation of hook_civicrm_install
 */
function civiexportexcel_civicrm_install() {
  return _civiexportexcel_civix_civicrm_install();
}

/**
 * Implementation of hook_civicrm_uninstall
 */
function civiexportexcel_civicrm_uninstall() {
  return _civiexportexcel_civix_civicrm_uninstall();
}

/**
 * Implementation of hook_civicrm_enable
 */
function civiexportexcel_civicrm_enable() {
  return _civiexportexcel_civix_civicrm_enable();
}

/**
 * Implementation of hook_civicrm_disable
 */
function civiexportexcel_civicrm_disable() {
  return _civiexportexcel_civix_civicrm_disable();
}

/**
 * Implementation of hook_civicrm_upgrade
 *
 * @param $op string, the type of operation being performed; 'check' or 'enqueue'
 * @param $queue CRM_Queue_Queue, (for 'enqueue') the modifiable list of pending up upgrade tasks
 *
 * @return mixed  based on op. for 'check', returns array(boolean) (TRUE if upgrades are pending)
 *                for 'enqueue', returns void
 */
function civiexportexcel_civicrm_upgrade($op, CRM_Queue_Queue $queue = NULL) {
  return _civiexportexcel_civix_civicrm_upgrade($op, $queue);
}

/**
 * Implementation of hook_civicrm_managed
 *
 * Generate a list of entities to create/deactivate/delete when this module
 * is installed, disabled, uninstalled.
 */
function civiexportexcel_civicrm_managed(&$entities) {
  return _civiexportexcel_civix_civicrm_managed($entities);
}

/**
 * Implements hook_civicrm_alterSettingsFolders().
 */
function civiexportexcel_civicrm_alterSettingsFolders(&$metaDataFolders = NULL) {
  _civiexportexcel_civix_civicrm_alterSettingsFolders($metaDataFolders);
}

/**
 * Implementation of hook_civicrm_buildForm().
 *
 * Used to add a 'Export to Excel' button in the Report forms.
 */
function civiexportexcel_civicrm_buildForm($formName, &$form) {
  // Reports extend the CRM_Report_Form class.
  // We use that to check whether we should inject the Excel export buttons.
  if (!is_subclass_of($form, 'CRM_Report_Form')) {
    return;
  }

  if (!$form->elementExists('task')) {
    return;
  }

  // Insert the "Export to Excel" task before "Export to CSV"
  if ($form->elementExists('task')) {
    $e = $form->getElement('task');

    $actions = CRM_Report_BAO_ReportInstance::getActionMetadata();
    $tasks = [];

    foreach ($actions as $key => $val) {
      // NB: ts() not E::ts(), because this is a core string.
      if ($key == 'report_instance.csv') {
        $tasks['report_instance.excel2007'] = [
          'title' => ts('Export to Excel', ['domain' => 'ca.bidon.civiexportexcel']),
        ];
      }

      $tasks[$key] = $val;
    }

    $form->removeElement('task');

    // Based on CRM_Report_BAO_ReportInstance
    $form->assign('taskMetaData', $tasks);
    $select = $form->add('select', 'task', NULL, array('' => ts('Actions')), FALSE, array(
      'class' => 'crm-select2 crm-action-menu fa-check-circle-o huge crm-search-result-actions')
    );

    foreach ($tasks as $key => $task) {
      $attributes = array();
      if (isset($task['data'])) {
        foreach ($task['data'] as $dataKey => $dataValue) {
          $attributes['data-' . $dataKey] = $dataValue;
        }
      }
      $select->addOption($task['title'], $key, $attributes);
    }

    $smarty = CRM_Core_Smarty::singleton();
    $vars = $smarty->get_template_vars();

    $form->_excelButtonName = $form->getButtonName('submit', 'excel');

    $label = (! empty($vars['instanceId']) ? E::ts('Export to Excel') : E::ts('Preview Excel'));
    $form->addElement('submit', $form->_excelButtonName, $label);

    CRM_Core_Region::instance('page-body')->add(array(
      'template' => 'CRM/Report/Form/Actions-civiexportexcel.tpl',
    ));

    // This is to preserve legacy behaviour, i.e. if core is not patched
    if (empty($form->supportsExportExcel) && CRM_Utils_Request::retrieveValue('task', 'String') == 'report_instance.excel2007') {
      civiexportexcel_legacyBuildFormExport($form);
    }
  }
}

/**
 * Legacy code for exporting report data, without a patch on CiviCRM core.
 *
 * @see civiexportexcel_civicrm_buildForm()
 * @deprecated
 */
function civiexportexcel_legacyBuildFormExport($form) {
  $output = CRM_Utils_Request::retrieve('output', 'String', CRM_Core_DAO::$_nullObject);
  $form->assign('printOnly', TRUE);
  $printOnly = TRUE;
  $form->assign('outputMode', 'excel2007');

  // FIXME: this duplicates part of CRM_Report_Form::postProcess()
  // since we do not have a place to hook into, we hi-jack the form process
  // before it gets into postProcess.

  // get ready with post process params
  $form->beginPostProcess();

  // build query
  $sql = $form->buildQuery(FALSE);

  // build array of result based on column headers. This method also allows
  // modifying column headers before using it to build result set i.e $rows.
  $rows = array();
  $form->buildRows($sql, $rows);

  // format result set.
  // This seems to cause more problems than it fixes.
  // $form->formatDisplay($rows);

  //Alter display for rows without formatDisplay
  $form->alterDisplay($rows);

  // assign variables to templates
  $form->doTemplateAssignment($rows);

  CRM_CiviExportExcel_Utils_Report::export2excel2007($form, $rows);
}

/**
 * Implements hook_civicrm_export().
 *
 * Called mostly to export search results.
 */
function civiexportexcel_civicrm_export($exportTempTable, $headerRows, $sqlColumns, $exportMode) {
  $writeHeader = true;

  $rows = array();

  $query = "SELECT * FROM $exportTempTable";
  $dao = CRM_Core_DAO::executeQuery($query);

  while ($dao->fetch()) {
    $row = array();
    foreach ($sqlColumns as $column => $dontCare) {
      $row[$column] = $dao->$column;
    }

    $rows[] = $row;
  }

  $dao->free();

  CRM_CiviExportExcel_Utils_SearchExport::export2excel2007($headerRows, $sqlColumns, $rows);
}

/**
 * Implements hook_civicrm_alterMailParams().
 *
 * Intercepts outgoing report emails, in order to attach the
 * excel2007 version of the report.
 *
 * TODO: we should really propose a patch to CRM_Report_Form::endPostProcess().
 */
function civiexportexcel_attach_to_email(&$form, &$rows, &$attachments) {
  $config = CRM_Core_Config::singleton();

  $filename = 'CiviReport.xlsx';
  $fullname = $config->templateCompileDir . CRM_Utils_File::makeFileName($filename);

  CRM_CiviExportExcel_Utils_Report::generateFile($form, $rows, $fullname);

  $attachments[] = array(
    'fullPath' => $fullname,
    'mime_type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'cleanName' => $filename,
  );
}
