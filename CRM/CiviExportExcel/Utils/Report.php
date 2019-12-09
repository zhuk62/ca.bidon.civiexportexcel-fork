<?php

use CRM_CiviExportExcel_ExtensionUtil as E;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Row;
use PhpOffice\PhpSpreadsheet\Cell\Cell;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

/**
 * @package civiexportexcel
 * @copyright Mathieu Lutfy (c) 2014-2015
 */
class CRM_CiviExportExcel_Utils_Report extends CRM_Core_Page {

  const ROW_PADDING = 5;
  const DEFAULT_CELL_WIDTH = 9.14;
  const DEFAULT_ROW_HEIGHT = 10;

  /**
   * Generates a XLS 2007 file and forces the browser to download it.
   *
   * @param Object $form
   * @param Array &$rows
   *
   * See @CRM_Report_Utils_Report::export2csv().
   */
  static function export2excel2007(&$form, &$rows) {
    // Force a download and name the file using the current timestamp.
    $datetime = date('Y-m-d H:i');
    $filename = $form->getTitle() . ' - ' . $datetime . '.xlsx';

    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="' . $filename . '"');
    header("Content-Description: " . $filename);
    header("Content-Transfer-Encoding: binary");
    header("Cache-Control: no-cache, must-revalidate, post-check=0, pre-check=0");

    // always modified
    header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT');

    self::generateFile($form, $rows);
    CRM_Utils_System::civiExit();
  }

  /**
   * Utility function for export2csv and CRM_Report_Form::endPostProcess
   * - make XLS file content and return as string.
   *
   * @param Object &$form CRM_Report_Form object.
   * @param Array &$rows Resulting rows from the report.
   * @param String Full path to the filename to write in (for mailing reports).
   *
   * See @CRM_Report_Utils_Report::makeCsv().
   */
  static function generateFile(&$form, &$rows, $filename = 'php://output') {
    $config = CRM_Core_Config::singleton();
    $csv = '';

    // Generate an array with { 0=>A, 1=>B, 2=>C, ... }
    $foo = array(0 => '', 1 => 'A', 2 => 'B', 3 => 'C', 4 => 'D', 5 => 'E', 6 => 'F', 7 => 'G', 8 => 'H', 9 => 'I', 10 => 'J', 11 => 'K', 12 => 'L', 13 => 'M');
    $a = ord('A');
    $cells = array();

    for ($i = 0; $i < count($foo); $i++) {
      for ($j = 0; $j < 26; $j++) {
        $cells[$j + ($i * 26)] = $foo[$i] . chr($j + $a);
      }
    }

    $objPHPExcel = new Spreadsheet();

    // Does magic things for date cells
    // https://phpexcel.codeplex.com/discussions/331005
    \PhpOffice\PhpSpreadsheet\Cell\Cell::setValueBinder(new \PhpOffice\PhpSpreadsheet\Cell\AdvancedValueBinder());

    // FIXME Set the locale of the XLS file
    // might not really be necessary (concerns mostly functions? not dates?)
    // $validLocale = \PhpOffice\PhpSpreadsheet\Settings::setLocale('fr');

    // Set document properties
    $contact_id = CRM_Core_Session::singleton()->get('userID');
    $display_name = civicrm_api3('Contact', 'getsingle', ['contact_id' => $contact_id])['display_name'];
    $domain_name = CRM_Core_DAO::getFieldValue('CRM_Core_DAO_Domain', CRM_Core_Config::domainID(), 'name');
    $title = $form->getTitle();

    $objPHPExcel->getProperties()
      ->setCreator($display_name)
      ->setLastModifiedBy($display_name)
      ->setTitle($form->getTitle())
      ->setDescription(E::ts('Report exported from CiviCRM (%1) by %2 at %3', [
        1 => $domain_name,
        2 => $display_name,
        3 => date('Y-m-d H:i:s'),
      ]));

    $sheet = $objPHPExcel->setActiveSheetIndex(0);
    $objPHPExcel->getActiveSheet()->setTitle('Report');

    // Add headers if this is the first row.
    $columnHeaders = array_keys($form->_columnHeaders);

    // Replace internal header names with friendly ones, where available.
    foreach ($columnHeaders as $header) {
      if (isset($form->_columnHeaders[$header])) {
        $headers[] = html_entity_decode(strip_tags($form->_columnHeaders[$header]['title']));
      }
    }

    // Row counter
    $cpt = 1;
    $first_data_row = 1;

    // Used for later calculating the column widths
    $widths = [];

    // Add some report metadata
    $add_blank_because_of_meta = FALSE;

    if (Civi::settings()->get('civiexportexcel_show_title')) {
      $cell = 'A' . $cpt;
      $objPHPExcel->getActiveSheet()
        ->setCellValue($cell, $title);

      $cpt++;
      $add_blank_because_of_meta = TRUE;
    }

    if (Civi::settings()->get('civiexportexcel_show_export_date')) {
      $cell = 'A' . $cpt;
      $objPHPExcel->getActiveSheet()
        ->setCellValue($cell, E::ts('Date prepared: %1', [1 => date('Y-m-d H:i')]));

      $cpt++;
      $add_blank_because_of_meta = TRUE;
    }

    if ($add_blank_because_of_meta) {
      $cpt++;
    }

    // Add the column headers.
    $col = 0;

    foreach ($headers as $h) {
      $cell = $cells[$col] . $cpt;

      $objPHPExcel->getActiveSheet()
        ->setCellValue($cell, $h)
        ->getStyle($cell)->applyFromArray(['font' => ['bold' => true]]);

      self::addValueLengthToColumnWidths($h, $cells[$col], $widths);

      $col++;
    }

    // Headers row
    $cpt++;

    // Used later for row height recalculations.
    $first_data_row = $cpt;

    // Exported data
    foreach ($rows as $row) {
      $displayRows = array();
      $col = 0;

      foreach ($columnHeaders as $k => $v) {
        $value = CRM_Utils_Array::value($v, $row);

        if (! isset($value)) {
          $col++;
          continue;
        }

        // Remove HTML, unencode entities
        $value = html_entity_decode(strip_tags($value));

        // Data transformation before adding it to the cell
        if (CRM_Utils_Array::value('type', $form->_columnHeaders[$v]) & CRM_Utils_Type::T_DATE) {
          $group_by = CRM_Utils_Array::value('group_by', $form->_columnHeaders[$v]);

          if ($group_by == 'MONTH' || $group_by == 'QUARTER') {
            $value = CRM_Utils_Date::customFormat($value, $config->dateformatPartial);
          }
          elseif ($group_by == 'YEAR') {
            $value = CRM_Utils_Date::customFormat($value, $config->dateformatYear);
          }
          else {
            $value = CRM_Utils_Date::customFormat($value, '%Y-%m-%d');
          }
        }

        $objPHPExcel->getActiveSheet()
          ->setCellValue($cells[$col] . $cpt, $value);

        self::addValueLengthToColumnWidths($value, $cells[$col], $widths);

        // Cell formats
        if (CRM_Utils_Array::value('type', $form->_columnHeaders[$v]) & CRM_Utils_Type::T_DATE) {
          $objPHPExcel->getActiveSheet()
            ->getStyle($cells[$col] . $cpt)
            ->getNumberFormat()
            ->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_DATE_YYYYMMDD);

          // Set autosize on date columns.
          // We only do it for dates because we know they have a fixed width, unlike strings.
          // For eco-friendlyness, this should only be done once, perhaps when processing the headers initially
          $objPHPExcel->getActiveSheet()->getColumnDimension($cells[$col])->setAutoSize(true);
        }
        elseif (CRM_Utils_Array::value('type', $form->_columnHeaders[$v]) & CRM_Utils_Type::T_MONEY) {
          $objPHPExcel->getActiveSheet()->getStyle($cells[$col])
            ->getNumberFormat()
            ->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_GENERAL);
        }

        $col++;
      }

      $cpt++;
    }

    // Re-adjust the column widths
    foreach ($cells as $key => $val) {
      // Only verify columns that have a header. The 'cells' variable is larger than necessary.
      if (empty($headers[$key])) {
        continue;
      }

      $w = self::getRecommendedColumnWidth($val, $headers[$key], $widths);

      // Ignore width that would be too small to be practical.
      if ($w < 5) {
        continue;
      }

      if ($w > 75) {
        $w = 75;

        // Ex: A0:A100
        $area = $val . '0:' . $val . $cpt;

        $objPHPExcel->getActiveSheet()->getStyle($area)
          ->getAlignment()->setWrapText(true);
      }

      $objPHPExcel->getActiveSheet()->getColumnDimension($val)->setWidth($w);
    }

    // Now set row heights (skip the header)
    $sheet = $objPHPExcel->getActiveSheet();

    for ($i = $first_data_row; $i < $cpt; $i++) {
      $row = new Row($sheet, $i);
      self::autofitRowHeight($row);
    }

    $objWriter = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($objPHPExcel, 'Xlsx');
    $objWriter->save($filename);

    return ''; // FIXME
  }

  /**
   * Helper function to log data widths. Later used by getRecommendedColumnWidth().
   * This might be a bit overkill, but worth testing.
   */
  static public function addValueLengthToColumnWidths($value, $column, &$widths) {
    $len = mb_strlen($value);

    // Ignore empty cells
    if (!$len) {
      return;
    }

    if (!isset($widths[$column])) {
      $widths[$column] = [];
    }

    $widths[$column][] = $len;
  }

  /**
   * Returns a recommended column width, based on 90th percentile of rows with data.
   *
   * Percentile function based on:
   * http://php.net/manual/en/function.stats-stat-percentile.php#79752
   */
  static public function getRecommendedColumnWidth($column, $header, &$widths) {
    $p = 0.9; // 90% percentile

    if (empty($widths[$column])) {
      return 0;
    }

    $data = $widths[$column];
    sort($data);

    $count = count($data);
    $allindex = ($count-1) * $p;
    $intvalindex = intval($allindex);
    $floatval = $allindex - $intvalindex;

    if (!is_float($floatval)) {
      $result = $data[$intvalindex];
    }
    else {
      if ($count > $intvalindex+1) {
        $result = $floatval * ($data[$intvalindex+1] - $data[$intvalindex]) + $data[$intvalindex];
      }
      else {
        $result = $data[$intvalindex];
      }
    }

    if ($result < mb_strlen($header)) {
      $result = mb_strlen($header);
    }

    // Add a bit of padding.
    $result += 2;

    return $result;
  }

  /**
   * Auto-fit the row height based on the largest cell
   *
   * Source: https://github.com/PHPOffice/PhpSpreadsheet/issues/333#issuecomment-385943533
   *
   * @param Row $start
   * @return Worksheet
   */
  public static function autofitRowHeight(Row $row, $rowPadding = self::ROW_PADDING) {
    $ws = $row->getWorksheet();
    $cellIterator = $row->getCellIterator();
    $cellIterator->setIterateOnlyExistingCells(true);

    $maxCellLines = 1;

    /* @var $cell Cell */
    foreach ($cellIterator as $cell) {
      $cellLength = mb_strlen($cell->getValue());
      $cellWidth = $ws->getColumnDimension($cell->getParent()->getCurrentColumn())->getWidth();

      // If no column width is set, set the default
      if ($cellWidth === -1) {
        // [ML] Let our own function do this
        // $ws->getColumnDimension($cell->getParent()->getCurrentColumn())->setWidth(self::DEFAULT_CELL_WIDTH);
        $cellWidth = $ws->getColumnDimension($cell->getParent()->getCurrentColumn())->getWidth();
      }

      // If the cell is in a merge range we need to determine the full width of the range
      if($cell->isInMergeRange()) {
        // We only need to do this for the master (first) cell in the range, the rest need to have a line height of 1
        if ($cell->isMergeRangeValueCell()) {
          $mergeRange = $cell->getMergeRange();
          if ($mergeRange) {
            $mergeWidth = 0;
            $mergeRefs = Coordinate::extractAllCellReferencesInRange($mergeRange);
            foreach ($mergeRefs as $cellRef) {
              $mergeCell = $ws->getCell($cellRef);
              $width = $ws->getColumnDimension($mergeCell->getParent()->getCurrentColumn())->getWidth();
              if($width === -1) {
                // [ML] Let our own function do this
                // $ws->getColumnDimension($mergeCell->getParent()->getCurrentColumn())->setWidth(self::DEFAULT_CELL_WIDTH);
                $width = $ws->getColumnDimension($mergeCell->getParent()->getCurrentColumn())->getWidth();
              }
              $mergeWidth += $width;
            }
            $cellWidth = $mergeWidth;
          }
          else {
            $cellWidth = 1;
          }
        } else {
          $cellWidth = 1;
        }
      }

      // Calculate the number of cell lines with a 10% additional margin
      $cellLines = ceil(($cellLength * 1.1) / $cellWidth);
      $maxCellLines = $cellLines > $maxCellLines ? $cellLines : $maxCellLines;
    }

    $rowDimension= $ws->getRowDimension($row->getRowIndex());
    $rowHeight = $rowDimension->getRowHeight();

    // If no row height is set, set the default
    if ($rowHeight === -1) {
      $rowDimension->setRowHeight(self::DEFAULT_ROW_HEIGHT);
      $rowHeight = $rowDimension->getRowHeight();
    }

    $rowLines = $maxCellLines <= 0 ? 1 : $maxCellLines;
    $rowDimension->setRowHeight(($rowHeight * $rowLines) + $rowPadding);

    return $ws;
  }

}
