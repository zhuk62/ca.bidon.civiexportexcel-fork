{*
 * Preserved for legacy behavior, when users expect
 * the 'Export to Excel' button.
 *
 * Remove at some point?
 *}

{* The nbps; are a mimic of what other buttons do in templates/CRM/Report/Form/Actions.tpl *}
{assign var=excel value="_qf_"|cat:$form.formName|cat:"_submit_excel"}
{$form.$excel.html}&nbsp;&nbsp;

{literal}
  <script>
    CRM.$(function($) {
      var form_id = '{/literal}{$form.$excel.id}{literal}';

      if ($('.crm-report-field-form-block .crm-submit-buttons').size() > 0) {
        $('input#' + form_id).appendTo('.crm-report-field-form-block .crm-submit-buttons');

        $('input#' + form_id).on('click', function(e) {
          e.preventDefault();
          $('select#task').val('report_instance.excel2007').trigger('change');
        });
      }
      else {
        // Do not show the button when running in a dashlet
        // FIXME: we should probably just not add the HTML in the first place.
        $('input#' + form_id).hide();
      }
    });
  </script>
{/literal}
