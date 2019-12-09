{foreach from=$elementNames item=elementName}
  <div class="crm-section">
    <div class="label">{$form.$elementName.label}</div>
    <div class="content">
      {$form.$elementName.html}
      {if $elementDescriptions.$elementName}
        <div class="description">{$elementDescriptions.$elementName}</div>
      {/if}
    </div>
    <div class="clear"></div>
  </div>
{/foreach}

<div class="crm-submit-buttons">
  {include file="CRM/common/formButtons.tpl" location="bottom"}
</div>
