'use strict';

/**
 * kpi-settings.js — Settings panel UI (open/close, all form builders, copy/paste).
 * Exports onto window.KPI.settingsPanel
 */

window.KPI = window.KPI || {};
window.KPI.settingsPanel = (function () {

  var utils = window.KPI.utils;
  var SETTINGS_KEYS = utils.SETTINGS_KEYS;
  var KPI_VERSION   = utils.KPI_VERSION;

  // These require Tableau — only called from the inline panel or the parent window.
  var loadSettings  = utils.loadSettings;
  var saveSettings  = utils.saveSettings;

  var _settingsOverlay = null;

  function closeSettingsPanel () {
    if (_settingsOverlay) {
      _settingsOverlay.remove();
      _settingsOverlay = null;
    }
  }

  // =====================================================================
  // Toast notification
  // =====================================================================

  function showToast (msg) {
    var toast = document.createElement('div');
    toast.className = 'settings-toast';
    toast.textContent = msg;
    document.body.appendChild(toast);
    setTimeout(function () { toast.remove(); }, 2000);
  }

  // =====================================================================
  // Shared form builder (used by inline panel)
  // =====================================================================

  /**
   * Builds the full settings form inside `panel`.
   *
   * @param {HTMLElement} panel - container to append form elements to
   * @param {object} settings - current settings values
   * @param {object} opts
   *   onInput(values)       — called (debounced) when any field changes
   *   onClose()             — called when close button is clicked
   *   onApplyPaste(parsed)  — called when pasted settings are applied
   *   kpiLabel              — default value label (from kpi.label)
   *   goal2Label            — default goal2 label
   *   sheetName             — worksheet name for defaults
   *   allFields             — array of { value, label } for field dropdowns
   *   measures              — array of { value, label } for measure dropdowns
   *
   * @returns {{ collectValues: function }}
   */
  function buildSettingsForm (panel, settings, opts) {

    var _debounceTimer = null;

    // Helper: collect current values from all form inputs
    function collectValues () {
      var inputs = panel.querySelectorAll('input[data-key], select[data-key]');
      var newSettings = {};
      inputs.forEach(function (inp) {
        if (inp.type === 'checkbox') {
          newSettings[inp.dataset.key] = inp.checked;
        } else if (inp.dataset.numVal !== undefined) {
          newSettings[inp.dataset.key] = inp.dataset.numVal;
        } else {
          newSettings[inp.dataset.key] = inp.value;
        }
      });
      return newSettings;
    }

    function onInput () {
      clearTimeout(_debounceTimer);
      _debounceTimer = setTimeout(function () {
        opts.onInput(collectValues());
      }, 300);
    }

    // ----- Section collapse persistence -----
    var SECTION_STORE_KEY = '_kpi_sections';
    function _getSectionStates () {
      try { return JSON.parse(sessionStorage.getItem(SECTION_STORE_KEY)) || {}; }
      catch (_) { return {}; }
    }
    function _saveSectionState (id, open) {
      var states = _getSectionStates();
      states[id] = open;
      try { sessionStorage.setItem(SECTION_STORE_KEY, JSON.stringify(states)); }
      catch (_) { /* ok */ }
    }

    // ----- Builders -----

    function addSection (title) {
      var sectionId = title.replace(/\s+/g, '_').toLowerCase();
      var savedStates = _getSectionStates();
      var isOpen = savedStates[sectionId] === true;

      var sec = document.createElement('div');
      sec.className = 'settings-section' + (isOpen ? '' : ' collapsed');
      var hdr = document.createElement('div');
      hdr.className = 'settings-section-header';

      var titleSpan = document.createElement('span');
      titleSpan.textContent = title;
      hdr.appendChild(titleSpan);

      var chevron = document.createElement('span');
      chevron.className = 'settings-section-chevron';
      chevron.textContent = '\u25BE';
      hdr.appendChild(chevron);

      var body = document.createElement('div');
      body.className = 'settings-section-body';
      body.style.maxHeight = isOpen ? 'none' : '0';

      hdr.addEventListener('click', function () {
        var nowCollapsed = sec.classList.toggle('collapsed');
        _saveSectionState(sectionId, !nowCollapsed);
        if (!nowCollapsed) {
          body.style.maxHeight = body.scrollHeight + 'px';
          setTimeout(function () { if (!sec.classList.contains('collapsed')) body.style.maxHeight = 'none'; }, 260);
        } else {
          body.style.maxHeight = body.scrollHeight + 'px';
          body.offsetHeight; // force reflow
          body.style.maxHeight = '0';
        }
      });

      sec.appendChild(hdr);
      sec.appendChild(body);
      panel.appendChild(sec);
      return body;
    }

    function addToggle (parent, labelText, key, hint) {
      var group = document.createElement('div');
      group.className = 'settings-group';
      var row = document.createElement('div');
      row.className = 'settings-toggle-row';
      var lbl = document.createElement('span');
      lbl.className = 'settings-toggle-label';
      lbl.textContent = labelText;
      row.appendChild(lbl);
      var sw = document.createElement('label');
      sw.className = 'toggle-switch';
      var cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = !!settings[key];
      cb.dataset.key = key;
      cb.addEventListener('change', onInput);
      sw.appendChild(cb);
      var track = document.createElement('span');
      track.className = 'toggle-track';
      sw.appendChild(track);
      row.appendChild(sw);
      group.appendChild(row);
      if (hint) {
        var hintEl = document.createElement('div');
        hintEl.className = 'settings-hint';
        hintEl.textContent = hint;
        group.appendChild(hintEl);
      }
      parent.appendChild(group);
    }

    function addField (parent, labelText, key, placeholder, hint, defaultValue) {
      var group = document.createElement('div');
      group.className = 'settings-group';
      var lbl = document.createElement('label');
      lbl.textContent = labelText;
      group.appendChild(lbl);
      var input = document.createElement('input');
      input.type = 'text';
      input.placeholder = placeholder;
      input.value = settings[key] !== undefined ? settings[key] : (defaultValue ?? '');
      input.dataset.key = key;
      input.addEventListener('input', onInput);
      group.appendChild(input);
      if (hint) {
        var hintEl = document.createElement('div');
        hintEl.className = 'settings-hint';
        hintEl.textContent = hint;
        group.appendChild(hintEl);
      }
      parent.appendChild(group);
    }

    function addColorField (parent, labelText, key, defaultColor, hint) {
      var group = document.createElement('div');
      group.className = 'settings-group';
      var row = document.createElement('div');
      row.style.display = 'flex';
      row.style.alignItems = 'center';
      row.style.gap = '8px';
      var lbl = document.createElement('label');
      lbl.textContent = labelText;
      lbl.style.flex = '1';
      lbl.style.marginBottom = '0';
      row.appendChild(lbl);
      var colorInp = document.createElement('input');
      colorInp.type = 'color';
      colorInp.value = settings[key] || defaultColor;
      colorInp.dataset.key = key;
      colorInp.style.width = '32px';
      colorInp.style.height = '28px';
      colorInp.style.border = '1px solid #ddd';
      colorInp.style.borderRadius = '6px';
      colorInp.style.cursor = 'pointer';
      colorInp.style.padding = '2px';
      colorInp.addEventListener('input', onInput);
      row.appendChild(colorInp);
      var btnDef = document.createElement('button');
      btnDef.className = 'settings-btn settings-btn-cancel';
      btnDef.innerHTML = '&#8634;';
      btnDef.title = 'Reset to default';
      btnDef.style.cssText = 'font-size:16px;padding:2px 4px;line-height:1;color:#a0aab4;cursor:pointer;';
      btnDef.addEventListener('click', function () {
        colorInp.value = defaultColor;
        colorInp.dispatchEvent(new Event('input'));
      });
      row.appendChild(btnDef);
      group.appendChild(row);
      if (hint) {
        var hintEl = document.createElement('div');
        hintEl.className = 'settings-hint';
        hintEl.textContent = hint;
        group.appendChild(hintEl);
      }
      parent.appendChild(group);
    }

    function addDropdown (parent, labelText, key, options, hint, allowNone) {
      var group = document.createElement('div');
      group.className = 'settings-group';
      var lbl = document.createElement('label');
      lbl.textContent = labelText;
      group.appendChild(lbl);
      var select = document.createElement('select');
      select.dataset.key = key;
      select.addEventListener('change', onInput);
      var currentValue = settings[key];
      if (allowNone) {
        var emptyOpt = document.createElement('option');
        emptyOpt.value = '';
        emptyOpt.textContent = '(None)';
        select.appendChild(emptyOpt);
      }
      for (var i = 0; i < options.length; i++) {
        var opt = options[i];
        var optEl = document.createElement('option');
        optEl.value = opt.value;
        optEl.textContent = opt.label;
        if (currentValue === opt.value) {
          optEl.selected = true;
        } else if (!allowNone && !currentValue && i === 0) {
          optEl.selected = true;
        }
        select.appendChild(optEl);
      }
      group.appendChild(select);
      if (hint) {
        var hintEl = document.createElement('div');
        hintEl.className = 'settings-hint';
        hintEl.textContent = hint;
        group.appendChild(hintEl);
      }
      parent.appendChild(group);
    }

    function addStepper (parent, labelText, key, defaultVal, min, max, step, autoLabel) {
      var group = document.createElement('div');
      group.className = 'settings-group';
      var lbl = document.createElement('label');
      lbl.textContent = labelText;
      group.appendChild(lbl);
      var row = document.createElement('div');
      row.style.cssText = 'display:flex;align-items:center;gap:6px;';
      var btnMinus = document.createElement('button');
      btnMinus.className = 'settings-btn settings-btn-cancel';
      btnMinus.style.cssText = 'padding:4px 10px;font-size:14px;line-height:1;';
      btnMinus.textContent = '\u2212';
      var display = document.createElement('input');
      display.type = 'text';
      display.style.cssText = 'width:62px;text-align:center;border:1.5px solid #e3e0ea;border-radius:8px;padding:5px 4px;font-size:13px;color:var(--brand-dark);outline:none;';
      var curVal = settings[key] !== undefined ? parseInt(settings[key], 10) : defaultVal;
      var resolvedVal = isNaN(curVal) ? defaultVal : curVal;
      display.value = (autoLabel && resolvedVal === min) ? autoLabel : resolvedVal;
      display.dataset.key = key;
      display.dataset.numVal = String(resolvedVal);
      var btnPlus = document.createElement('button');
      btnPlus.className = 'settings-btn settings-btn-cancel';
      btnPlus.style.cssText = 'padding:4px 10px;font-size:14px;line-height:1;';
      btnPlus.textContent = '+';

      function update (newVal) {
        var clamped = Math.max(min, Math.min(max, newVal));
        display.dataset.numVal = String(clamped);
        display.value = (autoLabel && clamped === min) ? autoLabel : clamped;
        onInput();
      }

      btnMinus.addEventListener('click', function () { update(parseInt(display.dataset.numVal, 10) - step); });
      btnPlus.addEventListener('click', function () { update(parseInt(display.dataset.numVal, 10) + step); });
      display.addEventListener('change', function () {
        var parsed = parseInt(display.value, 10);
        update(isNaN(parsed) ? defaultVal : parsed);
      });

      row.appendChild(btnMinus);
      row.appendChild(display);
      row.appendChild(btnPlus);

      if (arguments.length > 8 && arguments[8] === true) {
        var btnDef = document.createElement('button');
        btnDef.className = 'settings-btn settings-btn-cancel';
        btnDef.innerHTML = '&#8634;';
        btnDef.title = 'Reset to ' + defaultVal;
        btnDef.style.cssText = 'font-size:16px;padding:2px 4px;line-height:1;color:#a0aab4;cursor:pointer;';
        btnDef.addEventListener('click', function () { update(defaultVal); });
        row.appendChild(btnDef);
      }

      group.appendChild(row);
      parent.appendChild(group);
    }

    // ===================================================================
    // Build sections by component
    // ===================================================================

    var kpiLabel   = opts.kpiLabel   || '';
    var goal2Label = opts.goal2Label || 'Secondary Goal';
    var sheetName  = opts.sheetName  || '';
    var allFields  = opts.allFields  || [];
    var measures   = opts.measures   || [];

    // --- LAYOUT ---
    var secLayout = addSection('Layout');
    addDropdown(secLayout, 'Card Layout', 'cardLayout',
      [
        { value: 'standard', label: 'Standard' },
        { value: 'gauge',    label: 'Gauge' }
      ], 'Standard = list view with sparkline. Gauge = semi-circle goal dial.');
    addToggle(secLayout, 'Fill Worksheet', 'fillWorksheet', 'Scale the card to fill the entire view.');
    addToggle(secLayout, 'Compact Mode', 'compactMode', 'Reduce internal spacing for a tighter layout.');
    addStepper(secLayout, 'Margin Top (px)', 'padTop', 24, 0, 80, 2, undefined, true);
    addStepper(secLayout, 'Margin Left (px)', 'padLeft', 28, 0, 80, 2, undefined, true);
    addStepper(secLayout, 'Card Width (px)', 'cardWidth', 480, 200, 1200, 20, undefined, true);
    addColorField(secLayout, 'Accent Color', 'gradColor', '#D42F8A',
      'Top border gradient is generated from this color.');
    addDropdown(secLayout, 'Progress Bar Color', 'barColorMode',
      [
        { value: 'default', label: 'Default' },
        { value: 'accent',  label: 'Match accent' },
        { value: 'custom',  label: 'Custom' }
      ], 'Color style for goal progress bars.');
    addColorField(secLayout, 'Bar Color', 'barCustomColor', '#3b82f6',
      'Used when "Custom" is selected.');
    addToggle(secLayout, 'Open Animation', 'showAnimation', 'Short draw-in animation on load.');

    // --- TITLE & HEADER ---
    var secTitle = addSection('Title & Header');
    addToggle(secTitle, 'Show Title', 'showTitle', 'The sheet name displayed at the top.');
    addField(secTitle, 'Title Text', 'titleText',
      sheetName || 'Sheet name', 'Blank = use sheet name.', '');
    addStepper(secTitle, 'Title Size (px)', 'titleSize', 18, 10, 72, 1, undefined, true);
    addToggle(secTitle, 'Show Subtitle / Label', 'showHeader', 'Value label and date range row.');
    addDropdown(secTitle, 'Subtitle Alignment', 'subtitleAlign',
      [
        { value: 'left',   label: 'Left' },
        { value: 'center', label: 'Center' },
        { value: 'right',  label: 'Right' }
      ], 'When not using date field, align subtitle under the title. (Standard layout only.)');
    addField(secTitle, 'Value Label Override', 'valueLabel',
      kpiLabel || 'Auto-detected', 'Blank = hide label.', kpiLabel || '');
    addToggle(secTitle, 'Show Date Range', 'showDateRange', 'Period range in the subtitle.');

    // --- NUMBER FORMAT ---
    var secFmt = addSection('Number Format');
    addField(secFmt, 'Prefix', 'fmtPrefix', '', 'e.g. $ or \u20AC \u2014 applies to all numbers.', '');
    addField(secFmt, 'Suffix', 'fmtSuffix', '', 'e.g. % or units \u2014 applies to all numbers.', '');
    addStepper(secFmt, 'Decimal Places', 'fmtDecimals', 0, 0, 6, 1);
    addToggle(secFmt, 'Abbreviate (K / M / B)', 'fmtAbbreviate',
      'Show 120K instead of 120,000. Applies to all numbers.');

    // --- VALUE & DELTA ---
    var secValue = addSection('Value & Delta');
    addToggle(secValue, 'Show Value', 'showValue', 'The big number.');
    addStepper(secValue, 'Value Size (px)', 'valueSize', 40, 16, 72, 2, undefined, true);
    addToggle(secValue, 'Show Delta Badge', 'showDelta', 'Change vs previous period.');
    addField(secValue, 'Delta Label', 'deltaLabel', 'vs prev', 'Text after the percentage.', 'vs prev');
    addToggle(secValue, 'Reverse Delta Colors', 'reverseDelta', 'Higher = bad (costs, churn).');
    addStepper(secValue, 'Delta Decimal Places', 'fmtDeltaDecimals', 1, 0, 4, 1);
    addStepper(secValue, 'Badge Font Size (px)', 'deltaSize', 12, 8, 20, 1);

    // --- PRIMARY GOAL ---
    var secGoal = addSection('Primary Goal');
    addToggle(secGoal, 'Show Primary Goal', 'showGoal');
    addField(secGoal, 'Goal Label', 'goalLabel', 'Goal', 'Blank = hide bar.', 'Goal');

    // --- SECONDARY GOAL ---
    var secGoal2 = addSection('Secondary Goal');
    addToggle(secGoal2, 'Show Secondary Goal', 'showGoal2');
    addField(secGoal2, 'Goal Label', 'goal2Label',
      goal2Label || 'Secondary Goal', 'Blank = hide bar.', goal2Label || 'Secondary Goal');

    // --- SPARKLINE ---
    var secSpark = addSection('Sparkline Chart');
    addToggle(secSpark, 'Show Sparkline', 'showSparkline');
    addToggle(secSpark, 'Show Data Labels', 'showSparkLabels', 'Values above each data point.');
    addToggle(secSpark, 'Show Period Labels', 'showSparkPeriods', 'Period names along x-axis.');
    addToggle(secSpark, 'Show Legend', 'showLegend', 'Legend below the sparkline chart.');
    addStepper(secSpark, 'Chart Height (px)', 'sparkHeight', 130, 60, 300, 10);
    addDropdown(secSpark, 'Line & Dot Color', 'sparklineColorMode',
      [
        { value: 'default', label: 'Default' },
        { value: 'accent',  label: 'Match accent' },
        { value: 'custom',  label: 'Custom' }
      ], 'Color for sparkline lines and dots (labels stay neutral).');
    addColorField(secSpark, 'Custom Color', 'sparklineCustomColor', '#3b82f6',
      'Used when "Custom" is selected.');

    // --- SECONDARY COMPARISON ---
    var secComp = addSection('Secondary Comparison');
    addToggle(secComp, 'Enable Comparison', 'ptdEnabled',
      'Adds a badge and a dashed line on the sparkline.');
    addDropdown(secComp, 'Comparison Field', 'ptdFieldName', measures,
      'Select a calculated field to compare against (e.g. PTD, forecast).', true);
    addField(secComp, 'Badge Label', 'ptdLabel', 'vs prev PTD',
      'Text shown on the comparison badge.', 'vs prev PTD');
    addField(secComp, 'Legend Label', 'ptdLegendLabel', 'PTD Pace',
      'Label for the dashed line in the sparkline legend.', 'PTD Pace');

    // --- DATA FIELDS ---
    var secData = addSection('Data Fields');
    addDropdown(secData, 'Date / Period Field', 'dateFieldName', allFields,
      'Field to group by (quarters, months, etc).', true);

    // --- CUSTOM LINK ---
    var secLink = addSection('Custom Link');
    addToggle(secLink, 'Show Link', 'showLink');
    addField(secLink, 'Link Label', 'linkLabel', 'e.g. View Dashboard', 'Blank = hide link.');
    addField(secLink, 'Link URL', 'linkUrl', 'https://...');
    addField(secLink, 'Icon URL', 'linkIcon', 'https://...', 'Optional 16x16 icon.');

    // ===================================================================
    // Copy / Paste settings
    // ===================================================================

    var actionsRow = document.createElement('div');
    actionsRow.className = 'settings-actions';

    var copyBtn = document.createElement('button');
    copyBtn.className = 'settings-btn settings-btn-cancel';
    copyBtn.textContent = 'Copy Settings';
    copyBtn.addEventListener('click', async function () {
      var currentSettings = collectValues();
      try {
        await navigator.clipboard.writeText(JSON.stringify(currentSettings, null, 2));
        showToast('Settings copied to clipboard');
      } catch (_) {
        var ta = document.createElement('textarea');
        ta.value = JSON.stringify(currentSettings, null, 2);
        ta.style.cssText = 'position:fixed;left:-9999px';
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        ta.remove();
        showToast('Settings copied to clipboard');
      }
    });

    var pasteBtn = document.createElement('button');
    pasteBtn.className = 'settings-btn settings-btn-primary';
    pasteBtn.textContent = 'Paste Settings';

    var pasteArea = document.createElement('div');
    pasteArea.style.cssText = 'display:none;margin-top:10px;';

    var pasteTA = document.createElement('textarea');
    pasteTA.placeholder = 'Paste settings JSON here (Ctrl+V)...';
    pasteTA.style.cssText = 'width:100%;height:80px;border:1.5px solid #e3e0ea;border-radius:8px;padding:8px;font-size:11px;font-family:monospace;color:var(--brand-dark);outline:none;resize:vertical;';

    var applyBtn = document.createElement('button');
    applyBtn.className = 'settings-btn settings-btn-primary';
    applyBtn.style.cssText = 'margin-top:6px;width:100%;';
    applyBtn.textContent = 'Apply';
    applyBtn.addEventListener('click', function () {
      var text = pasteTA.value.trim();
      if (!text) { showToast('Paste your settings JSON first'); return; }
      var parsed;
      try { parsed = JSON.parse(text); }
      catch (_) { showToast('Invalid JSON format'); return; }
      if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
        showToast('Invalid settings data'); return;
      }
      var knownKeys = new Set(Object.keys(SETTINGS_KEYS));
      var validKeys = Object.keys(parsed).filter(function (k) { return knownKeys.has(k); });
      if (validKeys.length === 0) { showToast('No recognized settings found'); return; }
      if (opts.onApplyPaste) {
        opts.onApplyPaste(parsed);
      }
      showToast('Settings applied (' + validKeys.length + ' values)');
    });

    pasteArea.appendChild(pasteTA);
    pasteArea.appendChild(applyBtn);

    pasteBtn.addEventListener('click', function () {
      var visible = pasteArea.style.display !== 'none';
      pasteArea.style.display = visible ? 'none' : 'block';
      if (!visible) pasteTA.focus();
    });

    actionsRow.appendChild(copyBtn);
    actionsRow.appendChild(pasteBtn);
    panel.appendChild(actionsRow);
    panel.appendChild(pasteArea);

    // Bottom close button
    var closeBtnBottom = document.createElement('button');
    closeBtnBottom.className = 'settings-btn settings-btn-cancel';
    closeBtnBottom.style.cssText = 'margin-top:12px;width:100%;padding:8px 0;font-size:13px;font-weight:600;';
    closeBtnBottom.textContent = 'Close';
    closeBtnBottom.addEventListener('click', opts.onClose);
    panel.appendChild(closeBtnBottom);

    // Version label
    var versionEl = document.createElement('div');
    versionEl.style.cssText = 'text-align:center;font-size:10px;color:#c4c9d0;margin-top:10px;';
    versionEl.textContent = 'KPI Card v' + KPI_VERSION;
    panel.appendChild(versionEl);

    return { collectValues: collectValues };
  }

  // =====================================================================
  // Inline side panel (for worksheet view)
  // =====================================================================

  function openSettingsPanel (kpi, dataResult, onSave, sheetName) {
    if (_settingsOverlay) { closeSettingsPanel(); return; }

    var settings = loadSettings();

    // Available fields for dropdowns
    var allFields = (dataResult && dataResult.columns)
      ? dataResult.columns.map(function (c) { return { value: c.fieldName, label: c.fieldName }; })
      : [];
    var measures = (dataResult && dataResult.columns)
      ? dataResult.columns
          .filter(function (c) { return c.dataType === 'float' || c.dataType === 'int'; })
          .map(function (c) { return { value: c.fieldName, label: c.fieldName }; })
      : [];

    // Container
    var overlay = document.createElement('div');
    overlay.className = 'settings-overlay';
    _settingsOverlay = overlay;

    var panel = document.createElement('div');
    panel.className = 'settings-panel';

    // Title row
    var titleRow = document.createElement('div');
    titleRow.className = 'settings-title';
    titleRow.innerHTML = '<svg viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M11.49 3.17c-.38-1.56-2.6-1.56-2.98 0a1.532 1.532 0 01-2.286.948c-1.372-.836-2.942.734-2.106 2.106.54.886.061 2.042-.947 2.287-1.561.379-1.561 2.6 0 2.978a1.532 1.532 0 01.947 2.287c-.836 1.372.734 2.942 2.106 2.106a1.532 1.532 0 012.287.947c.379 1.561 2.6 1.561 2.978 0a1.533 1.533 0 012.287-.947c1.372.836 2.942-.734 2.106-2.106a1.533 1.533 0 01.947-2.287c1.561-.379 1.561-2.6 0-2.978a1.532 1.532 0 01-.947-2.287c.836-1.372-.734-2.942-2.106-2.106a1.532 1.532 0 01-2.287-.947zM10 13a3 3 0 100-6 3 3 0 000 6z" clip-rule="evenodd"/></svg> <span style="flex:1">Settings</span>';

    var closeBtn = document.createElement('button');
    closeBtn.className = 'settings-btn settings-btn-cancel';
    closeBtn.style.cssText = 'padding:4px 10px;font-size:16px;line-height:1;';
    closeBtn.textContent = '\u00D7';
    closeBtn.addEventListener('click', closeSettingsPanel);
    titleRow.appendChild(closeBtn);
    panel.appendChild(titleRow);

    // Build form
    buildSettingsForm(panel, settings, {
      onInput: function (values) {
        saveSettings(values, onSave);
      },
      onClose: closeSettingsPanel,
      onApplyPaste: function (parsed) {
        saveSettings(parsed, onSave);
        closeSettingsPanel();
        openSettingsPanel(kpi, dataResult, onSave, sheetName);
      },
      kpiLabel:   kpi ? kpi.label : '',
      goal2Label: kpi ? kpi.goal2Label : 'Secondary Goal',
      sheetName:  sheetName,
      allFields:  allFields,
      measures:   measures
    });

    overlay.appendChild(panel);
    document.body.appendChild(overlay);

    var firstInput = panel.querySelector('input');
    if (firstInput) firstInput.focus();
  }

  // =========================================================================
  // Public API
  // =========================================================================

  return {
    openSettingsPanel: openSettingsPanel,
    closeSettingsPanel: closeSettingsPanel
  };

})();
