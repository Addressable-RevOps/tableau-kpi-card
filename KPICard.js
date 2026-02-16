'use strict';

/**
 * KPICard.js — Entry point: bootstrap, render, and goal bar builder.
 * Depends on: kpi-utils.js, kpi-data.js, kpi-settings.js, kpi-sparkline.js
 */

(function () {

  // Aliases
  var utils = window.KPI.utils;
  var data  = window.KPI.data;
  var settingsPanel = window.KPI.settingsPanel;
  var sparkline     = window.KPI.sparkline;

  var loadSettings          = utils.loadSettings;
  var buildGradientFromColor = utils.buildGradientFromColor;
  var escapeHtml            = utils.escapeHtml;
  var formatNumberCompact   = utils.formatNumberCompact;

  var getSummaryDataTable = data.getSummaryDataTable;
  var getEncodingMap      = data.getEncodingMap;
  var resolveEncodings    = data.resolveEncodings;
  var computeKpi          = data.computeKpi;

  var openSettingsPanel  = settingsPanel.openSettingsPanel;
  var closeSettingsPanel = settingsPanel.closeSettingsPanel;

  var drawSparkline = sparkline.drawSparkline;
  var drawGauge     = window.KPI.gauge.drawGauge;

  // Module-level state
  var _renderLatest      = null;
  var _refetchAndRender  = null;
  var _prevShowAnimation = null;

  // =========================================================================
  // Bootstrap
  // =========================================================================

  window.onload = tableau.extensions.initializeAsync().then(function () {
    var worksheet = tableau.extensions.worksheetContent.worksheet;

    var cachedResult = { rows: [], columns: [] };
    var cachedEncodings = {};
    var sheetName = worksheet.name;

    var renderLatest = function () { render(cachedResult, cachedEncodings, sheetName); };
    _renderLatest = renderLatest;

    var updateDataAndRender = async function () {
      var results = await Promise.all([
        getSummaryDataTable(worksheet),
        getEncodingMap(worksheet)
      ]);
      cachedResult = results[0];
      cachedEncodings = results[1];
      renderLatest();
    };

    _refetchAndRender = updateDataAndRender;

    onresize = function () { utils.setSkipAnimation(true); renderLatest(); };

    worksheet.addEventListener(
      tableau.TableauEventType.SummaryDataChanged,
      updateDataAndRender
    );

    updateDataAndRender();
  });

  // =========================================================================
  // Render
  // =========================================================================

  function render (dataResult, encodings, sheetName) {
    var content = document.getElementById('content');
    content.innerHTML = '';

    var settings = loadSettings();

    // Animation: in edit mode only when user toggles "Open Animation" on to preview; in published use saved setting
    var justToggledOn = _prevShowAnimation === false && settings.showAnimation === true;
    var skipAnim = utils.getSkipAnimation();
    var isAuthoring = false;
    try { if (tableau.extensions.environment.mode === 'authoring') isAuthoring = true; } catch (_) { /* ok */ }
    if (isAuthoring) {
      settings._animate = justToggledOn; // edit mode: only animate when they turn the toggle on to preview
    } else {
      settings._animate = settings.showAnimation && (!skipAnim || justToggledOn); // published: normal
    }
    _prevShowAnimation = settings.showAnimation;
    utils.setSkipAnimation(false);

    // Apply page-level layout
    var padT = parseInt(settings.padTop, 10);
    var padR = parseInt(settings.padRight, 10);
    var padB = parseInt(settings.padBottom, 10);
    var padL = parseInt(settings.padLeft, 10);
    var top = padT >= 0 ? padT : 24;
    var right = padR >= 0 ? padR : 28;
    var bottom = padB >= 0 ? padB : 24;
    var left = padL >= 0 ? padL : 28;
    content.style.padding = top + 'px ' + right + 'px ' + bottom + 'px ' + left + 'px';

    // Fill worksheet mode
    var isFill = !!settings.fillWorksheet;

    var r = resolveEncodings(encodings, dataResult.columns, settings);

    if (!r.valueCol) {
      content.innerHTML =
        '<div class="kpi-empty">' +
        'Drag a <b>measure</b> onto the <b>Value</b> tile on the Marks card.' +
        '</div>';
      return;
    }

    var kpi = computeKpi(dataResult, encodings, settings);
    if (!kpi) {
      content.innerHTML = '<div class="kpi-empty">No data available.</div>';
      return;
    }

    var card = document.createElement('div');
    card.className = 'kpi-card';
    if (settings.compactMode) card.classList.add('kpi-card--compact');
    var cw = parseInt(settings.cardWidth, 10);
    if (cw > 0) card.style.maxWidth = cw + 'px';
    try { card.classList.add(tableau.ClassNameKey.Worksheet); } catch (e) { /* ok */ }

    // Custom top gradient
    var accentHex = settings.gradColor || '#D42F8A';
    var gradStops = buildGradientFromColor(accentHex);
    card.style.borderImage = 'linear-gradient(135deg, ' + gradStops.join(', ') + ') 1';

    // ---- Custom link (top-left) ----
    if (settings.showLink) {
      var linkUrl   = (settings.linkUrl   || '').trim();
      var linkLabel = (settings.linkLabel  || '').trim();
      var linkIcon  = (settings.linkIcon   || '').trim();
      if (linkUrl && (linkIcon || linkLabel)) {
        var linkEl = document.createElement('a');
        linkEl.className = 'kpi-link';
        linkEl.href = linkUrl;
        linkEl.target = '_blank';
        linkEl.rel = 'noopener noreferrer';
        if (linkLabel) linkEl.classList.add('has-label');
        if (linkIcon) {
          var iconEl = document.createElement('img');
          iconEl.className = 'kpi-link-icon';
          iconEl.src = linkIcon;
          iconEl.alt = '';
          iconEl.onerror = function () { this.style.display = 'none'; };
          linkEl.appendChild(iconEl);
        }
        if (linkLabel) {
          var labelSpan = document.createElement('span');
          labelSpan.textContent = linkLabel;
          linkEl.appendChild(labelSpan);
        }
        card.appendChild(linkEl);
      }
    }

    // ---- Settings gear (top-right, authoring mode only, worksheet only) ----
    var showGear = false;
    try {
      var mode = tableau.extensions.environment.mode;
      if (mode === 'authoring') {
        showGear = true;
        try {
          var active = tableau.extensions.workbook.activeSheet;
          if (active && active.sheetType && active.sheetType !== 'worksheet') showGear = false;
        } catch (_) { /* ok */ }
      }
    } catch (_) { /* ok */ }

    if (showGear) {
      var topActions = document.createElement('div');
      topActions.className = 'kpi-top-actions';

      var gearBtn = document.createElement('button');
      gearBtn.className = 'kpi-settings-btn';
      gearBtn.title = 'Settings';
      gearBtn.innerHTML = '&#9881;';
      gearBtn.addEventListener('click', function () {
        openSettingsPanel(kpi, dataResult, _renderLatest || function () { render(dataResult, encodings, sheetName); }, sheetName);
      });
      topActions.appendChild(gearBtn);
      card.appendChild(topActions);
    } else {
      closeSettingsPanel();
    }

    // ---- Title ----
    if (settings.showTitle) {
      var title = document.createElement('div');
      title.className = 'kpi-title';
      title.textContent = settings.titleText || sheetName || kpi.label;
      var ts = parseInt(settings.titleSize, 10);
      if (ts > 0 && ts !== 18) title.style.fontSize = ts + 'px';
      if (settings.cardLayout === 'gauge') title.style.textAlign = 'center';
      card.appendChild(title);
    }

    // ---- Header (label + period) ----
    if (settings.showHeader) {
      var valueLabelText = settings.valueLabel !== undefined
        ? settings.valueLabel
        : kpi.label;
      var showPeriod = settings.showDateRange && kpi.periodLabel;
      var isGaugeHeader = settings.cardLayout === 'gauge';

      if (valueLabelText !== '') {
        var header = document.createElement('div');
        header.className = 'kpi-header';
        if (isGaugeHeader) {
          header.style.justifyContent = 'center';
        } else if (!showPeriod) {
          var align = (settings.subtitleAlign || 'right');
          header.style.justifyContent = align === 'left' ? 'flex-start' : align === 'center' ? 'center' : 'flex-end';
        }
        header.innerHTML =
          '<span class="kpi-label">' + escapeHtml(valueLabelText) + '</span>' +
          (showPeriod
            ? '<span class="kpi-period">' + escapeHtml(kpi.periodLabel) + '</span>'
            : '');
        card.appendChild(header);
      } else if (showPeriod) {
        var header2 = document.createElement('div');
        header2.className = 'kpi-header';
        if (isGaugeHeader) header2.style.justifyContent = 'center';
        header2.innerHTML = '<span class="kpi-period">' + escapeHtml(kpi.periodLabel) + '</span>';
        card.appendChild(header2);
      }
    }

    // ---- Shared delta badge helpers ----
    var rev = settings.reverseDelta;
    var badgePx = parseInt(settings.deltaSize, 10) || 12;
    var badgeStyle = badgePx !== 12 ? 'font-size:' + badgePx + 'px' : '';
    var deltaDec = parseInt(settings.fmtDeltaDecimals, 10);
    var deltaDecPlaces = (deltaDec >= 0) ? deltaDec : 1;

    var isGauge = settings.cardLayout === 'gauge';

    if (isGauge) {
      // ================================================================
      // GAUGE LAYOUT
      // ================================================================

      // Gauge container (needs to be in DOM for width measurement)
      var gaugeSection = document.createElement('div');
      gaugeSection.className = 'kpi-gauge-section';
      card.appendChild(gaugeSection);

      // Delta badges below gauge
      var gaugeBadges = document.createElement('div');
      gaugeBadges.className = 'kpi-gauge-badges';

      if (settings.showDelta && kpi.delta !== null) {
        var deltaEl = document.createElement('span');
        var displayDelta = Math.abs(kpi.delta).toFixed(deltaDecPlaces);
        var isNeutral = parseFloat(displayDelta) === 0;
        var isUp = kpi.delta > 0;
        var sentiment = isNeutral ? 'neutral' : (isUp !== rev) ? 'positive' : 'negative';
        deltaEl.className = 'kpi-delta ' + sentiment;
        if (badgeStyle) deltaEl.style.cssText = badgeStyle;
        var arrow = isNeutral ? '\u2013' : isUp ? '\u25B2' : '\u25BC';
        deltaEl.innerHTML =
          '<span class="kpi-delta-arrow">' + arrow + '</span> ' +
          displayDelta + '% ' + escapeHtml(settings.deltaLabel || 'vs prev');
        gaugeBadges.appendChild(deltaEl);
      }

      var ptdLabelText = settings.ptdLabel !== undefined ? settings.ptdLabel : 'vs prev PTD';
      if (settings.ptdEnabled && kpi.ptdDelta !== null && ptdLabelText !== '') {
        var ptdEl = document.createElement('span');
        var displayPtdDelta = Math.abs(kpi.ptdDelta).toFixed(deltaDecPlaces);
        var ptdNeutral = parseFloat(displayPtdDelta) === 0;
        var ptdUp = kpi.ptdDelta > 0;
        var ptdSentiment = ptdNeutral ? 'neutral' : (ptdUp !== rev) ? 'positive' : 'negative';
        ptdEl.className = 'kpi-delta kpi-delta-ptd ' + ptdSentiment;
        if (badgeStyle) ptdEl.style.cssText = badgeStyle;
        var ptdArrow = ptdNeutral ? '\u2013' : ptdUp ? '\u25B2' : '\u25BC';
        ptdEl.innerHTML =
          '<span class="kpi-delta-arrow">' + ptdArrow + '</span> ' +
          displayPtdDelta + '% ' + escapeHtml(ptdLabelText);
        gaugeBadges.appendChild(ptdEl);
      }

      if (gaugeBadges.childNodes.length > 0) card.appendChild(gaugeBadges);

      // Append to DOM so gauge can measure container width
      content.appendChild(card);

      // Draw gauge
      drawGauge(gaugeSection, kpi, settings);

    } else {
      // ================================================================
      // STANDARD LAYOUT
      // ================================================================

      // ---- Value + delta row ----
      if (settings.showValue || settings.showDelta) {
        var valueRow = document.createElement('div');
        valueRow.className = 'kpi-value-row';

        if (settings.showValue) {
          var valueEl = document.createElement('span');
          valueEl.className = 'kpi-value';
          valueEl.textContent = kpi.formattedValue;
          var vs = parseInt(settings.valueSize, 10);
          if (vs > 0 && vs !== 40) valueEl.style.fontSize = vs + 'px';
          valueRow.appendChild(valueEl);
        }

        if (settings.showDelta && kpi.delta !== null) {
          var deltaEl2 = document.createElement('span');
          var displayDelta2 = Math.abs(kpi.delta).toFixed(deltaDecPlaces);
          var isNeutral2 = parseFloat(displayDelta2) === 0;
          var isUp2 = kpi.delta > 0;
          var sentiment2 = isNeutral2 ? 'neutral' : (isUp2 !== rev) ? 'positive' : 'negative';
          deltaEl2.className = 'kpi-delta ' + sentiment2;
          if (badgeStyle) deltaEl2.style.cssText = badgeStyle;
          var arrow2 = isNeutral2 ? '\u2013' : isUp2 ? '\u25B2' : '\u25BC';
          deltaEl2.innerHTML =
            '<span class="kpi-delta-arrow">' + arrow2 + '</span> ' +
            displayDelta2 + '% ' + escapeHtml(settings.deltaLabel || 'vs prev');
          valueRow.appendChild(deltaEl2);
        }

        // Secondary comparison badge
        var ptdLabelText2 = settings.ptdLabel !== undefined
          ? settings.ptdLabel
          : 'vs prev PTD';

        if (settings.ptdEnabled && kpi.ptdDelta !== null && ptdLabelText2 !== '') {
          var ptdEl2 = document.createElement('span');
          var displayPtdDelta2 = Math.abs(kpi.ptdDelta).toFixed(deltaDecPlaces);
          var ptdNeutral2 = parseFloat(displayPtdDelta2) === 0;
          var ptdUp2 = kpi.ptdDelta > 0;
          var ptdSentiment2 = ptdNeutral2 ? 'neutral' : (ptdUp2 !== rev) ? 'positive' : 'negative';
          ptdEl2.className = 'kpi-delta kpi-delta-ptd ' + ptdSentiment2;
          if (badgeStyle) ptdEl2.style.cssText = badgeStyle;
          var ptdArrow2 = ptdNeutral2 ? '\u2013' : ptdUp2 ? '\u25B2' : '\u25BC';
          ptdEl2.innerHTML =
            '<span class="kpi-delta-arrow">' + ptdArrow2 + '</span> ' +
            displayPtdDelta2 + '% ' + escapeHtml(ptdLabelText2);
          valueRow.appendChild(ptdEl2);
        }

        card.appendChild(valueRow);
      }

      // ---- Primary goal ----
      if (settings.showGoal) {
        var goalLabelText = settings.goalLabel !== undefined
          ? settings.goalLabel
          : 'Goal';
        if (kpi.goalPct !== null && goalLabelText !== '') {
          card.appendChild(buildGoalBar(kpi.goalPct, kpi.formattedGoal, goalLabelText, settings));
        }
      }

      // ---- Secondary goal ----
      if (settings.showGoal2) {
        var goal2LabelText = settings.goal2Label !== undefined
          ? settings.goal2Label
          : kpi.goal2Label;
        if (kpi.goal2Pct !== null && goal2LabelText !== '') {
          card.appendChild(buildGoalBar(kpi.goal2Pct, kpi.formattedGoal2, goal2LabelText, settings));
        }
      }

      // ---- Sparkline placeholder ----
      var sparkSection = null;
      if (settings.showSparkline && kpi.sparkData && kpi.sparkData.length > 0) {
        sparkSection = document.createElement('div');
        sparkSection.className = 'kpi-sparkline-section';
        card.appendChild(sparkSection);
      }

      // Append card to DOM first so sparkline can measure actual width
      content.appendChild(card);

      // Now draw sparkline
      if (sparkSection) {
        var ptdData = settings.ptdEnabled ? kpi.ptdSparkData : null;
        var height = parseInt(settings.sparkHeight, 10) || 130;
        drawSparkline(sparkSection, kpi.sparkData, ptdData, settings, height, kpi.globalFmt || formatNumberCompact);
      }
    }

    // ---- Fill worksheet: zoom card to fit available space ----
    if (isFill) {
      var availW = content.clientWidth - left - right;
      var availH = content.clientHeight - top - bottom;
      var cardW  = card.offsetWidth;
      var cardH  = card.offsetHeight;
      if (cardW > 0 && cardH > 0) {
        var zoomFactor = Math.min(availW / cardW, availH / cardH, 3); // cap at 3x
        if (zoomFactor !== 1) {
          card.style.zoom = zoomFactor;
          card.style.transformOrigin = 'top left';
        }
      }
    }
  }

  // =========================================================================
  // Goal bar builder
  // =========================================================================

  function buildGoalBar (goalPct, formattedGoal, label, settings) {
    var section = document.createElement('div');
    section.className = 'kpi-goal-section';

    var track = document.createElement('div');
    track.className = 'kpi-progress-track';

    var fill = document.createElement('div');
    fill.className = 'kpi-progress-fill';
    var pct = Math.max(0, Math.min(goalPct, 100));

    if (settings && settings._animate) {
      fill.style.width = '0%';
      fill.classList.add('animate');
      requestAnimationFrame(function () { fill.style.width = pct + '%'; });
    } else {
      fill.style.width = pct + '%';
    }

    var barMode = (settings && settings.barColorMode) || 'default';
    if (barMode === 'accent') {
      var accent = (settings && settings.gradColor) || '#D42F8A';
      var stops = buildGradientFromColor(accent);
      fill.style.background = 'linear-gradient(90deg, ' + stops.join(', ') + ')';
    } else if (barMode === 'custom') {
      var custom = (settings && settings.barCustomColor) || '#3b82f6';
      fill.style.background = custom;
    } else {
      fill.style.background = 'linear-gradient(90deg, var(--brand-orange), var(--brand-coral), var(--brand-magenta), var(--brand-purple))';
    }

    track.appendChild(fill);
    section.appendChild(track);

    var meta = document.createElement('div');
    meta.className = 'kpi-goal-meta';
    meta.innerHTML =
      '<span>' + escapeHtml(label) + ': ' + escapeHtml(formattedGoal) + '</span>' +
      '<span class="kpi-goal-pct">' + goalPct + '% of goal</span>';
    section.appendChild(meta);

    return section;
  }

})();
