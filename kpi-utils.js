'use strict';

/**
 * kpi-utils.js — Shared constants, settings management, number formatting, and utilities.
 * Exports onto window.KPI.utils
 */

window.KPI = window.KPI || {};
window.KPI.utils = (function () {

  const KPI_VERSION = '1.0.0';

  // =========================================================================
  // Settings keys & defaults
  // =========================================================================

  const SETTINGS_KEYS = {
    // Visibility toggles
    showTitle:        'kpi_showTitle',
    showHeader:       'kpi_showHeader',
    showValue:        'kpi_showValue',
    showDelta:        'kpi_showDelta',
    showGoal:         'kpi_showGoal',
    showGoal2:        'kpi_showGoal2',
    showSparkline:    'kpi_showSparkline',
    showSparkLabels:  'kpi_showSparkLabels',
    showSparkPeriods: 'kpi_showSparkPeriods',
    showLegend:       'kpi_showLegend',
    showLink:         'kpi_showLink',
    showAnimation:    'kpi_showAnimation',
    // Labels
    titleText:     'kpi_titleText',
    titleSize:     'kpi_titleSize',
    valueLabel:    'kpi_valueLabel',
    deltaLabel:    'kpi_deltaLabel',
    goalLabel:     'kpi_goalLabel',
    goal2Label:    'kpi_goal2Label',
    showDateRange: 'kpi_showDateRange',
    // Comparison
    reverseDelta:  'kpi_reverseDelta',
    // Secondary comparison
    ptdEnabled:    'kpi_ptdEnabled',
    ptdLabel:      'kpi_ptdLabel',
    ptdFieldName:  'kpi_ptdFieldName',
    ptdLegendLabel:'kpi_ptdLegendLabel',
    // Data fields
    dateFieldName: 'kpi_dateFieldName',
    // Sparkline
    sparkHeight:   'kpi_sparkHeight',
    // Badge size
    deltaSize:     'kpi_deltaSize',
    // Global number format
    fmtPrefix:     'kpi_fmtPrefix',
    fmtSuffix:     'kpi_fmtSuffix',
    fmtDecimals:   'kpi_fmtDecimals',
    fmtAbbreviate: 'kpi_fmtAbbreviate',
    fmtDeltaDecimals:'kpi_fmtDeltaDecimals',
    // Layout
    fillWorksheet: 'kpi_fillWorksheet',
    compactMode:   'kpi_compactMode',
    padTop:        'kpi_padTop',
    padLeft:       'kpi_padLeft',
    cardWidth:     'kpi_cardWidth',
    // Gradient
    gradColor:     'kpi_gradColor',
    // Progress bars
    barColorMode:  'kpi_barColorMode',
    barCustomColor:'kpi_barCustomColor',
    // Sparkline colors
    sparklineColorMode: 'kpi_sparklineColorMode',
    sparklineCustomColor: 'kpi_sparklineCustomColor',
    // Card layout
    cardLayout:    'kpi_cardLayout',
    // Value size
    valueSize:     'kpi_valueSize',
    // Link
    linkUrl:       'kpi_linkUrl',
    linkLabel:     'kpi_linkLabel',
    linkIcon:      'kpi_linkIcon'
  };

  const BOOL_DEFAULTS = {
    showTitle: true,
    showHeader: true,
    showValue: true,
    showDelta: true,
    showGoal: true,
    showGoal2: true,
    showSparkline: true,
    showSparkLabels: true,
    showSparkPeriods: true,
    showLegend: true,
    showLink: true,
    showAnimation: true,
    showDateRange: true,
    fillWorksheet: false,
    compactMode: false,
    reverseDelta: false,
    ptdEnabled: false,
    fmtAbbreviate: false
  };

  // =========================================================================
  // Settings load / save
  // =========================================================================

  function loadSettings () {
    var s = tableau.extensions.settings;
    var result = {};

    for (var _i = 0, _a = Object.entries(BOOL_DEFAULTS); _i < _a.length; _i++) {
      var key = _a[_i][0], def = _a[_i][1];
      var raw = s.get(SETTINGS_KEYS[key]);
      if (raw === null || raw === undefined) {
        result[key] = def;
      } else {
        result[key] = raw === 'true';
      }
    }

    var stringKeys = [
      'titleText', 'titleSize',
      'valueLabel', 'deltaLabel', 'goalLabel', 'goal2Label', 'ptdLabel',
      'ptdFieldName', 'ptdLegendLabel', 'dateFieldName', 'sparkHeight', 'deltaSize',
      'fmtPrefix', 'fmtSuffix', 'fmtDecimals', 'fmtDeltaDecimals',
      'padTop', 'padLeft', 'cardWidth',
      'gradColor', 'barColorMode', 'barCustomColor', 'sparklineColorMode', 'sparklineCustomColor',
      'cardLayout', 'valueSize',
      'linkUrl', 'linkLabel', 'linkIcon'
    ];
    for (var _j = 0; _j < stringKeys.length; _j++) {
      var k = stringKeys[_j];
      result[k] = s.get(SETTINGS_KEYS[k]) ?? undefined;
    }

    return result;
  }

  // Module-level animation flag — shared across settings & render
  var _skipAnimation = false;

  function setSkipAnimation (val) { _skipAnimation = val; }
  function getSkipAnimation ()    { return _skipAnimation; }

  async function saveSettings (values, renderCb) {
    var s = tableau.extensions.settings;
    for (var _i = 0, _a = Object.entries(SETTINGS_KEYS); _i < _a.length; _i++) {
      var key = _a[_i][0], settingKey = _a[_i][1];
      if (values[key] !== undefined) {
        s.set(settingKey, String(values[key]));
      } else {
        s.erase(settingKey);
      }
    }
    await s.saveAsync();
    if (renderCb) {
      _skipAnimation = true;
      renderCb();
    }
  }

  // =========================================================================
  // Number formatting
  // =========================================================================

  function abbreviateNumber (n, decimals, prefix, suffix) {
    if (n == null) return '';
    var abs = Math.abs(n);
    var short, tag;
    if (abs >= 1e9)      { short = n / 1e9; tag = 'B'; }
    else if (abs >= 1e6) { short = n / 1e6; tag = 'M'; }
    else if (abs >= 1e3) { short = n / 1e3; tag = 'K'; }
    else                 { short = n;        tag = '';  }
    var dec = (decimals >= 0) ? decimals : (tag ? 1 : 0);
    var num = short.toFixed(dec).replace(/\.0+$/, '');
    return (prefix || '') + num + tag + (suffix || '');
  }

  function buildGlobalFormatter (settings) {
    var prefix   = settings.fmtPrefix || '';
    var suffix   = settings.fmtSuffix || '';
    var decRaw   = parseInt(settings.fmtDecimals, 10);
    var decimals = (decRaw >= 0) ? decRaw : 0;
    var abbrev   = !!settings.fmtAbbreviate;

    return function (n) {
      if (n == null) return '';
      if (abbrev) {
        return abbreviateNumber(n, decimals, prefix, suffix);
      }
      var formatted = Number(n).toLocaleString(undefined, {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
        useGrouping: true
      });
      return (prefix || '') + formatted + (suffix || '');
    };
  }

  function formatNumberCompact (n) {
    if (n == null) return '';
    var abs = Math.abs(n);
    if (abs >= 1e6) return (n / 1e6).toFixed(1).replace(/\.0$/, '') + 'M';
    if (abs >= 1e3) return (n / 1e3).toFixed(1).replace(/\.0$/, '') + 'K';
    if (Number.isInteger(n)) return n.toLocaleString();
    return n.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  }

  // =========================================================================
  // Color gradient from single hex
  // =========================================================================

  function buildGradientFromColor (hex) {
    var r = parseInt(hex.slice(1, 3), 16) / 255;
    var g = parseInt(hex.slice(3, 5), 16) / 255;
    var b = parseInt(hex.slice(5, 7), 16) / 255;

    var max = Math.max(r, g, b), min = Math.min(r, g, b);
    var h, s, l = (max + min) / 2;
    if (max === min) { h = s = 0; }
    else {
      var d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
      else if (max === g) h = ((b - r) / d + 2) / 6;
      else h = ((r - g) / d + 4) / 6;
    }

    function hslToHex (hh, ss, ll) {
      hh = ((hh % 1) + 1) % 1;
      var hue2rgb = function (p, q, t) {
        if (t < 0) t += 1; if (t > 1) t -= 1;
        if (t < 1/6) return p + (q - p) * 6 * t;
        if (t < 1/2) return q;
        if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
        return p;
      };
      var rr, gg, bb;
      if (ss === 0) { rr = gg = bb = ll; }
      else {
        var q = ll < 0.5 ? ll * (1 + ss) : ll + ss - ll * ss;
        var p = 2 * ll - q;
        rr = hue2rgb(p, q, hh + 1/3);
        gg = hue2rgb(p, q, hh);
        bb = hue2rgb(p, q, hh - 1/3);
      }
      var toHex = function (v) { return Math.round(v * 255).toString(16).padStart(2, '0'); };
      return '#' + toHex(rr) + toHex(gg) + toHex(bb);
    }

    var c1 = hslToHex(h - 0.11, Math.min(1, s + 0.1), Math.min(0.65, l + 0.1));
    var c2 = hslToHex(h, s, l);
    var c3 = hslToHex(h + 0.11, Math.min(1, s + 0.05), Math.max(0.3, l - 0.08));
    return [c1, c2, c3];
  }

  // =========================================================================
  // HTML escape
  // =========================================================================

  function escapeHtml (str) {
    var div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  // =========================================================================
  // Public API
  // =========================================================================

  return {
    KPI_VERSION: KPI_VERSION,
    SETTINGS_KEYS: SETTINGS_KEYS,
    BOOL_DEFAULTS: BOOL_DEFAULTS,
    loadSettings: loadSettings,
    saveSettings: saveSettings,
    setSkipAnimation: setSkipAnimation,
    getSkipAnimation: getSkipAnimation,
    abbreviateNumber: abbreviateNumber,
    buildGlobalFormatter: buildGlobalFormatter,
    formatNumberCompact: formatNumberCompact,
    buildGradientFromColor: buildGradientFromColor,
    escapeHtml: escapeHtml
  };

})();
