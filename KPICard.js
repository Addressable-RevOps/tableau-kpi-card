'use strict';
/* global d3 */

(function () {
  const KPI_VERSION = '1.0.0';

  // Module-level reference so settings callbacks always use the latest data
  let _renderLatest = null;
  // Module-level refetch + render (for reload button)
  let _refetchAndRender = null;
  // Skip animation on settings-triggered re-renders
  let _skipAnimation = false;
  // Track previous showAnimation state to detect toggle-on
  let _prevShowAnimation = null;

  // =========================================================================
  // Bootstrap
  // =========================================================================

  /** Persistent settings keys */
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
    fmtPrefix:     'kpi_fmtPrefix',       // e.g. "$"
    fmtSuffix:     'kpi_fmtSuffix',       // e.g. "%"
    fmtDecimals:   'kpi_fmtDecimals',     // -1 = auto, 0-6 explicit
    fmtAbbreviate: 'kpi_fmtAbbreviate',   // true = 120K / 1.2M
    fmtDeltaDecimals:'kpi_fmtDeltaDecimals', // decimal places on delta % badge
    // Layout
    padTop:        'kpi_padTop',
    padLeft:       'kpi_padLeft',
    cardWidth:     'kpi_cardWidth',
    // Gradient
    gradColor:     'kpi_gradColor',
    // Progress bars
    barColorMode:  'kpi_barColorMode',   // 'default' | 'accent' | 'custom'
    barCustomColor:'kpi_barCustomColor',
    // Sparkline colors
    sparklineColorMode: 'kpi_sparklineColorMode', // 'default' | 'accent' | 'custom'
    sparklineCustomColor: 'kpi_sparklineCustomColor',
    // Value size
    valueSize:     'kpi_valueSize',
    // Link
    linkUrl:       'kpi_linkUrl',
    linkLabel:     'kpi_linkLabel',
    linkIcon:      'kpi_linkIcon'
  };

  /** Defaults for boolean toggles (all default to true except reverseDelta/ptdEnabled) */
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
    reverseDelta: false,
    ptdEnabled: false,
    fmtAbbreviate: false
  };

  /** Read all custom-label settings */
  function loadSettings () {
    const s = tableau.extensions.settings;
    const result = {};

    // Load booleans
    for (const [key, def] of Object.entries(BOOL_DEFAULTS)) {
      const raw = s.get(SETTINGS_KEYS[key]);
      if (raw === null || raw === undefined) {
        result[key] = def;
      } else {
        result[key] = raw === 'true';
      }
    }

    // Load strings
    const stringKeys = [
      'titleText', 'titleSize',
      'valueLabel', 'deltaLabel', 'goalLabel', 'goal2Label', 'ptdLabel',
      'ptdFieldName', 'ptdLegendLabel', 'dateFieldName', 'sparkHeight', 'deltaSize',
      'fmtPrefix', 'fmtSuffix', 'fmtDecimals', 'fmtDeltaDecimals',
      'padTop', 'padLeft', 'cardWidth',
      'gradColor', 'barColorMode', 'barCustomColor', 'sparklineColorMode', 'sparklineCustomColor',
      'valueSize',
      'linkUrl', 'linkLabel', 'linkIcon'
    ];
    for (const key of stringKeys) {
      result[key] = s.get(SETTINGS_KEYS[key]) ?? undefined;
    }

    return result;
  }

  /** Persist settings and re-render. */
  async function saveSettings (values, renderCb) {
    const s = tableau.extensions.settings;
    for (const [key, settingKey] of Object.entries(SETTINGS_KEYS)) {
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

  window.onload = tableau.extensions.initializeAsync().then(() => {
    const worksheet = tableau.extensions.worksheetContent.worksheet;

    let cachedResult = { rows: [], columns: [] };
    let cachedEncodings = {};
    const sheetName = worksheet.name;

    // Central render with latest cached data — used by settings onSave, resize, etc.
    const renderLatest = () => render(cachedResult, cachedEncodings, sheetName);
    _renderLatest = renderLatest;

    const updateDataAndRender = async () => {
      [cachedResult, cachedEncodings] = await Promise.all([
        getSummaryDataTable(worksheet),
        getEncodingMap(worksheet)
      ]);
      renderLatest();
    };

    _refetchAndRender = updateDataAndRender;

    onresize = () => { _skipAnimation = true; renderLatest(); };

    worksheet.addEventListener(
      tableau.TableauEventType.SummaryDataChanged,
      updateDataAndRender
    );

    updateDataAndRender();
  });

  // =========================================================================
  // Data helpers
  // =========================================================================

  function convertToListOfNamedRows (dataTablePage) {
    const rows = [];
    const columns = dataTablePage.columns;
    const data = dataTablePage.data;
    for (let i = 0; i < data.length; i++) {
      const row = {};
      for (let j = 0; j < columns.length; j++) {
        row[columns[j].fieldName] = data[i][columns[j].index];
      }
      rows.push(row);
    }
    return { rows, columns };
  }

  async function getSummaryDataTable (worksheet) {
    let allRows = [];
    let allColumns = [];
    const dataTableReader = await worksheet.getSummaryDataReaderAsync(
      undefined,
      { ignoreSelection: true }
    );
    for (let page = 0; page < dataTableReader.pageCount; page++) {
      const dataTablePage = await dataTableReader.getPageAsync(page);
      const result = convertToListOfNamedRows(dataTablePage);
      allRows = allRows.concat(result.rows);
      if (page === 0) allColumns = result.columns;
    }
    await dataTableReader.releaseAsync();
    return { rows: allRows, columns: allColumns };
  }

  async function getEncodingMap (worksheet) {
    const visualSpec = await worksheet.getVisualSpecificationAsync();
    const map = {};
    if (visualSpec.activeMarksSpecificationIndex < 0) return map;
    const marksCard =
      visualSpec.marksSpecifications[visualSpec.activeMarksSpecificationIndex];
    for (const enc of marksCard.encodings) {
      map[enc.id] = enc.field;
    }
    return map;
  }

  // =========================================================================
  // Resolve encodings -> data columns
  // =========================================================================

  function resolveEncodings (encodingMap, columns, settings) {
    const resolved = {
      valueCol: null, goalCol: null, goal2Col: null, dateCol: null, ptdCol: null
    };
    const colNameSet = new Set(columns.map(c => c.fieldName));

    function findCol (field) {
      if (!field) return null;
      if (colNameSet.has(field.name)) return field.name;
      for (const cn of colNameSet) {
        if (cn.includes(field.name) || field.name.includes(cn)) return cn;
      }
      return null;
    }

    // 1. Custom encoding IDs
    resolved.valueCol = findCol(encodingMap.value);
    resolved.goalCol = findCol(encodingMap.goal);
    resolved.goal2Col = findCol(encodingMap.goal2);
    resolved.dateCol = findCol(encodingMap.date);

    // 2. User-selected date field (overrides auto-detection)
    if (settings && settings.dateFieldName) {
      if (colNameSet.has(settings.dateFieldName)) {
        resolved.dateCol = settings.dateFieldName;
      } else {
        for (const cn of colNameSet) {
          if (cn.toLowerCase().includes(settings.dateFieldName.toLowerCase()) ||
              settings.dateFieldName.toLowerCase().includes(cn.toLowerCase())) {
            resolved.dateCol = cn;
            break;
          }
        }
      }
    }

    // 3. Fallback: find dimension on any encoding for date
    if (!resolved.dateCol) {
      for (const [id, field] of Object.entries(encodingMap)) {
        if (!field || id === 'value' || id === 'goal' || id === 'goal2') continue;
        if (field.role === 'dimension') {
          const col = findCol(field);
          if (col) {
            resolved.dateCol = col;
            break;
          }
        }
      }
    }

    // 4. Fallback: find extra measures for goals
    const usedMeasures = new Set(
      [resolved.valueCol, resolved.goalCol, resolved.goal2Col].filter(Boolean)
    );
    if (!resolved.goalCol && resolved.valueCol) {
      for (const [id, field] of Object.entries(encodingMap)) {
        if (!field || id === 'value') continue;
        if (field.role === 'measure') {
          const col = findCol(field);
          if (col && !usedMeasures.has(col)) {
            resolved.goalCol = col;
            usedMeasures.add(col);
            break;
          }
        }
      }
    }
    if (!resolved.goal2Col && resolved.valueCol) {
      for (const [id, field] of Object.entries(encodingMap)) {
        if (!field || id === 'value' || id === 'goal') continue;
        if (field.role === 'measure') {
          const col = findCol(field);
          if (col && !usedMeasures.has(col)) {
            resolved.goal2Col = col;
            usedMeasures.add(col);
            break;
          }
        }
      }
    }

    resolved.valueField = encodingMap.value || null;
    resolved.goal2Field = encodingMap.goal2 || null;

    // 5. PTD field from settings
    if (settings && settings.ptdFieldName) {
      if (colNameSet.has(settings.ptdFieldName)) {
        resolved.ptdCol = settings.ptdFieldName;
      } else {
        for (const cn of colNameSet) {
          const cnLower = cn.toLowerCase();
          const searchLower = settings.ptdFieldName.toLowerCase();
          if (cnLower.includes(searchLower) || searchLower.includes(cnLower)) {
            resolved.ptdCol = cn;
            break;
          }
        }
      }
    }

    // Detect the date column name for quarter formatting hints
    resolved.dateFieldName = '';
    if (encodingMap.date) resolved.dateFieldName = encodingMap.date.name || '';
    else {
      for (const [, field] of Object.entries(encodingMap)) {
        if (field && field.role === 'dimension') {
          resolved.dateFieldName = field.name || '';
          break;
        }
      }
    }

    return resolved;
  }

  // =========================================================================
  // Quarter label formatting
  // =========================================================================

  function isQuarterDateField (dateFieldName) {
    const n = (dateFieldName || '').toUpperCase();
    return n.includes('QUARTER') || n.includes('QTR');
  }

  function formatPeriodLabel (rawLabel, sortKey, dateFieldName) {
    const s = String(rawLabel).trim();

    const m1 = s.match(/^Q(\d)\s+(\d{4})$/i);
    if (m1) return m1[2] + ' Q' + m1[1];

    const m2 = s.match(/^(\d{4})\s+Q(\d)$/i);
    if (m2) return s;

    const m3 = s.match(/^(\d{4})-Q(\d)$/i);
    if (m3) return m3[1] + ' Q' + m3[2];

    if (/^[1-4]$/.test(s) && isQuarterDateField(dateFieldName)) {
      return 'Q' + s;
    }

    return s;
  }

  // =========================================================================
  // Period-to-date helpers
  // =========================================================================

  function getPeriodStartDate (periodLabel) {
    let year, quarter;
    const mA = periodLabel.match(/^(\d{4})\s*Q(\d)$/i);
    const mB = periodLabel.match(/^Q(\d)\s+(\d{4})$/i);
    const mC = periodLabel.match(/^(\d{4})-Q(\d)$/i);
    if (mA) { year = +mA[1]; quarter = +mA[2]; }
    else if (mB) { quarter = +mB[1]; year = +mB[2]; }
    else if (mC) { year = +mC[1]; quarter = +mC[2]; }

    if (year && quarter >= 1 && quarter <= 4) {
      return new Date(year, (quarter - 1) * 3, 1);
    }

    const mY = periodLabel.match(/^(\d{4})$/);
    if (mY) return new Date(+mY[1], 0, 1);

    const mMonth = periodLabel.match(/^(\d{4})-(\d{2})$/) ||
                   periodLabel.match(/^(\w+)\s+(\d{4})$/i);
    if (mMonth) {
      let mYear, mMon;
      if (/^\d{4}-\d{2}$/.test(periodLabel)) {
        mYear = +mMonth[1]; mMon = +mMonth[2] - 1;
      } else {
        const parsed = new Date(periodLabel + ' 1');
        if (!isNaN(parsed)) { mYear = parsed.getFullYear(); mMon = parsed.getMonth(); }
      }
      if (mYear && mMon != null) return new Date(mYear, mMon, 1);
    }
    return null;
  }

  function getPeriodEndDate (periodLabel) {
    let year, quarter;
    const mA = periodLabel.match(/^(\d{4})\s*Q(\d)$/i);
    const mB = periodLabel.match(/^Q(\d)\s+(\d{4})$/i);
    const mC = periodLabel.match(/^(\d{4})-Q(\d)$/i);
    if (mA) { year = +mA[1]; quarter = +mA[2]; }
    else if (mB) { quarter = +mB[1]; year = +mB[2]; }
    else if (mC) { year = +mC[1]; quarter = +mC[2]; }

    if (year && quarter >= 1 && quarter <= 4) {
      return new Date(year, (quarter - 1) * 3 + 3, 0);
    }

    const mY = periodLabel.match(/^(\d{4})$/);
    if (mY) return new Date(+mY[1], 11, 31);

    const mMonth = periodLabel.match(/^(\d{4})-(\d{2})$/) ||
                   periodLabel.match(/^(\w+)\s+(\d{4})$/i);
    if (mMonth) {
      let mYear, mMon;
      if (/^\d{4}-\d{2}$/.test(periodLabel)) {
        mYear = +mMonth[1]; mMon = +mMonth[2] - 1;
      } else {
        const parsed = new Date(periodLabel + ' 1');
        if (!isNaN(parsed)) { mYear = parsed.getFullYear(); mMon = parsed.getMonth(); }
      }
      if (mYear && mMon != null) return new Date(mYear, mMon + 1, 0);
    }
    return null;
  }

  // Infer previous period label (e.g., "2026 Q1" -> "2025 Q4")
  function inferPreviousPeriodLabel (currentLabel) {
    const s = String(currentLabel).trim();
    // Quarter format: "2026 Q1" or "2026 Q1" or "2026-Q1"
    const mQ = s.match(/^(\d{4})\s*Q(\d)$/i) || s.match(/^(\d{4})-Q(\d)$/i);
    if (mQ) {
      let year = +mQ[1];
      let quarter = +mQ[2];
      if (quarter === 1) {
        year--;
        quarter = 4;
      } else {
        quarter--;
      }
      return year + ' Q' + quarter;
    }
    // Year format: "2026"
    const mY = s.match(/^(\d{4})$/);
    if (mY) {
      return String(+mY[1] - 1);
    }
    // Month format: "2026-01" or "January 2026"
    const mM = s.match(/^(\d{4})-(\d{2})$/);
    if (mM) {
      let year = +mM[1];
      let month = +mM[2];
      if (month === 1) {
        year--;
        month = 12;
      } else {
        month--;
      }
      return year + '-' + String(month).padStart(2, '0');
    }
    // Fallback: just "Previous"
    return 'Previous';
  }

  // =========================================================================
  // KPI computation
  // =========================================================================

  function computeKpi (dataResult, encodings, settings) {
    const { rows: data, columns } = dataResult;
    const r = resolveEncodings(encodings, columns, settings);

    if (!r.valueCol) return null;

    const valueLabel = r.valueField ? r.valueField.name : r.valueCol;

    // ------------------------------------------------------------------
    // Build the global number formatter from settings
    // ------------------------------------------------------------------
    const s = settings || {};
    const globalFmt = buildGlobalFormatter(s);

    // ------------------------------------------------------------------
    // 1. No date column — simple aggregate
    // ------------------------------------------------------------------
    if (!r.dateCol) {
      let totalValue = 0, totalGoal = 0, totalGoal2 = 0;
      for (const row of data) {
        totalValue += Number(row[r.valueCol]?.value) || 0;
        if (r.goalCol) totalGoal += Number(row[r.goalCol]?.value) || 0;
        if (r.goal2Col) totalGoal2 += Number(row[r.goal2Col]?.value) || 0;
      }
      return {
        label: valueLabel,
        currentValue: totalValue,
        formattedValue: globalFmt(totalValue),
        previousValue: null,
        delta: null,
        ptdDelta: null,
        goalValue: r.goalCol ? totalGoal : null,
        goalPct: r.goalCol && totalGoal !== 0
          ? Math.round((totalValue / totalGoal) * 100) : null,
        formattedGoal: r.goalCol ? globalFmt(totalGoal) : null,
        goal2Value: r.goal2Col ? totalGoal2 : null,
        goal2Pct: r.goal2Col && totalGoal2 !== 0
          ? Math.round((totalValue / totalGoal2) * 100) : null,
        formattedGoal2: r.goal2Col ? globalFmt(totalGoal2) : null,
        goal2Label: r.goal2Field ? r.goal2Field.name : 'Secondary Goal',
        sparkData: null,
        periodLabel: null,
        globalFmt
      };
    }

    // ------------------------------------------------------------------
    // 2. Group rows by date period
    // ------------------------------------------------------------------
    const periodMap = new Map();

    for (const row of data) {
      const dateVal = row[r.dateCol];
      if (!dateVal) continue;
      const rawValue = dateVal.value;
      if (rawValue == null || rawValue === '' || String(rawValue).toLowerCase() === 'null') continue;
      const rawLabel = String(dateVal.formattedValue ?? rawValue);
      if (rawLabel === '' || rawLabel.toLowerCase() === 'null') continue;

      const sortKey = rawValue;
      const periodKey = rawLabel;

      let rowDate = null;
      if (rawValue instanceof Date) rowDate = new Date(rawValue);
      else if (typeof rawValue === 'number') rowDate = new Date(rawValue);
      else if (typeof rawValue === 'string') {
        const parsed = new Date(rawValue);
        if (!isNaN(parsed)) rowDate = parsed;
      }

      if (!periodMap.has(periodKey)) {
        periodMap.set(periodKey, {
          value: 0, goal: 0, goal2: 0, ptdValue: 0, sortKey,
          label: formatPeriodLabel(rawLabel, sortKey, r.dateFieldName),
          rows: []
        });
      }
      const bucket = periodMap.get(periodKey);

      bucket.rows.push({
        date: rowDate,
        value: Number(row[r.valueCol]?.value) || 0,
        goal: r.goalCol ? Number(row[r.goalCol]?.value) || 0 : 0,
        goal2: r.goal2Col ? Number(row[r.goal2Col]?.value) || 0 : 0
      });

      bucket.value += Number(row[r.valueCol]?.value) || 0;
      if (r.goalCol) bucket.goal += Number(row[r.goalCol]?.value) || 0;
      if (r.goal2Col) bucket.goal2 += Number(row[r.goal2Col]?.value) || 0;
      if (r.ptdCol) bucket.ptdValue += Number(row[r.ptdCol]?.value) || 0;
    }

    const allPeriods = Array.from(periodMap.values()).sort((a, b) => {
      if (a.sortKey < b.sortKey) return -1;
      if (a.sortKey > b.sortKey) return 1;
      return 0;
    });

    if (allPeriods.length === 0) return null;

    const periods = allPeriods.slice(-4);
    const current = periods[periods.length - 1];
    const previous = periods.length > 1 ? periods[periods.length - 2] : null;

    // ------------------------------------------------------------------
    // 3. PTD values (field-based or day-by-day fallback)
    // ------------------------------------------------------------------
    if (!r.ptdCol) {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const currentStart = getPeriodStartDate(current.label);
      let elapsedDays = null;
      if (currentStart) {
        const currentEnd = getPeriodEndDate(current.label);
        if (currentEnd && today <= currentEnd) {
          elapsedDays = Math.floor((today - currentStart) / 864e5) + 1;
        }
      }
      for (const p of periods) {
        p.ptdValue = 0;
        if (elapsedDays === null || !p.rows || p.rows.length === 0) {
          p.ptdValue = null; continue;
        }
        const periodStart = getPeriodStartDate(p.label);
        if (!periodStart) { p.ptdValue = null; continue; }
        for (const row of p.rows) {
          if (!row.date) continue;
          const dayOfPeriod = Math.floor((row.date - periodStart) / 864e5) + 1;
          if (dayOfPeriod >= 1 && dayOfPeriod <= elapsedDays) p.ptdValue += row.value;
        }
      }
    }

    // Delta vs previous period
    let delta = null;
    if (previous && previous.value !== 0) {
      delta = ((current.value - previous.value) / Math.abs(previous.value)) * 100;
    }

    // Goals
    let goalPct = null, goalValue = null, formattedGoal = null;
    if (r.goalCol) {
      goalValue = current.goal;
      formattedGoal = globalFmt(goalValue);
      if (goalValue !== 0) goalPct = Math.round((current.value / goalValue) * 100);
    }
    let goal2Pct = null, goal2Value = null, formattedGoal2 = null;
    if (r.goal2Col) {
      goal2Value = current.goal2;
      formattedGoal2 = globalFmt(goal2Value);
      if (goal2Value !== 0) goal2Pct = Math.round((current.value / goal2Value) * 100);
    }

    // PTD delta
    let ptdDelta = null;
    if (r.ptdCol && previous && previous.ptdValue !== 0) {
      ptdDelta = ((current.value - previous.ptdValue) / Math.abs(previous.ptdValue)) * 100;
    } else if (previous && current.ptdValue !== null && previous.ptdValue !== null && previous.ptdValue !== 0) {
      ptdDelta = ((current.ptdValue - previous.ptdValue) / Math.abs(previous.ptdValue)) * 100;
    }

    let sparkData = periods.map(p => ({ label: p.label, value: p.value }));

    // If only one data point, prepend a dummy previous period with 0
    if (sparkData.length === 1) {
      const prevLabel = inferPreviousPeriodLabel(sparkData[0].label);
      sparkData.unshift({ label: prevLabel, value: 0 });
    }

    let ptdSparkData = null;
    if (periods.every(p => p.ptdValue !== null && p.ptdValue !== undefined)) {
      ptdSparkData = periods.map(p => ({ label: p.label, value: p.ptdValue }));
      // If only one PTD data point, prepend dummy previous period
      if (ptdSparkData.length === 1) {
        const prevLabel = inferPreviousPeriodLabel(ptdSparkData[0].label);
        ptdSparkData.unshift({ label: prevLabel, value: 0 });
      }
    }

    const periodLabel = previous
      ? `${periods[0].label} — ${current.label}`
      : current.label;

    return {
      label: valueLabel,
      currentValue: current.value,
      formattedValue: globalFmt(current.value),
      previousValue: previous ? previous.value : null,
      delta,
      ptdDelta,
      goalValue, goalPct, formattedGoal,
      goal2Value, goal2Pct, formattedGoal2,
      goal2Label: r.goal2Field ? r.goal2Field.name : 'Secondary Goal',
      sparkData,
      ptdSparkData,
      periodLabel,
      globalFmt
    };
  }

  // =========================================================================
  // Number formatting
  // =========================================================================

  /**
   * Abbreviate a number with K / M / B suffixes.
   * Respects the given decimal places and wraps with prefix/suffix.
   */
  function abbreviateNumber (n, decimals, prefix, suffix) {
    if (n == null) return '';
    const abs = Math.abs(n);
    let short, tag;
    if (abs >= 1e9)      { short = n / 1e9; tag = 'B'; }
    else if (abs >= 1e6) { short = n / 1e6; tag = 'M'; }
    else if (abs >= 1e3) { short = n / 1e3; tag = 'K'; }
    else                 { short = n;        tag = '';  }
    const dec = (decimals >= 0) ? decimals : (tag ? 1 : 0);
    const num = short.toFixed(dec).replace(/\.0+$/, '');
    return (prefix || '') + num + tag + (suffix || '');
  }

  /**
   * Build the ONE global formatter used for every number in the KPI card.
   *
   * @param {object} settings    - user settings
   * @param {object|null} auto   - auto-detected pattern from Tableau data
   * @returns {Function} formatter(n) -> string
   */
  function buildGlobalFormatter (settings) {
    const prefix   = settings.fmtPrefix || '';
    const suffix   = settings.fmtSuffix || '';
    const decRaw   = parseInt(settings.fmtDecimals, 10);
    const decimals = (decRaw >= 0) ? decRaw : 0;
    const abbrev   = !!settings.fmtAbbreviate;

    return function (n) {
      if (n == null) return '';

      // Abbreviations: 120K, 1.2M, 3.5B
      if (abbrev) {
        return abbreviateNumber(n, decimals, prefix, suffix);
      }

      // Full number with locale grouping
      const formatted = Number(n).toLocaleString(undefined, {
        minimumFractionDigits: decimals,
        maximumFractionDigits: decimals,
        useGrouping: true
      });
      return (prefix || '') + formatted + (suffix || '');
    };
  }

  /** Compact fallback (K/M) — used when no settings configured */
  function formatNumberCompact (n) {
    if (n == null) return '';
    const abs = Math.abs(n);
    if (abs >= 1e6) return (n / 1e6).toFixed(1).replace(/\.0$/, '') + 'M';
    if (abs >= 1e3) return (n / 1e3).toFixed(1).replace(/\.0$/, '') + 'K';
    if (Number.isInteger(n)) return n.toLocaleString();
    return n.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 2 });
  }

  // =========================================================================
  // Settings panel
  // =========================================================================

  let _settingsOverlay = null;

  function closeSettingsPanel () {
    if (_settingsOverlay) {
      _settingsOverlay.remove();
      _settingsOverlay = null;
    }
  }

  function openSettingsPanel (kpi, dataResult, onSave, sheetName) {
    if (_settingsOverlay) { closeSettingsPanel(); return; }

    const settings = loadSettings();
    let _debounceTimer = null;

    // Available fields for dropdowns
    const allFields = (dataResult && dataResult.columns)
      ? dataResult.columns.map(c => ({ value: c.fieldName, label: c.fieldName }))
      : [];
    const measures = (dataResult && dataResult.columns)
      ? dataResult.columns
          .filter(c => c.dataType === 'float' || c.dataType === 'int')
          .map(c => ({ value: c.fieldName, label: c.fieldName }))
      : [];

    // Container
    const overlay = document.createElement('div');
    overlay.className = 'settings-overlay';
    _settingsOverlay = overlay;

    const panel = document.createElement('div');
    panel.className = 'settings-panel';

    // Title row
    const titleRow = document.createElement('div');
    titleRow.className = 'settings-title';
    titleRow.innerHTML = '<svg viewBox="0 0 20 20" fill="currentColor"><path fill-rule="evenodd" d="M11.49 3.17c-.38-1.56-2.6-1.56-2.98 0a1.532 1.532 0 01-2.286.948c-1.372-.836-2.942.734-2.106 2.106.54.886.061 2.042-.947 2.287-1.561.379-1.561 2.6 0 2.978a1.532 1.532 0 01.947 2.287c-.836 1.372.734 2.942 2.106 2.106a1.532 1.532 0 012.287.947c.379 1.561 2.6 1.561 2.978 0a1.533 1.533 0 012.287-.947c1.372.836 2.942-.734 2.106-2.106a1.533 1.533 0 01.947-2.287c1.561-.379 1.561-2.6 0-2.978a1.532 1.532 0 01-.947-2.287c.836-1.372-.734-2.942-2.106-2.106a1.532 1.532 0 01-2.287-.947zM10 13a3 3 0 100-6 3 3 0 000 6z" clip-rule="evenodd"/></svg> <span style="flex:1">Settings</span>';

    const closeBtn = document.createElement('button');
    closeBtn.className = 'settings-btn settings-btn-cancel';
    closeBtn.style.cssText = 'padding:4px 10px;font-size:16px;line-height:1;';
    closeBtn.textContent = '\u00D7';
    closeBtn.addEventListener('click', closeSettingsPanel);
    titleRow.appendChild(closeBtn);
    panel.appendChild(titleRow);

    // Live-update helper
    function onInput () {
      clearTimeout(_debounceTimer);
      _debounceTimer = setTimeout(async () => {
        const inputs = panel.querySelectorAll('input[data-key], select[data-key]');
        const newSettings = {};
        inputs.forEach(inp => {
          if (inp.type === 'checkbox') {
            newSettings[inp.dataset.key] = inp.checked;
          } else if (inp.dataset.numVal !== undefined) {
            // Stepper with possible "Auto" display — use numeric backing value
            newSettings[inp.dataset.key] = inp.dataset.numVal;
          } else {
            newSettings[inp.dataset.key] = inp.value;
          }
        });
        await saveSettings(newSettings, onSave);
      }, 300);
    }

    // ----- Builders -----

    // Section collapse state is remembered in sessionStorage so it
    // persists while the settings panel is opened/closed during a session.
    const SECTION_STORE_KEY = '_kpi_sections';
    function _getSectionStates () {
      try { return JSON.parse(sessionStorage.getItem(SECTION_STORE_KEY)) || {}; }
      catch (_) { return {}; }
    }
    function _saveSectionState (id, open) {
      const states = _getSectionStates();
      states[id] = open;
      try { sessionStorage.setItem(SECTION_STORE_KEY, JSON.stringify(states)); }
      catch (_) { /* ok */ }
    }

    function addSection (title) {
      // Derive a stable key from the title
      const sectionId = title.replace(/\s+/g, '_').toLowerCase();
      // Default: collapsed.  If user previously expanded it, restore that.
      const savedStates = _getSectionStates();
      const isOpen = savedStates[sectionId] === true;

      const sec = document.createElement('div');
      sec.className = 'settings-section' + (isOpen ? '' : ' collapsed');
      const hdr = document.createElement('div');
      hdr.className = 'settings-section-header';

      const titleSpan = document.createElement('span');
      titleSpan.textContent = title;
      hdr.appendChild(titleSpan);

      const chevron = document.createElement('span');
      chevron.className = 'settings-section-chevron';
      chevron.textContent = '\u25BE'; // ▾
      hdr.appendChild(chevron);

      const body = document.createElement('div');
      body.className = 'settings-section-body';
      body.style.maxHeight = isOpen ? 'none' : '0';

      hdr.addEventListener('click', () => {
        const nowCollapsed = sec.classList.toggle('collapsed');
        _saveSectionState(sectionId, !nowCollapsed);
        if (!nowCollapsed) {
          body.style.maxHeight = body.scrollHeight + 'px';
          setTimeout(() => { if (!sec.classList.contains('collapsed')) body.style.maxHeight = 'none'; }, 260);
        } else {
          body.style.maxHeight = body.scrollHeight + 'px';
          body.offsetHeight; // eslint-disable-line no-unused-expressions
          body.style.maxHeight = '0';
        }
      });

      sec.appendChild(hdr);
      sec.appendChild(body);
      panel.appendChild(sec);
      return body;
    }

    function addToggle (parent, labelText, key, hint) {
      const group = document.createElement('div');
      group.className = 'settings-group';

      const row = document.createElement('div');
      row.className = 'settings-toggle-row';

      const lbl = document.createElement('span');
      lbl.className = 'settings-toggle-label';
      lbl.textContent = labelText;
      row.appendChild(lbl);

      const sw = document.createElement('label');
      sw.className = 'toggle-switch';
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.checked = !!settings[key];
      cb.dataset.key = key;
      cb.addEventListener('change', onInput);
      sw.appendChild(cb);
      const track = document.createElement('span');
      track.className = 'toggle-track';
      sw.appendChild(track);
      row.appendChild(sw);

      group.appendChild(row);

      if (hint) {
        const hintEl = document.createElement('div');
        hintEl.className = 'settings-hint';
        hintEl.textContent = hint;
        group.appendChild(hintEl);
      }

      parent.appendChild(group);
    }

    function addField (parent, labelText, key, placeholder, hint, defaultValue) {
      const group = document.createElement('div');
      group.className = 'settings-group';

      const lbl = document.createElement('label');
      lbl.textContent = labelText;
      group.appendChild(lbl);

      const input = document.createElement('input');
      input.type = 'text';
      input.placeholder = placeholder;
      input.value = settings[key] !== undefined ? settings[key] : (defaultValue ?? '');
      input.dataset.key = key;
      input.addEventListener('input', onInput);
      group.appendChild(input);

      if (hint) {
        const hintEl = document.createElement('div');
        hintEl.className = 'settings-hint';
        hintEl.textContent = hint;
        group.appendChild(hintEl);
      }

      parent.appendChild(group);
    }

    function addColorField (parent, labelText, key, defaultColor, hint) {
      const group = document.createElement('div');
      group.className = 'settings-group';

      const row = document.createElement('div');
      row.style.display = 'flex';
      row.style.alignItems = 'center';
      row.style.gap = '8px';

      const lbl = document.createElement('label');
      lbl.textContent = labelText;
      lbl.style.flex = '1';
      lbl.style.marginBottom = '0';
      row.appendChild(lbl);

      const colorInp = document.createElement('input');
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

      // Reset button (icon)
      const btnDef = document.createElement('button');
      btnDef.className = 'settings-btn settings-btn-cancel';
      btnDef.innerHTML = '&#8634;';
      btnDef.title = 'Reset to default';
      btnDef.style.cssText = 'font-size:16px;padding:2px 4px;line-height:1;color:#a0aab4;cursor:pointer;';
      btnDef.addEventListener('click', () => {
        colorInp.value = defaultColor;
        colorInp.dispatchEvent(new Event('input'));
      });
      row.appendChild(btnDef);
      group.appendChild(row);

      if (hint) {
        const hintEl = document.createElement('div');
        hintEl.className = 'settings-hint';
        hintEl.textContent = hint;
        group.appendChild(hintEl);
      }

      parent.appendChild(group);
    }

    function addDropdown (parent, labelText, key, options, hint, allowNone) {
      const group = document.createElement('div');
      group.className = 'settings-group';

      const lbl = document.createElement('label');
      lbl.textContent = labelText;
      group.appendChild(lbl);

      const select = document.createElement('select');
      select.dataset.key = key;
      select.addEventListener('change', onInput);

      const currentValue = settings[key];

      // Field-selection dropdowns need a "(None)" option
      if (allowNone) {
        const emptyOpt = document.createElement('option');
        emptyOpt.value = '';
        emptyOpt.textContent = '(None)';
        select.appendChild(emptyOpt);
      }

      for (let i = 0; i < options.length; i++) {
        const opt = options[i];
        const optEl = document.createElement('option');
        optEl.value = opt.value;
        optEl.textContent = opt.label;
        if (currentValue === opt.value) {
          optEl.selected = true;
        } else if (!allowNone && !currentValue && i === 0) {
          // No "(None)" and no saved value: default to first option
          optEl.selected = true;
        }
        select.appendChild(optEl);
      }

      group.appendChild(select);

      if (hint) {
        const hintEl = document.createElement('div');
        hintEl.className = 'settings-hint';
        hintEl.textContent = hint;
        group.appendChild(hintEl);
      }

      parent.appendChild(group);
    }

    /**
     * @param {string} [autoLabel] - if provided, display this string when value === min (the "auto" sentinel).
     */
    function addStepper (parent, labelText, key, defaultVal, min, max, step, autoLabel) {
      const group = document.createElement('div');
      group.className = 'settings-group';

      const lbl = document.createElement('label');
      lbl.textContent = labelText;
      group.appendChild(lbl);

      const row = document.createElement('div');
      row.style.cssText = 'display:flex;align-items:center;gap:6px;';

      const btnMinus = document.createElement('button');
      btnMinus.className = 'settings-btn settings-btn-cancel';
      btnMinus.style.cssText = 'padding:4px 10px;font-size:14px;line-height:1;';
      btnMinus.textContent = '−';

      const display = document.createElement('input');
      display.type = 'text';
      display.style.cssText = 'width:62px;text-align:center;border:1.5px solid #e3e0ea;border-radius:8px;padding:5px 4px;font-size:13px;color:var(--brand-dark);outline:none;';
      const curVal = settings[key] !== undefined ? parseInt(settings[key], 10) : defaultVal;
      const resolvedVal = isNaN(curVal) ? defaultVal : curVal;
      display.value = (autoLabel && resolvedVal === min) ? autoLabel : resolvedVal;
      display.dataset.key = key;
      display.dataset.numVal = String(resolvedVal);

      const btnPlus = document.createElement('button');
      btnPlus.className = 'settings-btn settings-btn-cancel';
      btnPlus.style.cssText = 'padding:4px 10px;font-size:14px;line-height:1;';
      btnPlus.textContent = '+';

      function update (newVal) {
        const clamped = Math.max(min, Math.min(max, newVal));
        display.dataset.numVal = String(clamped);
        display.value = (autoLabel && clamped === min) ? autoLabel : clamped;
        onInput();
      }

      btnMinus.addEventListener('click', () => update(parseInt(display.dataset.numVal, 10) - step));
      btnPlus.addEventListener('click', () => update(parseInt(display.dataset.numVal, 10) + step));
      display.addEventListener('change', () => {
        const parsed = parseInt(display.value, 10);
        update(isNaN(parsed) ? defaultVal : parsed);
      });

      row.appendChild(btnMinus);
      row.appendChild(display);
      row.appendChild(btnPlus);

      if (arguments.length > 8 && arguments[8] === true) {
        const btnDef = document.createElement('button');
        btnDef.className = 'settings-btn settings-btn-cancel';
        btnDef.innerHTML = '&#8634;';
        btnDef.title = 'Reset to ' + defaultVal;
        btnDef.style.cssText = 'font-size:16px;padding:2px 4px;line-height:1;color:#a0aab4;cursor:pointer;';
        btnDef.addEventListener('click', () => update(defaultVal));
        row.appendChild(btnDef);
      }

      group.appendChild(row);
      parent.appendChild(group);
    }

    // ===================================================================
    // Build sections by component
    // ===================================================================

    // --- LAYOUT ---
    const secLayout = addSection('Layout');
    addStepper(secLayout, 'Margin Top (px)', 'padTop', 24, 0, 80, 2, undefined, true);
    addStepper(secLayout, 'Margin Left (px)', 'padLeft', 28, 0, 80, 2, undefined, true);
    addStepper(secLayout, 'Card Width (px)', 'cardWidth', 480, 200, 1200, 20, undefined, true);

    // Top gradient accent color
    addColorField(secLayout, 'Accent Color', 'gradColor', '#D42F8A',
      'Top border gradient is generated from this color.');

    // Progress bar color mode
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
    const secTitle = addSection('Title & Header');
    addToggle(secTitle, 'Show Title', 'showTitle', 'The sheet name displayed at the top.');
    addField(secTitle, 'Title Text', 'titleText',
      sheetName || 'Sheet name',
      'Blank = use sheet name.', '');
    addStepper(secTitle, 'Title Size (px)', 'titleSize', 18, 10, 36, 1, undefined, true);
    addToggle(secTitle, 'Show Subtitle / Label', 'showHeader', 'Value label and date range row.');
    addField(secTitle, 'Value Label Override', 'valueLabel',
      kpi ? kpi.label : 'Auto-detected',
      'Blank = hide label.',
      kpi ? kpi.label : '');
    addToggle(secTitle, 'Show Date Range', 'showDateRange', 'Period range in the subtitle.');

    // --- NUMBER FORMAT (global) ---
    const secFmt = addSection('Number Format');
    addField(secFmt, 'Prefix', 'fmtPrefix', '',
      'e.g. $ or € — applies to all numbers.', '');
    addField(secFmt, 'Suffix', 'fmtSuffix', '',
      'e.g. % or units — applies to all numbers.', '');
    addStepper(secFmt, 'Decimal Places', 'fmtDecimals', 0, 0, 6, 1);
    addToggle(secFmt, 'Abbreviate (K / M / B)', 'fmtAbbreviate',
      'Show 120K instead of 120,000. Applies to all numbers.');

    // --- VALUE & DELTA ---
    const secValue = addSection('Value & Delta');
    addToggle(secValue, 'Show Value', 'showValue', 'The big number.');
    addStepper(secValue, 'Value Size (px)', 'valueSize', 40, 16, 72, 2, undefined, true);
    addToggle(secValue, 'Show Delta Badge', 'showDelta', 'Change vs previous period.');
    addField(secValue, 'Delta Label', 'deltaLabel', 'vs prev',
      'Text after the percentage.', 'vs prev');
    addToggle(secValue, 'Reverse Delta Colors', 'reverseDelta', 'Higher = bad (costs, churn).');
    addStepper(secValue, 'Delta Decimal Places', 'fmtDeltaDecimals', 1, 0, 4, 1);
    addStepper(secValue, 'Badge Font Size (px)', 'deltaSize', 12, 8, 20, 1);

    // --- PRIMARY GOAL ---
    const secGoal = addSection('Primary Goal');
    addToggle(secGoal, 'Show Primary Goal', 'showGoal');
    addField(secGoal, 'Goal Label', 'goalLabel', 'Goal',
      'Blank = hide bar.', 'Goal');

    // --- SECONDARY GOAL (collapsed by default) ---
    const secGoal2 = addSection('Secondary Goal');
    addToggle(secGoal2, 'Show Secondary Goal', 'showGoal2');
    addField(secGoal2, 'Goal Label', 'goal2Label',
      kpi ? kpi.goal2Label : 'Secondary Goal',
      'Blank = hide bar.',
      kpi ? kpi.goal2Label : 'Secondary Goal');

    // --- SPARKLINE ---
    const secSpark = addSection('Sparkline Chart');
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

    // --- SECONDARY COMPARISON (collapsed by default) ---
    const secComp = addSection('Secondary Comparison');
    addToggle(secComp, 'Enable Comparison', 'ptdEnabled',
      'Adds a badge and a dashed line on the sparkline.');
    addDropdown(secComp, 'Comparison Field', 'ptdFieldName', measures,
      'Select a calculated field to compare against (e.g. PTD, forecast).', true);
    addField(secComp, 'Badge Label', 'ptdLabel', 'vs prev PTD',
      'Text shown on the comparison badge.', 'vs prev PTD');
    addField(secComp, 'Legend Label', 'ptdLegendLabel', 'PTD Pace',
      'Label for the dashed line in the sparkline legend.', 'PTD Pace');

    // --- DATA FIELDS ---
    const secData = addSection('Data Fields');
    addDropdown(secData, 'Date / Period Field', 'dateFieldName', allFields,
      'Field to group by (quarters, months, etc).', true);

    // --- CUSTOM LINK ---
    const secLink = addSection('Custom Link');
    addToggle(secLink, 'Show Link', 'showLink');
    addField(secLink, 'Link Label', 'linkLabel', 'e.g. View Dashboard',
      'Blank = hide link.');
    addField(secLink, 'Link URL', 'linkUrl', 'https://...');
    addField(secLink, 'Icon URL', 'linkIcon', 'https://...',
      'Optional 16x16 icon.');

    // --- COPY / PASTE SETTINGS ---
    const actionsRow = document.createElement('div');
    actionsRow.className = 'settings-actions';

    function showToast (msg) {
      const toast = document.createElement('div');
      toast.className = 'settings-toast';
      toast.textContent = msg;
      document.body.appendChild(toast);
      setTimeout(() => toast.remove(), 2000);
    }

    const copyBtn = document.createElement('button');
    copyBtn.className = 'settings-btn settings-btn-cancel';
    copyBtn.textContent = 'Copy Settings';
    copyBtn.addEventListener('click', async () => {
      const currentSettings = loadSettings();
      try {
        await navigator.clipboard.writeText(JSON.stringify(currentSettings, null, 2));
        showToast('Settings copied to clipboard');
      } catch (_) {
        // Fallback for environments where clipboard API is blocked
        const ta = document.createElement('textarea');
        ta.value = JSON.stringify(currentSettings, null, 2);
        ta.style.cssText = 'position:fixed;left:-9999px';
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        ta.remove();
        showToast('Settings copied to clipboard');
      }
    });

    const pasteBtn = document.createElement('button');
    pasteBtn.className = 'settings-btn settings-btn-primary';
    pasteBtn.textContent = 'Paste Settings';

    // Paste area (hidden until user clicks Paste)
    const pasteArea = document.createElement('div');
    pasteArea.style.cssText = 'display:none;margin-top:10px;';

    const pasteTA = document.createElement('textarea');
    pasteTA.placeholder = 'Paste settings JSON here (Ctrl+V)...';
    pasteTA.style.cssText = 'width:100%;height:80px;border:1.5px solid #e3e0ea;border-radius:8px;padding:8px;font-size:11px;font-family:monospace;color:var(--brand-dark);outline:none;resize:vertical;';

    const applyBtn = document.createElement('button');
    applyBtn.className = 'settings-btn settings-btn-primary';
    applyBtn.style.cssText = 'margin-top:6px;width:100%;';
    applyBtn.textContent = 'Apply';
    applyBtn.addEventListener('click', async () => {
      const text = pasteTA.value.trim();
      if (!text) { showToast('Paste your settings JSON first'); return; }
      let parsed;
      try {
        parsed = JSON.parse(text);
      } catch (_) {
        showToast('Invalid JSON format');
        return;
      }
      if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
        showToast('Invalid settings data');
        return;
      }
      const knownKeys = new Set(Object.keys(SETTINGS_KEYS));
      const validKeys = Object.keys(parsed).filter(k => knownKeys.has(k));
      if (validKeys.length === 0) {
        showToast('No recognized settings found');
        return;
      }
      await saveSettings(parsed, onSave);
      closeSettingsPanel();
      openSettingsPanel(kpi, dataResult, onSave, sheetName);
      showToast('Settings applied (' + validKeys.length + ' values)');
    });

    pasteArea.appendChild(pasteTA);
    pasteArea.appendChild(applyBtn);

    pasteBtn.addEventListener('click', () => {
      const visible = pasteArea.style.display !== 'none';
      pasteArea.style.display = visible ? 'none' : 'block';
      if (!visible) pasteTA.focus();
    });

    actionsRow.appendChild(copyBtn);
    actionsRow.appendChild(pasteBtn);
    panel.appendChild(actionsRow);
    panel.appendChild(pasteArea);

    // Version label
    const versionEl = document.createElement('div');
    versionEl.style.cssText = 'text-align:center;font-size:10px;color:#c4c9d0;margin-top:10px;';
    versionEl.textContent = 'KPI Card v' + KPI_VERSION;
    panel.appendChild(versionEl);

    overlay.appendChild(panel);
    document.body.appendChild(overlay);

    const firstInput = panel.querySelector('input');
    if (firstInput) firstInput.focus();
  }

  // =========================================================================
  // Render
  // =========================================================================

  function render (dataResult, encodings, sheetName) {
    const content = document.getElementById('content');
    content.innerHTML = '';

    const settings = loadSettings();
    // Animate on initial load, data change, or when toggle is switched on
    const justToggledOn = _prevShowAnimation === false && settings.showAnimation === true;
    settings._animate = settings.showAnimation && (!_skipAnimation || justToggledOn);
    _prevShowAnimation = settings.showAnimation;
    _skipAnimation = false;

    // Apply page-level layout (space between card and page edges)
    const padT = parseInt(settings.padTop, 10);
    const padL = parseInt(settings.padLeft, 10);
    const top = padT >= 0 ? padT : 24;
    const left = padL >= 0 ? padL : 28;
    content.style.padding = top + 'px ' + left + 'px';

    const r = resolveEncodings(encodings, dataResult.columns, settings);

    if (!r.valueCol) {
      content.innerHTML =
        '<div class="kpi-empty">' +
        'Drag a <b>measure</b> onto the <b>Value</b> tile on the Marks card.' +
        '</div>';
      return;
    }

    const kpi = computeKpi(dataResult, encodings, settings);
    if (!kpi) {
      content.innerHTML = '<div class="kpi-empty">No data available.</div>';
      return;
    }

    const card = document.createElement('div');
    card.className = 'kpi-card';
    const cw = parseInt(settings.cardWidth, 10);
    if (cw > 0) card.style.maxWidth = cw + 'px';
    try { card.classList.add(tableau.ClassNameKey.Worksheet); } catch (e) { /* ok */ }

    // Custom top gradient from single accent color
    const accentHex = settings.gradColor || '#D42F8A';
    const gradStops = buildGradientFromColor(accentHex);
    card.style.borderImage = `linear-gradient(135deg, ${gradStops.join(', ')}) 1`;

    // ---- Custom link (top-left) ----
    if (settings.showLink) {
      const linkUrl  = (settings.linkUrl  || '').trim();
      const linkLabel = (settings.linkLabel || '').trim();
      const linkIcon  = (settings.linkIcon  || '').trim();
      if (linkUrl && (linkIcon || linkLabel)) {
        const linkEl = document.createElement('a');
        linkEl.className = 'kpi-link';
        linkEl.href = linkUrl;
        linkEl.target = '_blank';
        linkEl.rel = 'noopener noreferrer';
        if (linkLabel) linkEl.classList.add('has-label');
        if (linkIcon) {
          const iconEl = document.createElement('img');
          iconEl.className = 'kpi-link-icon';
          iconEl.src = linkIcon;
          iconEl.alt = '';
          iconEl.onerror = function () { this.style.display = 'none'; };
          linkEl.appendChild(iconEl);
        }
        if (linkLabel) {
          const labelSpan = document.createElement('span');
          labelSpan.textContent = linkLabel;
          linkEl.appendChild(labelSpan);
        }
        card.appendChild(linkEl);
      }
    }

    // ---- Settings gear (top-right, authoring mode only) ----
    let showGear = false;
    try {
      const mode = tableau.extensions.environment.mode;
      if (mode === 'authoring') {
        showGear = true;
        try {
          const active = tableau.extensions.workbook.activeSheet;
          if (active && active.sheetType && active.sheetType !== 'worksheet') showGear = false;
        } catch (_) { /* ok */ }
      }
    } catch (_) { /* ok */ }

    if (showGear) {
      const topActions = document.createElement('div');
      topActions.className = 'kpi-top-actions';

      const gearBtn = document.createElement('button');
      gearBtn.className = 'kpi-settings-btn';
      gearBtn.title = 'Settings';
      gearBtn.innerHTML = '&#9881;';
      gearBtn.addEventListener('click', () => {
        openSettingsPanel(kpi, dataResult, _renderLatest || (() => render(dataResult, encodings, sheetName)), sheetName);
      });
      topActions.appendChild(gearBtn);
      card.appendChild(topActions);
    } else {
      closeSettingsPanel();
    }

    // ---- Title ----
    if (settings.showTitle) {
      const title = document.createElement('div');
      title.className = 'kpi-title';
      title.textContent = settings.titleText || sheetName || kpi.label;
      const ts = parseInt(settings.titleSize, 10);
      if (ts > 0 && ts !== 18) title.style.fontSize = ts + 'px';
      card.appendChild(title);
    }

    // ---- Header (label + period) ----
    if (settings.showHeader) {
      const valueLabelText = settings.valueLabel !== undefined
        ? settings.valueLabel
        : kpi.label;
      const showPeriod = settings.showDateRange && kpi.periodLabel;

      if (valueLabelText !== '') {
        const header = document.createElement('div');
        header.className = 'kpi-header';
        header.innerHTML =
          '<span class="kpi-label">' + escapeHtml(valueLabelText) + '</span>' +
          (showPeriod
            ? '<span class="kpi-period">' + escapeHtml(kpi.periodLabel) + '</span>'
            : '');
        card.appendChild(header);
      } else if (showPeriod) {
        const header = document.createElement('div');
        header.className = 'kpi-header';
        header.innerHTML = '<span class="kpi-period">' + escapeHtml(kpi.periodLabel) + '</span>';
        card.appendChild(header);
      }
    }

    // ---- Value + delta row ----
    if (settings.showValue || settings.showDelta) {
      const valueRow = document.createElement('div');
      valueRow.className = 'kpi-value-row';

      if (settings.showValue) {
        const valueEl = document.createElement('span');
        valueEl.className = 'kpi-value';
        valueEl.textContent = kpi.formattedValue;
        const vs = parseInt(settings.valueSize, 10);
        if (vs > 0 && vs !== 40) valueEl.style.fontSize = vs + 'px';
        valueRow.appendChild(valueEl);
      }

      const rev = settings.reverseDelta;

      // Badge font size (user-adjustable, default 12px from CSS)
      const badgePx = parseInt(settings.deltaSize, 10) || 12;
      const badgeStyle = badgePx !== 12 ? 'font-size:' + badgePx + 'px' : '';

      const deltaDec = parseInt(settings.fmtDeltaDecimals, 10);
      const deltaDecPlaces = (deltaDec >= 0) ? deltaDec : 1;

      if (settings.showDelta && kpi.delta !== null) {
        const deltaEl = document.createElement('span');
        const displayDelta = Math.abs(kpi.delta).toFixed(deltaDecPlaces);
        const isNeutral = parseFloat(displayDelta) === 0;
        const isUp = kpi.delta > 0;
        const sentiment = isNeutral ? 'neutral' : (isUp !== rev) ? 'positive' : 'negative';
        deltaEl.className = 'kpi-delta ' + sentiment;
        if (badgeStyle) deltaEl.style.cssText = badgeStyle;
        const arrow = isNeutral ? '–' : isUp ? '▲' : '▼';
        deltaEl.innerHTML =
          '<span class="kpi-delta-arrow">' + arrow + '</span> ' +
          displayDelta + '% ' + escapeHtml(settings.deltaLabel || 'vs prev');
        valueRow.appendChild(deltaEl);
      }

      // Secondary comparison badge
      const ptdLabelText = settings.ptdLabel !== undefined
        ? settings.ptdLabel
        : 'vs prev PTD';

      if (settings.ptdEnabled && kpi.ptdDelta !== null && ptdLabelText !== '') {
        const ptdEl = document.createElement('span');
        const displayPtdDelta = Math.abs(kpi.ptdDelta).toFixed(deltaDecPlaces);
        const ptdNeutral = parseFloat(displayPtdDelta) === 0;
        const ptdUp = kpi.ptdDelta > 0;
        const ptdSentiment = ptdNeutral ? 'neutral' : (ptdUp !== rev) ? 'positive' : 'negative';
        ptdEl.className = 'kpi-delta kpi-delta-ptd ' + ptdSentiment;
        if (badgeStyle) ptdEl.style.cssText = badgeStyle;
        const ptdArrow = ptdNeutral ? '–' : ptdUp ? '▲' : '▼';
        ptdEl.innerHTML =
          '<span class="kpi-delta-arrow">' + ptdArrow + '</span> ' +
          displayPtdDelta + '% ' + escapeHtml(ptdLabelText);
        valueRow.appendChild(ptdEl);
      }

      card.appendChild(valueRow);
    }

    // ---- Primary goal ----
    if (settings.showGoal) {
      const goalLabelText = settings.goalLabel !== undefined
        ? settings.goalLabel
        : 'Goal';
      if (kpi.goalPct !== null && goalLabelText !== '') {
        card.appendChild(buildGoalBar(kpi.goalPct, kpi.formattedGoal, goalLabelText, settings));
      }
    }

    // ---- Secondary goal ----
    if (settings.showGoal2) {
      const goal2LabelText = settings.goal2Label !== undefined
        ? settings.goal2Label
        : kpi.goal2Label;
      if (kpi.goal2Pct !== null && goal2LabelText !== '') {
        card.appendChild(buildGoalBar(kpi.goal2Pct, kpi.formattedGoal2, goal2LabelText, settings));
      }
    }

    // ---- Sparkline placeholder (drawn after card is in the DOM) ----
    let sparkSection = null;
    if (settings.showSparkline && kpi.sparkData && kpi.sparkData.length > 0) {
      sparkSection = document.createElement('div');
      sparkSection.className = 'kpi-sparkline-section';
      card.appendChild(sparkSection);
    }

    // Append card to DOM first so sparkline can measure actual width
    content.appendChild(card);

    // Now draw sparkline with real container width
    if (sparkSection) {
      const ptdData = settings.ptdEnabled ? kpi.ptdSparkData : null;
      const height = parseInt(settings.sparkHeight, 10) || 130;
      drawSparkline(sparkSection, kpi.sparkData, ptdData, settings, height, kpi.globalFmt || formatNumberCompact);
    }
  }

  // =========================================================================
  // Goal bar builder — shared gradient for both bars
  // =========================================================================

  function buildGoalBar (goalPct, formattedGoal, label, settings) {
    const section = document.createElement('div');
    section.className = 'kpi-goal-section';

    const track = document.createElement('div');
    track.className = 'kpi-progress-track';

    const fill = document.createElement('div');
    fill.className = 'kpi-progress-fill';
    const pct = Math.max(0, Math.min(goalPct, 100));

    if (settings && settings._animate) {
      fill.style.width = '0%';
      fill.classList.add('animate');
      requestAnimationFrame(() => { fill.style.width = pct + '%'; });
    } else {
      fill.style.width = pct + '%';
    }

    // Progress bar color
    const barMode = (settings && settings.barColorMode) || 'default';
    if (barMode === 'accent') {
      const accent = (settings && settings.gradColor) || '#D42F8A';
      const stops = buildGradientFromColor(accent);
      fill.style.background = 'linear-gradient(90deg, ' + stops.join(', ') + ')';
    } else if (barMode === 'custom') {
      const custom = (settings && settings.barCustomColor) || '#3b82f6';
      fill.style.background = custom;
    } else {
      // default: brand gradient
      fill.style.background = 'linear-gradient(90deg, var(--brand-orange), var(--brand-coral), var(--brand-magenta), var(--brand-purple))';
    }

    track.appendChild(fill);
    section.appendChild(track);

    const meta = document.createElement('div');
    meta.className = 'kpi-goal-meta';
    meta.innerHTML =
      '<span>' + escapeHtml(label) + ': ' + escapeHtml(formattedGoal) + '</span>' +
      '<span class="kpi-goal-pct">' + goalPct + '% of goal</span>';
    section.appendChild(meta);

    return section;
  }

  // =========================================================================
  // Sparkline (D3)
  // =========================================================================

  function drawSparkline (container, sparkData, ptdSparkData, settings, chartHeight, valueFmt) {
    const margin = { top: 28, right: 8, bottom: settings.showSparkPeriods ? 24 : 8, left: 8 };
    const width = container.offsetWidth || 400;
    const totalWidth = width - margin.left - margin.right;
    const totalHeight = chartHeight || 130;
    const chartH = totalHeight - margin.top - margin.bottom;

    const svg = d3.create('svg')
      .attr('width', width)
      .attr('height', totalHeight + margin.bottom)
      .attr('viewBox', [0, 0, width, totalHeight + margin.bottom])
      .style('overflow', 'visible');

    // Color mode: default (brand colors), accent, or custom
    const colorMode = (settings && settings.sparklineColorMode) || 'default';
    const accentColor = (settings && settings.gradColor) || '#D42F8A';
    const customColor = (settings && settings.sparklineCustomColor) || '#3b82f6';
    
    let lineColor, dotColor, ptdColor;
    if (colorMode === 'accent') {
      lineColor = accentColor;
      dotColor = accentColor;
      ptdColor = accentColor;
    } else if (colorMode === 'custom') {
      lineColor = customColor;
      dotColor = customColor;
      ptdColor = customColor;
    } else {
      // default: brand colors
      lineColor = '#EF5A5E'; // brand-coral
      dotColor = '#D42F8A';  // brand-magenta
      ptdColor = '#7B2FBE';  // brand-purple
    }
    
    const defs = svg.append('defs');
    const grad = defs.append('linearGradient')
      .attr('id', 'sparkGradient')
      .attr('x1', '0%').attr('y1', '0%')
      .attr('x2', '0%').attr('y2', '100%');
    // Area gradient: always subtle/neutral for readability
    grad.append('stop').attr('offset', '0%').attr('stop-color', lineColor).attr('stop-opacity', 0.08);
    grad.append('stop').attr('offset', '100%').attr('stop-color', lineColor).attr('stop-opacity', 0.02);

    const g = svg.append('g')
      .attr('transform', `translate(${margin.left},${margin.top})`);

    const x = d3.scalePoint()
      .domain(sparkData.map((_, i) => i))
      .range([0, totalWidth]);

    const allValues = sparkData.map(d => d.value);
    if (ptdSparkData) ptdSparkData.forEach(d => allValues.push(d.value));
    const yExtent = d3.extent(allValues);
    const yPad = (yExtent[1] - yExtent[0]) * 0.2 || 1;
    const y = d3.scaleLinear()
      .domain([yExtent[0] - yPad, yExtent[1] + yPad])
      .range([chartH, 0]);

    const area = d3.area()
      .x((_, i) => x(i))
      .y0(chartH)
      .y1(d => y(d.value))
      .curve(d3.curveMonotoneX);

    g.append('path')
      .datum(sparkData)
      .attr('class', 'spark-area')
      .attr('d', area);

    // --- Secondary comparison line (dashed) ---
    if (ptdSparkData) {
      const ptdLine = d3.line()
        .x((_, i) => x(i))
        .y(d => y(d.value))
        .curve(d3.curveMonotoneX);

      g.append('path')
        .datum(ptdSparkData)
        .attr('class', 'spark-line-ptd')
        .attr('d', ptdLine)
        .attr('stroke', ptdColor)
        .attr('opacity', 0.7);

      g.selectAll('.spark-dot-ptd')
        .data(ptdSparkData)
        .join('circle')
        .attr('class', 'spark-dot-ptd')
        .attr('cx', (_, i) => x(i))
        .attr('cy', d => y(d.value))
        .attr('r', 2.5)
        .attr('fill', ptdColor)
        .attr('opacity', 0.7);
    }

    // --- Actual line ---
    const line = d3.line()
      .x((_, i) => x(i))
      .y(d => y(d.value))
      .curve(d3.curveMonotoneX);

    g.append('path')
      .datum(sparkData)
      .attr('class', 'spark-line')
      .attr('d', line)
      .attr('stroke', lineColor);

    g.selectAll('.spark-dot')
      .data(sparkData)
      .join('circle')
      .attr('class', 'spark-dot')
      .attr('cx', (_, i) => x(i))
      .attr('cy', d => y(d.value))
      .attr('r', 3.5)
      .attr('fill', dotColor);

    // ---- Smart data labels with collision avoidance ----
    if (settings.showSparkLabels) {
      // Collect all label candidates with their natural y positions
      const labels = [];

      // Actual line labels (above the dot) - neutral color for readability
      sparkData.forEach((d, i) => {
        labels.push({
          x: x(i), naturalY: y(d.value) - 10,
          text: valueFmt(d.value),
          fill: '#1a1a2e', opacity: 1, fontSize: 11, series: 'actual'
        });
      });

      // Comparison line labels (below the dot) - neutral color for readability
      if (ptdSparkData) {
        ptdSparkData.forEach((d, i) => {
          // Skip if value is very close to actual (would overlap)
          const actualY = y(sparkData[i].value);
          const ptdY = y(d.value);
          if (Math.abs(actualY - ptdY) < 4 && Math.abs(d.value - sparkData[i].value) < 0.5) return;
          labels.push({
            x: x(i), naturalY: y(d.value) + 14,
            text: valueFmt(d.value),
            fill: '#4a5568', opacity: 0.8, fontSize: 10, series: 'ptd'
          });
        });
      }

      // Sort by x then by natural y for collision resolution
      labels.sort((a, b) => a.x - b.x || a.naturalY - b.naturalY);

      // Resolve vertical overlaps at each x position
      const minGap = 12; // minimum px between label centers
      const resolvedLabels = [];
      for (const lbl of labels) {
        let finalY = lbl.naturalY;
        // Check against already-placed labels at same or nearby x
        for (const placed of resolvedLabels) {
          if (Math.abs(placed.x - lbl.x) < 30 && Math.abs(finalY - placed.y) < minGap) {
            // Push this label further away from the collision
            if (lbl.series === 'ptd') {
              finalY = placed.y + minGap; // push down
            } else {
              finalY = placed.y - minGap; // push up
            }
          }
        }
        // Clamp so labels stay in bounds
        finalY = Math.max(-margin.top + 8, Math.min(chartH + 6, finalY));
        resolvedLabels.push({ ...lbl, y: finalY });
      }

      // Render all labels — anchor at edges to avoid clipping
      for (const lbl of resolvedLabels) {
        let anchor = 'middle';
        if (lbl.x <= 10) anchor = 'start';
        else if (lbl.x >= totalWidth - 10) anchor = 'end';

        g.append('text')
          .attr('x', lbl.x)
          .attr('y', lbl.y)
          .attr('text-anchor', anchor)
          .attr('font-size', lbl.fontSize + 'px')
          .attr('font-weight', '600')
          .attr('fill', lbl.fill)
          .attr('opacity', lbl.opacity)
          .text(lbl.text);
      }
    }

    // Period labels — adjust anchor at edges to prevent clipping
    if (settings.showSparkPeriods) {
      g.selectAll('.spark-label')
        .data(sparkData)
        .join('text')
        .attr('x', (_, i) => x(i))
        .attr('y', chartH + 16)
        .attr('text-anchor', (_, i) => {
          if (i === 0) return 'start';
          if (i === sparkData.length - 1) return 'end';
          return 'middle';
        })
        .attr('font-size', '10px')
        .attr('fill', '#a0aab4')
        .text(d => d.label);
    }

    // ---- Hover interaction ----
    // Vertical guide line
    const guideLine = g.append('line')
      .attr('class', 'spark-guide')
      .attr('y1', -margin.top + 4)
      .attr('y2', chartH + 4);

    // Highlight circles (actual + ptd)
    const hoverDot = g.append('circle')
      .attr('r', 6)
      .attr('fill', dotColor)
      .attr('stroke', '#fff')
      .attr('stroke-width', 2)
      .style('pointer-events', 'none')
      .style('opacity', 0);

    let hoverDotPtd = null;
    if (ptdSparkData) {
      hoverDotPtd = g.append('circle')
        .attr('r', 5)
        .attr('fill', ptdColor)
        .attr('stroke', '#fff')
        .attr('stroke-width', 2)
        .style('pointer-events', 'none')
        .style('opacity', 0);
    }

    // Tooltip group
    const tooltip = g.append('g')
      .style('pointer-events', 'none')
      .style('opacity', 0);

    const tooltipBg = tooltip.append('rect').attr('class', 'spark-tooltip-bg');
    const tooltipVal = tooltip.append('text').attr('class', 'spark-tooltip-val');
    const tooltipLbl = tooltip.append('text').attr('class', 'spark-tooltip-lbl');
    let tooltipPtd = null;
    if (ptdSparkData) {
      tooltipPtd = tooltip.append('text').attr('class', 'spark-tooltip-ptd');
    }

    // Invisible overlay for mouse events
    g.append('rect')
      .attr('width', totalWidth)
      .attr('height', chartH)
      .attr('fill', 'transparent')
      .style('cursor', 'crosshair')
      .on('mousemove', function (event) {
        const [mx] = d3.pointer(event);
        // Find nearest data point index
        let minDist = Infinity, idx = 0;
        sparkData.forEach((_, i) => {
          const dist = Math.abs(x(i) - mx);
          if (dist < minDist) { minDist = dist; idx = i; }
        });

        const px = x(idx);
        const py = y(sparkData[idx].value);

        // Guide line
        guideLine.attr('x1', px).attr('x2', px).classed('visible', true);

        // Highlight dot
        hoverDot.attr('cx', px).attr('cy', py).style('opacity', 1);

        // Ptd dot
        if (hoverDotPtd && ptdSparkData && ptdSparkData[idx]) {
          hoverDotPtd
            .attr('cx', px)
            .attr('cy', y(ptdSparkData[idx].value))
            .style('opacity', 1);
        }

        // Build tooltip content
        const valText = valueFmt(sparkData[idx].value);
        const lblText = sparkData[idx].label || '';
        tooltipVal.text(valText);
        tooltipLbl.text(lblText);

        let tooltipH = 42;
        const lineSpacing = 16;
        tooltipVal.attr('x', 10).attr('y', 18);
        tooltipLbl.attr('x', 10).attr('y', 18 + lineSpacing);

        if (tooltipPtd && ptdSparkData && ptdSparkData[idx]) {
          const ptdText = (settings.ptdBadgeLabel || 'Comp') + ': ' + valueFmt(ptdSparkData[idx].value);
          tooltipPtd.text(ptdText).attr('x', 10).attr('y', 18 + lineSpacing * 2);
          tooltipH = 42 + lineSpacing;
        }

        // Measure text width for bg
        const valBBox = tooltipVal.node().getBBox();
        const lblBBox = tooltipLbl.node().getBBox();
        let maxW = Math.max(valBBox.width, lblBBox.width);
        if (tooltipPtd && ptdSparkData && ptdSparkData[idx]) {
          const ptdBBox = tooltipPtd.node().getBBox();
          maxW = Math.max(maxW, ptdBBox.width);
        }
        const tooltipW = maxW + 20;

        tooltipBg.attr('width', tooltipW).attr('height', tooltipH);

        // Position tooltip — flip if near right edge
        let tx = px + 12;
        if (tx + tooltipW > totalWidth) tx = px - tooltipW - 12;
        let ty = py - tooltipH - 10;
        if (ty < -margin.top) ty = py + 16;

        tooltip.attr('transform', `translate(${tx},${ty})`).style('opacity', 1);
      })
      .on('mouseleave', function () {
        guideLine.classed('visible', false);
        hoverDot.style('opacity', 0);
        if (hoverDotPtd) hoverDotPtd.style('opacity', 0);
        tooltip.style('opacity', 0);
      });

    container.appendChild(svg.node());

    // ---- Open animation: draw lines + pop dots ----
    if (settings._animate) {
      const animDuration = 600; // ms

      // Animate line paths (actual + ptd) via stroke-dasharray
      svg.selectAll('.spark-line, .spark-line-ptd').each(function () {
        const path = d3.select(this);
        const len = this.getTotalLength();
        path
          .attr('stroke-dasharray', len)
          .attr('stroke-dashoffset', len)
          .transition()
          .duration(animDuration)
          .ease(d3.easeCubicOut)
          .attr('stroke-dashoffset', 0)
          .on('end', function () {
            // Remove dasharray so hover/resize don't break
            d3.select(this).attr('stroke-dasharray', null);
          });
      });

      // Fade in area fill
      svg.selectAll('.spark-area')
        .style('opacity', 0)
        .transition()
        .delay(animDuration * 0.3)
        .duration(animDuration * 0.7)
        .style('opacity', 1);

      // Pop in dots with staggered delay
      svg.selectAll('.spark-dot, .spark-dot-ptd').each(function (d, i) {
        d3.select(this)
          .attr('r', 0)
          .transition()
          .delay(animDuration * 0.3 + i * 60)
          .duration(300)
          .ease(d3.easeBackOut.overshoot(1.5))
          .attr('r', d3.select(this).classed('spark-dot-ptd') ? 2.5 : 3.5);
      });

      // Fade in data labels + period labels
      svg.selectAll('text')
        .style('opacity', 0)
        .transition()
        .delay(animDuration * 0.5)
        .duration(300)
        .style('opacity', null); // restore original opacity
    }

    // Legend (only when comparison line is shown and legend enabled)
    if (ptdSparkData && settings.showLegend) {
      const compLegend = settings.ptdLegendLabel || 'Comparison';
      const legend = document.createElement('div');
      legend.className = 'spark-legend';
      legend.innerHTML =
        '<span class="spark-legend-item"><span class="spark-legend-line solid"></span>Actual</span>' +
        '<span class="spark-legend-item"><span class="spark-legend-line dashed"></span>' + escapeHtml(compLegend) + '</span>';
      container.appendChild(legend);
    }
  }

  // Build a 3-stop gradient from a single hex color by shifting hue
  function buildGradientFromColor (hex) {
    // hex → rgb
    let r = parseInt(hex.slice(1, 3), 16) / 255;
    let g = parseInt(hex.slice(3, 5), 16) / 255;
    let b = parseInt(hex.slice(5, 7), 16) / 255;

    // rgb → hsl
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h, s, l = (max + min) / 2;
    if (max === min) { h = s = 0; }
    else {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      if (max === r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
      else if (max === g) h = ((b - r) / d + 2) / 6;
      else h = ((r - g) / d + 4) / 6;
    }

    // hsl → hex helper
    function hslToHex (hh, ss, ll) {
      hh = ((hh % 1) + 1) % 1; // wrap 0-1
      const hue2rgb = (p, q, t) => {
        if (t < 0) t += 1; if (t > 1) t -= 1;
        if (t < 1/6) return p + (q - p) * 6 * t;
        if (t < 1/2) return q;
        if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
        return p;
      };
      let rr, gg, bb;
      if (ss === 0) { rr = gg = bb = ll; }
      else {
        const q = ll < 0.5 ? ll * (1 + ss) : ll + ss - ll * ss;
        const p = 2 * ll - q;
        rr = hue2rgb(p, q, hh + 1/3);
        gg = hue2rgb(p, q, hh);
        bb = hue2rgb(p, q, hh - 1/3);
      }
      const toHex = v => Math.round(v * 255).toString(16).padStart(2, '0');
      return '#' + toHex(rr) + toHex(gg) + toHex(bb);
    }

    // Generate 3 stops: hue-40° (warmer, lighter), base, hue+40° (cooler, darker)
    const c1 = hslToHex(h - 0.11, Math.min(1, s + 0.1), Math.min(0.65, l + 0.1));
    const c2 = hslToHex(h, s, l);
    const c3 = hslToHex(h + 0.11, Math.min(1, s + 0.05), Math.max(0.3, l - 0.08));
    return [c1, c2, c3];
  }

  function escapeHtml (str) {
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }
})();
