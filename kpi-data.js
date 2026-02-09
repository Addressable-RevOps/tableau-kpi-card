'use strict';

/**
 * kpi-data.js — Data fetching, encoding resolution, period helpers, and KPI computation.
 * Exports onto window.KPI.data
 */

window.KPI = window.KPI || {};
window.KPI.data = (function () {

  var buildGlobalFormatter = window.KPI.utils.buildGlobalFormatter;

  // =========================================================================
  // Data helpers
  // =========================================================================

  function convertToListOfNamedRows (dataTablePage) {
    var rows = [];
    var columns = dataTablePage.columns;
    var data = dataTablePage.data;
    for (var i = 0; i < data.length; i++) {
      var row = {};
      for (var j = 0; j < columns.length; j++) {
        row[columns[j].fieldName] = data[i][columns[j].index];
      }
      rows.push(row);
    }
    return { rows: rows, columns: columns };
  }

  async function getSummaryDataTable (worksheet) {
    var allRows = [];
    var allColumns = [];
    var dataTableReader = await worksheet.getSummaryDataReaderAsync(
      undefined,
      { ignoreSelection: true }
    );
    for (var page = 0; page < dataTableReader.pageCount; page++) {
      var dataTablePage = await dataTableReader.getPageAsync(page);
      var result = convertToListOfNamedRows(dataTablePage);
      allRows = allRows.concat(result.rows);
      if (page === 0) allColumns = result.columns;
    }
    await dataTableReader.releaseAsync();
    return { rows: allRows, columns: allColumns };
  }

  async function getEncodingMap (worksheet) {
    var visualSpec = await worksheet.getVisualSpecificationAsync();
    var map = {};
    if (visualSpec.activeMarksSpecificationIndex < 0) return map;
    var marksCard =
      visualSpec.marksSpecifications[visualSpec.activeMarksSpecificationIndex];
    for (var _i = 0; _i < marksCard.encodings.length; _i++) {
      var enc = marksCard.encodings[_i];
      map[enc.id] = enc.field;
    }
    return map;
  }

  // =========================================================================
  // Resolve encodings -> data columns
  // =========================================================================

  function resolveEncodings (encodingMap, columns, settings) {
    var resolved = {
      valueCol: null, goalCol: null, goal2Col: null, dateCol: null, ptdCol: null
    };
    var colNameSet = new Set(columns.map(function (c) { return c.fieldName; }));

    function findCol (field) {
      if (!field) return null;
      if (colNameSet.has(field.name)) return field.name;
      for (var cn of colNameSet) {
        if (cn.includes(field.name) || field.name.includes(cn)) return cn;
      }
      return null;
    }

    resolved.valueCol = findCol(encodingMap.value);
    resolved.goalCol = findCol(encodingMap.goal);
    resolved.goal2Col = findCol(encodingMap.goal2);
    resolved.dateCol = findCol(encodingMap.date);

    if (settings && settings.dateFieldName) {
      if (colNameSet.has(settings.dateFieldName)) {
        resolved.dateCol = settings.dateFieldName;
      } else {
        for (var cn of colNameSet) {
          if (cn.toLowerCase().includes(settings.dateFieldName.toLowerCase()) ||
              settings.dateFieldName.toLowerCase().includes(cn.toLowerCase())) {
            resolved.dateCol = cn;
            break;
          }
        }
      }
    }

    if (!resolved.dateCol) {
      for (var _e of Object.entries(encodingMap)) {
        var id = _e[0], field = _e[1];
        if (!field || id === 'value' || id === 'goal' || id === 'goal2') continue;
        if (field.role === 'dimension') {
          var col = findCol(field);
          if (col) { resolved.dateCol = col; break; }
        }
      }
    }

    var usedMeasures = new Set(
      [resolved.valueCol, resolved.goalCol, resolved.goal2Col].filter(Boolean)
    );
    if (!resolved.goalCol && resolved.valueCol) {
      for (var _e2 of Object.entries(encodingMap)) {
        var id2 = _e2[0], field2 = _e2[1];
        if (!field2 || id2 === 'value') continue;
        if (field2.role === 'measure') {
          var col2 = findCol(field2);
          if (col2 && !usedMeasures.has(col2)) {
            resolved.goalCol = col2;
            usedMeasures.add(col2);
            break;
          }
        }
      }
    }
    if (!resolved.goal2Col && resolved.valueCol) {
      for (var _e3 of Object.entries(encodingMap)) {
        var id3 = _e3[0], field3 = _e3[1];
        if (!field3 || id3 === 'value' || id3 === 'goal') continue;
        if (field3.role === 'measure') {
          var col3 = findCol(field3);
          if (col3 && !usedMeasures.has(col3)) {
            resolved.goal2Col = col3;
            usedMeasures.add(col3);
            break;
          }
        }
      }
    }

    resolved.valueField = encodingMap.value || null;
    resolved.goal2Field = encodingMap.goal2 || null;

    if (settings && settings.ptdFieldName) {
      if (colNameSet.has(settings.ptdFieldName)) {
        resolved.ptdCol = settings.ptdFieldName;
      } else {
        for (var cn2 of colNameSet) {
          var cnLower = cn2.toLowerCase();
          var searchLower = settings.ptdFieldName.toLowerCase();
          if (cnLower.includes(searchLower) || searchLower.includes(cnLower)) {
            resolved.ptdCol = cn2;
            break;
          }
        }
      }
    }

    resolved.dateFieldName = '';
    if (encodingMap.date) resolved.dateFieldName = encodingMap.date.name || '';
    else {
      for (var _e4 of Object.entries(encodingMap)) {
        var field4 = _e4[1];
        if (field4 && field4.role === 'dimension') {
          resolved.dateFieldName = field4.name || '';
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
    var n = (dateFieldName || '').toUpperCase();
    return n.includes('QUARTER') || n.includes('QTR');
  }

  function formatPeriodLabel (rawLabel, sortKey, dateFieldName) {
    var s = String(rawLabel).trim();

    var m1 = s.match(/^Q(\d)\s+(\d{4})$/i);
    if (m1) return m1[2] + ' Q' + m1[1];

    var m2 = s.match(/^(\d{4})\s+Q(\d)$/i);
    if (m2) return s;

    var m3 = s.match(/^(\d{4})-Q(\d)$/i);
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
    var year, quarter;
    var mA = periodLabel.match(/^(\d{4})\s*Q(\d)$/i);
    var mB = periodLabel.match(/^Q(\d)\s+(\d{4})$/i);
    var mC = periodLabel.match(/^(\d{4})-Q(\d)$/i);
    if (mA) { year = +mA[1]; quarter = +mA[2]; }
    else if (mB) { quarter = +mB[1]; year = +mB[2]; }
    else if (mC) { year = +mC[1]; quarter = +mC[2]; }

    if (year && quarter >= 1 && quarter <= 4) {
      return new Date(year, (quarter - 1) * 3, 1);
    }

    var mY = periodLabel.match(/^(\d{4})$/);
    if (mY) return new Date(+mY[1], 0, 1);

    var mMonth = periodLabel.match(/^(\d{4})-(\d{2})$/) ||
                 periodLabel.match(/^(\w+)\s+(\d{4})$/i);
    if (mMonth) {
      var mYear, mMon;
      if (/^\d{4}-\d{2}$/.test(periodLabel)) {
        mYear = +mMonth[1]; mMon = +mMonth[2] - 1;
      } else {
        var parsed = new Date(periodLabel + ' 1');
        if (!isNaN(parsed)) { mYear = parsed.getFullYear(); mMon = parsed.getMonth(); }
      }
      if (mYear && mMon != null) return new Date(mYear, mMon, 1);
    }
    return null;
  }

  function getPeriodEndDate (periodLabel) {
    var year, quarter;
    var mA = periodLabel.match(/^(\d{4})\s*Q(\d)$/i);
    var mB = periodLabel.match(/^Q(\d)\s+(\d{4})$/i);
    var mC = periodLabel.match(/^(\d{4})-Q(\d)$/i);
    if (mA) { year = +mA[1]; quarter = +mA[2]; }
    else if (mB) { quarter = +mB[1]; year = +mB[2]; }
    else if (mC) { year = +mC[1]; quarter = +mC[2]; }

    if (year && quarter >= 1 && quarter <= 4) {
      return new Date(year, (quarter - 1) * 3 + 3, 0);
    }

    var mY = periodLabel.match(/^(\d{4})$/);
    if (mY) return new Date(+mY[1], 11, 31);

    var mMonth = periodLabel.match(/^(\d{4})-(\d{2})$/) ||
                 periodLabel.match(/^(\w+)\s+(\d{4})$/i);
    if (mMonth) {
      var mYear, mMon;
      if (/^\d{4}-\d{2}$/.test(periodLabel)) {
        mYear = +mMonth[1]; mMon = +mMonth[2] - 1;
      } else {
        var parsed = new Date(periodLabel + ' 1');
        if (!isNaN(parsed)) { mYear = parsed.getFullYear(); mMon = parsed.getMonth(); }
      }
      if (mYear && mMon != null) return new Date(mYear, mMon + 1, 0);
    }
    return null;
  }

  function inferPreviousPeriodLabel (currentLabel) {
    var s = String(currentLabel).trim();
    var mQ = s.match(/^(\d{4})\s*Q(\d)$/i) || s.match(/^(\d{4})-Q(\d)$/i);
    if (mQ) {
      var year = +mQ[1];
      var quarter = +mQ[2];
      if (quarter === 1) { year--; quarter = 4; }
      else { quarter--; }
      return year + ' Q' + quarter;
    }
    var mY = s.match(/^(\d{4})$/);
    if (mY) return String(+mY[1] - 1);
    var mM = s.match(/^(\d{4})-(\d{2})$/);
    if (mM) {
      var yr = +mM[1];
      var mo = +mM[2];
      if (mo === 1) { yr--; mo = 12; }
      else { mo--; }
      return yr + '-' + String(mo).padStart(2, '0');
    }
    return 'Previous';
  }

  // =========================================================================
  // KPI computation
  // =========================================================================

  function computeKpi (dataResult, encodings, settings) {
    var data = dataResult.rows;
    var columns = dataResult.columns;
    var r = resolveEncodings(encodings, columns, settings);

    if (!r.valueCol) return null;

    var valueLabel = r.valueField ? r.valueField.name : r.valueCol;
    var s = settings || {};
    var globalFmt = buildGlobalFormatter(s);

    // No date column — simple aggregate
    if (!r.dateCol) {
      var totalValue = 0, totalGoal = 0, totalGoal2 = 0;
      for (var _i = 0; _i < data.length; _i++) {
        var row = data[_i];
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
        globalFmt: globalFmt
      };
    }

    // Group rows by date period
    var periodMap = new Map();

    for (var _j = 0; _j < data.length; _j++) {
      var row2 = data[_j];
      var dateVal = row2[r.dateCol];
      if (!dateVal) continue;
      var rawValue = dateVal.value;
      if (rawValue == null || rawValue === '' || String(rawValue).toLowerCase() === 'null') continue;
      var rawLabel = String(dateVal.formattedValue ?? rawValue);
      if (rawLabel === '' || rawLabel.toLowerCase() === 'null') continue;

      var sortKey = rawValue;
      var periodKey = rawLabel;

      var rowDate = null;
      if (rawValue instanceof Date) rowDate = new Date(rawValue);
      else if (typeof rawValue === 'number') rowDate = new Date(rawValue);
      else if (typeof rawValue === 'string') {
        var parsed = new Date(rawValue);
        if (!isNaN(parsed)) rowDate = parsed;
      }

      if (!periodMap.has(periodKey)) {
        periodMap.set(periodKey, {
          value: 0, goal: 0, goal2: 0, ptdValue: 0, sortKey: sortKey,
          label: formatPeriodLabel(rawLabel, sortKey, r.dateFieldName),
          rows: []
        });
      }
      var bucket = periodMap.get(periodKey);

      bucket.rows.push({
        date: rowDate,
        value: Number(row2[r.valueCol]?.value) || 0,
        goal: r.goalCol ? Number(row2[r.goalCol]?.value) || 0 : 0,
        goal2: r.goal2Col ? Number(row2[r.goal2Col]?.value) || 0 : 0
      });

      bucket.value += Number(row2[r.valueCol]?.value) || 0;
      if (r.goalCol) bucket.goal += Number(row2[r.goalCol]?.value) || 0;
      if (r.goal2Col) bucket.goal2 += Number(row2[r.goal2Col]?.value) || 0;
      if (r.ptdCol) bucket.ptdValue += Number(row2[r.ptdCol]?.value) || 0;
    }

    var allPeriods = Array.from(periodMap.values()).sort(function (a, b) {
      if (a.sortKey < b.sortKey) return -1;
      if (a.sortKey > b.sortKey) return 1;
      return 0;
    });

    if (allPeriods.length === 0) return null;

    var periods = allPeriods.slice(-4);
    var current = periods[periods.length - 1];
    var previous = periods.length > 1 ? periods[periods.length - 2] : null;

    // PTD values
    if (!r.ptdCol) {
      var today = new Date();
      today.setHours(0, 0, 0, 0);
      var currentStart = getPeriodStartDate(current.label);
      var elapsedDays = null;
      if (currentStart) {
        var currentEnd = getPeriodEndDate(current.label);
        if (currentEnd && today <= currentEnd) {
          elapsedDays = Math.floor((today - currentStart) / 864e5) + 1;
        }
      }
      for (var _k = 0; _k < periods.length; _k++) {
        var p = periods[_k];
        p.ptdValue = 0;
        if (elapsedDays === null || !p.rows || p.rows.length === 0) {
          p.ptdValue = null; continue;
        }
        var periodStart = getPeriodStartDate(p.label);
        if (!periodStart) { p.ptdValue = null; continue; }
        for (var _l = 0; _l < p.rows.length; _l++) {
          var prow = p.rows[_l];
          if (!prow.date) continue;
          var dayOfPeriod = Math.floor((prow.date - periodStart) / 864e5) + 1;
          if (dayOfPeriod >= 1 && dayOfPeriod <= elapsedDays) p.ptdValue += prow.value;
        }
      }
    }

    var delta = null;
    if (previous && previous.value !== 0) {
      delta = ((current.value - previous.value) / Math.abs(previous.value)) * 100;
    }

    var goalPct = null, goalValue = null, formattedGoal = null;
    if (r.goalCol) {
      goalValue = current.goal;
      formattedGoal = globalFmt(goalValue);
      if (goalValue !== 0) goalPct = Math.round((current.value / goalValue) * 100);
    }
    var goal2Pct = null, goal2Value = null, formattedGoal2 = null;
    if (r.goal2Col) {
      goal2Value = current.goal2;
      formattedGoal2 = globalFmt(goal2Value);
      if (goal2Value !== 0) goal2Pct = Math.round((current.value / goal2Value) * 100);
    }

    var ptdDelta = null;
    if (r.ptdCol && previous && previous.ptdValue !== 0) {
      ptdDelta = ((current.value - previous.ptdValue) / Math.abs(previous.ptdValue)) * 100;
    } else if (previous && current.ptdValue !== null && previous.ptdValue !== null && previous.ptdValue !== 0) {
      ptdDelta = ((current.ptdValue - previous.ptdValue) / Math.abs(previous.ptdValue)) * 100;
    }

    var sparkData = periods.map(function (p) { return { label: p.label, value: p.value }; });
    if (sparkData.length === 1) {
      var prevLabel = inferPreviousPeriodLabel(sparkData[0].label);
      sparkData.unshift({ label: prevLabel, value: 0 });
    }

    var ptdSparkData = null;
    if (periods.every(function (p) { return p.ptdValue !== null && p.ptdValue !== undefined; })) {
      ptdSparkData = periods.map(function (p) { return { label: p.label, value: p.ptdValue }; });
      if (ptdSparkData.length === 1) {
        var prevLabel2 = inferPreviousPeriodLabel(ptdSparkData[0].label);
        ptdSparkData.unshift({ label: prevLabel2, value: 0 });
      }
    }

    var periodLabel = previous
      ? periods[0].label + ' \u2014 ' + current.label
      : current.label;

    return {
      label: valueLabel,
      currentValue: current.value,
      formattedValue: globalFmt(current.value),
      previousValue: previous ? previous.value : null,
      delta: delta,
      ptdDelta: ptdDelta,
      goalValue: goalValue, goalPct: goalPct, formattedGoal: formattedGoal,
      goal2Value: goal2Value, goal2Pct: goal2Pct, formattedGoal2: formattedGoal2,
      goal2Label: r.goal2Field ? r.goal2Field.name : 'Secondary Goal',
      sparkData: sparkData,
      ptdSparkData: ptdSparkData,
      periodLabel: periodLabel,
      globalFmt: globalFmt
    };
  }

  // =========================================================================
  // Public API
  // =========================================================================

  return {
    convertToListOfNamedRows: convertToListOfNamedRows,
    getSummaryDataTable: getSummaryDataTable,
    getEncodingMap: getEncodingMap,
    resolveEncodings: resolveEncodings,
    computeKpi: computeKpi,
    inferPreviousPeriodLabel: inferPreviousPeriodLabel
  };

})();
