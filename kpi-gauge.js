'use strict';
/* global d3 */

/**
 * kpi-gauge.js — Semi-circle gauge renderer using D3 arc generators.
 * Exports onto window.KPI.gauge
 */

window.KPI = window.KPI || {};
window.KPI.gauge = (function () {

  var buildGradientFromColor = window.KPI.utils.buildGradientFromColor;
  var escapeHtml             = window.KPI.utils.escapeHtml;

  /**
   * Draw a semi-circle (180 deg) gauge inside `container`.
   *
   * @param {HTMLElement} container   - DOM element to append SVG into
   * @param {object}      kpi        - computed KPI data
   * @param {object}      settings   - user settings
   */
  function drawGauge (container, kpi, settings) {
    // --- Dimensions ---
    var containerW = container.offsetWidth || 300;
    var radius     = Math.min(containerW * 0.44, 180);
    var thickness  = Math.max(16, radius * 0.18);
    var innerR     = radius - thickness;
    var svgW       = radius * 2 + 40;        // horizontal breathing room
    var svgH       = radius + thickness + 36; // half-circle + space for labels below

    var cx = svgW / 2;
    var cy = radius + 10; // center of arcs

    var svg = d3.create('svg')
      .attr('width', svgW)
      .attr('height', svgH)
      .attr('viewBox', '0 0 ' + svgW + ' ' + svgH)
      .style('overflow', 'visible');

    var g = svg.append('g')
      .attr('transform', 'translate(' + cx + ',' + cy + ')');

    // --- Arc generators ---
    // Semi-circle: from -PI/2 (left) to PI/2 (right)
    var startAngle = -Math.PI / 2;
    var endAngle   =  Math.PI / 2;

    var arcGen = d3.arc()
      .innerRadius(innerR)
      .outerRadius(radius)
      .cornerRadius(thickness / 2);

    // --- Track (gray background) ---
    g.append('path')
      .attr('d', arcGen({ startAngle: startAngle, endAngle: endAngle }))
      .attr('fill', '#f0edf5');

    // --- Fill arc (primary goal) ---
    var goalPct = kpi.goalPct;
    var hasGoal = goalPct !== null && goalPct !== undefined;
    var fillPct = hasGoal ? Math.max(0, Math.min(goalPct, 100)) / 100 : 0;
    var fillAngle = startAngle + fillPct * Math.PI;

    // Gradient definition for the fill
    var gradId = 'gaugeGrad_' + Math.random().toString(36).slice(2, 8);
    var defs = svg.append('defs');

    var barMode = (settings && settings.barColorMode) || 'default';
    var fillColor;

    if (barMode === 'accent') {
      var accent = (settings && settings.gradColor) || '#D42F8A';
      var stops = buildGradientFromColor(accent);
      var lg = defs.append('linearGradient')
        .attr('id', gradId)
        .attr('x1', '0%').attr('y1', '0%')
        .attr('x2', '100%').attr('y2', '0%');
      lg.append('stop').attr('offset', '0%').attr('stop-color', stops[0]);
      lg.append('stop').attr('offset', '50%').attr('stop-color', stops[1]);
      lg.append('stop').attr('offset', '100%').attr('stop-color', stops[2]);
      fillColor = 'url(#' + gradId + ')';
    } else if (barMode === 'custom') {
      fillColor = (settings && settings.barCustomColor) || '#3b82f6';
    } else {
      // default: brand gradient
      var lgd = defs.append('linearGradient')
        .attr('id', gradId)
        .attr('x1', '0%').attr('y1', '0%')
        .attr('x2', '100%').attr('y2', '0%');
      lgd.append('stop').attr('offset', '0%').attr('stop-color', '#F7941D');
      lgd.append('stop').attr('offset', '33%').attr('stop-color', '#EF5A5E');
      lgd.append('stop').attr('offset', '66%').attr('stop-color', '#D42F8A');
      lgd.append('stop').attr('offset', '100%').attr('stop-color', '#7B2FBE');
      fillColor = 'url(#' + gradId + ')';
    }

    if (hasGoal && fillPct > 0) {
      var fillPath = g.append('path')
        .attr('fill', fillColor);

      if (settings._animate) {
        // Animate: sweep from start to final angle
        fillPath
          .attr('d', arcGen({ startAngle: startAngle, endAngle: startAngle + 0.001 }))
          .transition()
          .duration(700)
          .ease(d3.easeCubicOut)
          .attrTween('d', function () {
            var interp = d3.interpolate(startAngle + 0.001, fillAngle);
            return function (t) {
              return arcGen({ startAngle: startAngle, endAngle: interp(t) });
            };
          });
      } else {
        fillPath.attr('d', arcGen({ startAngle: startAngle, endAngle: fillAngle }));
      }
    }

    // Secondary goal is hidden in gauge mode (not enough space to render cleanly)

    // --- Center text: value + goal percentage ---
    var valueSize = parseInt(settings.valueSize, 10);
    if (!(valueSize > 0)) valueSize = 42;
    // Scale value font to fit inside gauge — cap relative to inner radius
    var maxFontPx = Math.floor(innerR * 0.6);
    var actualFontPx = Math.min(valueSize, maxFontPx);

    g.append('text')
      .attr('class', 'kpi-gauge-value')
      .attr('x', 0)
      .attr('y', -10)
      .attr('text-anchor', 'middle')
      .attr('dominant-baseline', 'auto')
      .attr('font-size', actualFontPx + 'px')
      .text(kpi.formattedValue);

    if (hasGoal) {
      g.append('text')
        .attr('class', 'kpi-gauge-pct')
        .attr('x', 0)
        .attr('y', actualFontPx * 0.38)
        .attr('text-anchor', 'middle')
        .text(goalPct + '% of goal');
    }

    // --- Goal label underneath arc (left = 0%, right = goal value) ---
    if (hasGoal) {
      var goalLabel = settings.goalLabel || 'Goal';
      var formattedGoal = kpi.formattedGoal || '';

      g.append('text')
        .attr('class', 'kpi-gauge-goal-label')
        .attr('x', -radius + 2)
        .attr('y', 18)
        .attr('text-anchor', 'start')
        .text('0');

      g.append('text')
        .attr('class', 'kpi-gauge-goal-label')
        .attr('x', radius - 2)
        .attr('y', 18)
        .attr('text-anchor', 'end')
        .text(goalLabel + ': ' + formattedGoal);
    }

    // --- Animate center text fade-in ---
    if (settings._animate) {
      g.selectAll('text')
        .style('opacity', 0)
        .transition()
        .delay(350)
        .duration(300)
        .style('opacity', null);
    }

    container.appendChild(svg.node());
  }

  // =========================================================================
  // Public API
  // =========================================================================

  return {
    drawGauge: drawGauge
  };

})();
