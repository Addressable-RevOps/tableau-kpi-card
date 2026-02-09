'use strict';
/* global d3 */

/**
 * kpi-sparkline.js — D3 sparkline chart rendering (area, lines, dots, labels, hover, animation).
 * Exports onto window.KPI.sparkline
 */

window.KPI = window.KPI || {};
window.KPI.sparkline = (function () {

  var escapeHtml = window.KPI.utils.escapeHtml;

  function drawSparkline (container, sparkData, ptdSparkData, settings, chartHeight, valueFmt) {
    var margin = { top: 28, right: 8, bottom: settings.showSparkPeriods ? 24 : 8, left: 8 };
    var width = container.offsetWidth || 400;
    var totalWidth = width - margin.left - margin.right;
    var totalHeight = chartHeight || 130;
    var chartH = totalHeight - margin.top - margin.bottom;

    var svg = d3.create('svg')
      .attr('width', width)
      .attr('height', totalHeight + margin.bottom)
      .attr('viewBox', [0, 0, width, totalHeight + margin.bottom])
      .style('overflow', 'visible');

    // Color mode: default (brand colors), accent, or custom
    var colorMode = (settings && settings.sparklineColorMode) || 'default';
    var accentColor = (settings && settings.gradColor) || '#D42F8A';
    var customColor = (settings && settings.sparklineCustomColor) || '#3b82f6';

    var lineColor, dotColor, ptdColor;
    if (colorMode === 'accent') {
      lineColor = accentColor;
      dotColor = accentColor;
      ptdColor = accentColor;
    } else if (colorMode === 'custom') {
      lineColor = customColor;
      dotColor = customColor;
      ptdColor = customColor;
    } else {
      lineColor = '#EF5A5E';
      dotColor = '#D42F8A';
      ptdColor = '#7B2FBE';
    }

    var defs = svg.append('defs');
    var grad = defs.append('linearGradient')
      .attr('id', 'sparkGradient')
      .attr('x1', '0%').attr('y1', '0%')
      .attr('x2', '0%').attr('y2', '100%');
    grad.append('stop').attr('offset', '0%').attr('stop-color', lineColor).attr('stop-opacity', 0.08);
    grad.append('stop').attr('offset', '100%').attr('stop-color', lineColor).attr('stop-opacity', 0.02);

    var g = svg.append('g')
      .attr('transform', 'translate(' + margin.left + ',' + margin.top + ')');

    var x = d3.scalePoint()
      .domain(sparkData.map(function (_, i) { return i; }))
      .range([0, totalWidth]);

    var allValues = sparkData.map(function (d) { return d.value; });
    if (ptdSparkData) ptdSparkData.forEach(function (d) { allValues.push(d.value); });
    var yExtent = d3.extent(allValues);
    var yPad = (yExtent[1] - yExtent[0]) * 0.2 || 1;
    var y = d3.scaleLinear()
      .domain([yExtent[0] - yPad, yExtent[1] + yPad])
      .range([chartH, 0]);

    var area = d3.area()
      .x(function (_, i) { return x(i); })
      .y0(chartH)
      .y1(function (d) { return y(d.value); })
      .curve(d3.curveMonotoneX);

    g.append('path')
      .datum(sparkData)
      .attr('class', 'spark-area')
      .attr('d', area);

    // Secondary comparison line (dashed)
    if (ptdSparkData) {
      var ptdLine = d3.line()
        .x(function (_, i) { return x(i); })
        .y(function (d) { return y(d.value); })
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
        .attr('cx', function (_, i) { return x(i); })
        .attr('cy', function (d) { return y(d.value); })
        .attr('r', 2.5)
        .attr('fill', ptdColor)
        .attr('opacity', 0.7);
    }

    // Actual line
    var line = d3.line()
      .x(function (_, i) { return x(i); })
      .y(function (d) { return y(d.value); })
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
      .attr('cx', function (_, i) { return x(i); })
      .attr('cy', function (d) { return y(d.value); })
      .attr('r', 3.5)
      .attr('fill', dotColor);

    // Smart data labels with collision avoidance
    if (settings.showSparkLabels) {
      var labels = [];

      sparkData.forEach(function (d, i) {
        labels.push({
          x: x(i), naturalY: y(d.value) - 10,
          text: valueFmt(d.value),
          fill: '#1a1a2e', opacity: 1, fontSize: 11, series: 'actual'
        });
      });

      if (ptdSparkData) {
        ptdSparkData.forEach(function (d, i) {
          var actualY = y(sparkData[i].value);
          var ptdY = y(d.value);
          if (Math.abs(actualY - ptdY) < 4 && Math.abs(d.value - sparkData[i].value) < 0.5) return;
          labels.push({
            x: x(i), naturalY: y(d.value) + 14,
            text: valueFmt(d.value),
            fill: '#4a5568', opacity: 0.8, fontSize: 10, series: 'ptd'
          });
        });
      }

      labels.sort(function (a, b) { return a.x - b.x || a.naturalY - b.naturalY; });

      var minGap = 12;
      var resolvedLabels = [];
      for (var _i = 0; _i < labels.length; _i++) {
        var lbl = labels[_i];
        var finalY = lbl.naturalY;
        for (var _j = 0; _j < resolvedLabels.length; _j++) {
          var placed = resolvedLabels[_j];
          if (Math.abs(placed.x - lbl.x) < 30 && Math.abs(finalY - placed.y) < minGap) {
            if (lbl.series === 'ptd') { finalY = placed.y + minGap; }
            else { finalY = placed.y - minGap; }
          }
        }
        finalY = Math.max(-margin.top + 8, Math.min(chartH + 6, finalY));
        resolvedLabels.push(Object.assign({}, lbl, { y: finalY }));
      }

      for (var _k = 0; _k < resolvedLabels.length; _k++) {
        var rl = resolvedLabels[_k];
        var anchor = 'middle';
        if (rl.x <= 10) anchor = 'start';
        else if (rl.x >= totalWidth - 10) anchor = 'end';

        g.append('text')
          .attr('x', rl.x)
          .attr('y', rl.y)
          .attr('text-anchor', anchor)
          .attr('font-size', rl.fontSize + 'px')
          .attr('font-weight', '600')
          .attr('fill', rl.fill)
          .attr('opacity', rl.opacity)
          .text(rl.text);
      }
    }

    // Period labels
    if (settings.showSparkPeriods) {
      g.selectAll('.spark-label')
        .data(sparkData)
        .join('text')
        .attr('x', function (_, i) { return x(i); })
        .attr('y', chartH + 16)
        .attr('text-anchor', function (_, i) {
          if (i === 0) return 'start';
          if (i === sparkData.length - 1) return 'end';
          return 'middle';
        })
        .attr('font-size', '10px')
        .attr('fill', '#a0aab4')
        .text(function (d) { return d.label; });
    }

    // Hover interaction
    var guideLine = g.append('line')
      .attr('class', 'spark-guide')
      .attr('y1', -margin.top + 4)
      .attr('y2', chartH + 4);

    var hoverDot = g.append('circle')
      .attr('r', 6)
      .attr('fill', dotColor)
      .attr('stroke', '#fff')
      .attr('stroke-width', 2)
      .style('pointer-events', 'none')
      .style('opacity', 0);

    var hoverDotPtd = null;
    if (ptdSparkData) {
      hoverDotPtd = g.append('circle')
        .attr('r', 5)
        .attr('fill', ptdColor)
        .attr('stroke', '#fff')
        .attr('stroke-width', 2)
        .style('pointer-events', 'none')
        .style('opacity', 0);
    }

    var tooltip = g.append('g')
      .style('pointer-events', 'none')
      .style('opacity', 0);

    var tooltipBg = tooltip.append('rect').attr('class', 'spark-tooltip-bg');
    var tooltipVal = tooltip.append('text').attr('class', 'spark-tooltip-val');
    var tooltipLbl = tooltip.append('text').attr('class', 'spark-tooltip-lbl');
    var tooltipPtd = null;
    if (ptdSparkData) {
      tooltipPtd = tooltip.append('text').attr('class', 'spark-tooltip-ptd');
    }

    g.append('rect')
      .attr('width', totalWidth)
      .attr('height', chartH)
      .attr('fill', 'transparent')
      .style('cursor', 'crosshair')
      .on('mousemove', function (event) {
        var mx = d3.pointer(event)[0];
        var minDist = Infinity, idx = 0;
        sparkData.forEach(function (_, i) {
          var dist = Math.abs(x(i) - mx);
          if (dist < minDist) { minDist = dist; idx = i; }
        });

        var px = x(idx);
        var py = y(sparkData[idx].value);

        guideLine.attr('x1', px).attr('x2', px).classed('visible', true);
        hoverDot.attr('cx', px).attr('cy', py).style('opacity', 1);

        if (hoverDotPtd && ptdSparkData && ptdSparkData[idx]) {
          hoverDotPtd
            .attr('cx', px)
            .attr('cy', y(ptdSparkData[idx].value))
            .style('opacity', 1);
        }

        var valText = valueFmt(sparkData[idx].value);
        var lblText = sparkData[idx].label || '';
        tooltipVal.text(valText);
        tooltipLbl.text(lblText);

        var tooltipH = 42;
        var lineSpacing = 16;
        tooltipVal.attr('x', 10).attr('y', 18);
        tooltipLbl.attr('x', 10).attr('y', 18 + lineSpacing);

        if (tooltipPtd && ptdSparkData && ptdSparkData[idx]) {
          var ptdText = (settings.ptdBadgeLabel || 'Comp') + ': ' + valueFmt(ptdSparkData[idx].value);
          tooltipPtd.text(ptdText).attr('x', 10).attr('y', 18 + lineSpacing * 2);
          tooltipH = 42 + lineSpacing;
        }

        var valBBox = tooltipVal.node().getBBox();
        var lblBBox = tooltipLbl.node().getBBox();
        var maxW = Math.max(valBBox.width, lblBBox.width);
        if (tooltipPtd && ptdSparkData && ptdSparkData[idx]) {
          var ptdBBox = tooltipPtd.node().getBBox();
          maxW = Math.max(maxW, ptdBBox.width);
        }
        var tooltipW = maxW + 20;

        tooltipBg.attr('width', tooltipW).attr('height', tooltipH);

        var tx = px + 12;
        if (tx + tooltipW > totalWidth) tx = px - tooltipW - 12;
        var ty = py - tooltipH - 10;
        if (ty < -margin.top) ty = py + 16;

        tooltip.attr('transform', 'translate(' + tx + ',' + ty + ')').style('opacity', 1);
      })
      .on('mouseleave', function () {
        guideLine.classed('visible', false);
        hoverDot.style('opacity', 0);
        if (hoverDotPtd) hoverDotPtd.style('opacity', 0);
        tooltip.style('opacity', 0);
      });

    container.appendChild(svg.node());

    // Open animation: draw lines + pop dots
    if (settings._animate) {
      var animDuration = 600;

      svg.selectAll('.spark-line, .spark-line-ptd').each(function () {
        var path = d3.select(this);
        var len = this.getTotalLength();
        path
          .attr('stroke-dasharray', len)
          .attr('stroke-dashoffset', len)
          .transition()
          .duration(animDuration)
          .ease(d3.easeCubicOut)
          .attr('stroke-dashoffset', 0)
          .on('end', function () {
            d3.select(this).attr('stroke-dasharray', null);
          });
      });

      svg.selectAll('.spark-area')
        .style('opacity', 0)
        .transition()
        .delay(animDuration * 0.3)
        .duration(animDuration * 0.7)
        .style('opacity', 1);

      svg.selectAll('.spark-dot, .spark-dot-ptd').each(function (d, i) {
        d3.select(this)
          .attr('r', 0)
          .transition()
          .delay(animDuration * 0.3 + i * 60)
          .duration(300)
          .ease(d3.easeBackOut.overshoot(1.5))
          .attr('r', d3.select(this).classed('spark-dot-ptd') ? 2.5 : 3.5);
      });

      svg.selectAll('text')
        .style('opacity', 0)
        .transition()
        .delay(animDuration * 0.5)
        .duration(300)
        .style('opacity', null);
    }

    // Legend
    if (ptdSparkData && settings.showLegend) {
      var compLegend = settings.ptdLegendLabel || 'Comparison';
      var legend = document.createElement('div');
      legend.className = 'spark-legend';
      legend.innerHTML =
        '<span class="spark-legend-item"><span class="spark-legend-line solid"></span>Actual</span>' +
        '<span class="spark-legend-item"><span class="spark-legend-line dashed"></span>' + escapeHtml(compLegend) + '</span>';
      container.appendChild(legend);
    }
  }

  return {
    drawSparkline: drawSparkline
  };

})();
