(function () {
  'use strict';

  // ---------- Core calculation (mirrors the supplied Excel formulas) ----------
  function calculate(d) {
    var incomeAfter = d.incomeBefore + d.arrears;                        // B6 = B4 + B5
    var extraReceipt = Math.max(0, d.taxAfter - d.taxBefore);            // B8 = MAX(0, B7 - B3)
    var allocTotal = 0, extraEarlier = 0;
    var yearExtras = d.years.map(function (y) {
      allocTotal += y.alloc;                                             // B15 = SUM(B10:B14)
      var e = Math.max(0, y.taxWith - y.taxWithout);                     // E = MAX(0, D - C)
      extraEarlier += e;                                                 // B17 = SUM(E10:E14)
      return e;
    });
    var relief = Math.max(0, extraReceipt - extraEarlier);               // B18 = MAX(0, B8 - B17)
    var net = Math.max(0, extraReceipt - relief);                        // B19 = MAX(0, B8 - B18)
    return {
      incomeAfter: incomeAfter, extraReceipt: extraReceipt, allocTotal: allocTotal,
      extraEarlier: extraEarlier, relief: relief, net: net, yearExtras: yearExtras,
      allocDiff: d.arrears - allocTotal
    };
  }
  if (typeof module !== 'undefined') module.exports = { calculate: calculate };
  if (typeof document === 'undefined') return;

  // ---------- Helpers ----------
  var $ = function (id) { return document.getElementById(id); };
  var fmt = function (n) {
    return '₹ ' + (Math.round(n * 100) / 100).toLocaleString('en-IN', { maximumFractionDigits: 2 });
  };
  var num = function (el) { var v = parseFloat(el.value); return isNaN(v) || v < 0 ? 0 : v; };

  var yearRows = $('yearRows');
  var MAX_YEARS = 15;
  var counter = 0;

  function addYear(label) {
    if (yearRows.children.length >= MAX_YEARS) return;
    counter++;
    var div = document.createElement('div');
    div.className = 'yrow';
    div.innerHTML =
      '<div class="yhead"><input type="text" class="y-label" placeholder="Year e.g. 2022-23" value="' + (label || '') + '" aria-label="Financial year">' +
      '<button type="button" class="rm">Remove</button></div>' +
      '<div class="grid3">' +
      '<label class="field"><span class="lbl">Arrears belonging to this year (₹)</span><input type="number" class="y-alloc" min="0" step="any" inputmode="decimal" placeholder="0"></label>' +
      '<label class="field"><span class="lbl">Tax <em>without</em> this arrear (₹)</span><input type="number" class="y-without" min="0" step="any" inputmode="decimal" placeholder="0"></label>' +
      '<label class="field"><span class="lbl">Tax <em>with</em> this arrear (₹)</span><input type="number" class="y-with" min="0" step="any" inputmode="decimal" placeholder="0"></label>' +
      '</div><div class="yfoot">Extra tax for this year: <b class="y-extra">₹ 0</b></div>';
    div.querySelector('.rm').addEventListener('click', function () {
      div.remove();
      if (!yearRows.children.length) addYear();
      update();
    });
    yearRows.appendChild(div);
  }

  function readData() {
    var years = Array.prototype.map.call(yearRows.children, function (r) {
      return {
        alloc: num(r.querySelector('.y-alloc')),
        taxWithout: num(r.querySelector('.y-without')),
        taxWith: num(r.querySelector('.y-with'))
      };
    });
    return {
      taxBefore: num($('taxBefore')), incomeBefore: num($('incomeBefore')),
      arrears: num($('arrears')), taxAfter: num($('taxAfter')), years: years
    };
  }

  function update() {
    var d = readData(), r = calculate(d);
    $('incomeAfter').textContent = fmt(r.incomeAfter);
    Array.prototype.forEach.call(yearRows.children, function (row, i) {
      row.querySelector('.y-extra').textContent = fmt(r.yearExtras[i]);
    });

    // Allocation check
    var c = $('allocCheck');
    if (d.arrears === 0 && r.allocTotal === 0) {
      c.className = 'check';
      c.textContent = 'Allocated so far: ₹ 0. The amounts for all earlier years must add up to your total arrears.';
    } else if (Math.abs(r.allocDiff) < 0.005) {
      c.className = 'check ok';
      c.textContent = '✔ OK – allocated ' + fmt(r.allocTotal) + ' matches total arrears.';
    } else if (r.allocDiff > 0) {
      c.className = 'check warn';
      c.textContent = 'CHECK allocation – ' + fmt(r.allocDiff) + ' of arrears is still not allocated (allocated ' + fmt(r.allocTotal) + ' of ' + fmt(d.arrears) + ').';
    } else {
      c.className = 'check bad';
      c.textContent = 'CHECK allocation – you have allocated ' + fmt(-r.allocDiff) + ' more than total arrears.';
    }

    // Results
    $('rExtraReceipt').textContent = fmt(r.extraReceipt);
    $('rExtraEarlier').textContent = fmt(r.extraEarlier);
    $('rRelief').textContent = fmt(r.relief);
    $('rNet').textContent = fmt(r.net);
    $('reliefBig').textContent = fmt(r.relief);

    var m = $('resultMsg');
    var ready = d.arrears > 0 && d.taxAfter > 0;
    if (!ready) m.textContent = 'Fill in the details to see your result.';
    else if (Math.abs(r.allocDiff) >= 0.005) m.textContent = 'Result shown, but the allocation does not match the total arrears yet – please check Step 2.';
    else if (r.extraReceipt === 0) m.textContent = 'The arrears did not increase your tax this year, so there is nothing to relieve.';
    else if (r.relief > 0) m.textContent = 'Spreading the arrears over earlier years would have cost less tax. You may claim relief of ' + fmt(r.relief) + '.';
    else m.textContent = 'Earlier years would have cost the same or more tax, so no relief is indicated.';
  }

  function resetAll() {
    ['taxBefore', 'incomeBefore', 'arrears', 'taxAfter'].forEach(function (id) { $(id).value = ''; });
    yearRows.innerHTML = '';
    for (var i = 0; i < 3; i++) addYear();
    update();
  }

  document.addEventListener('input', update);
  $('addYear').addEventListener('click', function () { addYear(); update(); });
  $('resetBtn').addEventListener('click', resetAll);
  $('printBtn').addEventListener('click', function () { window.print(); });

  resetAll();
})();
