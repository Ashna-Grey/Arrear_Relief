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

  // ---------- Translations ----------
  var I18N = {
    en: {
      title: 'Arrear Relief Calculator',
      badge: 'Income-tax Act, 2025 · Section 157 · Form 39',
      lead: 'Received salary arrears (like DA arrears) in one year that belong to earlier years? Find out how much tax relief you may claim. Fill in the boxes below – the answer updates instantly.',
      step1: 'Enter your tax for the <b>year you received</b> the arrears',
      step2: 'Split the arrears across the <b>earlier years</b> they belong to',
      step3: 'See your <b>estimated relief</b> instantly',
      c1title: 'Year in which arrears were received',
      c1hint: 'Use the financial year in which the arrears were actually credited to you.',
      f1l: 'Tax payable <em>before</em> adding arrears (₹)',
      f1s: 'Total tax, including rebate &amp; cess. Do not subtract TDS.',
      f2l: 'Taxable income <em>before</em> adding arrears (₹)',
      f2s: 'Income after deductions – <b>not</b> your gross salary.',
      f3l: 'Total arrears received (₹)',
      f3s: 'Sum of all arrears credited this year.',
      f4l: 'Tax payable <em>after</em> adding arrears (₹)',
      f4s: 'Recalculate your tax with the arrears included.',
      auto: 'Taxable income after arrears:',
      c2title: 'Earlier years the arrears belong to',
      c2hint: 'For each earlier year, enter the part of the arrears that belongs to it, and your tax for that year <em>without</em> and <em>with</em> that part added.',
      addYear: '+ Add another year',
      reset: 'Reset all',
      print: 'Print / Save as PDF',
      yPh: 'Year e.g. 2022-23',
      yAlloc: 'Arrears belonging to this year (₹)',
      yWithout: 'Tax <em>without</em> this arrear (₹)',
      yWith: 'Tax <em>with</em> this arrear (₹)',
      yExtra: 'Extra tax for this year:',
      remove: 'Remove',
      rTitle: 'Your result',
      rBig: 'Estimated relief under Section 157',
      r1: 'Extra tax in receipt year',
      r2: 'Extra tax if spread over earlier years',
      r3: 'Estimated relief',
      r4: 'Net extra tax after relief',
      msgEmpty: 'Fill in the details to see your result.',
      msgMismatch: 'Result shown, but the allocation does not match the total arrears yet – please check Step 2.',
      msgNoExtra: 'The arrears did not increase your tax this year, so there is nothing to relieve.',
      msgRelief: 'Spreading the arrears over earlier years would have cost less tax. You may claim relief of {0}.',
      msgNone: 'Earlier years would have cost the same or more tax, so no relief is indicated.',
      chkEmpty: 'Allocated so far: ₹ 0. The amounts for all earlier years must add up to your total arrears.',
      chkOk: '✔ OK – allocated {0} matches total arrears.',
      chkLess: 'CHECK allocation – {0} of arrears is still not allocated (allocated {1} of {2}).',
      chkMore: 'CHECK allocation – you have allocated {0} more than total arrears.',
      gTitle: 'What do these terms mean?',
      g1t: 'Receipt year', g1d: 'The financial year in which the arrears were actually paid into your account – not the year they were earned.',
      g2t: 'Arrears', g2d: 'Salary or DA that was due to you in earlier years but paid to you late, in a lump sum.',
      g3t: 'Taxable income (not gross salary)', g3d: 'What is left of your income after standard deduction and other allowed deductions. Use the figure on which your tax is calculated.',
      g4t: 'Tax payable', g4d: 'Total tax including applicable rebate and cess. Do <b>not</b> subtract TDS already deducted – we compare actual tax, not what was paid.',
      g5t: 'Extra tax in receipt year', g5d: 'How much more tax you pay this year just because the arrears were added (tax after arrears minus tax before arrears).',
      g6t: 'Amount belonging to the earlier year', g6d: 'The part of your arrears that was really due in that year. It should come from your arrears statement. All parts must add up to the total arrears.',
      g7t: 'Tax without / with that year\'s arrear', g7d: 'Your actual tax for that earlier year as it was, and what it would have been if that year\'s share of arrears had been included.',
      g8t: 'Extra tax if spread over earlier years', g8d: 'The total of extra tax each earlier year would have had if the arrears had been received on time.',
      g9t: 'Estimated relief', g9d: 'How much of this year\'s extra tax you may save because the arrears are spread back over earlier years. It is never below zero.',
      g10t: 'Net extra tax after relief', g10d: 'The extra tax you still pay in the receipt year after the relief is deducted.',
      nTitle: 'Good to know',
      n1: '<b>How it works:</b> Relief = (extra tax in receipt year) − (extra tax in earlier years). If the earlier years would have cost the same or more, relief is ₹0.',
      n2: 'Include rebate and cess in all tax figures. Exclude TDS.',
      n3: 'You must calculate your tax figures yourself (or with your DDO / tax adviser) using the correct slab rates for each year. This tool does the comparison for you.',
      n4: 'Your data is never uploaded. Everything is calculated in your browser.',
      disc: '<b>Important:</b> This is an indicative comparison only. It does not confirm your eligibility for relief or the legally correct amount. Confirm the applicable provision, year-wise computations and Form 39 / documentation with your DDO or a tax adviser before filing.',
      foot: 'Arrear Relief Calculator · For guidance only, not tax advice',
      pageTitle: 'Arrear Relief Calculator – Section 157 (Form 39)'
    },
    bn: {
      title: 'বকেয়া ত্রাণ (Arrear Relief) ক্যালকুলেটর',
      badge: 'আয়কর আইন, ২০২৫ · ধারা ১৫৭ · ফর্ম ৩৯',
      lead: 'আগের বছরগুলির বকেয়া বেতন (যেমন ডিএ বকেয়া) কি এক বছরে একসঙ্গে পেয়েছেন? আপনি কত কর-ত্রাণ দাবি করতে পারেন তা জেনে নিন। নিচের ঘরগুলি পূরণ করুন – উত্তর সঙ্গে সঙ্গে দেখা যাবে।',
      step1: 'যে বছরে বকেয়া <b>পেয়েছেন</b>, সেই বছরের কর লিখুন',
      step2: 'বকেয়া যে <b>আগের বছরগুলির</b>, সেগুলিতে ভাগ করুন',
      step3: 'সঙ্গে সঙ্গে <b>আনুমানিক ত্রাণ</b> দেখুন',
      c1title: 'যে বছরে বকেয়া পেয়েছেন',
      c1hint: 'যে অর্থবর্ষে বকেয়া আপনার অ্যাকাউন্টে জমা পড়েছে, সেটি ধরুন।',
      f1l: 'বকেয়া যোগ করার <em>আগে</em> প্রদেয় কর (₹)',
      f1s: 'রিবেট ও সেস সহ মোট কর। TDS বাদ দেবেন না।',
      f2l: 'বকেয়া যোগ করার <em>আগে</em> করযোগ্য আয় (₹)',
      f2s: 'ছাড়ের পরের আয় – <b>মোট বেতন নয়</b>।',
      f3l: 'মোট বকেয়া (₹)',
      f3s: 'এই বছরে জমা পড়া সব বকেয়ার যোগফল।',
      f4l: 'বকেয়া যোগ করার <em>পরে</em> প্রদেয় কর (₹)',
      f4s: 'বকেয়া ধরে আবার কর হিসাব করুন।',
      auto: 'বকেয়া যোগের পর করযোগ্য আয়:',
      c2title: 'যে আগের বছরগুলির বকেয়া',
      c2hint: 'প্রতিটি আগের বছরের জন্য সেই বছরের ভাগের বকেয়া লিখুন, এবং ওই ভাগ <em>ছাড়া</em> ও <em>সহ</em> সেই বছরের কর লিখুন।',
      addYear: '+ আরেকটি বছর যোগ করুন',
      reset: 'সব মুছে নতুন করুন',
      print: 'প্রিন্ট / PDF সংরক্ষণ',
      yPh: 'বছর, যেমন ২০২২-২৩',
      yAlloc: 'এই বছরের ভাগের বকেয়া (₹)',
      yWithout: 'এই বকেয়া <em>ছাড়া</em> কর (₹)',
      yWith: 'এই বকেয়া <em>সহ</em> কর (₹)',
      yExtra: 'এই বছরের অতিরিক্ত কর:',
      remove: 'সরান',
      rTitle: 'আপনার ফলাফল',
      rBig: 'ধারা ১৫৭ অনুযায়ী আনুমানিক ত্রাণ',
      r1: 'প্রাপ্তির বছরে অতিরিক্ত কর',
      r2: 'আগের বছরগুলিতে ভাগ করলে অতিরিক্ত কর',
      r3: 'আনুমানিক ত্রাণ',
      r4: 'ত্রাণের পর নিট অতিরিক্ত কর',
      msgEmpty: 'ফলাফল দেখতে তথ্য পূরণ করুন।',
      msgMismatch: 'ফলাফল দেখানো হয়েছে, কিন্তু ভাগ করা বকেয়া এখনও মোট বকেয়ার সঙ্গে মিলছে না – ধাপ ২ দেখে নিন।',
      msgNoExtra: 'বকেয়ার কারণে এই বছরে আপনার কর বাড়েনি, তাই ত্রাণের প্রশ্ন নেই।',
      msgRelief: 'বকেয়া আগের বছরগুলিতে ভাগ করলে কর কম হত। আপনি {0} ত্রাণ দাবি করতে পারেন।',
      msgNone: 'আগের বছরগুলিতে কর সমান বা বেশি হত, তাই কোনো ত্রাণ দেখা যাচ্ছে না।',
      chkEmpty: 'এখনও পর্যন্ত ভাগ করা হয়েছে: ₹ 0। সব আগের বছরের অঙ্কের যোগফল মোট বকেয়ার সমান হতে হবে।',
      chkOk: '✔ ঠিক আছে – ভাগ করা {0} মোট বকেয়ার সঙ্গে মিলেছে।',
      chkLess: 'ভাগ যাচাই করুন – {0} বকেয়া এখনও ভাগ করা হয়নি ({2}-এর মধ্যে ভাগ করা হয়েছে {1})।',
      chkMore: 'ভাগ যাচাই করুন – আপনি মোট বকেয়ার চেয়ে {0} বেশি ভাগ করেছেন।',
      gTitle: 'এই শব্দগুলির মানে কী?',
      g1t: 'প্রাপ্তির বছর (Receipt year)', g1d: 'যে অর্থবর্ষে বকেয়া সত্যিই আপনার অ্যাকাউন্টে এসেছে – যে বছরের জন্য পাওনা ছিল সেই বছর নয়।',
      g2t: 'বকেয়া (Arrears)', g2d: 'আগের বছরগুলিতে আপনার প্রাপ্য বেতন বা ডিএ, যা দেরিতে একসঙ্গে দেওয়া হয়েছে।',
      g3t: 'করযোগ্য আয় (মোট বেতন নয়)', g3d: 'স্ট্যান্ডার্ড ডিডাকশন ও অন্যান্য অনুমোদিত ছাড়ের পর আপনার আয় যা থাকে। যে অঙ্কের উপর কর হিসাব হয়, সেটিই দিন।',
      g4t: 'প্রদেয় কর (Tax payable)', g4d: 'প্রযোজ্য রিবেট ও সেস সহ মোট কর। আগে কাটা TDS <b>বাদ দেবেন না</b> – আমরা পরিশোধিত অঙ্ক নয়, আসল কর তুলনা করছি।',
      g5t: 'প্রাপ্তির বছরে অতিরিক্ত কর', g5d: 'শুধু বকেয়া যোগ হওয়ার কারণে এই বছরে আপনাকে কতটা বেশি কর দিতে হচ্ছে (বকেয়া সহ কর − বকেয়া ছাড়া কর)।',
      g6t: 'আগের বছরের ভাগের বকেয়া', g6d: 'বকেয়ার যে অংশ আসলে সেই বছরের প্রাপ্য ছিল। এটি আপনার বকেয়ার হিসাব-বিবরণী থেকে নিন। সব ভাগের যোগফল মোট বকেয়ার সমান হতে হবে।',
      g7t: 'সেই বছরের বকেয়া ছাড়া / সহ কর', g7d: 'ওই আগের বছরে আপনার আসল কর কত ছিল, এবং সেই বছরের ভাগের বকেয়া ধরলে কত হত।',
      g8t: 'আগের বছরগুলিতে ভাগ করলে অতিরিক্ত কর', g8d: 'বকেয়া সময়মতো পেলে প্রতিটি আগের বছরে যে অতিরিক্ত কর লাগত, তার মোট যোগফল।',
      g9t: 'আনুমানিক ত্রাণ (Relief)', g9d: 'বকেয়া আগের বছরগুলিতে ভাগ করার ফলে এই বছরের অতিরিক্ত করের কতটা আপনি বাঁচাতে পারেন। এটি কখনও শূন্যের কম হয় না।',
      g10t: 'ত্রাণের পর নিট অতিরিক্ত কর', g10d: 'ত্রাণ বাদ দেওয়ার পরেও প্রাপ্তির বছরে আপনাকে যে অতিরিক্ত কর দিতে হবে।',
      nTitle: 'জেনে রাখুন',
      n1: '<b>কীভাবে কাজ করে:</b> ত্রাণ = (প্রাপ্তির বছরে অতিরিক্ত কর) − (আগের বছরগুলিতে অতিরিক্ত কর)। আগের বছরগুলিতে কর সমান বা বেশি হলে ত্রাণ ₹0।',
      n2: 'সব করের অঙ্কে রিবেট ও সেস ধরুন। TDS ধরবেন না।',
      n3: 'প্রতিটি বছরের সঠিক স্ল্যাব অনুযায়ী করের অঙ্ক আপনাকে নিজে (বা DDO / ট্যাক্স উপদেষ্টার সাহায্যে) হিসাব করতে হবে। এই টুল শুধু তুলনা করে দেয়।',
      n4: 'আপনার তথ্য কোথাও আপলোড হয় না। সব হিসাব আপনার ব্রাউজারেই হয়।',
      disc: '<b>গুরুত্বপূর্ণ:</b> এটি শুধুমাত্র একটি আনুমানিক তুলনা। এটি আপনার ত্রাণের যোগ্যতা বা আইনত সঠিক অঙ্ক নিশ্চিত করে না। ফাইল করার আগে প্রযোজ্য ধারা, বছরভিত্তিক হিসাব এবং ফর্ম ৩৯ / প্রয়োজনীয় নথি আপনার DDO বা ট্যাক্স উপদেষ্টার সঙ্গে যাচাই করে নিন।',
      foot: 'বকেয়া ত্রাণ ক্যালকুলেটর · শুধুমাত্র দিকনির্দেশের জন্য, ট্যাক্স পরামর্শ নয়',
      pageTitle: 'বকেয়া ত্রাণ ক্যালকুলেটর – ধারা ১৫৭ (ফর্ম ৩৯)'
    }
  };

  var lang = 'bn';
  try { var saved = localStorage.getItem('arrearLang'); if (saved === 'en' || saved === 'bn') lang = saved; } catch (e) {}

  function t(key) {
    var s = I18N[lang][key];
    for (var i = 1; i < arguments.length; i++) s = s.replace('{' + (i - 1) + '}', arguments[i]);
    return s;
  }

  // ---------- Helpers ----------
  var $ = function (id) { return document.getElementById(id); };
  var fmt = function (n) {
    return '₹ ' + (Math.round(n * 100) / 100).toLocaleString('en-IN', { maximumFractionDigits: 2 });
  };
  var num = function (el) { var v = parseFloat(el.value); return isNaN(v) || v < 0 ? 0 : v; };

  var yearRows = $('yearRows');
  var MAX_YEARS = 15;

  function applyLang() {
    document.documentElement.lang = lang;
    document.title = t('pageTitle');
    Array.prototype.forEach.call(document.querySelectorAll('[data-i18n]'), function (el) {
      el.innerHTML = t(el.getAttribute('data-i18n'));
    });
    Array.prototype.forEach.call(document.querySelectorAll('[data-i18n-ph]'), function (el) {
      el.placeholder = t(el.getAttribute('data-i18n-ph'));
    });
    Array.prototype.forEach.call(document.querySelectorAll('[data-i18n-aria]'), function (el) {
      el.setAttribute('aria-label', t(el.getAttribute('data-i18n-aria')));
    });
    $('langBn').classList.toggle('on', lang === 'bn');
    $('langEn').classList.toggle('on', lang === 'en');
  }

  function addYear() {
    if (yearRows.children.length >= MAX_YEARS) return;
    var div = document.createElement('div');
    div.className = 'yrow';
    div.innerHTML =
      '<div class="yhead"><input type="text" class="y-label" data-i18n-ph="yPh" data-i18n-aria="yPh">' +
      '<button type="button" class="rm" data-i18n="remove"></button></div>' +
      '<div class="grid3">' +
      '<label class="field"><span class="lbl" data-i18n="yAlloc"></span><input type="number" class="y-alloc" min="0" step="any" inputmode="decimal" placeholder="0"></label>' +
      '<label class="field"><span class="lbl" data-i18n="yWithout"></span><input type="number" class="y-without" min="0" step="any" inputmode="decimal" placeholder="0"></label>' +
      '<label class="field"><span class="lbl" data-i18n="yWith"></span><input type="number" class="y-with" min="0" step="any" inputmode="decimal" placeholder="0"></label>' +
      '</div><div class="yfoot"><span data-i18n="yExtra"></span> <b class="y-extra">₹ 0</b></div>';
    div.querySelector('.rm').addEventListener('click', function () {
      div.remove();
      if (!yearRows.children.length) addYear();
      applyLang(); update();
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

    var c = $('allocCheck');
    if (d.arrears === 0 && r.allocTotal === 0) {
      c.className = 'check'; c.textContent = t('chkEmpty');
    } else if (Math.abs(r.allocDiff) < 0.005) {
      c.className = 'check ok'; c.textContent = t('chkOk', fmt(r.allocTotal));
    } else if (r.allocDiff > 0) {
      c.className = 'check warn'; c.textContent = t('chkLess', fmt(r.allocDiff), fmt(r.allocTotal), fmt(d.arrears));
    } else {
      c.className = 'check bad'; c.textContent = t('chkMore', fmt(-r.allocDiff));
    }

    $('rExtraReceipt').textContent = fmt(r.extraReceipt);
    $('rExtraEarlier').textContent = fmt(r.extraEarlier);
    $('rRelief').textContent = fmt(r.relief);
    $('rNet').textContent = fmt(r.net);
    $('reliefBig').textContent = fmt(r.relief);

    var m = $('resultMsg');
    var ready = d.arrears > 0 && d.taxAfter > 0;
    if (!ready) m.textContent = t('msgEmpty');
    else if (Math.abs(r.allocDiff) >= 0.005) m.textContent = t('msgMismatch');
    else if (r.extraReceipt === 0) m.textContent = t('msgNoExtra');
    else if (r.relief > 0) m.textContent = t('msgRelief', fmt(r.relief));
    else m.textContent = t('msgNone');
  }

  function resetAll() {
    ['taxBefore', 'incomeBefore', 'arrears', 'taxAfter'].forEach(function (id) { $(id).value = ''; });
    yearRows.innerHTML = '';
    for (var i = 0; i < 3; i++) addYear();
    applyLang(); update();
  }

  function setLang(l) {
    lang = l;
    try { localStorage.setItem('arrearLang', l); } catch (e) {}
    applyLang(); update();
  }

  document.addEventListener('input', update);
  $('addYear').addEventListener('click', function () { addYear(); applyLang(); update(); });
  $('resetBtn').addEventListener('click', resetAll);
  $('printBtn').addEventListener('click', function () { window.print(); });
  $('langBn').addEventListener('click', function () { setLang('bn'); });
  $('langEn').addEventListener('click', function () { setLang('en'); });

  resetAll();
})();
