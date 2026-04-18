/* =============================================================================
 * Admin utility rail — live month calendar + clock
 * -----------------------------------------------------------------------------
 *  Expects markup in admin.html: #adminRailCalMonth, #adminRailCalGrid,
 *  #adminRailCalPrev, #adminRailCalNext, #adminRailClock.
 *  Week starts Monday (labels M T W T F S S).
 * ========================================================================== */
(function () {
  'use strict';

  var monthEl = document.getElementById('adminRailCalMonth');
  var gridEl = document.getElementById('adminRailCalGrid');
  var prevBtn = document.getElementById('adminRailCalPrev');
  var nextBtn = document.getElementById('adminRailCalNext');
  var clockEl = document.getElementById('adminRailClock');

  if (!monthEl || !gridEl || !prevBtn || !nextBtn) return;

  /** Monday = 0 … Sunday = 6 */
  function mondayIndex(date) {
    return (date.getDay() + 6) % 7;
  }

  var view = new Date();
  view.setDate(1);

  function renderCal() {
    var y = view.getFullYear();
    var m = view.getMonth();
    monthEl.textContent = view.toLocaleDateString(undefined, { month: 'long', year: 'numeric' });

    var first = new Date(y, m, 1);
    var lastDay = new Date(y, m + 1, 0).getDate();
    var startPad = mondayIndex(first);
    var prevLast = new Date(y, m, 0).getDate();

    var today = new Date();
    var cells = [];
    var i;
    var d;

    for (i = 0; i < startPad; i++) {
      d = prevLast - startPad + i + 1;
      cells.push({ n: d, muted: true, today: false });
    }
    for (d = 1; d <= lastDay; d++) {
      var isToday = today.getFullYear() === y && today.getMonth() === m && today.getDate() === d;
      cells.push({ n: d, muted: false, today: isToday });
    }
    var nextNum = 1;
    while (cells.length < 42) {
      cells.push({ n: nextNum++, muted: true, today: false });
    }

    gridEl.innerHTML = cells.map(function (c) {
      var cls = 'admin-rail-day';
      if (c.muted) cls += ' is-muted';
      if (c.today) cls += ' is-today';
      return '<span class="' + cls + '">' + c.n + '</span>';
    }).join('');
  }

  prevBtn.addEventListener('click', function () {
    view.setMonth(view.getMonth() - 1);
    renderCal();
  });
  nextBtn.addEventListener('click', function () {
    view.setMonth(view.getMonth() + 1);
    renderCal();
  });

  function tickClock() {
    if (!clockEl) return;
    var n = new Date();
    clockEl.textContent = n.toLocaleString(undefined, {
      weekday: 'short',
      month: 'short',
      day: 'numeric',
      hour: 'numeric',
      minute: '2-digit',
      second: '2-digit',
      hour12: true,
    });
  }

  renderCal();
  tickClock();
  setInterval(tickClock, 1000);

  document.addEventListener('visibilitychange', function () {
    if (!document.hidden) {
      tickClock();
      renderCal();
    }
  });
})();
