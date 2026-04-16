/* =============================================================================
 * PRS (Planning, Research & Statistics) — Camp Monitoring UI / API helper
 * -----------------------------------------------------------------------------
 *  - Exposes window.Prs        : thin API wrappers around Api.call()
 *  - Exposes window.PrsAdmin   : admin dashboard renderers (slot into admin.html)
 *  - Public (staff-facing) helpers live on Prs.Staff and are used by
 *    prs-camp.html (QR landing page).
 *
 *  This module is fully self-contained — it does NOT read or mutate any HRM
 *  state. Staff identity in PRS is maintained purely by phone + name.
 * =============================================================================
 */
(function (global) {
  'use strict';

  // ---------------------------------------------------------------------------
  // 1. API WRAPPERS
  // ---------------------------------------------------------------------------
  //
  // Admin calls always need { key: adminKey }. Public (staff-facing) calls do
  // not. Every wrapper is small and predictable — no caching at this layer,
  // the backend handles CacheService semantics itself.

  function call(action, payload, opts) {
    if (typeof Api === 'undefined' || !Api || typeof Api.call !== 'function') {
      return Promise.reject(new Error('Api module not loaded.'));
    }
    return Api.call(action, payload || {}, opts || { skipCache: true });
  }

  // Auto-derive the full frontend base URL (origin + directory path), e.g.
  //   https://my-site.netlify.app            → https://my-site.netlify.app
  //   https://my-site.netlify.app/admin.html → https://my-site.netlify.app
  //   https://my-site.netlify.app/nysc/      → https://my-site.netlify.app/nysc
  // This is passed to prsGetCampQr so the QR URL is always a full, scannable
  // URL with protocol + host, even when no Script Property is set on the
  // Apps Script project.
  function currentFrontendBaseUrl() {
    try {
      const origin = window.location.origin || '';
      const dir = (window.location.pathname || '').replace(/\/[^/]*$/, '');
      return (origin + dir).replace(/\/+$/, '');
    } catch (e) { return ''; }
  }

  const PrsApi = {
    // ---- Events ----
    createEvent: (key, body) => call('prsCreateEvent', Object.assign({ key }, body)),
    updateEvent: (key, body) => call('prsUpdateEvent', Object.assign({ key }, body)),
    closeEvent:  (key, eventId) => call('prsCloseEvent', { key, eventId }),
    listEvents:  (key, status) => call('prsListEvents', { key, status: status || '' }),
    getEvent:    (key, eventId) => call('prsGetEvent', { key, eventId }),

    // ---- Camps ----
    addCamp:             (key, body) => call('prsAddCamp', Object.assign({ key }, body)),
    updateCamp:          (key, body) => call('prsUpdateCamp', Object.assign({ key }, body)),
    deleteCamp:          (key, campId) => call('prsDeleteCamp', { key, campId }),
    listCamps:           (key, eventId) => call('prsListCamps', { key, eventId: eventId || '' }),
    regenerateCampToken: (key, campId) => call('prsRegenerateCampToken', { key, campId }),
    getCampQr:           (key, campId) => call('prsGetCampQr', {
      key,
      campId,
      baseUrl: currentFrontendBaseUrl(),
    }),

    // ---- Assignments ----
    assignStaff:      (key, body) => call('prsAssignStaff', Object.assign({ key }, body)),
    updateAssignment: (key, body) => call('prsUpdateAssignment', Object.assign({ key }, body)),
    deleteAssignment: (key, assignmentId) => call('prsDeleteAssignment', { key, assignmentId }),
    listAssignments:  (key, eventId, campId) => call('prsListAssignments', {
      key, eventId, campId: campId || ''
    }),

    // ---- Dashboard ----
    getDashboardStats: (key) => call('prsGetDashboardStats', { key }),
    getAttendanceLogs: (key, body) => call('prsGetAttendanceLogs', Object.assign({ key }, body || {})),
    getDailySummary:   (key, body) => call('prsGetDailySummary',   Object.assign({ key }, body || {})),

    // ---- Admin bootstrap (SUPER_ADMIN) ----
    initializeSheets: (key) => call('prsInitializeSheets', { key }),
  };

  // Public staff-facing (no adminKey). Used by prs-camp.html.
  const PrsStaff = {
    validateQr:        (body) => call('prsValidateCampQr',    body),
    resolveAssignment: (body) => call('prsResolveAssignment', body),
    getStaffDashboard: (body) => call('prsGetStaffDashboard', body),
    signIn:            (body) => call('prsSignIn',            body),
    signOut:           (body) => call('prsSignOut',           body),
  };

  global.Prs = { Api: PrsApi, Staff: PrsStaff };

  // ---------------------------------------------------------------------------
  // 2. ADMIN UI — renders into the #prsWorkspace in admin.html
  // ---------------------------------------------------------------------------

  // Rendering context pushed by ui.js on every module open.
  let ctx = { adminKey: null, adminRole: null };

  function setContext(c) {
    ctx = Object.assign({ adminKey: null, adminRole: null }, c || {});
  }

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function fmtDate(v) {
    if (!v) return '';
    const d = v instanceof Date ? v : new Date(v);
    if (isNaN(d.getTime())) return esc(v);
    return d.toLocaleString();
  }

  function fmtDateOnly(v) {
    if (!v) return '';
    const d = v instanceof Date ? v : new Date(v);
    if (isNaN(d.getTime())) return esc(v);
    return d.toISOString().slice(0, 10);
  }

  function toast(msg, type) {
    if (typeof Swal !== 'undefined') {
      Swal.fire({
        icon: type || 'info', title: msg,
        toast: true, position: 'top', timer: 2500, showConfirmButton: false,
      });
    } else {
      console.log('[PRS]', msg);
    }
  }

  function err(e) {
    const msg = (e && e.message) || String(e);
    if (typeof Swal !== 'undefined') {
      Swal.fire({ icon: 'error', title: 'Error', text: msg });
    } else {
      alert(msg);
    }
  }

  // ---- Action bar (contextual tabs) ----
  const PRS_VIEWS = [
    { view: 'dashboard',   label: 'Overview',    icon: '📊' },
    { view: 'events',      label: 'Events',      icon: '🗓️' },
    { view: 'camps',       label: 'Camps',       icon: '🏕️' },
    { view: 'assignments', label: 'Assignments', icon: '👥' },
    { view: 'logs',        label: 'Attendance',  icon: '📋' },
    { view: 'summary',     label: 'Daily Summary', icon: '📈' },
  ];

  function renderActions(container, onViewSelect) {
    if (!container) return;
    container.innerHTML = PRS_VIEWS.map(v => `
      <button class="action-btn${v.view === 'dashboard' ? ' active' : ''}" data-view="${v.view}">
        <span class="action-btn-icon">${v.icon}</span>
        <span>${v.label}</span>
      </button>
    `).join('');

    container.querySelectorAll('.action-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const view = btn.getAttribute('data-view');
        container.querySelectorAll('.action-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        if (typeof onViewSelect === 'function') onViewSelect(view);
      });
    });
  }

  // ---- View dispatcher ----
  async function renderView(view, container) {
    if (!container) return;
    container.innerHTML = '<div class="info info-muted">Loading…</div>';
    try {
      if (view === 'dashboard')       return await renderDashboard(container);
      if (view === 'events')          return await renderEvents(container);
      if (view === 'camps')           return await renderCamps(container);
      if (view === 'assignments')     return await renderAssignments(container);
      if (view === 'logs')            return await renderLogs(container);
      if (view === 'summary')         return await renderSummary(container);
      container.innerHTML = '<div class="info info-muted">Unknown view.</div>';
    } catch (e) {
      container.innerHTML = `<div class="info info-error">Failed: ${esc(e.message || e)}</div>`;
    }
  }

  // ---------------------------------------------------------------------------
  // 2a. Dashboard
  // ---------------------------------------------------------------------------
  async function renderDashboard(container) {
    const res = await PrsApi.getDashboardStats(ctx.adminKey);
    const d = (res && res.data) || {};
    const events = d.events || [];
    const activeEvents = events.filter(e => String(e.status).toUpperCase() === 'ACTIVE').length;
    const today = (d.today || {});

    container.innerHTML = `
      <section class="card card-full-width">
        <h2>PRS Camp Monitoring — Overview</h2>
        <div class="dashboard-stats" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:1rem;margin-top:1rem;">
          <div class="stat-card"><div class="stat-value">${activeEvents}</div><div class="stat-label">Active Events</div></div>
          <div class="stat-card"><div class="stat-value">${events.length}</div><div class="stat-label">Total Events</div></div>
          <div class="stat-card"><div class="stat-value">${d.totalCamps || 0}</div><div class="stat-label">Camps</div></div>
          <div class="stat-card"><div class="stat-value">${d.totalAssignments || 0}</div><div class="stat-label">Assignments</div></div>
          <div class="stat-card"><div class="stat-value">${today.signIns || 0}</div><div class="stat-label">Sign-Ins Today</div></div>
          <div class="stat-card"><div class="stat-value">${today.signOuts || 0}</div><div class="stat-label">Sign-Outs Today</div></div>
        </div>
        <p class="info info-muted" style="margin-top:1rem;">Date: ${esc(today.date || '')}</p>
      </section>

      <section class="card card-full-width">
        <h3>Events</h3>
        <div class="table-wrapper">
          <table class="data-table">
            <thead><tr><th>Event</th><th>Status</th></tr></thead>
            <tbody>
              ${events.length === 0
                ? '<tr><td colspan="2" style="text-align:center;color:var(--text-muted);">No events yet.</td></tr>'
                : events.map(e => `<tr><td>${esc(e.eventName)}</td><td>${esc(e.status)}</td></tr>`).join('')}
            </tbody>
          </table>
        </div>
      </section>
    `;
  }

  // ---------------------------------------------------------------------------
  // 2b. Events
  // ---------------------------------------------------------------------------
  async function renderEvents(container) {
    const res = await PrsApi.listEvents(ctx.adminKey, '');
    const events = (res && res.data && res.data.events) || [];

    container.innerHTML = `
      <section class="card card-full-width">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:1rem;flex-wrap:wrap;gap:1rem;">
          <h2>Camp Events</h2>
          <button class="btn btn-primary" id="prsCreateEventBtn">➕ New Event</button>
        </div>
        <div class="table-wrapper">
          <table class="data-table">
            <thead>
              <tr><th>Name</th><th>Start</th><th>End</th><th>Status</th><th>Actions</th></tr>
            </thead>
            <tbody>
              ${events.length === 0
                ? '<tr><td colspan="5" style="text-align:center;color:var(--text-muted);">No events. Click "New Event" to create one.</td></tr>'
                : events.map(e => `
                  <tr data-event-id="${esc(e.eventId)}">
                    <td>${esc(e.eventName)}</td>
                    <td>${esc(fmtDateOnly(e.startDate))}</td>
                    <td>${esc(fmtDateOnly(e.endDate))}</td>
                    <td><span class="badge badge-${String(e.status).toLowerCase() === 'active' ? 'success' : 'muted'}">${esc(e.status)}</span></td>
                    <td>
                      <button class="btn btn-xs btn-secondary prs-edit-event">Edit</button>
                      ${String(e.status).toUpperCase() === 'ACTIVE'
                        ? '<button class="btn btn-xs btn-danger prs-close-event">Close</button>'
                        : ''}
                    </td>
                  </tr>
                `).join('')}
            </tbody>
          </table>
        </div>
      </section>
    `;

    document.getElementById('prsCreateEventBtn').addEventListener('click', () => showEventModal(null, container));

    container.querySelectorAll('.prs-edit-event').forEach(btn => {
      btn.addEventListener('click', () => {
        const tr = btn.closest('tr');
        const eventId = tr.getAttribute('data-event-id');
        const ev = events.find(x => String(x.eventId) === eventId);
        showEventModal(ev, container);
      });
    });
    container.querySelectorAll('.prs-close-event').forEach(btn => {
      btn.addEventListener('click', async () => {
        const tr = btn.closest('tr');
        const eventId = tr.getAttribute('data-event-id');
        const c = await Swal.fire({
          icon: 'warning', title: 'Close this event?',
          text: 'Closed events cannot accept new sign-ins. Camps remain read-only.',
          showCancelButton: true, confirmButtonText: 'Close Event',
        });
        if (!c.isConfirmed) return;
        try {
          await PrsApi.closeEvent(ctx.adminKey, eventId);
          toast('Event closed.', 'success');
          renderEvents(container);
        } catch (e) { err(e); }
      });
    });
  }

  async function showEventModal(existing, container) {
    const isEdit = !!existing;
    const { value: formValues } = await Swal.fire({
      title: isEdit ? 'Edit Event' : 'New Camp Event',
      html: `
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Event Name</label>
        <input id="prsEvName" class="swal2-input" placeholder="e.g. Batch C Stream 2 2025"
          value="${esc(existing ? existing.eventName : '')}" />
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Start Date</label>
        <input id="prsEvStart" type="date" class="swal2-input"
          value="${esc(existing ? fmtDateOnly(existing.startDate) : '')}" />
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">End Date</label>
        <input id="prsEvEnd" type="date" class="swal2-input"
          value="${esc(existing ? fmtDateOnly(existing.endDate) : '')}" />
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Notes (optional)</label>
        <textarea id="prsEvNotes" class="swal2-textarea" placeholder="Internal notes">${esc(existing && existing.notes || '')}</textarea>
      `,
      focusConfirm: false,
      showCancelButton: true,
      confirmButtonText: isEdit ? 'Save' : 'Create',
      preConfirm: () => {
        const eventName = document.getElementById('prsEvName').value.trim();
        const startDate = document.getElementById('prsEvStart').value;
        const endDate = document.getElementById('prsEvEnd').value;
        const notes = document.getElementById('prsEvNotes').value.trim();
        if (!eventName || !startDate || !endDate) {
          Swal.showValidationMessage('Name, start date and end date are required.');
          return false;
        }
        if (new Date(endDate) < new Date(startDate)) {
          Swal.showValidationMessage('End date must be on or after start date.');
          return false;
        }
        return { eventName, startDate, endDate, notes };
      },
    });
    if (!formValues) return;

    try {
      if (isEdit) {
        await PrsApi.updateEvent(ctx.adminKey, Object.assign({ eventId: existing.eventId }, formValues));
        toast('Event updated.', 'success');
      } else {
        await PrsApi.createEvent(ctx.adminKey, formValues);
        toast('Event created.', 'success');
      }
      renderEvents(container);
    } catch (e) { err(e); }
  }

  // ---------------------------------------------------------------------------
  // 2c. Camps  (select an event first, then list/add camps for that event)
  // ---------------------------------------------------------------------------
  async function renderCamps(container) {
    const evRes = await PrsApi.listEvents(ctx.adminKey, '');
    const events = (evRes && evRes.data && evRes.data.events) || [];
    if (events.length === 0) {
      container.innerHTML = `
        <section class="card">
          <h2>Camps</h2>
          <p class="info info-muted">Create an event first before adding camps.</p>
        </section>`;
      return;
    }

    container.innerHTML = `
      <section class="card card-full-width">
        <h2>Camp Locations</h2>
        <label style="display:block;margin:0.5rem 0 0.25rem 0;font-weight:600;">Event</label>
        <select id="prsCampEventSelect" class="swal2-select" style="width:100%;max-width:400px;padding:0.5rem;">
          ${events.map(e => `<option value="${esc(e.eventId)}">${esc(e.eventName)} (${esc(e.status)})</option>`).join('')}
        </select>
        <div style="margin:1rem 0;">
          <button class="btn btn-primary" id="prsAddCampBtn">➕ Add Camp</button>
          <button class="btn btn-secondary" id="prsRefreshCampsBtn">Refresh</button>
        </div>
        <div class="table-wrapper" id="prsCampsTableWrap"></div>
      </section>
    `;

    const eventSelect = document.getElementById('prsCampEventSelect');
    const tableWrap = document.getElementById('prsCampsTableWrap');

    async function reload() {
      const eventId = eventSelect.value;
      tableWrap.innerHTML = '<div class="info info-muted">Loading…</div>';
      try {
        const res = await PrsApi.listCamps(ctx.adminKey, eventId);
        const camps = (res && res.data && res.data.camps) || [];
        tableWrap.innerHTML = `
          <table class="data-table">
            <thead>
              <tr><th>Name</th><th>State</th><th>GPS</th><th>Radius (m)</th><th>QR Token</th><th>Actions</th></tr>
            </thead>
            <tbody>
              ${camps.length === 0
                ? '<tr><td colspan="6" style="text-align:center;color:var(--text-muted);">No camps for this event yet.</td></tr>'
                : camps.map(c => `
                  <tr data-camp-id="${esc(c.campId)}">
                    <td>${esc(c.campName)}</td>
                    <td>${esc(c.state)}</td>
                    <td>${Number(c.latitude).toFixed(5)}, ${Number(c.longitude).toFixed(5)}</td>
                    <td>${esc(c.radiusMeters)}</td>
                    <td><code style="font-size:0.75rem;">${esc(c.qrToken.slice(0, 10))}…</code></td>
                    <td>
                      <button class="btn btn-xs btn-secondary prs-camp-qr">QR</button>
                      <button class="btn btn-xs btn-secondary prs-camp-edit">Edit</button>
                      <button class="btn btn-xs btn-secondary prs-camp-regen">Regen Token</button>
                      <button class="btn btn-xs btn-danger prs-camp-del">Delete</button>
                    </td>
                  </tr>
                `).join('')}
            </tbody>
          </table>
        `;
        tableWrap.querySelectorAll('.prs-camp-qr').forEach(b => b.addEventListener('click', () => showCampQr(b.closest('tr').dataset.campId)));
        tableWrap.querySelectorAll('.prs-camp-edit').forEach(b => b.addEventListener('click', () => {
          const camp = camps.find(x => String(x.campId) === b.closest('tr').dataset.campId);
          showCampModal(eventId, camp, reload);
        }));
        tableWrap.querySelectorAll('.prs-camp-regen').forEach(b => b.addEventListener('click', async () => {
          const c = await Swal.fire({ icon: 'warning', title: 'Regenerate QR token?', text: 'The old QR code will stop working.', showCancelButton: true });
          if (!c.isConfirmed) return;
          try {
            await PrsApi.regenerateCampToken(ctx.adminKey, b.closest('tr').dataset.campId);
            toast('Token regenerated.', 'success');
            reload();
          } catch (e) { err(e); }
        }));
        tableWrap.querySelectorAll('.prs-camp-del').forEach(b => b.addEventListener('click', async () => {
          const c = await Swal.fire({ icon: 'warning', title: 'Delete this camp?', text: 'Only possible if there are no attendance logs yet.', showCancelButton: true });
          if (!c.isConfirmed) return;
          try {
            await PrsApi.deleteCamp(ctx.adminKey, b.closest('tr').dataset.campId);
            toast('Camp deleted.', 'success');
            reload();
          } catch (e) { err(e); }
        }));
      } catch (e) {
        tableWrap.innerHTML = `<div class="info info-error">${esc(e.message)}</div>`;
      }
    }

    eventSelect.addEventListener('change', reload);
    document.getElementById('prsRefreshCampsBtn').addEventListener('click', reload);
    document.getElementById('prsAddCampBtn').addEventListener('click', () => {
      showCampModal(eventSelect.value, null, reload);
    });

    reload();
  }

  async function showCampModal(eventId, existing, onDone) {
    const isEdit = !!existing;
    const { value } = await Swal.fire({
      title: isEdit ? 'Edit Camp' : 'Add Camp',
      html: `
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Camp Name</label>
        <input id="prsCampName" class="swal2-input" value="${esc(existing ? existing.campName : '')}" placeholder="e.g. Iseyin Camp" />
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">State</label>
        <input id="prsCampState" class="swal2-input" value="${esc(existing ? existing.state : '')}" placeholder="e.g. Oyo" />
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Latitude</label>
        <input id="prsCampLat" class="swal2-input" type="number" step="any" value="${esc(existing ? existing.latitude : '')}" placeholder="7.9606" />
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Longitude</label>
        <input id="prsCampLng" class="swal2-input" type="number" step="any" value="${esc(existing ? existing.longitude : '')}" placeholder="3.5939" />
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Radius (meters, optional)</label>
        <input id="prsCampRadius" class="swal2-input" type="number" min="50" value="${esc(existing ? existing.radiusMeters : 500)}" />
        <button id="prsUseMyLocation" type="button" class="btn btn-secondary" style="margin-top:0.75rem;">📍 Use my current GPS</button>
      `,
      focusConfirm: false,
      showCancelButton: true,
      confirmButtonText: isEdit ? 'Save' : 'Add',
      didOpen: () => {
        document.getElementById('prsUseMyLocation').addEventListener('click', () => {
          if (!navigator.geolocation) return toast('Geolocation not supported.', 'error');
          navigator.geolocation.getCurrentPosition(p => {
            document.getElementById('prsCampLat').value = p.coords.latitude;
            document.getElementById('prsCampLng').value = p.coords.longitude;
            toast('Location captured.', 'success');
          }, () => toast('Could not fetch location.', 'error'), { enableHighAccuracy: true, timeout: 10000 });
        });
      },
      preConfirm: () => {
        const campName = document.getElementById('prsCampName').value.trim();
        const state = document.getElementById('prsCampState').value.trim();
        const latitude = parseFloat(document.getElementById('prsCampLat').value);
        const longitude = parseFloat(document.getElementById('prsCampLng').value);
        const radiusMeters = parseInt(document.getElementById('prsCampRadius').value, 10) || 500;
        if (!campName || !state) { Swal.showValidationMessage('Camp name and state are required.'); return false; }
        if (isNaN(latitude) || isNaN(longitude)) { Swal.showValidationMessage('Valid latitude and longitude required.'); return false; }
        return { campName, state, latitude, longitude, radiusMeters };
      },
    });
    if (!value) return;
    try {
      if (isEdit) {
        await PrsApi.updateCamp(ctx.adminKey, Object.assign({ campId: existing.campId }, value));
        toast('Camp updated.', 'success');
      } else {
        await PrsApi.addCamp(ctx.adminKey, Object.assign({ eventId }, value));
        toast('Camp added.', 'success');
      }
      if (typeof onDone === 'function') onDone();
    } catch (e) { err(e); }
  }

  async function showCampQr(campId) {
    try {
      const res = await PrsApi.getCampQr(ctx.adminKey, campId);
      const d = (res && res.data) || {};
      const url = d.qrUrl || '';
      const qrImg = 'https://api.qrserver.com/v1/create-qr-code/?size=320x320&data=' + encodeURIComponent(url);
      await Swal.fire({
        title: esc(d.campName || 'Camp QR Code'),
        html: `
          <p style="margin-bottom:0.75rem;">Scan this code at the camp to sign in / out.</p>
          <img src="${qrImg}" alt="QR" style="max-width:100%;height:auto;" />
          <p style="margin-top:0.75rem;font-size:0.85rem;word-break:break-all;"><a href="${esc(url)}" target="_blank">${esc(url)}</a></p>
          <button id="prsCopyQr" class="btn btn-secondary" style="margin-top:0.75rem;">Copy link</button>
          <button id="prsDownloadQr" class="btn btn-secondary" style="margin-top:0.75rem;">Download PNG</button>
        `,
        showConfirmButton: true, confirmButtonText: 'Close',
        didOpen: () => {
          document.getElementById('prsCopyQr').addEventListener('click', () => {
            navigator.clipboard.writeText(url).then(() => toast('Copied.', 'success'));
          });
          document.getElementById('prsDownloadQr').addEventListener('click', () => {
            const a = document.createElement('a'); a.href = qrImg;
            a.download = 'camp-' + campId + '.png'; a.target = '_blank'; a.click();
          });
        },
      });
    } catch (e) { err(e); }
  }

  // ---------------------------------------------------------------------------
  // 2d. Assignments
  // ---------------------------------------------------------------------------
  async function renderAssignments(container) {
    const evRes = await PrsApi.listEvents(ctx.adminKey, '');
    const events = (evRes && evRes.data && evRes.data.events) || [];
    if (events.length === 0) {
      container.innerHTML = `<section class="card"><h2>Assignments</h2><p class="info info-muted">Create an event and add camps first.</p></section>`;
      return;
    }

    container.innerHTML = `
      <section class="card card-full-width">
        <h2>Staff Assignments</h2>
        <div style="display:flex;gap:1rem;flex-wrap:wrap;align-items:end;margin-bottom:1rem;">
          <div style="flex:1;min-width:200px;">
            <label style="display:block;font-weight:600;margin-bottom:0.25rem;">Event</label>
            <select id="prsAsgEvent" class="swal2-select" style="width:100%;padding:0.5rem;">
              ${events.map(e => `<option value="${esc(e.eventId)}">${esc(e.eventName)}</option>`).join('')}
            </select>
          </div>
          <div style="flex:1;min-width:200px;">
            <label style="display:block;font-weight:600;margin-bottom:0.25rem;">Camp (optional)</label>
            <select id="prsAsgCamp" class="swal2-select" style="width:100%;padding:0.5rem;"><option value="">All camps</option></select>
          </div>
          <button class="btn btn-primary" id="prsAsgAddBtn">➕ Assign Staff</button>
        </div>
        <div class="table-wrapper" id="prsAsgTableWrap"></div>
      </section>
    `;

    const evSel = document.getElementById('prsAsgEvent');
    const campSel = document.getElementById('prsAsgCamp');
    const tableWrap = document.getElementById('prsAsgTableWrap');

    async function refreshCampSel() {
      try {
        const r = await PrsApi.listCamps(ctx.adminKey, evSel.value);
        const camps = (r && r.data && r.data.camps) || [];
        campSel.innerHTML = '<option value="">All camps</option>' +
          camps.map(c => `<option value="${esc(c.campId)}">${esc(c.campName)} — ${esc(c.state)}</option>`).join('');
        campSel.dataset.camps = JSON.stringify(camps);
      } catch (e) { /* ignore */ }
    }

    async function reload() {
      tableWrap.innerHTML = '<div class="info info-muted">Loading…</div>';
      try {
        const r = await PrsApi.listAssignments(ctx.adminKey, evSel.value, campSel.value);
        const rows = (r && r.data && r.data.assignments) || [];

        // Prefer names joined server-side (a.eventName / a.campName). Fall back
        // to the local camps dataset only if the backend didn't send a name.
        const camps = JSON.parse(campSel.dataset.camps || '[]');
        const campLabel = a => {
          if (a.campName) {
            return a.campState ? `${a.campName} (${a.campState})` : a.campName;
          }
          const c = camps.find(x => String(x.campId) === String(a.campId));
          return c ? `${c.campName} (${c.state})` : '';
        };

        tableWrap.innerHTML = `
          <table class="data-table">
            <thead>
              <tr><th>Name</th><th>Phone</th><th>Department</th><th>Event</th><th>Camp</th><th>Required Days</th><th>Status</th><th>Actions</th></tr>
            </thead>
            <tbody>
              ${rows.length === 0
                ? '<tr><td colspan="8" style="text-align:center;color:var(--text-muted);">No assignments yet.</td></tr>'
                : rows.map(a => `
                  <tr data-asg-id="${esc(a.assignmentId)}">
                    <td>${esc(a.staffName)}</td>
                    <td>${esc(a.phone)}</td>
                    <td>${esc(a.department)}</td>
                    <td>${esc(a.eventName || '')}</td>
                    <td>${esc(campLabel(a))}</td>
                    <td>${esc(a.requiredDays)}</td>
                    <td><span class="badge badge-${String(a.status).toLowerCase() === 'active' ? 'success' : 'muted'}">${esc(a.status)}</span></td>
                    <td>
                      <button class="btn btn-xs btn-secondary prs-asg-edit">Edit</button>
                      <button class="btn btn-xs btn-danger prs-asg-del">Remove</button>
                    </td>
                  </tr>
                `).join('')}
            </tbody>
          </table>
        `;

        tableWrap.querySelectorAll('.prs-asg-edit').forEach(b => b.addEventListener('click', () => {
          const a = rows.find(x => String(x.assignmentId) === b.closest('tr').dataset.asgId);
          showAssignmentModal(evSel.value, JSON.parse(campSel.dataset.camps || '[]'), a, reload);
        }));
        tableWrap.querySelectorAll('.prs-asg-del').forEach(b => b.addEventListener('click', async () => {
          const c = await Swal.fire({ icon: 'warning', title: 'Remove assignment?', text: 'If the staff has sign-ins, it will be marked REVOKED instead of deleted.', showCancelButton: true });
          if (!c.isConfirmed) return;
          try {
            await PrsApi.deleteAssignment(ctx.adminKey, b.closest('tr').dataset.asgId);
            toast('Removed.', 'success');
            reload();
          } catch (e) { err(e); }
        }));
      } catch (e) {
        tableWrap.innerHTML = `<div class="info info-error">${esc(e.message)}</div>`;
      }
    }

    evSel.addEventListener('change', async () => { await refreshCampSel(); reload(); });
    campSel.addEventListener('change', reload);
    document.getElementById('prsAsgAddBtn').addEventListener('click', async () => {
      const camps = JSON.parse(campSel.dataset.camps || '[]');
      if (camps.length === 0) return toast('No camps in this event — add one first.', 'info');
      showAssignmentModal(evSel.value, camps, null, reload);
    });

    await refreshCampSel();
    await reload();
  }

  async function showAssignmentModal(eventId, camps, existing, onDone) {
    const isEdit = !!existing;
    const { value } = await Swal.fire({
      title: isEdit ? 'Edit Assignment' : 'Assign Staff to Camp',
      html: `
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Camp</label>
        <select id="prsAsgCampSel" class="swal2-select" style="width:100%;padding:0.5rem;">
          ${camps.map(c => `<option value="${esc(c.campId)}" ${existing && String(existing.campId) === String(c.campId) ? 'selected' : ''}>${esc(c.campName)} — ${esc(c.state)}</option>`).join('')}
        </select>
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Staff Name</label>
        <input id="prsAsgName" class="swal2-input" value="${esc(existing ? existing.staffName : '')}" placeholder="Full name" />
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Phone (Nigerian format)</label>
        <input id="prsAsgPhone" class="swal2-input" value="${esc(existing ? existing.phone : '')}" placeholder="0803…" ${isEdit ? '' : ''} />
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Department</label>
        <input id="prsAsgDept" class="swal2-input" value="${esc(existing ? existing.department : '')}" placeholder="e.g. Planning" />
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Required Days</label>
        <input id="prsAsgDays" class="swal2-input" type="number" min="1" value="${esc(existing ? existing.requiredDays : 21)}" />
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-size:0.85rem;color:#666;">External Reference (optional)</label>
        <input id="prsAsgRef" class="swal2-input" value="${esc(existing ? existing.staffSystemRecordId : '')}" placeholder="Leave blank to auto-generate" />
      `,
      focusConfirm: false,
      showCancelButton: true,
      confirmButtonText: isEdit ? 'Save' : 'Assign',
      preConfirm: () => {
        const campId = document.getElementById('prsAsgCampSel').value;
        const staffName = document.getElementById('prsAsgName').value.trim();
        const phone = document.getElementById('prsAsgPhone').value.trim();
        const department = document.getElementById('prsAsgDept').value.trim();
        const requiredDays = parseInt(document.getElementById('prsAsgDays').value, 10) || 0;
        const staffSystemRecordId = document.getElementById('prsAsgRef').value.trim();
        if (!campId || !staffName || !phone) {
          Swal.showValidationMessage('Camp, name and phone are required.'); return false;
        }
        return { campId, staffName, phone, department, requiredDays, staffSystemRecordId };
      },
    });
    if (!value) return;

    try {
      if (isEdit) {
        await PrsApi.updateAssignment(ctx.adminKey, Object.assign({ assignmentId: existing.assignmentId }, value));
        toast('Assignment updated.', 'success');
      } else {
        await PrsApi.assignStaff(ctx.adminKey, Object.assign({ eventId }, value));
        toast('Staff assigned.', 'success');
      }
      if (typeof onDone === 'function') onDone();
    } catch (e) { err(e); }
  }

  // ---------------------------------------------------------------------------
  // 2e. Attendance logs
  // ---------------------------------------------------------------------------
  async function renderLogs(container) {
    const evRes = await PrsApi.listEvents(ctx.adminKey, '');
    const events = (evRes && evRes.data && evRes.data.events) || [];

    container.innerHTML = `
      <section class="card card-full-width">
        <h2>Attendance Logs</h2>
        <div style="display:flex;gap:1rem;flex-wrap:wrap;align-items:end;margin-bottom:1rem;">
          <div style="flex:1;min-width:200px;">
            <label style="display:block;font-weight:600;margin-bottom:0.25rem;">Event (optional)</label>
            <select id="prsLogEvent" class="swal2-select" style="width:100%;padding:0.5rem;">
              <option value="">All events</option>
              ${events.map(e => `<option value="${esc(e.eventId)}">${esc(e.eventName)}</option>`).join('')}
            </select>
          </div>
          <div>
            <label style="display:block;font-weight:600;margin-bottom:0.25rem;">Date (optional)</label>
            <input type="date" id="prsLogDate" class="swal2-input" style="margin:0;" />
          </div>
          <button class="btn btn-secondary" id="prsLogRefresh">Refresh</button>
        </div>
        <div class="table-wrapper" id="prsLogWrap"></div>
      </section>
    `;

    const evSel = document.getElementById('prsLogEvent');
    const dtSel = document.getElementById('prsLogDate');
    const wrap = document.getElementById('prsLogWrap');

    async function reload() {
      wrap.innerHTML = '<div class="info info-muted">Loading…</div>';
      try {
        const r = await PrsApi.getAttendanceLogs(ctx.adminKey, {
          eventId: evSel.value || '',
          dateStr: dtSel.value || '',
          limit: 500,
        });
        const rows = (r && r.data && r.data.logs) || [];
        wrap.innerHTML = `
          <table class="data-table">
            <thead>
              <tr><th>When</th><th>Action</th><th>Staff</th><th>Phone</th><th>Camp</th><th>Event</th><th>Distance (m)</th></tr>
            </thead>
            <tbody>
              ${rows.length === 0
                ? '<tr><td colspan="7" style="text-align:center;color:var(--text-muted);">No logs.</td></tr>'
                : rows.map(l => {
                    const campLabel = l.campName
                      ? (l.campState ? `${l.campName} (${l.campState})` : l.campName)
                      : (l.campId || '');
                    const eventLabel = l.eventName || l.eventId || '';
                    return `
                      <tr>
                        <td>${esc(fmtDate(l.timestamp))}</td>
                        <td><span class="badge badge-${l.action === 'SIGN_IN' ? 'success' : 'muted'}">${esc(l.action)}</span></td>
                        <td>${esc(l.staffName || '')}</td>
                        <td>${esc(l.phone)}</td>
                        <td>${esc(campLabel)}</td>
                        <td>${esc(eventLabel)}</td>
                        <td>${l.distanceMeters != null ? Math.round(Number(l.distanceMeters)) : ''}</td>
                      </tr>
                    `;
                  }).join('')}
            </tbody>
          </table>
        `;
      } catch (e) {
        wrap.innerHTML = `<div class="info info-error">${esc(e.message)}</div>`;
      }
    }

    document.getElementById('prsLogRefresh').addEventListener('click', reload);
    evSel.addEventListener('change', reload);
    dtSel.addEventListener('change', reload);
    reload();
  }

  // ---------------------------------------------------------------------------
  // 2f. Daily Summary
  // ---------------------------------------------------------------------------
  async function renderSummary(container) {
    const evRes = await PrsApi.listEvents(ctx.adminKey, '');
    const events = (evRes && evRes.data && evRes.data.events) || [];
    if (events.length === 0) {
      container.innerHTML = `<section class="card"><h2>Daily Summary</h2><p class="info info-muted">No events yet.</p></section>`;
      return;
    }

    container.innerHTML = `
      <section class="card card-full-width">
        <h2>Daily Summary</h2>
        <div style="display:flex;gap:1rem;flex-wrap:wrap;align-items:end;margin-bottom:1rem;">
          <div style="flex:1;min-width:200px;">
            <label style="display:block;font-weight:600;margin-bottom:0.25rem;">Event</label>
            <select id="prsSumEvent" class="swal2-select" style="width:100%;padding:0.5rem;">
              ${events.map(e => `<option value="${esc(e.eventId)}">${esc(e.eventName)}</option>`).join('')}
            </select>
          </div>
          <div>
            <label style="display:block;font-weight:600;margin-bottom:0.25rem;">Date (defaults to today)</label>
            <input type="date" id="prsSumDate" class="swal2-input" style="margin:0;" />
          </div>
          <button class="btn btn-secondary" id="prsSumRefresh">Compute</button>
        </div>
        <div id="prsSumResult"></div>
      </section>
    `;

    const evSel = document.getElementById('prsSumEvent');
    const dtSel = document.getElementById('prsSumDate');
    const out = document.getElementById('prsSumResult');

    async function reload() {
      out.innerHTML = '<div class="info info-muted">Computing…</div>';
      try {
        const r = await PrsApi.getDailySummary(ctx.adminKey, {
          eventId: evSel.value,
          dateStr: dtSel.value || '',
        });
        const d = (r && r.data) || {};
        out.innerHTML = `
          <div class="dashboard-stats" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:1rem;">
            <div class="stat-card"><div class="stat-value">${d.totalAssigned || 0}</div><div class="stat-label">Assigned</div></div>
            <div class="stat-card"><div class="stat-value">${d.present || 0}</div><div class="stat-label">Present</div></div>
            <div class="stat-card"><div class="stat-value">${d.absent || 0}</div><div class="stat-label">Absent</div></div>
            <div class="stat-card"><div class="stat-value">${d.completedDays || 0}</div><div class="stat-label">Completed Days</div></div>
          </div>
          <p class="info info-muted" style="margin-top:1rem;">Date: ${esc(d.dateStr || '')}</p>
        `;
      } catch (e) {
        out.innerHTML = `<div class="info info-error">${esc(e.message)}</div>`;
      }
    }

    document.getElementById('prsSumRefresh').addEventListener('click', reload);
    evSel.addEventListener('change', reload);
    reload();
  }

  // ---------------------------------------------------------------------------
  // Export
  // ---------------------------------------------------------------------------
  global.PrsAdmin = {
    setContext,
    renderActions,
    renderView,
    VIEWS: PRS_VIEWS,
  };
})(window);
