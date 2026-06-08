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

    // ---- v2: Assignment Definitions ----
    createAssignmentDef:    (key, body) => call('prsCreateAssignmentDef',    Object.assign({ key }, body)),
    updateAssignmentDef:    (key, body) => call('prsUpdateAssignmentDef',    Object.assign({ key }, body)),
    closeAssignmentDef:     (key, assignmentDefId) => call('prsCloseAssignmentDef', { key, assignmentDefId }),
    listAssignmentDefs:     (key, eventId, campId) => call('prsListAssignmentDefs', {
      key, eventId, campId: campId || ''
    }),
    getAssignmentDef:       (key, assignmentDefId) => call('prsGetAssignmentDef', { key, assignmentDefId }),
    linkAssignmentToCamps:  (key, assignmentDefId, campIds) => call('prsLinkAssignmentToCamps', {
      key, assignmentDefId, campIds: campIds || [],
    }),
    unlinkAssignmentFromCamp: (key, assignmentDefId, campId) => call('prsUnlinkAssignmentFromCamp', {
      key, assignmentDefId, campId,
    }),
    listAssignmentCamps:    (key, assignmentDefId) => call('prsListAssignmentCamps', { key, assignmentDefId }),

    // ---- v2: Roster (per Assignment Definition × Camp) ----
    addStaffToAssignment:     (key, body) => call('prsAddStaffToAssignment',     Object.assign({ key }, body)),
    bulkAddStaffToAssignment: (key, body) => call('prsBulkAddStaffToAssignment', Object.assign({ key }, body)),
    listAssignmentRoster:     (key, assignmentDefId, campId) => call('prsListAssignmentRoster', {
      key, assignmentDefId, campId: campId || '',
    }),

    // ---- v2: Assignment Materials ----
    addAssignmentMaterial:        (key, body) => call('prsAddAssignmentMaterial',    Object.assign({ key }, body)),
    updateAssignmentMaterial:     (key, body) => call('prsUpdateAssignmentMaterial', Object.assign({ key }, body)),
    deleteAssignmentMaterial:     (key, materialId) => call('prsDeleteAssignmentMaterial', { key, materialId }),
    listAssignmentMaterials:      (key, assignmentDefId) => call('prsListAssignmentMaterials',  { key, assignmentDefId }),
    releaseAssignmentMaterialNow: (key, materialId) => call('prsReleaseAssignmentMaterialNow', { key, materialId }),
    getMaterialReleaseStats:      (key, assignmentDefId) => call('prsGetMaterialReleaseStats', { key, assignmentDefId }),

    // ---- Legacy single-roster surface (kept so older code keeps working) ----
    assignStaff:      (key, body) => call('prsAssignStaff', Object.assign({ key }, body)),
    updateAssignment: (key, body) => call('prsUpdateAssignment', Object.assign({ key }, body)),
    deleteAssignment: (key, assignmentId) => call('prsDeleteAssignment', { key, assignmentId }),
    listAssignments:  (key, eventId, campId) => call('prsListAssignments', {
      key, eventId, campId: campId || ''
    }),
    downloadAssignmentsExcel: (key, eventId) => call('prsDownloadAssignmentsExcel', {
      key,
      eventId,
    }),

    // ---- Dashboard ----
    getDashboardStats: (key) => call('prsGetDashboardStats', { key }),
    getAttendanceLogs: (key, body) => call('prsGetAttendanceLogs', Object.assign({ key }, body || {})),
    getDailySummary:   (key, body) => call('prsGetDailySummary',   Object.assign({ key }, body || {})),

    // ---- Admin bootstrap (SUPER_ADMIN) ----
    initializeSheets:    (key) => call('prsInitializeSheets', { key }),
    migrateLegacyRoster: (key) => call('prsMigrateLegacyRoster', { key }),

    // ---- Reuse / copy from previous events ----
    copyCampsFromEvent:      (key, body) => call('prsCopyCampsFromEvent',      Object.assign({ key }, body)),
    duplicateAssignmentDef:  (key, body) => call('prsDuplicateAssignmentDef',  Object.assign({ key }, body)),
    copyEventSetup:          (key, body) => call('prsCopyEventSetup',          Object.assign({ key }, body)),
    copyRosterFromAssignment:(key, body) => call('prsCopyRosterFromAssignment', Object.assign({ key }, body)),

    // ---- SUPER_ADMIN exclusive: mark attendance on behalf of a staff ----
    adminMarkAttendance: (key, body) => call('prsAdminMarkAttendance', Object.assign({ key }, body)),
  };

  // Public staff-facing (no adminKey). Used by prs-camp.html.
  const PrsStaff = {
    validateQr:           (body) => call('prsValidateCampQr',       body),
    resolveAssignment:    (body) => call('prsResolveAssignment',    body),
    listStaffAssignments: (body) => call('prsListStaffAssignments', body),
    getStaffDashboard:    (body) => call('prsGetStaffDashboard',    body),
    signIn:               (body) => call('prsSignIn',               body),
    signOut:              (body) => call('prsSignOut',              body),
    listStaffMaterials:   (body) => call('prsListStaffMaterials',   body),
    logMaterialDownload:  (body) => call('prsLogMaterialDownload',  body),
  };

  global.Prs = { Api: PrsApi, Staff: PrsStaff, formatDateTime: fmtDate };

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

  /** 12-hour clock for all PRS date+time display (admin + staff UI). */
  const PRS_DATETIME_OPTS = {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit',
    second: '2-digit',
    hour12: true,
  };

  function fmtDate(v) {
    if (!v) return '';
    const d = v instanceof Date ? v : new Date(v);
    if (isNaN(d.getTime())) return esc(v);
    return d.toLocaleString(undefined, PRS_DATETIME_OPTS);
  }

  function fmtDateOnly(v) {
    if (!v) return '';
    const d = v instanceof Date ? v : new Date(v);
    if (isNaN(d.getTime())) return esc(v);
    return d.toISOString().slice(0, 10);
  }

  async function loadAllEvents() {
    const res = await PrsApi.listEvents(ctx.adminKey, '');
    return (res && res.data && res.data.events) || [];
  }

  async function pickSourceEvent(events, excludeEventId, title) {
    const choices = (events || []).filter(e => String(e.eventId) !== String(excludeEventId || ''));
    if (!choices.length) {
      toast('No other events available to copy from.', 'info');
      return null;
    }
    const options = {};
    choices.forEach(e => { options[e.eventId] = `${e.eventName} (${e.status})`; });
    const result = await Swal.fire({
      title: title || 'Select source event',
      input: 'select',
      inputOptions: options,
      inputPlaceholder: 'Choose an event…',
      showCancelButton: true,
      inputValidator: (v) => (!v ? 'Please select an event.' : undefined),
    });
    return result.isConfirmed && result.value ? result.value : null;
  }

  async function pickSourceActivity(sourceEventId, title) {
    const r = await PrsApi.listAssignmentDefs(ctx.adminKey, sourceEventId, '');
    const defs = (r && r.data && r.data.assignmentDefs) || [];
    const active = defs.filter(d => String(d.status).toUpperCase() === 'ACTIVE');
    if (!active.length) {
      toast('That event has no activities to copy from.', 'info');
      return null;
    }
    const options = {};
    active.forEach(d => {
      options[d.assignmentDefId] = `${d.title} (${d.rosterCount || 0} staff, ${d.materialsCount || 0} materials)`;
    });
    const result = await Swal.fire({
      title: title || 'Select source activity',
      input: 'select',
      inputOptions: options,
      showCancelButton: true,
      inputValidator: (v) => (!v ? 'Please select an activity.' : undefined),
    });
    return result.isConfirmed && result.value ? result.value : null;
  }

  async function showCopyOptionsModal(defaults) {
    const d = Object.assign({ camps: true, activities: true, rosters: true, materials: true }, defaults || {});
    const result = await Swal.fire({
      title: 'What to reuse?',
      html: `
        <p style="text-align:left;font-size:0.9rem;color:#555;margin-bottom:0.75rem;">
          Camps are matched by <strong>name + state</strong> when possible (existing camps are reused, not duplicated).
          New QR codes are generated only for camps that are actually copied.
        </p>
        <label style="display:flex;align-items:center;gap:0.5rem;margin:0.45rem 0;text-align:left;">
          <input type="checkbox" id="prsCopyCamps" ${d.camps ? 'checked' : ''} /> Camp locations &amp; GPS addresses
        </label>
        <label style="display:flex;align-items:center;gap:0.5rem;margin:0.45rem 0;text-align:left;">
          <input type="checkbox" id="prsCopyActivities" ${d.activities ? 'checked' : ''} /> Activities (assignments)
        </label>
        <label style="display:flex;align-items:center;gap:0.5rem;margin:0.45rem 0;text-align:left;">
          <input type="checkbox" id="prsCopyRosters" ${d.rosters ? 'checked' : ''} /> Staff rosters
        </label>
        <label style="display:flex;align-items:center;gap:0.5rem;margin:0.45rem 0;text-align:left;">
          <input type="checkbox" id="prsCopyMaterials" ${d.materials ? 'checked' : ''} /> Materials &amp; download links
        </label>
      `,
      focusConfirm: false,
      showCancelButton: true,
      confirmButtonText: 'Copy selected',
      preConfirm: () => {
        const opts = {
          camps: document.getElementById('prsCopyCamps').checked,
          activities: document.getElementById('prsCopyActivities').checked,
          rosters: document.getElementById('prsCopyRosters').checked,
          materials: document.getElementById('prsCopyMaterials').checked,
        };
        if (!opts.camps && !opts.activities && !opts.rosters && !opts.materials) {
          Swal.showValidationMessage('Tick at least one item to copy.');
          return false;
        }
        return opts;
      },
    });
    return result.isConfirmed && result.value ? result.value : null;
  }

  // ---------------------------------------------------------------------------
  // Reusable paginated table view (search + S/N + 10-per-page pagination).
  //
  // Usage:
  //   const view = paginatedTableView({
  //     wrap:             <HTMLElement container>,
  //     rows:             <Array of row objects>,
  //     columns:          [{ label, render(row, serialNumber) -> html }],
  //     searchFields:     (row) -> string used for case-insensitive substring match,
  //     searchPlaceholder:'Search…',
  //     emptyText:        'No entries.',
  //     rowAttrs:         (row) -> extra attrs for <tr> (optional),
  //     onAfterRender:    (tbody, pageRows) -> wire up row action buttons (optional),
  //   });
  //   view.setRows(newRows);   // swap dataset, resets to page 1
  //   view.refresh();          // re-render current state
  // ---------------------------------------------------------------------------
  const PRS_PAGE_SIZE = 10;

  function paginatedTableView(opts) {
    const wrap = opts.wrap;
    if (!wrap) throw new Error('paginatedTableView: wrap element is required.');

    let currentPage = 1;
    let searchQuery = '';
    let rows = Array.isArray(opts.rows) ? opts.rows.slice() : [];

    wrap.innerHTML = `
      <div class="prs-table-toolbar" style="display:flex;justify-content:space-between;align-items:center;gap:0.75rem;flex-wrap:wrap;margin-bottom:0.75rem;">
        <input type="search" class="prs-table-search"
          placeholder="${esc(opts.searchPlaceholder || 'Search…')}"
          style="flex:1;min-width:220px;max-width:400px;padding:0.5rem 0.75rem;border:1px solid #d1d5db;border-radius:6px;font-size:0.9rem;" />
        <div class="prs-table-meta" style="font-size:0.85rem;color:var(--text-muted,#6b7280);"></div>
      </div>
      <div class="prs-table-body table-wrapper"></div>
      <div class="prs-table-pagination" style="display:flex;justify-content:center;align-items:center;flex-wrap:wrap;gap:0.4rem;margin-top:1rem;"></div>
    `;

    const searchInput = wrap.querySelector('.prs-table-search');
    const metaEl      = wrap.querySelector('.prs-table-meta');
    const bodyEl      = wrap.querySelector('.prs-table-body');
    const pagEl       = wrap.querySelector('.prs-table-pagination');

    function matches(row) {
      if (!searchQuery) return true;
      const val = (typeof opts.searchFields === 'function' ? opts.searchFields(row) : '') || '';
      return String(val).toLowerCase().indexOf(searchQuery) !== -1;
    }

    function render() {
      const filtered  = rows.filter(matches);
      const total     = filtered.length;
      const totalPages = Math.max(1, Math.ceil(total / PRS_PAGE_SIZE));
      if (currentPage > totalPages) currentPage = totalPages;
      if (currentPage < 1) currentPage = 1;
      const startIdx  = (currentPage - 1) * PRS_PAGE_SIZE;
      const pageRows  = filtered.slice(startIdx, startIdx + PRS_PAGE_SIZE);

      const colCount  = (opts.columns || []).length + 1; // +1 for S/N
      const theadHtml = '<thead><tr><th style="width:60px;">S/N</th>'
        + (opts.columns || []).map(c => `<th>${c.label}</th>`).join('')
        + '</tr></thead>';

      let tbodyHtml;
      if (pageRows.length === 0) {
        const msg = searchQuery
          ? 'No entries match your search.'
          : (opts.emptyText || 'No entries.');
        tbodyHtml = `<tr><td colspan="${colCount}" style="text-align:center;color:var(--text-muted);padding:1rem;">${esc(msg)}</td></tr>`;
      } else {
        tbodyHtml = pageRows.map((row, i) => {
          const sn = startIdx + i + 1;
          const attrs = (typeof opts.rowAttrs === 'function') ? (opts.rowAttrs(row) || '') : '';
          const cells = (opts.columns || []).map(c => `<td>${c.render(row, sn)}</td>`).join('');
          return `<tr ${attrs}><td>${sn}</td>${cells}</tr>`;
        }).join('');
      }
      bodyEl.innerHTML = `<table class="data-table">${theadHtml}<tbody>${tbodyHtml}</tbody></table>`;

      // Summary meta
      if (total === 0) {
        metaEl.textContent = searchQuery ? 'No matches' : '0 entries';
      } else {
        const from = startIdx + 1;
        const to   = startIdx + pageRows.length;
        metaEl.textContent = `Showing ${from}–${to} of ${total}${searchQuery ? ' (filtered)' : ''}`;
      }

      // Pagination controls
      pagEl.innerHTML = '';
      if (total > 0 && totalPages > 1) {
        const btn = (label, page, disabled, active) =>
          `<button type="button" class="btn btn-sm ${active ? 'btn-primary' : 'btn-secondary'}"`
          + ` ${disabled ? 'disabled' : ''} data-prs-page="${page}"`
          + `${active ? ' style="background:var(--nysc-green,#059669);color:#fff;"' : ''}>${label}</button>`;

        let html = '';
        html += btn('« Prev', currentPage - 1, currentPage === 1, false);

        const windowStart = Math.max(1, currentPage - 2);
        const windowEnd   = Math.min(totalPages, currentPage + 2);
        if (windowStart > 1) {
          html += btn('1', 1, false, false);
          if (windowStart > 2) html += '<span style="padding:0 0.25rem;color:var(--text-muted);">…</span>';
        }
        for (let p = windowStart; p <= windowEnd; p++) {
          html += btn(String(p), p, false, p === currentPage);
        }
        if (windowEnd < totalPages) {
          if (windowEnd < totalPages - 1) html += '<span style="padding:0 0.25rem;color:var(--text-muted);">…</span>';
          html += btn(String(totalPages), totalPages, false, false);
        }
        html += btn('Next »', currentPage + 1, currentPage === totalPages, false);
        html += `<span style="margin-left:0.5rem;color:var(--text-muted,#6b7280);font-size:0.85rem;">Page ${currentPage} of ${totalPages}</span>`;

        pagEl.innerHTML = html;
        pagEl.querySelectorAll('button[data-prs-page]').forEach(b => {
          if (b.disabled) return;
          b.addEventListener('click', () => {
            const p = parseInt(b.getAttribute('data-prs-page'), 10);
            if (!isNaN(p) && p !== currentPage) { currentPage = p; render(); }
          });
        });
      }

      // Hook for wiring per-row action buttons
      const tbody = bodyEl.querySelector('tbody');
      if (tbody && typeof opts.onAfterRender === 'function') {
        opts.onAfterRender(tbody, pageRows);
      }
    }

    // Debounced search
    let searchTimer = null;
    searchInput.addEventListener('input', () => {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(() => {
        const q = searchInput.value.trim().toLowerCase();
        if (q === searchQuery) return;
        searchQuery = q;
        currentPage = 1;
        render();
      }, 150);
    });

    render();

    return {
      setRows(newRows) {
        rows = Array.isArray(newRows) ? newRows.slice() : [];
        currentPage = 1;
        render();
      },
      refresh: render,
    };
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

  function isSuperAdmin_() {
    return String(ctx.adminRole || '').trim() === 'SUPER_ADMIN';
  }

  /** True when the backend reports missing PRS v2 sheets. */
  function isPrsSheetSetupError(e) {
    const msg = String((e && e.message) || '');
    const reason = String((e && e.reason) || (e && e.raw && e.raw.reason) || '');
    return msg.indexOf('PRS_SHEET_MISSING') !== -1
      || reason.indexOf('PRS_SHEET_MISSING') !== -1
      || msg.indexOf('run prsInitializeSheets') !== -1;
  }

  /**
   * Render a setup card with a one-click sheet initializer (SUPER_ADMIN only).
   * @param {HTMLElement} container
   * @param {string} errorMsg
   * @param {function} [onSuccess] called after init succeeds
   */
  function renderPrsSetupPrompt(container, errorMsg, onSuccess) {
    const canInit = isSuperAdmin_();
    container.innerHTML = `
      <section class="card card-full-width">
        <h2>PRS setup required</h2>
        <p class="info info-error">${esc(errorMsg)}</p>
        <p class="info info-muted" style="margin-top:0.75rem;">
          PRS v2 adds sheets such as <strong>PRS_Assignments</strong>, <strong>PRS_AssignmentCamps</strong>,
          <strong>PRS_AssignmentMaterials</strong>, and <strong>PRS_MaterialReleases</strong>.
          Initializing is <em>safe and idempotent</em> — your existing events, camps, and roster rows are kept.
        </p>
        ${canInit ? `
          <div style="display:flex;gap:0.5rem;flex-wrap:wrap;margin-top:1rem;">
            <button type="button" class="btn btn-primary" id="prsRunInitSheetsBtn">Initialize PRS sheets</button>
            <button type="button" class="btn btn-secondary" id="prsRunMigrateBtn" title="After init, back-fill AssignmentDefID on legacy roster rows">Migrate legacy roster</button>
          </div>
        ` : `
          <p class="info info-muted" style="margin-top:1rem;">
            A <strong>SUPER_ADMIN</strong> must run <em>Initialize PRS sheets</em> once before Activities can load.
          </p>
        `}
      </section>
    `;

    if (!canInit) return;

    document.getElementById('prsRunInitSheetsBtn').addEventListener('click', async () => {
      const btn = document.getElementById('prsRunInitSheetsBtn');
      btn.disabled = true;
      btn.textContent = 'Initializing…';
      try {
        const res = await PrsApi.initializeSheets(ctx.adminKey);
        const d = (res && res.data) || {};
        const created = (d.created || []).join(', ') || 'none';
        toast('PRS sheets ready. Created: ' + created, 'success');
        if (typeof onSuccess === 'function') onSuccess();
      } catch (e) {
        err(e);
      } finally {
        btn.disabled = false;
        btn.textContent = 'Initialize PRS sheets';
      }
    });

    document.getElementById('prsRunMigrateBtn').addEventListener('click', async () => {
      const btn = document.getElementById('prsRunMigrateBtn');
      btn.disabled = true;
      btn.textContent = 'Migrating…';
      try {
        await PrsApi.initializeSheets(ctx.adminKey);
        const res = await PrsApi.migrateLegacyRoster(ctx.adminKey);
        toast((res && res.message) || 'Legacy roster migrated.', 'success');
        if (typeof onSuccess === 'function') onSuccess();
      } catch (e) {
        err(e);
      } finally {
        btn.disabled = false;
        btn.textContent = 'Migrate legacy roster';
      }
    });
  }

  function downloadBase64File(base64, fileName, mimeType) {
    if (!base64) throw new Error('Empty file payload.');
    const byteChars = atob(base64);
    const byteArrays = [];
    const sliceSize = 1024;
    for (let offset = 0; offset < byteChars.length; offset += sliceSize) {
      const slice = byteChars.slice(offset, offset + sliceSize);
      const bytes = new Array(slice.length);
      for (let i = 0; i < slice.length; i++) bytes[i] = slice.charCodeAt(i);
      byteArrays.push(new Uint8Array(bytes));
    }
    const blob = new Blob(byteArrays, { type: mimeType || 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName || 'download';
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  // ---- Action bar (contextual tabs) ----
  const PRS_VIEWS = [
    { view: 'dashboard',   label: 'Overview',      icon: '📊' },
    { view: 'events',      label: 'Events',        icon: '🗓️' },
    { view: 'camps',       label: 'Camps',         icon: '🏕️' },
    { view: 'assignments', label: 'Activities',    icon: '🧩' },
    { view: 'logs',        label: 'Attendance',    icon: '📋' },
    { view: 'summary',     label: 'Daily Summary', icon: '📈' },
  ];

  // Whitelist of attendance modes + release rules surfaced in admin pickers.
  // Must mirror the backend's PRS_CONFIG values exactly.
  const PRS_ATTENDANCE_MODES = [
    { value: 'STANDARD',       label: 'Standard — daily sign-in + sign-out on the final day' },
    { value: 'MATERIAL_GATED', label: 'Material-gated — same as Standard, but materials are unlocked by sign-in' },
    { value: 'SIGN_IN_ONLY',   label: 'Sign-in only — no sign-out, one sign-in per required day' },
  ];
  const PRS_RELEASE_RULES = [
    { value: 'IMMEDIATE',          label: 'Immediate — released the moment staff opens the assignment' },
    { value: 'ON_SIGN_IN',         label: 'On sign-in — released after the first qualifying sign-in' },
    { value: 'ON_SIGN_IN_DAY_N',   label: 'On sign-in day N — released on/after the Nth distinct sign-in day' },
    { value: 'ON_DATE',            label: 'On date — released on or after a fixed calendar date (still needs same-day sign-in)' },
    { value: 'MANUAL',             label: 'Manual — released only when an admin clicks "Release now"' },
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
    if (typeof ScrollTop !== 'undefined') ScrollTop.toTop();
    container.innerHTML = '<div class="info info-muted">Loading…</div>';
    try {
      if (view === 'dashboard')    return await renderDashboard(container);
      if (view === 'events')       return await renderEvents(container);
      if (view === 'camps')        return await renderCamps(container);
      if (view === 'assignments')  return await renderAssignments(container);
      if (view === 'logs')         return await renderLogs(container);
      if (view === 'summary')      return await renderSummary(container);
      container.innerHTML = '<div class="info info-muted">Unknown view.</div>';
    } catch (e) {
      if (isPrsSheetSetupError(e)) {
        renderPrsSetupPrompt(container, e.message || String(e), () => renderView(view, container));
        return;
      }
      container.innerHTML = `<div class="info info-error">Failed: ${esc(e.message || e)}</div>`;
    } finally {
      if (typeof ScrollTop !== 'undefined') ScrollTop.afterRender();
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
        <div id="prsDashEventsWrap"></div>
      </section>
    `;

    paginatedTableView({
      wrap: document.getElementById('prsDashEventsWrap'),
      rows: events,
      searchPlaceholder: 'Search by event name or status…',
      emptyText: 'No events yet.',
      searchFields: e => [e.eventName, e.status].filter(Boolean).join(' '),
      columns: [
        { label: 'Event',  render: e => esc(e.eventName) },
        { label: 'Status', render: e => esc(e.status) },
      ],
    });
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
        <div id="prsEventsWrap"></div>
      </section>
    `;

    document.getElementById('prsCreateEventBtn').addEventListener('click', () => showEventModal(null, container));

    paginatedTableView({
      wrap: document.getElementById('prsEventsWrap'),
      rows: events,
      searchPlaceholder: 'Search by event name or status…',
      emptyText: 'No events. Click "New Event" to create one.',
      searchFields: e => [e.eventName, e.status, fmtDateOnly(e.startDate), fmtDateOnly(e.endDate)].filter(Boolean).join(' '),
      rowAttrs: e => `data-event-id="${esc(e.eventId)}"`,
      columns: [
        { label: 'Name',   render: e => esc(e.eventName) },
        { label: 'Start',  render: e => esc(fmtDateOnly(e.startDate)) },
        { label: 'End',    render: e => esc(fmtDateOnly(e.endDate)) },
        { label: 'Status', render: e => `<span class="badge badge-${String(e.status).toLowerCase() === 'active' ? 'success' : 'muted'}">${esc(e.status)}</span>` },
        { label: 'Actions', render: e => `
            <button class="btn btn-xs btn-secondary prs-edit-event">Edit</button>
            ${String(e.status).toUpperCase() === 'ACTIVE'
              ? '<button class="btn btn-xs btn-danger prs-close-event">Close</button>'
              : ''}
          ` },
      ],
      onAfterRender: (tbody) => {
        tbody.querySelectorAll('.prs-edit-event').forEach(btn => {
          btn.addEventListener('click', () => {
            const eventId = btn.closest('tr').getAttribute('data-event-id');
            const ev = events.find(x => String(x.eventId) === eventId);
            showEventModal(ev, container);
          });
        });
        tbody.querySelectorAll('.prs-close-event').forEach(btn => {
          btn.addEventListener('click', async () => {
            const eventId = btn.closest('tr').getAttribute('data-event-id');
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
      },
    });
  }

  async function showEventModal(existing, container) {
    const isEdit = !!existing;
    const allEvents = isEdit ? [] : await loadAllEvents();
    const copyEventOptions = allEvents.length
      ? otherEvents.map(e => `<option value="${esc(e.eventId)}">${esc(e.eventName)} (${esc(e.status)})</option>`).join('')
      : '';

    const { value: formValues } = await Swal.fire({
      title: isEdit ? 'Edit Event' : 'New Camp Event',
      width: 640,
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
        ${!isEdit && copyEventOptions ? `
          <hr style="margin:1rem 0;border:none;border-top:1px solid #e5e7eb;" />
          <label style="display:flex;align-items:center;gap:0.5rem;text-align:left;font-weight:600;margin-bottom:0.5rem;">
            <input type="checkbox" id="prsEvCopyEnable" /> Reuse setup from a previous event
          </label>
          <div id="prsEvCopyPanel" style="display:none;text-align:left;">
            <label style="display:block;margin:0.35rem 0 0.25rem 0;">Copy from</label>
            <select id="prsEvCopyFrom" class="swal2-select" style="width:100%;padding:0.5rem;">
              <option value="">— Select event —</option>
              ${copyEventOptions}
            </select>
            <p style="font-size:0.85rem;color:#666;margin:0.5rem 0;">You can choose camps, activities, rosters, and materials on the next step.</p>
          </div>
        ` : ''}
      `,
      focusConfirm: false,
      showCancelButton: true,
      confirmButtonText: isEdit ? 'Save' : 'Create',
      didOpen: () => {
        const en = document.getElementById('prsEvCopyEnable');
        const panel = document.getElementById('prsEvCopyPanel');
        if (en && panel) {
          en.addEventListener('change', () => {
            panel.style.display = en.checked ? 'block' : 'none';
          });
        }
      },
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
        const out = { eventName, startDate, endDate, notes };
        const copyEn = document.getElementById('prsEvCopyEnable');
        const copyFrom = document.getElementById('prsEvCopyFrom');
        if (copyEn && copyEn.checked) {
          if (!copyFrom || !copyFrom.value) {
            Swal.showValidationMessage('Select an event to copy from, or untick reuse.');
            return false;
          }
          out.copyFromEventId = copyFrom.value;
        }
        return out;
      },
    });
    if (!formValues) return;

    try {
      if (isEdit) {
        await PrsApi.updateEvent(ctx.adminKey, Object.assign({ eventId: existing.eventId }, formValues));
        toast('Event updated.', 'success');
      } else {
        const copyFromEventId = formValues.copyFromEventId;
        const createPayload = {
          eventName: formValues.eventName,
          startDate: formValues.startDate,
          endDate: formValues.endDate,
          notes: formValues.notes,
        };
        const createRes = await PrsApi.createEvent(ctx.adminKey, createPayload);
        const newEventId = createRes && createRes.data && createRes.data.eventId;
        toast('Event created.', 'success');

        if (copyFromEventId && newEventId) {
          const copyOpts = await showCopyOptionsModal();
          if (copyOpts) {
            Swal.fire({ title: 'Copying setup…', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
            try {
              const copyRes = await PrsApi.copyEventSetup(ctx.adminKey, {
                sourceEventId: copyFromEventId,
                targetEventId: newEventId,
                options: {
                  camps: copyOpts.camps,
                  activities: copyOpts.activities,
                  rosters: copyOpts.rosters,
                  materials: copyOpts.materials,
                },
              });
              Swal.close();
              const cd = (copyRes && copyRes.data) || {};
              const nActs = (cd.activities || []).length;
              const nCamps = ((cd.camps && cd.camps.copied) || []).length + ((cd.camps && cd.camps.reused) || []).length;
              toast(`Copied setup: ${nCamps} camp(s), ${nActs} activit${nActs === 1 ? 'y' : 'ies'}. Download new QR codes from Camps.`, 'success');
            } catch (copyErr) {
              Swal.close();
              err(copyErr);
            }
          }
        }
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
        <div style="margin:1rem 0;display:flex;gap:0.5rem;flex-wrap:wrap;">
          <button class="btn btn-primary" id="prsAddCampBtn">➕ Add Camp</button>
          <button class="btn btn-secondary" id="prsCopyCampsBtn">📋 Copy from another event</button>
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
        tableWrap.innerHTML = '<div id="prsCampsViewWrap"></div>';

        paginatedTableView({
          wrap: document.getElementById('prsCampsViewWrap'),
          rows: camps,
          searchPlaceholder: 'Search by camp name or state…',
          emptyText: 'No camps for this event yet.',
          searchFields: c => [c.campName, c.state].filter(Boolean).join(' '),
          rowAttrs: c => `data-camp-id="${esc(c.campId)}"`,
          columns: [
            { label: 'Name',  render: c => esc(c.campName) },
            { label: 'State', render: c => esc(c.state) },
            { label: 'GPS',   render: c => `${Number(c.latitude).toFixed(5)}, ${Number(c.longitude).toFixed(5)}` },
            { label: 'Radius (m)', render: c => esc(c.radiusMeters) },
            { label: 'QR Token',   render: c => `<code style="font-size:0.75rem;">${esc(String(c.qrToken || '').slice(0, 10))}…</code>` },
            { label: 'Actions', render: () => `
                <button class="btn btn-xs btn-secondary prs-camp-qr">QR</button>
                <button class="btn btn-xs btn-secondary prs-camp-edit">Edit</button>
                <button class="btn btn-xs btn-secondary prs-camp-regen">Regen Token</button>
                <button class="btn btn-xs btn-danger prs-camp-del">Delete</button>
              ` },
          ],
          onAfterRender: (tbody) => {
            tbody.querySelectorAll('.prs-camp-qr').forEach(b => b.addEventListener('click', () => showCampQr(b.closest('tr').dataset.campId)));
            tbody.querySelectorAll('.prs-camp-edit').forEach(b => b.addEventListener('click', () => {
              const camp = camps.find(x => String(x.campId) === b.closest('tr').dataset.campId);
              showCampModal(eventId, camp, reload);
            }));
            tbody.querySelectorAll('.prs-camp-regen').forEach(b => b.addEventListener('click', async () => {
              const c = await Swal.fire({ icon: 'warning', title: 'Regenerate QR token?', text: 'The old QR code will stop working.', showCancelButton: true });
              if (!c.isConfirmed) return;
              try {
                await PrsApi.regenerateCampToken(ctx.adminKey, b.closest('tr').dataset.campId);
                toast('Token regenerated. Open QR to download the new code.', 'success');
                reload();
                showCampQr(b.closest('tr').dataset.campId);
              } catch (e) { err(e); }
            }));
            tbody.querySelectorAll('.prs-camp-del').forEach(b => b.addEventListener('click', async () => {
              const c = await Swal.fire({ icon: 'warning', title: 'Delete this camp?', text: 'Only possible if there are no attendance logs yet.', showCancelButton: true });
              if (!c.isConfirmed) return;
              try {
                await PrsApi.deleteCamp(ctx.adminKey, b.closest('tr').dataset.campId);
                toast('Camp deleted.', 'success');
                reload();
              } catch (e) { err(e); }
            }));
          },
        });
      } catch (e) {
        tableWrap.innerHTML = `<div class="info info-error">${esc(e.message)}</div>`;
      }
    }

    eventSelect.addEventListener('change', reload);
    document.getElementById('prsRefreshCampsBtn').addEventListener('click', reload);
    document.getElementById('prsAddCampBtn').addEventListener('click', () => {
      showCampModal(eventSelect.value, null, reload);
    });
    document.getElementById('prsCopyCampsBtn').addEventListener('click', async () => {
      const targetEventId = eventSelect.value;
      if (!targetEventId) return;
      const allEv = await loadAllEvents();
      const sourceEventId = await pickSourceEvent(allEv, targetEventId, 'Copy camps from…');
      if (!sourceEventId) return;
      const confirm = await Swal.fire({
        icon: 'question',
        title: 'Copy camp locations?',
        text: 'Camp name, state, GPS coordinates and radius will be copied. Each camp gets a new QR code and token. Old QR codes from the previous event will not work.',
        showCancelButton: true,
        confirmButtonText: 'Copy camps',
      });
      if (!confirm.isConfirmed) return;
      try {
        Swal.fire({ title: 'Copying camps…', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
        const res = await PrsApi.copyCampsFromEvent(ctx.adminKey, {
          sourceEventId,
          targetEventId,
          skipExisting: false,
        });
        Swal.close();
        const d = (res && res.data) || {};
        toast(`${(d.copied || []).length} copied, ${(d.reused || []).length} matched existing.`, 'success');
        reload();
      } catch (e) {
        Swal.close();
        err(e);
      }
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
  // 2d. Activities (Assignment Definitions) — v2 multi-assignment workspace
  // ---------------------------------------------------------------------------
  // Two-level UI:
  //   Level 1: pick event → list of activities for that event (table with
  //            badges showing # camps, # staff, # materials, mode).
  //   Level 2: drill into a single activity → tabbed sub-workspace for
  //            Camps, Roster, Materials.

  // Per-activity navigation state. Cleared whenever the user backs out of
  // an activity or switches events.
  const asgNav = { selectedDefId: null };

  async function renderAssignments(container) {
    const evRes = await PrsApi.listEvents(ctx.adminKey, '');
    const events = (evRes && evRes.data && evRes.data.events) || [];
    if (events.length === 0) {
      container.innerHTML = `<section class="card"><h2>Activities</h2><p class="info info-muted">Create an event and add camps first.</p></section>`;
      return;
    }

    container.innerHTML = `
      <section class="card card-full-width prs-asg-list-panel">
        <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:1rem;flex-wrap:wrap;gap:1rem;">
          <h2 style="margin:0;">Activities <span style="font-size:0.85rem;color:var(--text-muted);font-weight:400;">(per-event assignments)</span></h2>
          <div style="display:flex;gap:0.5rem;flex-wrap:wrap;">
            <button class="btn btn-secondary" id="prsAsgCopyFromEventBtn">📋 Copy from previous event</button>
            <button class="btn btn-secondary" id="prsAsgDownloadBtn">⬇ Download Records (Excel)</button>
            <button class="btn btn-primary"   id="prsAsgNewDefBtn">➕ New Activity</button>
          </div>
        </div>
        <div style="display:flex;gap:1rem;flex-wrap:wrap;align-items:end;margin-bottom:1rem;">
          <div style="flex:1;min-width:220px;">
            <label style="display:block;font-weight:600;margin-bottom:0.25rem;">Event</label>
            <select id="prsAsgEvent" class="swal2-select" style="width:100%;padding:0.5rem;">
              ${events.map(e => `<option value="${esc(e.eventId)}">${esc(e.eventName)}</option>`).join('')}
            </select>
          </div>
          <div style="flex:1;min-width:220px;">
            <label style="display:block;font-weight:600;margin-bottom:0.25rem;">Camp (filter, optional)</label>
            <select id="prsAsgCampFilter" class="swal2-select" style="width:100%;padding:0.5rem;">
              <option value="">All camps</option>
            </select>
          </div>
        </div>
        <div id="prsAsgDefsList"></div>
      </section>
      <div id="prsAsgDefDetail" class="prs-asg-detail-host" aria-live="polite"></div>
    `;

    const evSel = document.getElementById('prsAsgEvent');
    const campFilter = document.getElementById('prsAsgCampFilter');
    const listEl = document.getElementById('prsAsgDefsList');
    const detailEl = document.getElementById('prsAsgDefDetail');

    async function refreshCampFilter() {
      try {
        const r = await PrsApi.listCamps(ctx.adminKey, evSel.value);
        const camps = (r && r.data && r.data.camps) || [];
        campFilter.innerHTML = '<option value="">All camps</option>' +
          camps.map(c => `<option value="${esc(c.campId)}">${esc(c.campName)} — ${esc(c.state)}</option>`).join('');
        campFilter.dataset.camps = JSON.stringify(camps);
      } catch (e) { /* ignore */ }
    }

    async function reloadDefs() {
      listEl.innerHTML = '<div class="info info-muted">Loading activities…</div>';
      try {
        const r = await PrsApi.listAssignmentDefs(ctx.adminKey, evSel.value, campFilter.value);
        const defs = (r && r.data && r.data.assignmentDefs) || [];
        listEl.innerHTML = '<div id="prsAsgDefsTableWrap"></div>';
        paginatedTableView({
          wrap: document.getElementById('prsAsgDefsTableWrap'),
          rows: defs,
          searchPlaceholder: 'Search activities by title, mode or camp…',
          emptyText: 'No activities yet. Click "New Activity" to add one (e.g. "Distribution of Kits", "Verification of Foreign Credentials"…).',
          searchFields: d => [
            d.title, d.attendanceMode, d.status, d.description,
            ...((d.camps || []).map(c => c.campName + ' ' + c.state)),
          ].filter(Boolean).join(' '),
          rowAttrs: d => `data-def-id="${esc(d.assignmentDefId)}"`,
          columns: [
            { label: 'Activity', render: d => `
              <div style="font-weight:600;">${esc(d.title)}</div>
              <div style="font-size:0.8rem;color:var(--text-muted);">${esc((d.description || '').slice(0, 80))}${(d.description || '').length > 80 ? '…' : ''}</div>
            ` },
            { label: 'Mode', render: d => `<span class="badge badge-info" title="${esc(d.attendanceMode)}">${esc(d.attendanceMode)}</span>` },
            { label: 'Days',  render: d => esc(d.requiredDays) },
            { label: 'Camps', render: d => (d.camps && d.camps.length)
              ? d.camps.map(c => `<span class="badge badge-muted" style="margin-right:0.25rem;">${esc(c.campName || '—')}</span>`).join('')
              : `<span class="info-muted">No camps yet</span>` },
            { label: 'Staff',     render: d => esc(d.rosterCount) },
            { label: 'Materials', render: d => esc(d.materialsCount) },
            { label: 'Status', render: d => `<span class="badge badge-${String(d.status).toLowerCase() === 'active' ? 'success' : 'muted'}">${esc(d.status)}</span>` },
            { label: 'Actions', render: () => `
              <button class="btn btn-xs btn-primary prs-def-open">Open</button>
              <button class="btn btn-xs btn-secondary prs-def-edit">Edit</button>
              <button class="btn btn-xs btn-secondary prs-def-dup" title="Duplicate this activity">Duplicate</button>
            ` },
          ],
          onAfterRender: (tbody) => {
            tbody.querySelectorAll('.prs-def-open').forEach(b => b.addEventListener('click', () => {
              const defId = b.closest('tr').dataset.defId;
              const def = defs.find(x => String(x.assignmentDefId) === defId);
              if (def) openAssignmentDefDetail(def, detailEl, JSON.parse(campFilter.dataset.camps || '[]'), () => reloadDefs());
            }));
            tbody.querySelectorAll('.prs-def-edit').forEach(b => b.addEventListener('click', () => {
              const defId = b.closest('tr').dataset.defId;
              const def = defs.find(x => String(x.assignmentDefId) === defId);
              if (def) showAssignmentDefModal(evSel.value, JSON.parse(campFilter.dataset.camps || '[]'), def, () => reloadDefs());
            }));
            tbody.querySelectorAll('.prs-def-dup').forEach(b => b.addEventListener('click', () => {
              const defId = b.closest('tr').dataset.defId;
              const def = defs.find(x => String(x.assignmentDefId) === defId);
              if (def) duplicateActivityPrompt(def, evSel.value, () => reloadDefs());
            }));
          },
        });
      } catch (e) {
        if (isPrsSheetSetupError(e)) {
          renderPrsSetupPrompt(listEl, e.message || String(e), reloadDefs);
          return;
        }
        listEl.innerHTML = `<div class="info info-error">${esc(e.message)}</div>`;
      }
    }

    evSel.addEventListener('change', async () => {
      asgNav.selectedDefId = null;
      detailEl.innerHTML = '';
      await refreshCampFilter();
      reloadDefs();
    });
    campFilter.addEventListener('change', reloadDefs);
    document.getElementById('prsAsgNewDefBtn').addEventListener('click', () => {
      const camps = JSON.parse(campFilter.dataset.camps || '[]');
      showAssignmentDefModal(evSel.value, camps, null, () => reloadDefs());
    });
    document.getElementById('prsAsgCopyFromEventBtn').addEventListener('click', async () => {
      const targetEventId = evSel.value;
      if (!targetEventId) return;
      const allEv = await loadAllEvents();
      const sourceEventId = await pickSourceEvent(allEv, targetEventId, 'Copy activities from…');
      if (!sourceEventId) return;
      const copyOpts = await showCopyOptionsModal({ camps: true, activities: true, rosters: true, materials: true });
      if (!copyOpts) return;
      try {
        Swal.fire({ title: 'Copying…', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
        const res = await PrsApi.copyEventSetup(ctx.adminKey, {
          sourceEventId,
          targetEventId,
          options: copyOpts,
        });
        Swal.close();
        const d = (res && res.data) || {};
        toast(`Copied ${(d.activities || []).length} activit${(d.activities || []).length === 1 ? 'y' : 'ies'}.`, 'success');
        await refreshCampFilter();
        reloadDefs();
      } catch (e) {
        Swal.close();
        err(e);
      }
    });
    document.getElementById('prsAsgDownloadBtn').addEventListener('click', async () => {
      try {
        const eventId = evSel.value;
        if (!eventId) return toast('Please select an event first.', 'info');
        const res = await PrsApi.downloadAssignmentsExcel(ctx.adminKey, eventId);
        const data = (res && res.data) || {};
        if (data.fileBase64) {
          downloadBase64File(
            data.fileBase64,
            data.fileName || 'prs_assignments.xlsx',
            data.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          );
        } else if (data.downloadUrl) {
          const a = document.createElement('a');
          a.href = data.downloadUrl; a.target = '_blank'; a.rel = 'noopener';
          a.download = data.fileName || 'prs_assignments.xlsx';
          a.click();
        } else {
          throw new Error('Could not generate spreadsheet for download.');
        }
        toast('Spreadsheet generated. Download starting…', 'success');
      } catch (e) { err(e); }
    });

    await refreshCampFilter();
    await reloadDefs();
  }

  // ---- Modal: create / edit Assignment Definition ----
  async function duplicateActivityPrompt(sourceDef, targetEventId, onDone) {
    const sameEvent = String(sourceDef.eventId || '') === String(targetEventId);
    const result = await Swal.fire({
      title: 'Duplicate activity',
      width: 560,
      html: `
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem;font-weight:600;">New title</label>
        <input id="prsDupTitle" class="swal2-input" value="${esc(sourceDef.title + (sameEvent ? ' (Copy)' : ''))}" />
        <label style="display:flex;align-items:center;gap:0.5rem;margin:0.65rem 0;text-align:left;">
          <input type="checkbox" id="prsDupRoster" checked /> Copy staff roster
        </label>
        <label style="display:flex;align-items:center;gap:0.5rem;margin:0.65rem 0;text-align:left;">
          <input type="checkbox" id="prsDupMaterials" checked /> Copy materials &amp; download links
        </label>
        ${!sameEvent ? `
        <label style="display:flex;align-items:center;gap:0.5rem;margin:0.65rem 0;text-align:left;">
          <input type="checkbox" id="prsDupAutoCamp" checked /> Copy missing camps from source event
        </label>
        ` : ''}
      `,
      focusConfirm: false,
      showCancelButton: true,
      confirmButtonText: 'Duplicate',
      preConfirm: () => {
        const title = document.getElementById('prsDupTitle').value.trim();
        if (!title) { Swal.showValidationMessage('Title is required.'); return false; }
        return {
          title,
          copyRoster: document.getElementById('prsDupRoster').checked,
          copyMaterials: document.getElementById('prsDupMaterials').checked,
          autoCopyMissingCamps: !sameEvent && document.getElementById('prsDupAutoCamp')
            ? document.getElementById('prsDupAutoCamp').checked
            : false,
        };
      },
    });
    if (!result.isConfirmed || !result.value) return;
    try {
      Swal.fire({ title: 'Duplicating…', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
      const res = await PrsApi.duplicateAssignmentDef(ctx.adminKey, {
        sourceAssignmentDefId: sourceDef.assignmentDefId,
        targetEventId,
        title: result.value.title,
        copyRoster: result.value.copyRoster,
        copyMaterials: result.value.copyMaterials,
        autoCopyMissingCamps: result.value.autoCopyMissingCamps,
      });
      Swal.close();
      const d = (res && res.data) || {};
      toast(`Duplicated: ${d.rosterAdded || 0} staff, ${d.materialsAdded || 0} materials.`, 'success');
      if (typeof onDone === 'function') onDone();
    } catch (e) {
      Swal.close();
      err(e);
    }
  }

  async function showAssignmentDefModal(eventId, camps, existing, onDone) {
    const isEdit = !!existing;
    const modeOpts = PRS_ATTENDANCE_MODES.map(m =>
      `<option value="${esc(m.value)}" ${existing && existing.attendanceMode === m.value ? 'selected' : ''}>${esc(m.label)}</option>`
    ).join('');
    const linkedCampIds = isEdit ? new Set((existing.camps || []).map(c => String(c.campId))) : new Set();
    const campsHtml = camps.length
      ? camps.map(c => `
          <label style="display:flex;align-items:center;gap:0.5rem;padding:0.35rem 0.5rem;border:1px solid #e5e7eb;border-radius:6px;background:#fff;">
            <input type="checkbox" class="prs-def-camp" value="${esc(c.campId)}" ${linkedCampIds.has(String(c.campId)) ? 'checked' : ''} />
            <span>${esc(c.campName)} — ${esc(c.state)}</span>
          </label>
        `).join('')
      : '<p class="info info-muted">No camps in this event yet. Add one first.</p>';

    const { value } = await Swal.fire({
      title: isEdit ? 'Edit Activity' : 'New Activity',
      width: 720,
      html: `
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-weight:600;">Title</label>
        <input id="prsDefTitle" class="swal2-input" placeholder="e.g. Distribution of Kits Items"
          value="${esc(existing ? existing.title : '')}" />
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-weight:600;">Description / Brief</label>
        <textarea id="prsDefDesc" class="swal2-textarea" placeholder="What is this activity about? Who is it for?">${esc(existing && existing.description || '')}</textarea>
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-weight:600;">Attendance mode</label>
        <select id="prsDefMode" class="swal2-select" style="width:100%;padding:0.5rem;">
          ${modeOpts}
        </select>
        <div style="display:flex;gap:0.75rem;flex-wrap:wrap;">
          <div style="flex:1;min-width:140px;">
            <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-weight:600;">Required days</label>
            <input id="prsDefDays" class="swal2-input" type="number" min="1" value="${esc(existing ? existing.requiredDays : 1)}" />
          </div>
          <div style="flex:1;min-width:140px;">
            <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-weight:600;">Start date (optional)</label>
            <input id="prsDefStart" class="swal2-input" type="date" value="${esc(existing ? fmtDateOnly(existing.startDate) : '')}" />
          </div>
          <div style="flex:1;min-width:140px;">
            <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-weight:600;">End date (optional)</label>
            <input id="prsDefEnd" class="swal2-input" type="date" value="${esc(existing ? fmtDateOnly(existing.endDate) : '')}" />
          </div>
        </div>
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-weight:600;">Run at camps</label>
        <div id="prsDefCamps" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:0.5rem;text-align:left;">
          ${campsHtml}
        </div>
      `,
      focusConfirm: false,
      showCancelButton: true,
      confirmButtonText: isEdit ? 'Save' : 'Create',
      preConfirm: () => {
        const title = document.getElementById('prsDefTitle').value.trim();
        const description = document.getElementById('prsDefDesc').value.trim();
        const attendanceMode = document.getElementById('prsDefMode').value;
        const requiredDays = parseInt(document.getElementById('prsDefDays').value, 10) || 1;
        const startDate = document.getElementById('prsDefStart').value;
        const endDate = document.getElementById('prsDefEnd').value;
        const campIds = Array.from(document.querySelectorAll('.prs-def-camp:checked')).map(el => el.value);
        if (!title) { Swal.showValidationMessage('Title is required.'); return false; }
        if (!campIds.length && !isEdit) { Swal.showValidationMessage('Pick at least one camp to run this activity at.'); return false; }
        return { title, description, attendanceMode, requiredDays, startDate, endDate, campIds };
      },
    });
    if (!value) return;
    try {
      if (isEdit) {
        await PrsApi.updateAssignmentDef(ctx.adminKey, {
          assignmentDefId: existing.assignmentDefId,
          title: value.title,
          description: value.description,
          attendanceMode: value.attendanceMode,
          requiredDays: value.requiredDays,
          startDate: value.startDate,
          endDate: value.endDate,
        });
        // Reconcile camp links separately so existing links don't get rewritten unintentionally.
        const currentLinked = new Set((existing.camps || []).map(c => String(c.campId)));
        const newLinked = new Set(value.campIds.map(String));
        const toAdd = [...newLinked].filter(c => !currentLinked.has(c));
        const toRemove = [...currentLinked].filter(c => !newLinked.has(c));
        if (toAdd.length) {
          await PrsApi.linkAssignmentToCamps(ctx.adminKey, existing.assignmentDefId, toAdd);
        }
        for (const c of toRemove) {
          await PrsApi.unlinkAssignmentFromCamp(ctx.adminKey, existing.assignmentDefId, c);
        }
        toast('Activity updated.', 'success');
      } else {
        await PrsApi.createAssignmentDef(ctx.adminKey, Object.assign({ eventId }, value));
        toast('Activity created.', 'success');
      }
      if (typeof onDone === 'function') onDone();
    } catch (e) { err(e); }
  }

  // ---- Drill-in: one Assignment Definition (tabs: Camps / Roster / Materials) ----
  async function openAssignmentDefDetail(def, mountEl, allCamps, onParentReload) {
    asgNav.selectedDefId = def.assignmentDefId;
    mountEl.innerHTML = `
      <section class="card card-full-width prs-asg-detail-panel">
        <div class="prs-asg-detail-head">
          <div class="prs-asg-detail-title">
            <h3 style="margin:0;">${esc(def.title)}</h3>
            <div class="prs-asg-detail-meta">
              <span class="badge badge-info">${esc(def.attendanceMode)}</span>
              <span>Required days: <strong>${esc(def.requiredDays)}</strong></span>
              <span>Status: <strong>${esc(def.status)}</strong></span>
            </div>
          </div>
          <div class="prs-asg-detail-actions">
            <button type="button" class="btn btn-sm btn-secondary" id="prsDefEditBtn">Edit</button>
            ${def.status === 'ACTIVE' ? '<button type="button" class="btn btn-sm btn-danger" id="prsDefCloseBtn">Close Activity</button>' : ''}
            <button type="button" class="btn btn-sm btn-secondary" id="prsDefBackBtn">← Back</button>
          </div>
        </div>
        <div class="contextual-actions prs-asg-detail-tabs" id="prsDefSubTabs"></div>
        <div id="prsDefSubBody" class="prs-asg-detail-body"></div>
      </section>
    `;

    const subTabs = [
      { id: 'roster',    label: '👥 Roster' },
      { id: 'camps',     label: '🏕️ Camps' },
      { id: 'materials', label: '📎 Materials' },
    ];
    const subTabsEl = document.getElementById('prsDefSubTabs');
    const subBodyEl = document.getElementById('prsDefSubBody');
    subTabsEl.innerHTML = subTabs.map((t, i) =>
      `<button class="action-btn${i === 0 ? ' active' : ''}" data-subtab="${t.id}"><span>${t.label}</span></button>`
    ).join('');

    function activate(id) {
      if (typeof ScrollTop !== 'undefined') ScrollTop.toTop();
      subTabsEl.querySelectorAll('.action-btn').forEach(b => b.classList.toggle('active', b.dataset.subtab === id));
      var p;
      if (id === 'roster')    p = renderDefRoster(def, subBodyEl, allCamps);
      else if (id === 'camps')     p = renderDefCamps(def, subBodyEl, allCamps);
      else if (id === 'materials') p = renderDefMaterials(def, subBodyEl);
      if (p && typeof p.then === 'function') {
        p.finally(function () {
          if (typeof ScrollTop !== 'undefined') ScrollTop.afterRender();
        });
      } else if (typeof ScrollTop !== 'undefined') {
        ScrollTop.afterRender();
      }
      return p;
    }
    subTabsEl.querySelectorAll('.action-btn').forEach(b => b.addEventListener('click', () => activate(b.dataset.subtab)));

    document.getElementById('prsDefEditBtn').addEventListener('click', () => {
      showAssignmentDefModal(def.eventId, allCamps, def, () => {
        if (typeof onParentReload === 'function') onParentReload();
        mountEl.innerHTML = '';
      });
    });
    const closeBtn = document.getElementById('prsDefCloseBtn');
    if (closeBtn) closeBtn.addEventListener('click', async () => {
      const c = await Swal.fire({
        icon: 'warning',
        title: 'Close activity?',
        text: 'Closed activities no longer accept new sign-ins. Existing roster and logs remain readable.',
        showCancelButton: true,
        confirmButtonText: 'Close Activity',
      });
      if (!c.isConfirmed) return;
      try {
        await PrsApi.closeAssignmentDef(ctx.adminKey, def.assignmentDefId);
        toast('Activity closed.', 'success');
        if (typeof onParentReload === 'function') onParentReload();
        mountEl.innerHTML = '';
      } catch (e) { err(e); }
    });
    document.getElementById('prsDefBackBtn').addEventListener('click', () => {
      asgNav.selectedDefId = null;
      mountEl.innerHTML = '';
      if (typeof ScrollTop !== 'undefined') ScrollTop.afterRender();
    });

    if (typeof ScrollTop !== 'undefined') ScrollTop.toTop();
    activate('roster');
  }

  // ---- Sub-tab: Roster ----
  async function renderDefRoster(def, body, allCamps) {
    const linkedCamps = (def.camps && def.camps.length) ? def.camps : [];
    body.innerHTML = `
      <div style="display:flex;gap:1rem;flex-wrap:wrap;align-items:end;margin-bottom:0.75rem;">
        <div style="flex:1;min-width:220px;">
          <label style="display:block;font-weight:600;margin-bottom:0.25rem;">Filter by camp</label>
          <select id="prsRosterCampFilter" class="swal2-select" style="width:100%;padding:0.5rem;">
            <option value="">All linked camps</option>
            ${linkedCamps.map(c => `<option value="${esc(c.campId)}">${esc(c.campName)} — ${esc(c.state)}</option>`).join('')}
          </select>
        </div>
        <button class="btn btn-primary"   id="prsRosterAddBtn">➕ Add staff</button>
        <button class="btn btn-secondary" id="prsRosterBulkBtn">📥 Bulk add to selected camps</button>
        <button class="btn btn-secondary" id="prsRosterImportBtn">📋 Import from previous activity</button>
      </div>
      <div id="prsRosterWrap"></div>
    `;
    const filter = document.getElementById('prsRosterCampFilter');
    const wrap = document.getElementById('prsRosterWrap');

    async function reload() {
      wrap.innerHTML = '<div class="info info-muted">Loading roster…</div>';
      try {
        const r = await PrsApi.listAssignmentRoster(ctx.adminKey, def.assignmentDefId, filter.value);
        const rows = (r && r.data && r.data.roster) || [];
        wrap.innerHTML = '<div id="prsRosterTableWrap"></div>';
        paginatedTableView({
          wrap: document.getElementById('prsRosterTableWrap'),
          rows: rows,
          searchPlaceholder: 'Search by name, phone, department or camp…',
          emptyText: linkedCamps.length === 0
            ? 'Link this activity to one or more camps first (see the Camps tab).'
            : 'No staff on the roster yet.',
          searchFields: a => [a.staffName, a.phone, a.department, a.campName, a.campState, a.status].filter(Boolean).join(' '),
          rowAttrs: a => `data-asg-id="${esc(a.assignmentId)}"`,
          columns: [
            { label: 'Name',       render: a => esc(a.staffName) },
            { label: 'Phone',      render: a => esc(a.phone) },
            { label: 'Department', render: a => esc(a.department) },
            { label: 'Camp',       render: a => esc((a.campName || '—') + (a.campState ? ' (' + a.campState + ')' : '')) },
            { label: 'Days',       render: a => esc(a.requiredDays) },
            { label: 'Status',     render: a => `<span class="badge badge-${String(a.status).toLowerCase() === 'active' ? 'success' : 'muted'}">${esc(a.status)}</span>` },
            { label: 'Actions',    render: () => `<button class="btn btn-xs btn-danger prs-roster-del">Remove</button>` },
          ],
          onAfterRender: (tbody) => {
            tbody.querySelectorAll('.prs-roster-del').forEach(b => b.addEventListener('click', async () => {
              const c = await Swal.fire({ icon: 'warning', title: 'Remove from roster?', text: 'If sign-ins exist this will be soft-revoked, not deleted.', showCancelButton: true });
              if (!c.isConfirmed) return;
              try {
                await PrsApi.deleteAssignment(ctx.adminKey, b.closest('tr').dataset.asgId);
                toast('Removed.', 'success');
                reload();
              } catch (e) { err(e); }
            }));
          },
        });
      } catch (e) {
        wrap.innerHTML = `<div class="info info-error">${esc(e.message)}</div>`;
      }
    }

    filter.addEventListener('change', reload);
    document.getElementById('prsRosterAddBtn').addEventListener('click', async () => {
      if (linkedCamps.length === 0) return toast('Link the activity to at least one camp first.', 'info');
      const campOpts = linkedCamps.map(c => `<option value="${esc(c.campId)}">${esc(c.campName)} — ${esc(c.state)}</option>`).join('');
      const { value } = await Swal.fire({
        title: 'Add staff to activity',
        html: `
          <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Camp</label>
          <select id="prsRosCamp" class="swal2-select" style="width:100%;padding:0.5rem;">${campOpts}</select>
          <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Staff name</label>
          <input id="prsRosName" class="swal2-input" placeholder="Full name" />
          <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Phone (Nigerian format)</label>
          <input id="prsRosPhone" class="swal2-input" placeholder="0803…" />
          <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;">Department</label>
          <input id="prsRosDept" class="swal2-input" placeholder="e.g. Planning" />
          <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-size:0.85rem;color:#666;">Required-days override (optional)</label>
          <input id="prsRosDays" class="swal2-input" type="number" min="0" placeholder="Leave blank to use activity default (${esc(def.requiredDays)})" />
        `,
        focusConfirm: false,
        showCancelButton: true,
        confirmButtonText: 'Add',
        preConfirm: () => {
          const campId = document.getElementById('prsRosCamp').value;
          const staffName = document.getElementById('prsRosName').value.trim();
          const phone = document.getElementById('prsRosPhone').value.trim();
          const department = document.getElementById('prsRosDept').value.trim();
          const rDays = parseInt(document.getElementById('prsRosDays').value, 10);
          if (!campId || !staffName || !phone) {
            Swal.showValidationMessage('Camp, name and phone are required.'); return false;
          }
          return {
            campId, staffName, phone, department,
            requiredDaysOverride: isNaN(rDays) ? 0 : rDays,
          };
        },
      });
      if (!value) return;
      try {
        await PrsApi.addStaffToAssignment(ctx.adminKey, Object.assign({ assignmentDefId: def.assignmentDefId }, value));
        toast('Staff added.', 'success');
        reload();
      } catch (e) { err(e); }
    });

    document.getElementById('prsRosterBulkBtn').addEventListener('click', async () => {
      if (linkedCamps.length === 0) return toast('Link the activity to at least one camp first.', 'info');
      const campOpts = linkedCamps.map(c =>
        `<label style="display:flex;align-items:center;gap:0.5rem;padding:0.35rem 0.5rem;border:1px solid #e5e7eb;border-radius:6px;background:#fff;">
          <input type="checkbox" class="prs-bulk-camp" value="${esc(c.campId)}" checked />
          <span>${esc(c.campName)} — ${esc(c.state)}</span>
        </label>`
      ).join('');
      const { value } = await Swal.fire({
        title: 'Bulk-add staff to selected camps',
        width: 720,
        html: `
          <p style="text-align:left;font-size:0.85rem;color:#555;">
            Paste one staff per line in <strong>name, phone, department</strong> format (commas).
            Department is optional. Each entry will be added to every camp ticked below.
          </p>
          <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-weight:600;">Camps</label>
          <div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:0.5rem;text-align:left;">${campOpts}</div>
          <label style="display:block;text-align:left;margin:0.75rem 0 0.25rem 0;font-weight:600;">Staff list (one per line)</label>
          <textarea id="prsBulkBox" class="swal2-textarea" rows="8" placeholder="Jane Doe, 08031234567, Planning
John Roe, 07045678901, Operations"></textarea>
        `,
        focusConfirm: false,
        showCancelButton: true,
        confirmButtonText: 'Add',
        preConfirm: () => {
          const campIds = Array.from(document.querySelectorAll('.prs-bulk-camp:checked')).map(el => el.value);
          if (!campIds.length) { Swal.showValidationMessage('Tick at least one camp.'); return false; }
          const raw = document.getElementById('prsBulkBox').value || '';
          const staff = raw.split(/\r?\n/).map(l => l.trim()).filter(Boolean).map(line => {
            const parts = line.split(',').map(p => p.trim());
            return { staffName: parts[0] || '', phone: parts[1] || '', department: parts[2] || '' };
          }).filter(s => s.staffName && s.phone);
          if (!staff.length) { Swal.showValidationMessage('At least one valid "name, phone" line is required.'); return false; }
          return { campIds, staff };
        },
      });
      if (!value) return;
      try {
        const r = await PrsApi.bulkAddStaffToAssignment(ctx.adminKey, Object.assign({ assignmentDefId: def.assignmentDefId }, value));
        const d = (r && r.data) || {};
        toast(`Added ${d.added || 0} • Skipped ${((d.skipped || []).length) || 0}`, (d.skipped && d.skipped.length) ? 'info' : 'success');
        reload();
      } catch (e) { err(e); }
    });

    document.getElementById('prsRosterImportBtn').addEventListener('click', async () => {
      const allEv = await loadAllEvents();
      const sourceEventId = await pickSourceEvent(allEv, def.eventId, 'Import roster from event…');
      if (!sourceEventId) return;
      const sourceDefId = await pickSourceActivity(sourceEventId, 'Import roster from activity…');
      if (!sourceDefId) return;
      const confirm = await Swal.fire({
        icon: 'question',
        title: 'Import staff roster?',
        html: `<p style="text-align:left;font-size:0.9rem;">Staff are matched to camps by <strong>camp name + state</strong>. Duplicates (same phone at same camp) are skipped. Missing camps can be copied automatically.</p>`,
        showCancelButton: true,
        confirmButtonText: 'Import roster',
      });
      if (!confirm.isConfirmed) return;
      try {
        Swal.fire({ title: 'Importing roster…', allowOutsideClick: false, didOpen: () => Swal.showLoading() });
        const res = await PrsApi.copyRosterFromAssignment(ctx.adminKey, {
          sourceAssignmentDefId: sourceDefId,
          targetAssignmentDefId: def.assignmentDefId,
          autoCopyMissingCamps: true,
        });
        Swal.close();
        const d = (res && res.data) || {};
        toast(`${d.rosterAdded || 0} imported, ${d.rosterSkipped || 0} skipped.`, 'success');
        reload();
      } catch (e) {
        Swal.close();
        err(e);
      }
    });

    reload();
  }

  // ---- Sub-tab: Camps ----
  async function renderDefCamps(def, body, allCamps) {
    body.innerHTML = `
      <p class="info info-muted">Tick camps where this activity should run. Unticking removes the link (soft delete — roster/logs are preserved).</p>
      <div id="prsDefCampGrid" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:0.5rem;margin-bottom:1rem;">
        ${allCamps.length ? allCamps.map(c => {
          const linked = (def.camps || []).some(x => String(x.campId) === String(c.campId));
          return `
            <label style="display:flex;align-items:center;gap:0.5rem;padding:0.4rem 0.6rem;border:1px solid #e5e7eb;border-radius:6px;background:#fff;">
              <input type="checkbox" class="prs-def-detail-camp" value="${esc(c.campId)}" ${linked ? 'checked' : ''} />
              <span><strong>${esc(c.campName)}</strong> — ${esc(c.state)}</span>
            </label>
          `;
        }).join('') : '<p class="info info-muted">No camps in this event yet.</p>'}
      </div>
      <button class="btn btn-primary" id="prsDefCampSaveBtn">Save camp links</button>
    `;
    document.getElementById('prsDefCampSaveBtn').addEventListener('click', async () => {
      const desired = new Set(Array.from(document.querySelectorAll('.prs-def-detail-camp:checked')).map(el => el.value));
      const current = new Set((def.camps || []).map(c => String(c.campId)));
      const toAdd = [...desired].filter(c => !current.has(c));
      const toRemove = [...current].filter(c => !desired.has(c));
      try {
        if (toAdd.length) {
          await PrsApi.linkAssignmentToCamps(ctx.adminKey, def.assignmentDefId, toAdd);
        }
        for (const c of toRemove) {
          await PrsApi.unlinkAssignmentFromCamp(ctx.adminKey, def.assignmentDefId, c);
        }
        toast('Camp links saved.', 'success');
        // Refresh the in-memory def so subsequent renders reflect new links.
        const r = await PrsApi.listAssignmentCamps(ctx.adminKey, def.assignmentDefId);
        def.camps = (r && r.data && r.data.camps) || [];
      } catch (e) { err(e); }
    });
  }

  // ---- Sub-tab: Materials ----
  async function renderDefMaterials(def, body) {
    body.innerHTML = `
      <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:0.75rem;margin-bottom:0.75rem;">
        <p class="info info-muted" style="margin:0;">Materials are downloadable resources (forms, manuals, guides). The Release Rule decides when staff can claim them via the QR page.</p>
        <button class="btn btn-primary" id="prsMatAddBtn">➕ Add material</button>
      </div>
      <div id="prsMatWrap"></div>
    `;
    const wrap = document.getElementById('prsMatWrap');

    async function reload() {
      wrap.innerHTML = '<div class="info info-muted">Loading materials…</div>';
      try {
        const [r, s] = await Promise.all([
          PrsApi.listAssignmentMaterials(ctx.adminKey, def.assignmentDefId),
          PrsApi.getMaterialReleaseStats(ctx.adminKey, def.assignmentDefId),
        ]);
        const mats = (r && r.data && r.data.materials) || [];
        const stats = (s && s.data && s.data.stats) || {};
        wrap.innerHTML = '<div id="prsMatTableWrap"></div>';
        paginatedTableView({
          wrap: document.getElementById('prsMatTableWrap'),
          rows: mats,
          searchPlaceholder: 'Search materials by title, rule or type…',
          emptyText: 'No materials yet. Paste a hosted file URL (Drive, Dropbox, NYSC portal…).',
          searchFields: m => [m.title, m.releaseRule, m.fileType, m.status].filter(Boolean).join(' '),
          rowAttrs: m => `data-mat-id="${esc(m.materialId)}"`,
          columns: [
            { label: 'Title', render: m => `
              <div style="font-weight:600;">${esc(m.title)}</div>
              <div style="font-size:0.8rem;"><a href="${esc(m.fileUrl)}" target="_blank" rel="noopener" style="word-break:break-all;">${esc(m.fileUrl)}</a></div>
            ` },
            { label: 'Type',  render: m => esc(m.fileType || '') },
            { label: 'Rule',  render: m => `<span class="badge badge-info" title="${esc(prsReleaseRuleLabel(m))}">${esc(m.releaseRule)}</span>` },
            { label: 'When',  render: m => prsReleaseTiming(m) },
            { label: 'Released?', render: m => m.releaseRule === 'MANUAL'
                ? (m.releasedAt
                    ? `<span class="badge badge-success">Released ${esc(fmtDate(m.releasedAt))}</span>`
                    : `<span class="badge badge-warning">Pending admin release</span>`)
                : '<span class="badge badge-muted">Auto</span>' },
            { label: 'Downloads', render: m => {
              const st = stats[m.materialId] || { downloads: 0, distinctPhones: 0 };
              return `${esc(st.downloads)} <span style="color:var(--text-muted);font-size:0.8rem;">(${esc(st.distinctPhones)} staff)</span>`;
            } },
            { label: 'Status', render: m => `<span class="badge badge-${m.status === 'ACTIVE' ? 'success' : 'muted'}">${esc(m.status)}</span>` },
            { label: 'Actions', render: m => `
              ${m.releaseRule === 'MANUAL' && !m.releasedAt && m.status === 'ACTIVE'
                ? '<button class="btn btn-xs btn-primary prs-mat-release">Release now</button>' : ''}
              <button class="btn btn-xs btn-secondary prs-mat-edit">Edit</button>
              <button class="btn btn-xs btn-danger prs-mat-del">Remove</button>
            ` },
          ],
          onAfterRender: (tbody) => {
            tbody.querySelectorAll('.prs-mat-edit').forEach(b => b.addEventListener('click', () => {
              const id = b.closest('tr').dataset.matId;
              const m = mats.find(x => String(x.materialId) === id);
              showMaterialModal(def, m, reload);
            }));
            tbody.querySelectorAll('.prs-mat-release').forEach(b => b.addEventListener('click', async () => {
              const id = b.closest('tr').dataset.matId;
              try {
                await PrsApi.releaseAssignmentMaterialNow(ctx.adminKey, id);
                toast('Material released.', 'success');
                reload();
              } catch (e) { err(e); }
            }));
            tbody.querySelectorAll('.prs-mat-del').forEach(b => b.addEventListener('click', async () => {
              const c = await Swal.fire({ icon: 'warning', title: 'Remove material?', text: 'This soft-deletes the material; download history is preserved.', showCancelButton: true });
              if (!c.isConfirmed) return;
              try {
                await PrsApi.deleteAssignmentMaterial(ctx.adminKey, b.closest('tr').dataset.matId);
                toast('Removed.', 'success');
                reload();
              } catch (e) { err(e); }
            }));
          },
        });
      } catch (e) {
        wrap.innerHTML = `<div class="info info-error">${esc(e.message)}</div>`;
      }
    }

    document.getElementById('prsMatAddBtn').addEventListener('click', () => showMaterialModal(def, null, reload));
    reload();
  }

  function prsReleaseRuleLabel(m) {
    const r = PRS_RELEASE_RULES.find(x => x.value === m.releaseRule);
    return r ? r.label : m.releaseRule;
  }

  function prsReleaseTiming(m) {
    if (m.releaseRule === 'IMMEDIATE') return 'Immediately';
    if (m.releaseRule === 'ON_SIGN_IN') return 'On first sign-in';
    if (m.releaseRule === 'ON_SIGN_IN_DAY_N') return 'On sign-in day ' + esc(m.releaseAfterSignInDay || 1);
    if (m.releaseRule === 'ON_DATE') return m.releaseDate ? esc(fmtDateOnly(m.releaseDate)) : '<span class="info-muted">No date set</span>';
    if (m.releaseRule === 'MANUAL') return 'Admin click';
    return '';
  }

  async function showMaterialModal(def, existing, onDone) {
    const isEdit = !!existing;
    const ruleOpts = PRS_RELEASE_RULES.map(r =>
      `<option value="${esc(r.value)}" ${existing && existing.releaseRule === r.value ? 'selected' : ''}>${esc(r.label)}</option>`
    ).join('');
    const { value } = await Swal.fire({
      title: isEdit ? 'Edit material' : 'Add material',
      width: 680,
      html: `
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-weight:600;">Title</label>
        <input id="prsMatTitle" class="swal2-input" placeholder="e.g. Foreign Credential Verification Form"
          value="${esc(existing ? existing.title : '')}" />
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-weight:600;">Description</label>
        <textarea id="prsMatDesc" class="swal2-textarea" placeholder="What is this file? Any notes for the staff?">${esc(existing && existing.description || '')}</textarea>
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-weight:600;">File URL</label>
        <input id="prsMatUrl" class="swal2-input" placeholder="https://drive.google.com/…"
          value="${esc(existing ? existing.fileUrl : '')}" />
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-weight:600;">File type (optional)</label>
        <input id="prsMatType" class="swal2-input" placeholder="pdf / docx / xlsx / image / zip…"
          value="${esc(existing ? existing.fileType : '')}" />
        <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-weight:600;">Release rule</label>
        <select id="prsMatRule" class="swal2-select" style="width:100%;padding:0.5rem;">${ruleOpts}</select>
        <div id="prsMatDateWrap" style="display:none;">
          <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-weight:600;">Release date</label>
          <input id="prsMatDate" class="swal2-input" type="date" value="${esc(existing && existing.releaseDate ? fmtDateOnly(existing.releaseDate) : '')}" />
        </div>
        <div id="prsMatDayNWrap" style="display:none;">
          <label style="display:block;text-align:left;margin:0.5rem 0 0.25rem 0;font-weight:600;">Release after sign-in day N</label>
          <input id="prsMatDayN" class="swal2-input" type="number" min="1" value="${esc(existing && existing.releaseAfterSignInDay || 1)}" />
        </div>
      `,
      focusConfirm: false,
      showCancelButton: true,
      confirmButtonText: isEdit ? 'Save' : 'Add',
      didOpen: () => {
        const ruleSel = document.getElementById('prsMatRule');
        const dateWrap = document.getElementById('prsMatDateWrap');
        const dayNWrap = document.getElementById('prsMatDayNWrap');
        function sync() {
          dateWrap.style.display = ruleSel.value === 'ON_DATE' ? 'block' : 'none';
          dayNWrap.style.display = ruleSel.value === 'ON_SIGN_IN_DAY_N' ? 'block' : 'none';
        }
        ruleSel.addEventListener('change', sync);
        sync();
      },
      preConfirm: () => {
        const title = document.getElementById('prsMatTitle').value.trim();
        const description = document.getElementById('prsMatDesc').value.trim();
        const fileUrl = document.getElementById('prsMatUrl').value.trim();
        const fileType = document.getElementById('prsMatType').value.trim();
        const releaseRule = document.getElementById('prsMatRule').value;
        const releaseDate = document.getElementById('prsMatDate').value || '';
        const releaseAfterSignInDay = parseInt(document.getElementById('prsMatDayN').value, 10) || 0;
        if (!title || !fileUrl) {
          Swal.showValidationMessage('Title and file URL are required.'); return false;
        }
        if (!/^https?:\/\//i.test(fileUrl)) {
          Swal.showValidationMessage('File URL must start with http:// or https://'); return false;
        }
        return { title, description, fileUrl, fileType, releaseRule, releaseDate, releaseAfterSignInDay };
      },
    });
    if (!value) return;
    try {
      if (isEdit) {
        await PrsApi.updateAssignmentMaterial(ctx.adminKey, Object.assign({ materialId: existing.materialId }, value));
        toast('Material updated.', 'success');
      } else {
        await PrsApi.addAssignmentMaterial(ctx.adminKey, Object.assign({ assignmentDefId: def.assignmentDefId }, value));
        toast('Material added.', 'success');
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

    const campLabelFor = l => l.campName
      ? (l.campState ? `${l.campName} (${l.campState})` : l.campName)
      : '— (camp removed)';
    const eventLabelFor = l => l.eventName || '— (event removed)';

    async function reload() {
      wrap.innerHTML = '<div class="info info-muted">Loading…</div>';
      try {
        const r = await PrsApi.getAttendanceLogs(ctx.adminKey, {
          eventId: evSel.value || '',
          dateStr: dtSel.value || '',
          limit: 500,
        });
        const rows = (r && r.data && r.data.logs) || [];
        wrap.innerHTML = '<div id="prsLogsViewWrap"></div>';

        paginatedTableView({
          wrap: document.getElementById('prsLogsViewWrap'),
          rows: rows,
          searchPlaceholder: 'Search by staff name, phone, camp, event or action…',
          emptyText: 'No logs.',
          searchFields: l => [
            l.staffName, l.phone, l.action, l.campName, l.campState, l.eventName,
            campLabelFor(l), eventLabelFor(l),
          ].filter(Boolean).join(' '),
          columns: [
            { label: 'When',   render: l => esc(fmtDate(l.timestamp)) },
            { label: 'Action', render: l => `<span class="badge badge-${l.action === 'SIGN_IN' ? 'success' : 'muted'}">${esc(l.action)}</span>` },
            { label: 'Staff',  render: l => esc(l.staffName || '') },
            { label: 'Phone',  render: l => esc(l.phone) },
            { label: 'Camp',   render: l => esc(campLabelFor(l)) },
            { label: 'Event',  render: l => esc(eventLabelFor(l)) },
          ],
        });
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
