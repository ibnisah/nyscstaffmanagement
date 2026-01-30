// Simple UI helpers and page controllers with SweetAlert2 integration.

const UI = (function () {
  // Check if SweetAlert2 is available
  const Swal = window.Swal || null;

  function setStatus(elementId, message, type) {
    const el = document.getElementById(elementId);
    if (!el) return;
    el.className = 'status ' + (type || '');
    el.textContent = message || '';
  }

  function show(elementId) {
    const el = document.getElementById(elementId);
    if (el) el.classList.remove('hidden');
  }

  function hide(elementId) {
    const el = document.getElementById(elementId);
    if (el) el.classList.add('hidden');
  }

  // SweetAlert2 wrapper functions for better UX
  function showSuccess(title, message, timer = 3000) {
    if (Swal) {
      return Swal.fire({
        icon: 'success',
        title: title || 'Success',
        text: message,
        timer: timer,
        timerProgressBar: true,
        showConfirmButton: timer === 0,
        confirmButtonColor: '#005B3E'
      });
    } else {
      alert(title + ': ' + message);
    }
  }

  function showError(title, message) {
    if (Swal) {
      return Swal.fire({
        icon: 'error',
        title: title || 'Error',
        text: message,
        confirmButtonColor: '#005B3E'
      });
    } else {
      alert(title + ': ' + message);
    }
  }

  function showWarning(title, message) {
    if (Swal) {
      return Swal.fire({
        icon: 'warning',
        title: title || 'Warning',
        text: message,
        confirmButtonColor: '#005B3E'
      });
    } else {
      alert(title + ': ' + message);
    }
  }

  function showInfo(title, message) {
    if (Swal) {
      return Swal.fire({
        icon: 'info',
        title: title || 'Information',
        text: message,
        confirmButtonColor: '#005B3E'
      });
    } else {
      alert(title + ': ' + message);
    }
  }

  async function confirmAction(title, message, confirmText = 'Yes, proceed', cancelText = 'Cancel') {
    if (Swal) {
      const result = await Swal.fire({
        icon: 'question',
        title: title || 'Confirm Action',
        text: message,
        showCancelButton: true,
        confirmButtonText: confirmText,
        cancelButtonText: cancelText,
        confirmButtonColor: '#005B3E',
        cancelButtonColor: '#6b7280',
        reverseButtons: true
      });
      return result.isConfirmed;
    } else {
      return window.confirm(title + ': ' + message);
    }
  }

  async function showLoading(title = 'Processing...', message = 'Please wait') {
    if (Swal) {
      Swal.fire({
        title: title,
        text: message,
        allowOutsideClick: false,
        allowEscapeKey: false,
        showConfirmButton: false,
        didOpen: () => {
          Swal.showLoading();
        }
      });
    }
  }

  function closeLoading() {
    if (Swal) {
      Swal.close();
    }
  }

  return {
    setStatus,
    show,
    hide,
    showSuccess,
    showError,
    showWarning,
    showInfo,
    confirmAction,
    showLoading,
    closeLoading,
  };
})();

// Attendance page controller (verify.html)
const AttendancePage = (function () {
  let formationId = '';

  async function init() {
    const params = new URLSearchParams(window.location.search);
    const token = params.get('t') || params.get('token') || '';
    formationId = params.get('f') || params.get('formationId') || '';

    document.getElementById('token').value = token;

    document
      .getElementById('startAttendanceBtn')
      .addEventListener('click', handleStart);
    document
      .getElementById('verifyAttendanceForm')
      .addEventListener('submit', handleSubmit);
  }

  async function handleStart() {
    UI.hide('pre-location-warning');
    UI.show('attendance-form');
  }

  async function handleSubmit(e) {
    e.preventDefault();

    if (!formationId) {
      await UI.showError(
        'Invalid QR Code',
        'Missing formation information. Please rescan the official QR code for your office.'
      );
      return;
    }

    UI.showLoading('Capturing Location', 'Please wait while we capture your location...');

    let location;
    try {
      location = await Geo.getLocation({ requiredAccuracyMeters: 50 });
    } catch (err) {
      UI.closeLoading();
      if (err.code === 'DENIED' && Geo.isIosSafari()) {
        await UI.showError(
          'Location Permission Denied',
          'On iPhone: Settings → Safari → Location → Allow While Using App, then reload this page.'
        );
      } else {
        await UI.showError('Location Error', err.message || 'Unable to capture location. Please enable location services.');
      }
      return;
    }

    let deviceHash;
    try {
      deviceHash = await Fingerprint.getHashedFingerprint();
    } catch (err) {
      UI.closeLoading();
      await UI.showError(
        'Device Identification Failed',
        'Cannot securely identify this device. Please use a newer browser that supports device fingerprinting.'
      );
      return;
    }

    const employeeId = document.getElementById('employeeId').value.trim();
    const token = document.getElementById('token').value.trim();

    if (!employeeId || !token) {
      UI.closeLoading();
      await UI.showError('Missing Information', 'Employee ID and token are required.');
      return;
    }

    try {
      UI.showLoading('Marking Attendance', 'Please wait while we process your attendance...');
      const result = await Api.call('markAttendance', {
        employeeId,
        token,
        formationId,
        deviceHash,
        location,
      });
      UI.closeLoading();
      await UI.showSuccess(
        'Attendance Marked!',
        result.message || 'Your attendance has been recorded successfully.',
        4000
      );
      document.getElementById('employeeId').value = '';
    } catch (err) {
      UI.closeLoading();
      await UI.showError(
        'Attendance Failed',
        err.message || 'Failed to mark attendance. Please contact admin if this persists.'
      );
    }
  }

  return {
    init,
  };
})();

// Device registration controller (register-device.html)
const RegisterDevicePage = (function () {
  async function init() {
    document
      .getElementById('registerDeviceForm')
      .addEventListener('submit', handleSubmit);

    // Optional: check if registration mode is enabled
    try {
      UI.showLoading('Checking', 'Verifying registration status...');
      const res = await Api.call('getRegistrationStatus', {});
      UI.closeLoading();
      if (!res.data || !res.data.enabled) {
        UI.show('registration-closed');
        UI.hide('registration-form-wrapper');
      }
    } catch (e) {
      UI.closeLoading();
      // Fail closed: do not allow registration if backend is unreachable
      UI.show('registration-closed');
      UI.hide('registration-form-wrapper');
    }
  }

  async function handleSubmit(e) {
    e.preventDefault();
    UI.showLoading('Capturing Location', 'Please wait while we capture your location...');

    let location;
    try {
      location = await Geo.getLocation({ requiredAccuracyMeters: 50 });
    } catch (err) {
      UI.closeLoading();
      if (err.code === 'DENIED' && Geo.isIosSafari()) {
        await UI.showError(
          'Location Permission Denied',
          'On iPhone: Settings → Safari → Location → Allow While Using App, then reload this page.'
        );
      } else {
        await UI.showError('Location Error', err.message || 'Unable to capture location. Please enable location services.');
      }
      return;
    }

    let deviceHash;
    try {
      deviceHash = await Fingerprint.getHashedFingerprint();
    } catch (err) {
      UI.closeLoading();
      await UI.showError(
        'Device Identification Failed',
        'Cannot securely identify this device. Please use a newer browser that supports device fingerprinting.'
      );
      return;
    }

    const employeeId = document.getElementById('employeeId').value.trim();
    if (!employeeId) {
      UI.closeLoading();
      await UI.showError('Missing Information', 'Employee ID is required.');
      return;
    }

    try {
      UI.showLoading('Registering Device', 'Please wait while we register your device...');
      const res = await Api.call('registerDevice', {
        employeeId,
        deviceHash,
        location,
      });
      UI.closeLoading();
      await UI.showSuccess(
        'Device Registered!',
        res.message || 'Your device has been successfully registered.',
        4000
      );
      document.getElementById('employeeId').value = '';
    } catch (err) {
      UI.closeLoading();
      await UI.showError(
        'Registration Failed',
        err.message || 'Failed to register device. Please try again.'
      );
    }
  }

  return {
    init,
  };
})();

// Visitor page controller (visitor.html)
const VisitorPage = (function () {
  let formationId = '';
  let subUnitId = '';

  async function init() {
    const params = new URLSearchParams(window.location.search);
    formationId = params.get('f') || params.get('formationId') || '';
    subUnitId = params.get('s') || params.get('subUnitId') || '';

    document
      .getElementById('visitorForm')
      .addEventListener('submit', handleSubmit);
  }

  async function handleSubmit(e) {
    e.preventDefault();

    if (!formationId || !subUnitId) {
      await UI.showError(
        'Invalid QR Code',
        'Missing formation or sub-unit information. Please rescan the official visitor QR code for this office.'
      );
      return;
    }

    UI.showLoading('Submitting Request', 'Please wait while we process your visit request...');

    const payload = {
      name: document.getElementById('visitorName').value.trim(),
      phone: document.getElementById('phone').value.trim(),
      purpose: document.getElementById('purpose').value.trim(),
      staffToSee: document.getElementById('staffToSee').value.trim(),
      formationId,
      subUnitId,
    };

    if (!payload.name || !payload.phone || !payload.purpose || !payload.staffToSee) {
      await UI.showError('Missing Information', 'All fields are required. Please fill in all details.');
      return;
    }

    try {
      UI.showLoading('Submitting Request', 'Please wait while we process your visit request...');
      const res = await Api.call('createVisit', payload);

      UI.closeLoading();
      await UI.showSuccess(
        'Visit Request Submitted!',
        res.message || 'Your visit request has been submitted. You will be notified once approved.',
        5000
      );
      e.target.reset();
    } catch (err) {
      UI.closeLoading();
      await UI.showError(
        'Submission Failed',
        err.message || 'Failed to submit visit request. Please try again.'
      );
    }
  }

  return {
    init,
  };
})();

// Admin page controller (admin.html)
const AdminPage = (function () {
  let adminToken = null; // not used for auth, kept for future use
  let adminKey = null;
  let adminRole = null;
  let adminFormationId = null;
  let adminDepartmentId = null;
  let currentFormationId = null; // for SUPER_ADMIN formation selection

  let currentModule = null;

  async function init() {
    document
      .getElementById('adminAuthForm')
      .addEventListener('submit', handleAuth);

    // Setup back to modules buttons (sidebar and header)
    const backToModulesBtn = document.getElementById('backToModulesBtn');
    if (backToModulesBtn) {
      backToModulesBtn.addEventListener('click', (e) => {
        e.preventDefault();
        showModuleSelector();
      });
    }

    // Setup header back buttons (using event delegation for all workspace headers)
    // Use event delegation on the document to catch clicks on back buttons
    // This works even if buttons are in hidden elements initially
    document.addEventListener('click', (e) => {
      // Check if clicked element or its parent has the back-to-modules-header class
      const backButton = e.target.closest('.back-to-modules-header');
      if (backButton) {
        e.preventDefault();
        e.stopPropagation();
        console.log('Back to modules clicked');
        showModuleSelector();
        return false;
      }
    }, true); // Use capture phase to catch events earlier

    // Setup logout button
    const logoutBtn = document.getElementById('adminLogoutBtn');
    if (logoutBtn) {
      logoutBtn.addEventListener('click', handleLogout);
    }

    // Setup mobile sidebar toggle
    const sidebarToggle = document.getElementById('sidebarToggle');
    const sidebarOverlay = document.getElementById('sidebarOverlay');
    const adminSidebar = document.getElementById('adminSidebar');

    if (sidebarToggle && adminSidebar) {
      sidebarToggle.addEventListener('click', () => {
        adminSidebar.classList.toggle('mobile-open');
        if (sidebarOverlay) {
          sidebarOverlay.classList.toggle('active');
        }
      });
    }

    if (sidebarOverlay && adminSidebar) {
      sidebarOverlay.addEventListener('click', () => {
        adminSidebar.classList.remove('mobile-open');
        sidebarOverlay.classList.remove('active');
      });
    }
    // Helper function to safely add event listeners
    function safeAddEventListener(id, event, handler) {
      const el = document.getElementById(id);
      if (el) {
        el.addEventListener(event, handler);
      }
    }

    // These elements are created dynamically, so we'll attach listeners when views load
    // For now, we'll use event delegation or attach listeners when elements are created
    safeAddEventListener('refreshEmployeesBtn', 'click', loadEmployees);
    safeAddEventListener('addEmployeeBtn', 'click', showAddEmployeeModal);
    safeAddEventListener('exportAttendanceBtn', 'click', () => adminAction('exportAttendance'));
    safeAddEventListener('exportVisitorsBtn', 'click', () => adminAction('exportVisitors'));
    safeAddEventListener('forceNewTokenBtn', 'click', () => adminAction('forceNewToken'));
    safeAddEventListener('registrationModeToggle', 'change', handleRegistrationToggle);

    // Super Admin buttons
    safeAddEventListener('refreshDashboardBtn', 'click', loadNationwideDashboard);
    safeAddEventListener('formationWizardBtn', 'click', () => showFormationWizard());
    safeAddEventListener('createFormationBtn', 'click', () => showFormationModal());
    safeAddEventListener('createAdminBtn', 'click', () => showAdminModal());
    safeAddEventListener('createHrmAdminBtn', 'click', showCreateHrmAdminModal);
    safeAddEventListener('refreshHrmAdminsBtn', 'click', loadHrmAdminsTable);

    // HRM Audit Logs buttons
    safeAddEventListener('hrmLogsSearchBtn', 'click', async () => {
      try {
        await loadHrmAuditLogs(1);
      } catch (err) {
        console.error('Error loading HRM audit logs:', err);
      }
    });
    safeAddEventListener('hrmLogsRefreshBtn', 'click', async () => {
      try {
        await loadHrmAuditLogs(1);
      } catch (err) {
        console.error('Error refreshing HRM audit logs:', err);
      }
    });
    safeAddEventListener('hrmLogsExportBtn', 'click', async () => {
      try {
        await exportHrmAuditLogs();
      } catch (err) {
        console.error('Error exporting HRM audit logs:', err);
      }
    });
    safeAddEventListener('runArchivalBtn', 'click', runArchival);
    safeAddEventListener('refreshRetentionBtn', 'click', loadRetentionPolicy);
    safeAddEventListener('setupTriggersBtn', 'click', setupArchivalTriggers);

    // HRM Module buttons
    safeAddEventListener('hrmRefreshBtn', 'click', loadHrmStats);
    safeAddEventListener('hrmViewProfilesBtn', 'click', () => {
      const profilesSection = document.getElementById('hrmProfilesSection');
      const leavesSection = document.getElementById('hrmLeavesSection');
      const performanceSection = document.getElementById('hrmPerformanceSection');
      const transfersSection = document.getElementById('hrmTransfersSection');
      if (profilesSection) profilesSection.classList.remove('hidden');
      if (leavesSection) leavesSection.classList.add('hidden');
      if (performanceSection) performanceSection.classList.add('hidden');
      if (transfersSection) transfersSection.classList.add('hidden');
      loadHrmProfiles();
    });
    safeAddEventListener('hrmManageLeavesBtn', 'click', () => {
      const profilesSection = document.getElementById('hrmProfilesSection');
      const leavesSection = document.getElementById('hrmLeavesSection');
      const performanceSection = document.getElementById('hrmPerformanceSection');
      const transfersSection = document.getElementById('hrmTransfersSection');
      if (profilesSection) profilesSection.classList.add('hidden');
      if (leavesSection) leavesSection.classList.remove('hidden');
      if (performanceSection) performanceSection.classList.add('hidden');
      if (transfersSection) transfersSection.classList.add('hidden');
      loadHrmLeaves();
    });
    safeAddEventListener('hrmPerformanceBtn', 'click', () => {
      const profilesSection = document.getElementById('hrmProfilesSection');
      const leavesSection = document.getElementById('hrmLeavesSection');
      const performanceSection = document.getElementById('hrmPerformanceSection');
      const transfersSection = document.getElementById('hrmTransfersSection');
      if (profilesSection) profilesSection.classList.add('hidden');
      if (leavesSection) leavesSection.classList.add('hidden');
      if (performanceSection) performanceSection.classList.remove('hidden');
      if (transfersSection) transfersSection.classList.add('hidden');
      loadHrmPerformance();
    });
    safeAddEventListener('hrmTransfersBtn', 'click', () => {
      const profilesSection = document.getElementById('hrmProfilesSection');
      const leavesSection = document.getElementById('hrmLeavesSection');
      const performanceSection = document.getElementById('hrmPerformanceSection');
      const transfersSection = document.getElementById('hrmTransfersSection');
      if (profilesSection) profilesSection.classList.add('hidden');
      if (leavesSection) leavesSection.classList.add('hidden');
      if (performanceSection) performanceSection.classList.add('hidden');
      if (transfersSection) transfersSection.classList.remove('hidden');
      loadHrmTransfers();
    });
    safeAddEventListener('hrmCreatePerformanceBtn', 'click', () => showCreatePerformanceModal());

    // HRM Staff Management buttons
    safeAddEventListener('hrmAddStaffBtn', 'click', showAddStaffModal);
    safeAddEventListener('hrmSearchStaffBtn', 'click', () => {
      const dashboardSection = document.getElementById('hrmStaffDashboardSection');
      const listSection = document.getElementById('hrmStaffListSection');
      const searchInput = document.getElementById('hrmStaffSearchInput');
      if (dashboardSection) dashboardSection.classList.add('hidden');
      if (listSection) listSection.classList.remove('hidden');
      if (searchInput) searchInput.value = '';
      loadStaffList();
    });
    safeAddEventListener('hrmViewAllStaffBtn', 'click', () => {
      const dashboardSection = document.getElementById('hrmStaffDashboardSection');
      const listSection = document.getElementById('hrmStaffListSection');
      const searchInput = document.getElementById('hrmStaffSearchInput');
      if (dashboardSection) dashboardSection.classList.add('hidden');
      if (listSection) listSection.classList.remove('hidden');
      if (searchInput) searchInput.value = '';
      loadStaffList();
    });
    safeAddEventListener('hrmRefreshStaffBtn', 'click', loadHrmStaffStats);
    safeAddEventListener('hrmStaffSearchBtn', 'click', () => {
      loadStaffList(1); // Reset to page 1 on new search
    });
    safeAddEventListener('hrmStaffClearSearchBtn', 'click', () => {
      const searchInput = document.getElementById('hrmStaffSearchInput');
      if (searchInput) searchInput.value = '';
      loadStaffList(1);
    });
    safeAddEventListener('hrmIncludeArchivedCheck', 'change', () => {
      loadStaffList(1); // Reset to page 1 when filter changes
    });
  }

  function handleLogout() {
    // Clear admin session
    adminToken = null;
    adminKey = null;
    adminRole = null;
    adminFormationId = null;
    adminDepartmentId = null;
    currentFormationId = null;
    currentModule = null;

    // Hide logout button
    const logoutBtn = document.getElementById('adminLogoutBtn');
    if (logoutBtn) {
      logoutBtn.style.display = 'none';
    }

    // Hide admin layout and module selector
    const adminLayout = document.getElementById('adminLayout');
    const moduleSelector = document.getElementById('moduleSelector');
    if (adminLayout) adminLayout.classList.add('hidden');
    if (moduleSelector) moduleSelector.classList.add('hidden');

    // Show login form
    const loginSection = document.getElementById('loginSection');
    if (loginSection) {
      loginSection.classList.remove('hidden');
      // Clear the admin key input
      const adminKeyInput = document.getElementById('adminKey');
      if (adminKeyInput) adminKeyInput.value = '';
    }

    // Show success message
    UI.showSuccess('Logged Out', 'You have been successfully logged out.', 2000);
  }

  async function handleAuth(e) {
    e.preventDefault();
    const key = document.getElementById('adminKey').value.trim();

    if (!key) {
      await UI.showError('Missing Admin Key', 'Please enter your admin key.');
      return;
    }

    try {
      UI.showLoading('Authenticating', 'Please wait while we verify your credentials...');

      console.log('Attempting admin login with key:', key);
      const res = await Api.call('adminLogin', { key });
      console.log('Login response:', res);

      UI.closeLoading();

      if (!res || !res.data) {
        throw new Error('Invalid response from server.');
      }

      adminToken = res.data.adminToken;
      adminKey = key;
      adminRole = res.data.role;
      adminFormationId = res.data.formationId;
      adminDepartmentId = res.data.departmentId;
      currentFormationId = adminFormationId;

      // Show logout button
      const logoutBtn = document.getElementById('adminLogoutBtn');
      if (logoutBtn) {
        logoutBtn.style.display = 'block';
      }

      await updateAdminContextUI();

      // Hide login form
      const loginSection = document.getElementById('loginSection');
      if (loginSection) {
        loginSection.classList.add('hidden');
      }

      // For HRM_ADMIN and HRM_VIEWER, directly open HRM module without showing module selector
      if (adminRole === 'HRM_ADMIN' || adminRole === 'HRM_VIEWER') {
        // Hide module selector immediately
        const moduleSelector = document.getElementById('moduleSelector');
        if (moduleSelector) {
          moduleSelector.classList.add('hidden');
        }

        // Ensure admin layout is visible
        const adminLayout = document.getElementById('adminLayout');
        if (adminLayout) {
          adminLayout.classList.remove('hidden');
        }

        // Load available modules in background (to set up module cards for "Back to Modules" button)
        // But don't await it - we'll open the module immediately
        loadAvailableModules().catch(err => console.error('Error loading modules:', err));

        // Small delay to ensure DOM is ready, then open HRM module (lowercase 'hrm')
        setTimeout(() => {
          openModule('hrm');
        }, 100);
      } else {
        // Show module selector for other roles
        const moduleSelector = document.getElementById('moduleSelector');
        if (moduleSelector) {
          moduleSelector.classList.remove('hidden');

          // Load and display only modules the admin has access to
          await loadAvailableModules();

          // Setup module card click handlers
          setupModuleSelectors();
        }
      }

      // Show success message
      await UI.showSuccess('Login Successful', `Welcome! You are logged in as ${adminRole}.`, 2000);

      // Show HRM Staff Dashboard only for HRM_ADMIN and SUPER_ADMIN
      const hrmStaffDashboard = document.getElementById('hrmStaffDashboardSection');
      if (hrmStaffDashboard) {
        if (adminRole === 'HRM_ADMIN' || adminRole === 'SUPER_ADMIN') {
          hrmStaffDashboard.classList.remove('hidden');
          await loadHrmStaffStats();
        } else {
          hrmStaffDashboard.classList.add('hidden');
        }
      }

      // If SUPER_ADMIN, load formations for selection and Super Admin sections; otherwise just refresh
      if (adminRole === 'SUPER_ADMIN') {
        try {
          await Promise.allSettled([
            loadFormations().catch(err => console.error('Error loading formations:', err)),
            loadNationwideDashboard().catch(err => console.error('Error loading nationwide dashboard:', err)),
            loadFormationsTable().catch(err => console.error('Error loading formations table:', err)),
            loadAdminsTable().catch(err => console.error('Error loading admins table:', err)),
            loadRetentionPolicy().catch(err => console.error('Error loading retention policy:', err))
          ]);
        } catch (err) {
          console.error('Error loading SUPER_ADMIN data:', err);
        }
      }

      // Load dashboard data (with error handling to prevent stuck loading)
      try {
        await refreshAll();
      } catch (err) {
        console.error('Error refreshing dashboard:', err);
      }

      // Close loading spinner
      UI.closeLoading();
    } catch (err) {
      UI.closeLoading();
      console.error('Admin login error:', err);

      // Show detailed error message
      let errorMessage = 'Failed to authenticate. ';

      if (err.message) {
        errorMessage += err.message;
      } else if (err.reason) {
        errorMessage += `Reason: ${err.reason}`;
      } else {
        errorMessage += 'Please check your admin key and try again.';
      }

      // Check for common error scenarios
      if (err.message && err.message.includes('Network error')) {
        errorMessage = 'Cannot connect to server. Please check:\n' +
          '1. Your internet connection\n' +
          '2. The API URL is correct in api.js\n' +
          '3. The Apps Script is deployed as a Web App';
      } else if (err.message && err.message.includes('Invalid admin key')) {
        errorMessage = 'Invalid admin key. Please:\n' +
          '1. Check the admin key is correct\n' +
          '2. Ensure the admin exists in the Admins sheet\n' +
          '3. Verify the admin is active (active = TRUE)';
      }

      await UI.showError('Login Failed', errorMessage);
    }
  }

  async function adminAction(action) {
    if (!adminKey) return;
    try {
      UI.showLoading('Processing', 'Please wait...');
      const payload = { key: adminKey };
      // For formation-scoped actions, include effective formationId
      const formationId = getEffectiveFormationId();
      if (formationId) {
        payload.formationId = formationId;
      }
      const res = await Api.call(action, payload);
      UI.closeLoading();
      UI.setStatus(
        'statusArea',
        (res && res.message) || 'Action completed.',
        'success'
      );
    } catch (err) {
      UI.closeLoading();
      UI.setStatus(
        'statusArea',
        err.message || 'Admin action failed.',
        'error'
      );
    }
  }

  async function handleRegistrationToggle(e) {
    if (!adminKey) return;
    const enabled = e.target.checked;
    try {
      UI.showLoading('Updating', 'Changing registration mode...');
      await Api.call('setRegistrationStatus', { key: adminKey, enabled });
      UI.closeLoading();
    } catch (err) {
      UI.closeLoading();
      e.target.checked = !enabled;
      UI.setStatus(
        'statusArea',
        err.message || 'Failed to change registration mode.',
        'error'
      );
    }
  }

  async function loadEmployees() {
    if (!adminKey) return;
    const container = document.getElementById('employeesTable');
    try {
      const formationId = getEffectiveFormationId();

      // For SUPER_ADMIN, require formation to be selected
      if (adminRole === 'SUPER_ADMIN' && !formationId) {
        if (container) {
          container.textContent = 'Please select a formation from the dropdown above.';
        }
        return;
      }

      // For other roles, formationId should always be present
      if (!formationId) {
        console.warn('loadEmployees: No formationId available');
        return;
      }

      if (container) {
        container.innerHTML = '<div class="info info-muted">Loading employees...</div>';
      }
      UI.showLoading('Loading', 'Fetching employee list...');

      let res;
      if (adminRole === 'DEPARTMENT_ADMIN') {
        res = await Api.call('listEmployeesByDepartment', {
          key: adminKey,
          formationId,
          departmentId: adminDepartmentId,
        });
      } else {
        res = await Api.call('listEmployeesByFormation', {
          key: adminKey,
          formationId,
        });
      }
      UI.closeLoading();
      const employees = (res && res.data && res.data.employees) || [];
      if (!container) return;
      container.innerHTML = '';
      if (!employees.length) {
        container.textContent = 'No employees found.';
        return;
      }
      const table = document.createElement('table');
      const thead = document.createElement('thead');
      thead.innerHTML =
        '<tr><th>ID</th><th>Name</th><th>Device Bound</th><th>Actions</th></tr>';
      table.appendChild(thead);
      const tbody = document.createElement('tbody');
      employees.forEach((emp) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${emp.id}</td>
          <td>${emp.name}</td>
          <td>${emp.deviceBound ? 'Yes' : 'No'}</td>
          <td>
            <button class="btn btn-secondary btn-xs" data-empid="${emp.id}">
              Reset Device
            </button>
          </td>
        `;
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      container.appendChild(table);

      container.querySelectorAll('button[data-empid]').forEach((btn) => {
        btn.addEventListener('click', async () => {
          const empId = btn.getAttribute('data-empid');
          const confirmed = await UI.confirmAction(
            'Reset Device',
            `Are you sure you want to reset the device for employee "${empId}"? This action will be logged in the audit trail.`,
            'Yes, Reset Device',
            'Cancel'
          );
          if (!confirmed) {
            return;
          }
          try {
            UI.showLoading('Resetting Device', 'Please wait...');
            await Api.call('resetDeviceByFormation', {
              key: adminKey,
              formationId: getEffectiveFormationId(),
              employeeId: empId,
            });
            UI.closeLoading();
            await UI.showSuccess('Device Reset', `Device for employee "${empId}" has been reset successfully.`);
            loadEmployees();
          } catch (err) {
            UI.closeLoading();
            await UI.showError('Reset Failed', err.message || 'Failed to reset device.');
          }
        });
      });
    } catch (err) {
      UI.closeLoading();
      const container = document.getElementById('employeesTable');
      if (container) {
        container.textContent = err.message || 'Failed to load employees.';
      }
      UI.setStatus(
        'statusArea',
        err.message || 'Failed to load employees.',
        'error'
      );
    }
  }

  async function showAddEmployeeModal() {
    if (!adminKey) {
      await UI.showError('Error', 'Please log in first.');
      return;
    }

    const formationId = getEffectiveFormationId();
    if (!formationId) {
      await UI.showError('Error', 'Please select a formation first.');
      return;
    }

    if (!window.Swal) {
      await UI.showError('Unavailable', 'SweetAlert2 is not loaded.');
      return;
    }

    // Load departments for the formation
    let departments = [];
    try {
      UI.showLoading('Loading', 'Fetching departments...');
      const deptRes = await Api.call('listDepartments', {
        key: adminKey,
        formationId: formationId
      });
      UI.closeLoading();
      if (deptRes && deptRes.success && deptRes.data && deptRes.data.departments) {
        departments = deptRes.data.departments;
      }
    } catch (err) {
      UI.closeLoading();
      // Departments are optional, so we can continue
      console.log('Could not load departments:', err);
    }

    const result = await Swal.fire({
      title: 'Add Employee',
      html: `
        <p style="text-align: left; color: #666; font-size: 0.9rem; margin-bottom: 1rem;">
          <strong>Note:</strong> Employee ID will be automatically generated by the system.
        </p>
        <input id="swalEmployeeName" class="swal2-input" placeholder="Full Name *" required>
        <input id="swalEmployeeEmail" class="swal2-input" type="email" placeholder="Email (optional)">
        <select id="swalEmployeeDepartment" class="swal2-select" style="width: 100%; margin-bottom: 1rem;">
          <option value="">Select Department (optional)</option>
          ${departments.map(dept => `<option value="${dept.departmentId}">${dept.name || dept.departmentId}</option>`).join('')}
        </select>
        <label class="swal2-checkbox" style="display: flex; align-items: center; margin-top: 1rem;">
          <input type="checkbox" id="swalEmployeeActive" checked>
          <span style="margin-left: 0.5rem;">Active</span>
        </label>
      `,
      focusConfirm: false,
      showCancelButton: true,
      confirmButtonText: 'Create Employee',
      cancelButtonText: 'Cancel',
      confirmButtonColor: '#059669',
      preConfirm: async () => {
        const name = document.getElementById('swalEmployeeName').value.trim();
        const email = document.getElementById('swalEmployeeEmail').value.trim();
        const departmentId = document.getElementById('swalEmployeeDepartment').value.trim();
        const active = document.getElementById('swalEmployeeActive').checked;

        if (!name) {
          Swal.showValidationMessage('Name is required.');
          return false;
        }

        try {
          UI.showLoading('Creating', 'Creating employee...');
          const res = await Api.call('createEmployee', {
            key: adminKey,
            name: name,
            email: email || '',
            formationId: formationId,
            departmentId: departmentId || '',
            active: active
            // employeeId is not sent - will be auto-generated by backend
          });
          UI.closeLoading();

          if (res && res.success) {
            return {
              success: true,
              employeeId: res.data?.employeeId || 'N/A',
              name: name
            };
          } else {
            throw new Error(res.message || 'Failed to create employee.');
          }
        } catch (err) {
          UI.closeLoading();
          Swal.showValidationMessage(err.message || 'Failed to create employee.');
          return false;
        }
      }
    });

    if (result.isConfirmed && result.value && result.value.success) {
      const employeeId = result.value.employeeId;
      await UI.showSuccess(
        'Success',
        `Employee created successfully!<br><strong>Employee ID: ${employeeId}</strong>`
      );
      await loadEmployees();
    }
  }

  async function refreshAll() {
    // Only load formation-scoped data if formation is available
    // For SUPER_ADMIN, this requires a formation to be selected
    const formationId = getEffectiveFormationId();

    // Only load if we have a formationId, or if not SUPER_ADMIN (other roles always have formationId)
    if (formationId || adminRole !== 'SUPER_ADMIN') {
      try {
        await Promise.allSettled([
          loadEmployees().catch(err => console.error('Error loading employees:', err)),
          loadAttendanceLogs().catch(err => console.error('Error loading attendance logs:', err)),
          loadVisitorLogs().catch(err => console.error('Error loading visitor logs:', err)),
          loadHrmStats().catch(err => console.error('Error loading HRM stats:', err))
        ]);
      } catch (err) {
        console.error('Error in refreshAll:', err);
      }
    } else {
      // SUPER_ADMIN needs to select a formation first
      console.log('refreshAll: Waiting for formation selection for SUPER_ADMIN');
    }
  }

  function getEffectiveFormationId() {
    // SUPER_ADMIN can switch formations via dropdown
    if (adminRole === 'SUPER_ADMIN') {
      const select = document.getElementById('formationSelect');
      if (select && select.value) {
        return select.value;
      }
      return currentFormationId;
    }
    return adminFormationId;
  }

  /**
   * Load available modules for the current admin and filter the module selector
   */
  async function loadAvailableModules() {
    if (!adminKey || !adminRole) {
      console.error('Cannot load modules: admin not authenticated');
      return;
    }

    try {
      const formationId = adminFormationId || getEffectiveFormationId();
      if (!formationId && adminRole !== 'SUPER_ADMIN') {
        console.error('Cannot load modules: formationId required');
        return;
      }

      // For SUPER_ADMIN, get all modules (they can access everything)
      // For others, get modules available for their role and formation
      let modules = [];

      if (adminRole === 'SUPER_ADMIN') {
        // SUPER_ADMIN can see all modules
        modules = [
          { id: 'ATTENDANCE', name: 'Attendance', description: 'Track staff attendance and manage devices', icon: '📊' },
          { id: 'VISITORS', name: 'Visitors', description: 'Handle visitor sign-ins and approvals', icon: '👥' },
          { id: 'HRM', name: 'Staff Records', description: 'Manage staff information, leaves, and documents', icon: '👔' },
          { id: 'SYSTEM_ADMIN', name: 'System Settings', description: 'Manage formations, admins, and system configuration', icon: '⚙️' }
        ];
      } else {
        // For other roles, call the backend to get available modules
        try {
          UI.showLoading('Loading', 'Fetching available modules...');
          const res = await Api.call('getAvailableModules', {
            key: adminKey,
            formationId: formationId
          });
          UI.closeLoading();

          if (res && res.data && res.data.modules) {
            modules = res.data.modules;
          } else {
            console.warn('No modules returned from getAvailableModules');
            // Fallback: show basic modules based on role
            if (adminRole === 'HRM_ADMIN' || adminRole === 'HRM_VIEWER') {
              modules = [{ id: 'HRM', name: 'Staff Records', description: 'Manage staff information', icon: '👔' }];
            } else {
              modules = [
                { id: 'ATTENDANCE', name: 'Attendance', description: 'Track staff attendance', icon: '📊' },
                { id: 'VISITORS', name: 'Visitors', description: 'Handle visitor sign-ins', icon: '👥' }
              ];
            }
          }
        } catch (apiError) {
          console.error('Error fetching available modules:', apiError);
          // Fallback: show basic modules based on role
          if (adminRole === 'HRM_ADMIN' || adminRole === 'HRM_VIEWER') {
            modules = [{ id: 'HRM', name: 'Staff Records', description: 'Manage staff information', icon: '👔' }];
          } else {
            modules = [
              { id: 'ATTENDANCE', name: 'Attendance', description: 'Track staff attendance', icon: '📊' },
              { id: 'VISITORS', name: 'Visitors', description: 'Handle visitor sign-ins', icon: '👥' }
            ];
          }
        }
      }

      // Map module IDs to frontend module names
      const moduleMap = {
        'ATTENDANCE': 'attendance',
        'VISITORS': 'visitors',
        'HRM': 'hrm',
        'SYSTEM_ADMIN': 'system'
      };

      // Hide all module cards first
      const moduleCards = document.querySelectorAll('.module-card');
      moduleCards.forEach(card => {
        card.style.display = 'none';
      });

      // Show only modules the admin has access to
      modules.forEach(module => {
        const frontendModule = moduleMap[module.id];
        if (frontendModule) {
          const card = document.querySelector(`.module-card[data-module="${frontendModule}"]`);
          if (card) {
            card.style.display = 'block';
            // Update card content if module info is available
            if (module.name) {
              const titleEl = card.querySelector('.module-title');
              if (titleEl) titleEl.textContent = module.name;
            }
            if (module.description) {
              const descEl = card.querySelector('.module-description');
              if (descEl) descEl.textContent = module.description;
            }
            if (module.icon) {
              const iconEl = card.querySelector('.module-icon');
              if (iconEl) iconEl.textContent = module.icon;
            }
          }
        }
      });

      // If no modules available, show message
      const moduleSelector = document.querySelector('.module-selector');
      if (modules.length === 0 && moduleSelector) {
        moduleSelector.innerHTML = '<p style="text-align: center; color: var(--text-muted); padding: 2rem;">No modules available for your role.</p>';
      }
    } catch (err) {
      console.error('Error loading available modules:', err);
      // On error, show basic modules as fallback
      const moduleCards = document.querySelectorAll('.module-card');
      moduleCards.forEach(card => {
        const module = card.getAttribute('data-module');
        if (module === 'system' && adminRole !== 'SUPER_ADMIN') {
          card.style.display = 'none';
        } else {
          card.style.display = 'block';
        }
      });
    }
  }

  async function updateAdminContextUI() {
    const infoEl = document.getElementById('adminContextInfo');
    if (!infoEl) return;
    let text = '';
    if (adminRole === 'SUPER_ADMIN') {
      text = 'Role: SUPER ADMIN – can view all formations.';
    } else if (adminRole === 'FORMATION_ADMIN') {
      text = 'Role: FORMATION ADMIN – Formation: ' + (adminFormationId || '');
    } else if (adminRole === 'DEPARTMENT_ADMIN') {
      text =
        'Role: DEPARTMENT ADMIN – Formation: ' +
        (adminFormationId || '') +
        ', Department: ' +
        (adminDepartmentId || '');
    } else {
      text = 'Role: ' + (adminRole || 'UNKNOWN');
    }
    infoEl.textContent = text;

    const wrapper = document.getElementById('formationSelectorWrapper');
    if (!wrapper) return;
    if (adminRole === 'SUPER_ADMIN') {
      wrapper.classList.remove('hidden');
    } else {
      wrapper.classList.add('hidden');
    }

    // Show/hide Super Admin sections
    const superAdminSection = document.getElementById('superAdminSection');
    const formationSection = document.getElementById('formationManagementSection');
    const adminSection = document.getElementById('adminAssignmentSection');
    const retentionSection = document.getElementById('retentionSection');
    const hrmAdminSection = document.getElementById('hrmAdminManagementSection');
    const hrmAuditLogsSection = document.getElementById('hrmAuditLogsSection');
    const addEmployeeBtn = document.getElementById('addEmployeeBtn');

    if (adminRole === 'SUPER_ADMIN') {
      if (superAdminSection) superAdminSection.classList.remove('hidden');
      if (formationSection) formationSection.classList.remove('hidden');
      if (adminSection) adminSection.classList.remove('hidden');
      if (hrmAdminSection) hrmAdminSection.classList.remove('hidden');
      if (hrmAuditLogsSection) hrmAuditLogsSection.classList.remove('hidden');
      await loadHrmAdminsTable();
      await loadHrmAuditLogs();
      if (retentionSection) retentionSection.classList.remove('hidden');
      // Show Add Employee button for SUPER_ADMIN
      if (addEmployeeBtn) addEmployeeBtn.classList.remove('hidden');
    } else if (adminRole === 'HRM_ADMIN' || adminRole === 'FORMATION_ADMIN') {
      // Show Add Employee button for HRM_ADMIN and FORMATION_ADMIN
      if (addEmployeeBtn) addEmployeeBtn.classList.remove('hidden');
      if (superAdminSection) superAdminSection.classList.add('hidden');
      if (formationSection) formationSection.classList.add('hidden');
      if (adminSection) adminSection.classList.add('hidden');
      if (retentionSection) retentionSection.classList.add('hidden');
    } else {
      if (superAdminSection) superAdminSection.classList.add('hidden');
      if (formationSection) formationSection.classList.add('hidden');
      if (adminSection) adminSection.classList.add('hidden');
      if (retentionSection) retentionSection.classList.add('hidden');
      // Hide Add Employee button for other roles
      if (addEmployeeBtn) addEmployeeBtn.classList.add('hidden');
    }
  }

  async function loadFormations() {
    const select = document.getElementById('formationSelect');
    if (!select || !adminKey) return;
    try {
      UI.showLoading('Loading', 'Fetching formations...');
      const res = await Api.call('listFormations', { key: adminKey });
      UI.closeLoading();
      const formations = (res && res.data && res.data.formations) || [];
      select.innerHTML = '';
      formations.forEach((f) => {
        const opt = document.createElement('option');
        opt.value = f.formationId;
        opt.textContent = f.name || f.formationId;
        select.appendChild(opt);
      });
      if (formations.length) {
        currentFormationId = formations[0].formationId;
      }
      select.addEventListener('change', () => {
        currentFormationId = select.value;
        refreshAll();
      });
    } catch (err) {
      UI.closeLoading();
      // If formations cannot be loaded, fall back to adminFormationId if any
      UI.setStatus(
        'statusArea',
        err.message || 'Failed to load formations list.',
        'error'
      );
    }
  }

  async function loadAttendanceLogs() {
    const container = document.getElementById('attendanceLogs');
    if (!container || !adminKey) return;
    try {
      const formationId = getEffectiveFormationId();

      // For SUPER_ADMIN, require formation to be selected
      if (adminRole === 'SUPER_ADMIN' && !formationId) {
        container.textContent = 'Please select a formation from the dropdown above.';
        return;
      }

      // For other roles, formationId should always be present
      if (!formationId) {
        console.warn('loadAttendanceLogs: No formationId available');
        container.textContent = 'Formation information not available.';
        return;
      }

      container.innerHTML = '<div class="info info-muted">Loading attendance logs...</div>';
      UI.showLoading('Loading', 'Fetching attendance records...');

      const payload = {
        key: adminKey,
        formationId,
      };
      if (adminRole === 'DEPARTMENT_ADMIN') {
        payload.departmentId = adminDepartmentId;
      }
      const res = await Api.call('getAttendanceLogs', payload);
      UI.closeLoading();
      const logs = (res && res.data && res.data.logs) || [];
      container.innerHTML = '';
      if (!logs.length) {
        container.textContent = 'No attendance records found.';
        return;
      }
      const table = document.createElement('table');
      const thead = document.createElement('thead');
      thead.innerHTML =
        '<tr><th>Date</th><th>Employee</th><th>Checked In</th><th>Distance (m)</th><th>Accuracy (m)</th></tr>';
      table.appendChild(thead);
      const tbody = document.createElement('tbody');
      logs.forEach((row) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${row.date}</td>
          <td>${row.employeeName || 'N/A'}</td>
          <td>${row.checkInTimestamp}</td>
          <td>${Math.round(row.distanceMeters || 0)}</td>
          <td>${Math.round(row.accuracy || 0)}</td>
        `;
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      container.appendChild(table);
    } catch (err) {
      UI.closeLoading();
      container.textContent =
        err.message || 'Failed to load attendance logs for this formation.';
    }
  }

  async function loadVisitorLogs() {
    const container = document.getElementById('visitorLogs');
    if (!container || !adminKey) return;
    try {
      const formationId = getEffectiveFormationId();

      // For SUPER_ADMIN, require formation to be selected
      if (adminRole === 'SUPER_ADMIN' && !formationId) {
        container.textContent = 'Please select a formation from the dropdown above.';
        return;
      }

      // For other roles, formationId should always be present
      if (!formationId) {
        console.warn('loadVisitorLogs: No formationId available');
        container.textContent = 'Formation information not available.';
        return;
      }

      container.innerHTML = '<div class="info info-muted">Loading visitor logs...</div>';
      UI.showLoading('Loading', 'Fetching visitor records...');

      const payload = {
        key: adminKey,
        formationId,
      };
      if (adminRole === 'DEPARTMENT_ADMIN') {
        payload.departmentId = adminDepartmentId;
      }
      const res = await Api.call('getVisitorLogs', payload);
      UI.closeLoading();
      const logs = (res && res.data && res.data.logs) || [];
      container.innerHTML = '';
      if (!logs.length) {
        container.textContent = 'No visitor records found.';
        return;
      }
      const table = document.createElement('table');
      const thead = document.createElement('thead');
      thead.innerHTML =
        '<tr><th>Requested At</th><th>Visitor</th><th>Phone</th><th>Purpose</th><th>Staff</th><th>Status</th></tr>';
      table.appendChild(thead);
      const tbody = document.createElement('tbody');
      logs.forEach((row) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${row.requestedAt}</td>
          <td>${row.visitorName}</td>
          <td>${row.phone}</td>
          <td>${row.purpose}</td>
          <td>${row.staffToSee}</td>
          <td>${row.status}</td>
        `;
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      container.appendChild(table);
    } catch (err) {
      UI.closeLoading();
      container.textContent =
        err.message || 'Failed to load visitor logs for this formation.';
    }
  }

  // Super Admin functions
  async function loadNationwideDashboard() {
    if (!adminKey || adminRole !== 'SUPER_ADMIN') return;
    try {
      UI.showLoading('Loading', 'Fetching nationwide statistics...');
      const res = await Api.call('nationwideDashboard', { key: adminKey });
      UI.closeLoading();
      const data = res && res.data;
      if (data) {
        document.getElementById('statTotalFormations').textContent = data.totalFormations || 0;
        document.getElementById('statTotalEmployees').textContent = data.totalEmployees || 0;
        document.getElementById('statAttendanceToday').textContent = data.totalAttendanceToday || 0;
        document.getElementById('statPendingVisits').textContent = data.pendingVisits || 0;

        // Update formation table with staff counts if available
        if (data.formationSummaries && data.formationSummaries.length) {
          // Update the formations table to show staff counts
          const container = document.getElementById('formationsTable');
          if (container) {
            const rows = container.querySelectorAll('tbody tr');
            rows.forEach(row => {
              const formationId = row.querySelector('td').textContent;
              const summary = data.formationSummaries.find(s => s.formationId === formationId);
              if (summary) {
                // Add staff count to the row (if not already present)
                const cells = row.querySelectorAll('td');
                if (cells.length === 6) {
                  // Insert staff count before Actions column
                  const staffCell = document.createElement('td');
                  staffCell.textContent = summary.staffCount || 0;
                  row.insertBefore(staffCell, cells[5]);
                } else if (cells.length === 7) {
                  // Update existing staff count cell
                  cells[5].textContent = summary.staffCount || 0;
                }
              }
            });
          }
        }
      }
    } catch (err) {
      UI.closeLoading();
      console.error('Failed to load nationwide dashboard:', err);
    }
  }

  async function loadFormationsTable() {
    const container = document.getElementById('formationsTable');
    if (!container || !adminKey || adminRole !== 'SUPER_ADMIN') return;
    try {
      container.innerHTML = '<div class="info info-muted">Loading formations...</div>';
      UI.showLoading('Loading', 'Fetching formations...');
      const res = await Api.call('listFormations', { key: adminKey });
      UI.closeLoading();
      const formations = (res && res.data && res.data.formations) || [];
      container.innerHTML = '';
      if (!formations.length) {
        container.textContent = 'No formations found.';
        return;
      }
      const table = document.createElement('table');
      const thead = document.createElement('thead');
      thead.innerHTML = '<tr><th>ID</th><th>Name</th><th>Location</th><th>Radius (m)</th><th>Staff</th><th>Status</th><th>Actions</th></tr>';
      table.appendChild(thead);
      const tbody = document.createElement('tbody');
      formations.forEach((f) => {
        const tr = document.createElement('tr');
        const location = f.lat && f.lng ? `${f.lat}, ${f.lng}` : 'Not set';
        tr.innerHTML = `
          <td>${f.formationId}</td>
          <td>${f.name || ''}</td>
          <td>${location}</td>
          <td>${f.radiusMeters || ''}</td>
          <td>-</td>
          <td>${f.active !== false ? 'Active' : 'Inactive'}</td>
          <td>
            <button class="btn btn-secondary btn-xs" data-formation-id="${f.formationId}" data-action="edit">Edit</button>
            <button class="btn btn-secondary btn-xs" data-formation-id="${f.formationId}" data-action="deactivate">${f.active !== false ? 'Deactivate' : 'Activate'}</button>
          </td>
        `;
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      container.appendChild(table);

      // Add event listeners
      container.querySelectorAll('button[data-formation-id]').forEach((btn) => {
        btn.addEventListener('click', async () => {
          const formationId = btn.getAttribute('data-formation-id');
          const action = btn.getAttribute('data-action');
          if (action === 'edit') {
            showFormationModal(formationId);
          } else if (action === 'deactivate') {
            const formation = formations.find(f => f.formationId === formationId);
            const confirmMsg = formation && formation.active !== false
              ? `Deactivate formation "${formation.name || formationId}"?`
              : `Activate formation "${formation.name || formationId}"?`;
            const confirmed = await UI.confirmAction(
              'Formation Status',
              confirmMsg,
              'Yes, Continue',
              'Cancel'
            );
            if (!confirmed) return;
            try {
              UI.showLoading('Processing', 'Updating formation status...');
              await Api.call('deactivateFormation', {
                key: adminKey,
                formationId,
                active: formation && formation.active !== false ? false : true,
              });
              UI.closeLoading();
              await loadFormationsTable();
              await loadFormations(); // Refresh dropdown
            } catch (err) {
              UI.closeLoading();
              await UI.showError('Update Failed', err.message || 'Failed to update formation.');
            }
          }
        });
      });
    } catch (err) {
      UI.closeLoading();
      container.textContent = err.message || 'Failed to load formations.';
    }
  }

  async function loadAdminsTable() {
    const container = document.getElementById('adminsTable');
    if (!container || !adminKey || adminRole !== 'SUPER_ADMIN') return;
    try {
      container.innerHTML = '<div class="info info-muted">Loading admins...</div>';
      UI.showLoading('Loading', 'Fetching admin list...');
      const res = await Api.call('listAdmins', { key: adminKey });
      UI.closeLoading();
      if (!res || !res.data) {
        container.textContent = 'Failed to load admins. Invalid response from server.';
        return;
      }
      const admins = (res && res.data && res.data.admins) || [];
      container.innerHTML = '';
      if (!admins.length) {
        container.textContent = 'No admins found.';
        return;
      }
      const table = document.createElement('table');
      const thead = document.createElement('thead');
      thead.innerHTML = '<tr><th>S/N</th><th>Admin Key</th><th>Name</th><th>Role</th><th>Formation</th><th>Department</th><th>Status</th><th>Actions</th></tr>';
      table.appendChild(thead);
      const tbody = document.createElement('tbody');
      admins.forEach((admin, idx) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td>${idx + 1}</td>
          <td>${admin.adminKey || ''}</td>
          <td>${admin.name || '-'}</td>
          <td>${admin.role || ''}</td>
          <td>${admin.formationId || '-'}</td>
          <td>${admin.departmentId || '-'}</td>
          <td>${admin.active !== false ? 'Active' : 'Inactive'}</td>
          <td>
            <button class="btn btn-secondary btn-xs" data-admin-key="${admin.adminKey}" data-action="edit" title="Change role and assignment">Update Role</button>
            ${admin.active !== false
            ? `<button class="btn btn-danger btn-xs" data-admin-key="${admin.adminKey}" data-action="deactivate" title="Remove access">Delete</button>`
            : `<button class="btn btn-secondary btn-xs" data-admin-key="${admin.adminKey}" data-action="activate">Activate</button>`
          }
          </td>
        `;
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      container.appendChild(table);

      // Add event listeners
      container.querySelectorAll('button[data-admin-key]').forEach((btn) => {
        btn.addEventListener('click', async () => {
          const adminKeyValue = btn.getAttribute('data-admin-key');
          const action = btn.getAttribute('data-action');
          if (action === 'edit') {
            const admin = admins.find(a => a.adminKey === adminKeyValue);
            showAdminModal(admin);
          } else if (action === 'deactivate') {
            const admin = admins.find(a => a.adminKey === adminKeyValue);
            const confirmed = await UI.confirmAction(
              'Deactivate Admin',
              `Are you sure you want to deactivate admin "${admin.name || admin.adminKey}" (${admin.role})? This action will be logged.`,
              'Yes, Deactivate',
              'Cancel'
            );
            if (!confirmed) return;
            try {
              UI.showLoading('Processing', 'Deactivating admin...');
              await Api.call('deactivateAdmin', {
                key: adminKey,
                targetKey: adminKeyValue,
              });
              UI.closeLoading();
              await loadAdminsTable();
              await UI.showSuccess('Admin Deactivated', 'Admin has been deactivated successfully.');
            } catch (err) {
              UI.closeLoading();
              await UI.showError('Deactivation Failed', err.message || 'Failed to deactivate admin.');
            }
          } else if (action === 'activate') {
            const admin = admins.find(a => a.adminKey === adminKeyValue);
            const confirmed = await UI.confirmAction(
              'Activate Admin',
              `Are you sure you want to activate admin "${admin.name || admin.adminKey}" (${admin.role})?`,
              'Yes, Activate',
              'Cancel'
            );
            if (!confirmed) return;
            try {
              UI.showLoading('Processing', 'Activating admin...');
              await Api.call('activateAdmin', {
                key: adminKey,
                targetKey: adminKeyValue,
              });
              UI.closeLoading();
              await loadAdminsTable();
              await UI.showSuccess('Admin Activated', 'Admin has been activated successfully.');
            } catch (err) {
              UI.closeLoading();
              await UI.showError('Activation Failed', err.message || 'Failed to activate admin.');
            }
          }
        });
      });
    } catch (err) {
      UI.closeLoading();
      container.textContent = err.message || 'Failed to load admins.';
    }
  }

  async function showFormationModal(formationId = null) {
    if (!window.Swal) {
      return UI.showError('Unavailable', 'SweetAlert2 is not loaded.');
    }

    const isEdit = !!formationId;
    let formation = null;

    if (isEdit) {
      try {
        UI.showLoading('Loading', 'Fetching formation details...');
        const res = await Api.call('listFormations', { key: adminKey });
        UI.closeLoading();
        const formations = (res && res.data && res.data.formations) || [];
        formation = formations.find(f => f.formationId === formationId) || null;
      } catch (err) {
        UI.closeLoading();
        console.error('Failed to load formation:', err);
      }
    }

    const result = await Swal.fire({
      title: isEdit ? 'Edit Formation' : 'Create Formation',
      html: `
        <input id="swalFormationId" class="swal2-input" placeholder="Formation ID *" ${isEdit ? 'readonly' : ''} value="${formation ? (formation.formationId || '') : (formationId || '')}">
        <input id="swalFormationName" class="swal2-input" placeholder="Formation Name *" value="${formation ? (formation.name || '') : ''}">
        <input id="swalFormationLat" class="swal2-input" type="number" step="any" placeholder="Latitude" value="${formation && formation.lat ? formation.lat : ''}">
        <input id="swalFormationLng" class="swal2-input" type="number" step="any" placeholder="Longitude" value="${formation && formation.lng ? formation.lng : ''}">
        <input id="swalFormationRadius" class="swal2-input" type="number" min="0" placeholder="Radius (meters)" value="${formation && formation.radiusMeters ? formation.radiusMeters : ''}">
      `,
      focusConfirm: false,
      showCancelButton: true,
      confirmButtonText: isEdit ? 'Update' : 'Create',
      cancelButtonText: 'Cancel',
      confirmButtonColor: '#059669',
      preConfirm: async () => {
        const id = document.getElementById('swalFormationId').value.trim();
        const name = document.getElementById('swalFormationName').value.trim();
        const lat = document.getElementById('swalFormationLat').value.trim();
        const lng = document.getElementById('swalFormationLng').value.trim();
        const radius = document.getElementById('swalFormationRadius').value.trim();

        if (!id || !name) {
          Swal.showValidationMessage('Formation ID and Name are required.');
          return false;
        }

        const payload = { key: adminKey, formationId: id, name };
        if (lat) payload.lat = lat;
        if (lng) payload.lng = lng;
        if (radius) payload.radiusMeters = radius;

        try {
          UI.showLoading('Saving', isEdit ? 'Updating formation...' : 'Creating formation...');
          if (isEdit) {
            await Api.call('updateFormation', payload);
          } else {
            await Api.call('createFormation', payload);
          }
          UI.closeLoading();
          return true;
        } catch (err) {
          UI.closeLoading();
          Swal.showValidationMessage(err.message || 'Failed to save formation.');
          return false;
        }
      },
    });

    if (result.isConfirmed) {
      await loadFormationsTable();
      await loadFormations();
      await UI.showSuccess('Success', `Formation ${isEdit ? 'updated' : 'created'} successfully.`);
    }
  }

  async function showAdminModal(admin = null) {
    if (!window.Swal) {
      return UI.showError('Unavailable', 'SweetAlert2 is not loaded.');
    }

    const isEdit = admin !== null;

    const result = await Swal.fire({
      title: isEdit ? 'Edit Admin' : 'Assign Admin',
      html: `
        <input id="swalAdminKey" class="swal2-input" placeholder="Admin Key *" ${isEdit ? 'readonly' : ''} value="${isEdit ? (admin.adminKey || '') : ''}">
        <input id="swalAdminName" class="swal2-input" placeholder="Name (optional)" value="${isEdit ? (admin.name || '') : ''}">
        <select id="swalAdminRole" class="swal2-select">
          <option value="">Select Role *</option>
              <option value="FORMATION_ADMIN"${isEdit && admin.role === 'FORMATION_ADMIN' ? ' selected' : ''}>Formation Admin</option>
              <option value="DEPARTMENT_ADMIN"${isEdit && admin.role === 'DEPARTMENT_ADMIN' ? ' selected' : ''}>Department Admin</option>
              <option value="EMPLOYEE"${isEdit && admin.role === 'EMPLOYEE' ? ' selected' : ''}>Employee</option>
              ${adminRole === 'SUPER_ADMIN' ? `
              <option value="HRM_ADMIN"${isEdit && admin.role === 'HRM_ADMIN' ? ' selected' : ''}>HRM Admin</option>
              <option value="HRM_VIEWER"${isEdit && admin.role === 'HRM_VIEWER' ? ' selected' : ''}>HRM Viewer</option>
              ` : ''}
        </select>
        <input id="swalAdminFormationId" class="swal2-input" placeholder="Formation ID *" value="${isEdit ? (admin.formationId || '') : ''}">
        <input id="swalAdminDepartmentId" class="swal2-input" placeholder="Department ID (for Dept Admin)" value="${isEdit ? (admin.departmentId || '') : ''}">
      `,
      focusConfirm: false,
      showCancelButton: true,
      confirmButtonText: isEdit ? 'Update' : 'Create',
      cancelButtonText: 'Cancel',
      confirmButtonColor: '#059669',
      preConfirm: async () => {
        const keyVal = document.getElementById('swalAdminKey').value.trim();
        const nameVal = document.getElementById('swalAdminName').value.trim();
        const roleVal = document.getElementById('swalAdminRole').value;
        const formationVal = document.getElementById('swalAdminFormationId').value.trim();
        const deptVal = document.getElementById('swalAdminDepartmentId').value.trim();

        if (!keyVal || !roleVal) {
          Swal.showValidationMessage('Admin Key and Role are required.');
          return false;
        }

        if (roleVal === 'DEPARTMENT_ADMIN' && (!formationVal || !deptVal)) {
          Swal.showValidationMessage('Formation ID and Department ID are required for Department Admin.');
          return false;
        }

        if ((roleVal === 'FORMATION_ADMIN' || roleVal === 'EMPLOYEE') && !formationVal) {
          Swal.showValidationMessage('Formation ID is required for ' + roleVal + '.');
          return false;
        }

        try {
          UI.showLoading('Saving', isEdit ? 'Updating admin...' : 'Creating admin...');
          if (isEdit) {
            const payload = {
              key: adminKey,
              targetKey: keyVal,
              role: roleVal,
              formationId: formationVal,
              departmentId: deptVal || undefined,
              name: nameVal || undefined,
            };
            await Api.call('updateAdminRole', payload);
          } else {
            const payload = {
              key: adminKey,
              newKey: keyVal,
              role: roleVal,
              formationId: formationVal,
              departmentId: deptVal || undefined,
              name: nameVal || undefined,
            };
            await Api.call('createAdmin', payload);
          }
          UI.closeLoading();
          return true;
        } catch (err) {
          UI.closeLoading();
          Swal.showValidationMessage(err.message || `Failed to ${isEdit ? 'update' : 'create'} admin.`);
          return false;
        }
      },
    });

    if (result.isConfirmed) {
      await loadAdminsTable();
      await UI.showSuccess(
        'Success',
        `Admin ${isEdit ? 'updated' : 'created'} successfully.`
      );
    }
  }

  // Formation Setup Wizard
  async function showFormationWizard() {
    if (adminRole !== 'SUPER_ADMIN') {
      await UI.showError('Access Denied', 'Only Super Admin can access the Formation Setup Wizard.');
      return;
    }
    if (!window.Swal) {
      await UI.showError('Unavailable', 'SweetAlert2 is not loaded.');
      return;
    }

    const wizardData = {
      type: '',
      formationId: '',
      name: '',
      address: '',
      lat: '',
      lng: '',
      radiusMeters: '',
      departments: [],
      adminKey: '',
      adminName: '',
    };

    const steps = ['1', '2', '3', '4', '5'];
    const Wizard = Swal.mixin({
      confirmButtonColor: '#059669',
      cancelButtonColor: '#6b7280',
      progressSteps: steps,
      showCancelButton: true,
      allowOutsideClick: false,
      allowEscapeKey: false,
      reverseButtons: true,
    });

    let currentStep = 0;

    // STEP 1: Type & ID
    const step1 = await Wizard.fire({
      title: 'Step 1 – Formation Type & ID',
      currentProgressStep: currentStep.toString(),
      html: `
        <select id="wizType" class="swal2-select">
          <option value="">Select type...</option>
          <option value="HQ">HQ</option>
          <option value="STATE_SECRETARIAT">State Secretariat</option>
          <option value="ORIENTATION_CAMP">Orientation Camp</option>
          <option value="LOCAL_GOVT">Local Govt Office</option>
        </select>
        <input type="text" id="wizFormationId" class="swal2-input" placeholder="Formation ID (e.g. FCT_HQ)">
      `,
      preConfirm: () => {
        const type = document.getElementById('wizType').value;
        const id = document.getElementById('wizFormationId').value.trim();
        if (!type) {
          Swal.showValidationMessage('Please select a formation type.');
          return false;
        }
        if (!id) {
          Swal.showValidationMessage('Formation ID is required.');
          return false;
        }
        if (!/^[A-Z0-9_]+$/.test(id)) {
          Swal.showValidationMessage('Formation ID must contain only uppercase letters, numbers, and underscores.');
          return false;
        }
        wizardData.type = type;
        wizardData.formationId = id;
        return true;
      },
    });
    if (!step1.isConfirmed) return;
    currentStep++;

    // STEP 2: Basic details
    const step2 = await Wizard.fire({
      title: 'Step 2 – Formation Details',
      currentProgressStep: currentStep.toString(),
      html: `
        <input type="text" id="wizName" class="swal2-input" placeholder="Formation Name *">
        <textarea id="wizAddress" class="swal2-textarea" placeholder="Address / Description"></textarea>
      `,
      preConfirm: () => {
        const name = document.getElementById('wizName').value.trim();
        const address = document.getElementById('wizAddress').value.trim();
        if (!name) {
          Swal.showValidationMessage('Formation name is required.');
          return false;
        }
        wizardData.name = name;
        wizardData.address = address;
        return true;
      },
    });
    if (!step2.isConfirmed) return;
    currentStep++;

    // STEP 3: Location
    const step3 = await Wizard.fire({
      title: 'Step 3 – GPS Location & Radius',
      currentProgressStep: currentStep.toString(),
      html: `
        <div style="display:flex; gap:0.5rem;">
          <input type="number" id="wizLat" class="swal2-input" step="any" placeholder="Latitude *" style="flex:1;">
          <input type="number" id="wizLng" class="swal2-input" step="any" placeholder="Longitude *" style="flex:1;">
        </div>
        <input type="number" id="wizRadius" class="swal2-input" min="10" max="1000" placeholder="Attendance Radius (meters) *" value="100">
        <button type="button" id="wizCaptureLocation" class="swal2-confirm" style="margin-top:0.5rem; background:#10b981;">Use Current Location</button>
      `,
      didOpen: () => {
        const btn = document.getElementById('wizCaptureLocation');
        if (btn) {
          btn.addEventListener('click', async () => {
            try {
              await UI.showLoading('Getting Location', 'Please allow location access...');
              const loc = await Geo.getLocation({ requiredAccuracyMeters: 50 });
              UI.closeLoading();
              document.getElementById('wizLat').value = loc.latitude;
              document.getElementById('wizLng').value = loc.longitude;
              await UI.showSuccess('Location Captured', 'Coordinates filled.');
            } catch (err) {
              UI.closeLoading();
              await UI.showError('Location Error', err.message || 'Failed to get location.');
            }
          });
        }
      },
      preConfirm: () => {
        const lat = parseFloat(document.getElementById('wizLat').value);
        const lng = parseFloat(document.getElementById('wizLng').value);
        const radius = parseInt(document.getElementById('wizRadius').value, 10);

        if (isNaN(lat) || lat < -90 || lat > 90) {
          Swal.showValidationMessage('Valid latitude is required (-90 to 90).');
          return false;
        }
        if (isNaN(lng) || lng < -180 || lng > 180) {
          Swal.showValidationMessage('Valid longitude is required (-180 to 180).');
          return false;
        }
        if (isNaN(radius) || radius < 10 || radius > 1000) {
          Swal.showValidationMessage('Attendance radius must be between 10 and 1000 meters.');
          return false;
        }

        wizardData.lat = lat;
        wizardData.lng = lng;
        wizardData.radiusMeters = radius;
        return true;
      },
    });
    if (!step3.isConfirmed) return;
    currentStep++;

    // STEP 4: Departments
    const step4 = await Wizard.fire({
      title: 'Step 4 – Departments',
      currentProgressStep: currentStep.toString(),
      html: `
        <div id="wizDeptContainer" style="text-align:left;">
          <div class="department-item" style="display:flex; gap:0.5rem; margin-bottom:0.5rem;">
            <input type="text" class="swal2-input wizDeptId" placeholder="Dept ID (e.g. HR)" style="flex:1;">
            <input type="text" class="swal2-input wizDeptName" placeholder="Dept Name (e.g. Human Resources)" style="flex:2;">
          </div>
        </div>
        <button type="button" id="wizAddDept" class="swal2-confirm" style="margin-top:0.5rem; background:#10b981;">+ Add Department</button>
      `,
      didOpen: () => {
        const addBtn = document.getElementById('wizAddDept');
        const container = document.getElementById('wizDeptContainer');
        if (addBtn && container) {
          addBtn.addEventListener('click', () => {
            const wrapper = document.createElement('div');
            wrapper.className = 'department-item';
            wrapper.style = 'display:flex; gap:0.5rem; margin-bottom:0.5rem;';
            wrapper.innerHTML = `
              <input type="text" class="swal2-input wizDeptId" placeholder="Dept ID" style="flex:1;">
              <input type="text" class="swal2-input wizDeptName" placeholder="Dept Name" style="flex:2;">
            `;
            container.appendChild(wrapper);
          });
        }
      },
      preConfirm: () => {
        const ids = Array.from(document.querySelectorAll('.wizDeptId')).map(i => i.value.trim());
        const names = Array.from(document.querySelectorAll('.wizDeptName')).map(i => i.value.trim());

        const departments = [];
        for (let i = 0; i < ids.length; i++) {
          const id = ids[i];
          const name = names[i];
          if (!id && !name) continue;
          if (!id || !name) {
            Swal.showValidationMessage('Both Department ID and Name are required if one is provided.');
            return false;
          }
          if (!/^[A-Z0-9_]+$/.test(id)) {
            Swal.showValidationMessage(`Department ID "${id}" must contain only uppercase letters, numbers, and underscores.`);
            return false;
          }
          departments.push({ departmentId: id, name });
        }

        if (!departments.length) {
          Swal.showValidationMessage('Please add at least one department.');
          return false;
        }

        wizardData.departments = departments;
        return true;
      },
    });
    if (!step4.isConfirmed) return;
    currentStep++;

    // STEP 5: Assign Formation Admin
    const step5 = await Wizard.fire({
      title: 'Step 5 – Assign Formation Admin',
      currentProgressStep: currentStep.toString(),
      html: `
        <input type="text" id="wizAdminKey" class="swal2-input" placeholder="Admin Key *">
        <input type="text" id="wizAdminName" class="swal2-input" placeholder="Admin Name (optional)">
      `,
      preConfirm: () => {
        const adminKey = document.getElementById('wizAdminKey').value.trim();
        const adminName = document.getElementById('wizAdminName').value.trim();
        if (!adminKey) {
          Swal.showValidationMessage('Admin key is required.');
          return false;
        }
        wizardData.adminKey = adminKey;
        wizardData.adminName = adminName;
        return true;
      },
    });
    if (!step5.isConfirmed) return;

    // Final: Create formation + admin via backend
    try {
      await UI.showLoading('Creating Formation', 'Finalizing setup. Please wait...');

      const createFormationPayload = {
        key: adminKey,
        formationId: wizardData.formationId,
        name: wizardData.name,
        type: wizardData.type,
        address: wizardData.address || undefined,
        lat: wizardData.lat,
        lng: wizardData.lng,
        radiusMeters: wizardData.radiusMeters,
        departments: wizardData.departments,
      };
      const formationRes = await Api.call('createFormation', createFormationPayload);
      const newFormationId = (formationRes && formationRes.data && formationRes.data.formationId) || wizardData.formationId;

      await Api.call('createAdmin', {
        key: adminKey,
        newKey: wizardData.adminKey,
        role: 'FORMATION_ADMIN',
        formationId: newFormationId,
        name: wizardData.adminName || undefined,
      });

      UI.closeLoading();
      await loadFormationsTable();
      await loadFormations();
      await loadAdminsTable();

      await UI.showSuccess('Formation Created', 'Formation and Formation Admin have been created successfully!');
    } catch (err) {
      UI.closeLoading();
      await UI.showError('Setup Failed', err.message || 'Failed to complete formation setup. Please try again.');
    }
  }

  // Data Retention & Archival
  async function loadRetentionPolicy() {
    if (!adminKey || adminRole !== 'SUPER_ADMIN') return;
    try {
      UI.showLoading('Loading', 'Fetching retention policy...');
      const res = await Api.call('getRetentionPolicy', { key: adminKey });
      UI.closeLoading();
      if (!res || !res.data) {
        console.error('loadRetentionPolicy: Invalid response', res);
        return;
      }
      const policy = res && res.data && res.data.policy;
      if (policy) {
        const container = document.getElementById('retentionPolicy');
        if (container) {
          const attendanceCutoff = policy.attendanceCutoff
            ? new Date(policy.attendanceCutoff).toLocaleDateString()
            : 'Indefinite';
          const visitorsCutoff = policy.visitorsCutoff
            ? new Date(policy.visitorsCutoff).toLocaleDateString()
            : 'Indefinite';

          container.innerHTML = `
            <strong>Retention Policy:</strong><br/>
            • Attendance Records: ${policy.attendanceYears} years (cutoff: ${attendanceCutoff})<br/>
            • Visitor Records: ${policy.visitorsYears} years (cutoff: ${visitorsCutoff})<br/>
            • Audit Logs: Indefinite (never archived)<br/>
            • Archival: ${policy.archiveEnabled ? 'Enabled' : 'Disabled'}
          `;
        }
      }
    } catch (err) {
      UI.closeLoading();
      console.error('Failed to load retention policy:', err);
    }
  }

  /**
   * Load system-wide activity/audit logs for Super Admin.
   * Shows who did what and when across Attendance, Visitors, HRM, and Admin management.
   */
  async function loadSystemAuditLogs(filters = {}) {
    const container = document.getElementById('systemAuditLogsTable');
    if (!container || !adminKey || adminRole !== 'SUPER_ADMIN') return;
    try {
      container.innerHTML = '<div class="info info-muted">Loading activity logs...</div>';
      UI.showLoading('Loading', 'Fetching activity logs...');
      const payload = {
        key: adminKey,
        limit: filters.limit || 300,
        formationId: filters.formationId || undefined,
        actionType: filters.actionType || undefined,
        actor: filters.actor || undefined,
        status: filters.status || undefined,
        startDate: filters.startDate || undefined,
        endDate: filters.endDate || undefined,
      };
      const res = await Api.call('getAuditLogs', payload);
      UI.closeLoading();
      if (!res) {
        container.textContent = 'Failed to load activity logs. No response from server.';
        return;
      }
      if (!res.success) {
        container.textContent = res.message || 'Failed to load activity logs. Invalid response from server.';
        return;
      }
      const logs = (res.data && res.data.logs) ? res.data.logs : [];
      container.innerHTML = '';
      if (!logs.length) {
        container.textContent = 'No activity logs found for the selected filters. Use Apply Filters to change the date range or other filters.';
        return;
      }
      const table = document.createElement('table');
      table.className = 'data-table';
      const thead = document.createElement('thead');
      thead.innerHTML = '<tr><th>When</th><th>Who (Actor)</th><th>Role</th><th>Action</th><th>Formation</th><th>Status</th><th>Details</th></tr>';
      table.appendChild(thead);
      const tbody = document.createElement('tbody');
      logs.forEach((log) => {
        const tr = document.createElement('tr');
        const when = log.timestamp ? (typeof log.timestamp === 'string' ? new Date(log.timestamp) : log.timestamp) : null;
        const whenStr = when && !isNaN(when.getTime())
          ? when.toLocaleString(undefined, { dateStyle: 'short', timeStyle: 'short' })
          : (log.timestamp || '-');
        let details = '-';
        if (log.extra && typeof log.extra === 'object') {
          const parts = [];
          if (log.extra.extra) parts.push(String(log.extra.extra));
          if (log.extra.targetId) parts.push('Target: ' + log.extra.targetId);
          if (log.extra.targetRole) parts.push('Role: ' + log.extra.targetRole);
          if (log.extra.oldValue !== undefined) parts.push('From: ' + log.extra.oldValue);
          if (log.extra.newValue !== undefined) parts.push('To: ' + log.extra.newValue);
          if (parts.length) details = parts.join(' | ');
        } else if (log.reason) {
          details = log.reason;
        }
        tr.innerHTML = `
          <td>${whenStr}</td>
          <td>${escapeHtml(String(log.actor || '-'))}</td>
          <td>${escapeHtml(String(log.actorRole || '-'))}</td>
          <td>${escapeHtml(String(log.actionType || '-'))}</td>
          <td>${escapeHtml(String(log.formationId || '-'))}</td>
          <td><span class="badge ${log.status === 'SUCCESS' ? 'badge-success' : log.status === 'FAILED' || log.status === 'BLOCKED' ? 'badge-danger' : 'badge-warning'}">${escapeHtml(String(log.status || '-'))}</span></td>
          <td style="max-width: 280px; word-break: break-word;" title="${escapeHtml(details)}">${escapeHtml(details.length > 60 ? details.substring(0, 60) + '…' : details)}</td>
        `;
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      container.appendChild(table);
    } catch (err) {
      UI.closeLoading();
      container.textContent = err.message || 'Failed to load activity logs.';
    }
  }

  function escapeHtml(str) {
    if (!str) return '';
    const div = document.createElement('div');
    div.textContent = str;
    return div.innerHTML;
  }

  async function runArchival() {
    if (!adminKey || adminRole !== 'SUPER_ADMIN') return;

    const statusEl = document.getElementById('archivalStatus');
    if (!statusEl) return;

    const confirmed = await UI.confirmAction(
      'Run Archival',
      'This will move old records to archive sheets based on retention policy. This process may take a few minutes. Continue?',
      'Yes, Run Archival',
      'Cancel'
    );
    if (!confirmed) {
      return;
    }

    UI.setStatus('archivalStatus', 'Running archival...', 'info');

    try {
      const res = await Api.call('runArchival', {
        key: adminKey,
        dataType: 'all', // 'attendance', 'visitors', or 'all'
      });

      const summary = res && res.data && res.data.summary;
      if (summary) {
        const attendanceCount = summary.attendance.archived || 0;
        const visitorsCount = summary.visitors.archived || 0;
        const message = `Archival completed: ${attendanceCount} attendance records, ${visitorsCount} visitor records archived.`;
        UI.setStatus('archivalStatus', message, 'success');
      } else {
        UI.setStatus('archivalStatus', res.message || 'Archival completed.', 'success');
      }

      // Refresh retention policy to show updated cutoffs
      await loadRetentionPolicy();
    } catch (err) {
      UI.closeLoading();
      await UI.showError('Archival Failed', err.message || 'Failed to run archival.');
    }
  }

  async function setupArchivalTriggers() {
    if (!adminKey || adminRole !== 'SUPER_ADMIN') return;

    const confirmed = await UI.confirmAction(
      'Setup Auto-Archival',
      'This will create a weekly trigger to automatically archive old records based on retention policy. Continue?',
      'Yes, Setup Triggers',
      'Cancel'
    );
    if (!confirmed) {
      return;
    }

    UI.showLoading('Setting Up Triggers', 'Please wait...');

    try {
      const res = await Api.call('setupArchivalTriggers', {
        key: adminKey,
      });

      UI.closeLoading();
      if (res && res.success) {
        await UI.showSuccess(
          'Triggers Setup Complete',
          'Archival triggers set up successfully! The system will automatically archive old records weekly.'
        );
      } else {
        await UI.showInfo('Trigger Setup', res.message || 'Trigger setup completed.');
      }
    } catch (err) {
      UI.closeLoading();
      await UI.showError(
        'Trigger Setup Failed',
        err.message || 'Failed to set up triggers. You may need to set them up manually in Apps Script.'
      );
    }
  }

  // ============================================================================
  // HRM MODULE FUNCTIONS
  // ============================================================================

  async function loadHrmStats() {
    if (!adminKey || !currentFormationId) return;

    try {
      UI.showLoading('Loading', 'Fetching HRM statistics...');
      const res = await Api.call('getHrmStats', {
        key: adminKey,
        formationId: currentFormationId,
      });
      UI.closeLoading();

      if (res && res.success && res.data && res.data.stats) {
        const stats = res.data.stats;
        document.getElementById('hrmTotalEmployees').textContent = stats.totalEmployees || 0;
        document.getElementById('hrmPendingLeaves').textContent = stats.pendingLeaveRequests || 0;
        document.getElementById('hrmPendingTransfers').textContent = stats.pendingTransfers || 0;
        document.getElementById('hrmOnLeave').textContent = stats.onLeave || 0;
      }
    } catch (err) {
      UI.closeLoading();
      console.error('Failed to load HRM stats:', err);
    }
  }

  async function loadHrmProfiles() {
    if (!adminKey || !currentFormationId) return;

    const container = document.getElementById('hrmProfilesTable');
    if (!container) return;

    container.innerHTML = '<div class="info info-muted">Loading profiles...</div>';
    UI.showLoading('Loading', 'Fetching employee profiles...');

    try {
      // Get employees list
      const empRes = await Api.call('listEmployeesByFormation', {
        key: adminKey,
        formationId: currentFormationId,
      });
      UI.closeLoading();

      if (!empRes || !empRes.success || !empRes.data || !empRes.data.employees) {
        container.textContent = 'No employees found.';
        return;
      }

      const employees = empRes.data.employees;
      let html = '<table><thead><tr><th>S/N</th><th>Name</th><th>Department</th><th>Status</th><th>Actions</th></tr></thead><tbody>';

      for (let idx = 0; idx < employees.length; idx++) {
        const emp = employees[idx];
        html += `<tr>
          <td>${idx + 1}</td>
          <td>${emp.name || ''}</td>
          <td>${emp.departmentName || emp.departmentId || ''}</td>
          <td>${emp.status || 'ACTIVE'}</td>
          <td><button class="btn btn-xs btn-secondary" onclick="AdminPage.viewEmployeeProfile('${emp.employeeId}')">View Profile</button></td>
        </tr>`;
      }

      html += '</tbody></table>';
      container.innerHTML = html;
    } catch (err) {
      UI.closeLoading();
      container.textContent = err.message || 'Failed to load profiles.';
    }
  }

  async function loadHrmLeaves() {
    if (!adminKey || !currentFormationId) return;

    const container = document.getElementById('hrmLeavesTable');
    if (!container) return;

    container.innerHTML = '<div class="info info-muted">Loading leave requests...</div>';
    UI.showLoading('Loading', 'Fetching leave requests...');

    try {
      const res = await Api.call('listLeaveRequests', {
        key: adminKey,
        formationId: currentFormationId,
        status: 'PENDING', // Show pending by default
      });
      UI.closeLoading();

      if (!res || !res.success || !res.data || !res.data.requests) {
        container.textContent = 'No pending leave requests.';
        return;
      }

      const requests = res.data.requests;
      let html = '<table><thead><tr><th>S/N</th><th>Employee Name</th><th>Leave Type</th><th>Start Date</th><th>End Date</th><th>Reason</th><th>Actions</th></tr></thead><tbody>';

      for (let idx = 0; idx < requests.length; idx++) {
        const req = requests[idx];
        html += `<tr>
          <td>${idx + 1}</td>
          <td>${req.employeeName || req.employeeId || ''}</td>
          <td>${req.leaveType || ''}</td>
          <td>${req.startDate || ''}</td>
          <td>${req.endDate || ''}</td>
          <td>${req.reason || ''}</td>
          <td>
            <button class="btn btn-xs btn-primary" onclick="AdminPage.approveLeave('${req.requestId}')">Approve</button>
            <button class="btn btn-xs btn-secondary" onclick="AdminPage.rejectLeave('${req.requestId}')">Reject</button>
          </td>
        </tr>`;
      }

      html += '</tbody></table>';
      container.innerHTML = html;
    } catch (err) {
      UI.closeLoading();
      container.textContent = err.message || 'Failed to load leave requests.';
    }
  }

  async function loadHrmPerformance() {
    if (!adminKey || !currentFormationId) return;

    const container = document.getElementById('hrmPerformanceTable');
    if (!container) return;

    container.innerHTML = '<div class="info info-muted">Loading performance reviews...</div>';
    UI.showLoading('Loading', 'Fetching performance reviews...');

    try {
      const res = await Api.call('listPerformanceReviews', {
        key: adminKey,
        formationId: currentFormationId,
      });
      UI.closeLoading();

      if (!res || !res.success || !res.data || !res.data.reviews) {
        container.textContent = 'No performance reviews found.';
        return;
      }

      const reviews = res.data.reviews;
      let html = '<table><thead><tr><th>S/N</th><th>Employee Name</th><th>Review Period</th><th>Rating</th><th>Comments</th><th>Reviewed By</th><th>Date</th></tr></thead><tbody>';

      for (let idx = 0; idx < reviews.length; idx++) {
        const review = reviews[idx];
        html += `<tr>
          <td>${idx + 1}</td>
          <td>${review.employeeName || review.employeeId || ''}</td>
          <td>${review.reviewPeriod || ''}</td>
          <td>${review.rating || ''}/5</td>
          <td>${review.comments || ''}</td>
          <td>${review.reviewedBy || ''}</td>
          <td>${review.createdAt ? new Date(review.createdAt).toLocaleDateString() : ''}</td>
        </tr>`;
      }

      html += '</tbody></table>';
      container.innerHTML = html;
    } catch (err) {
      UI.closeLoading();
      container.textContent = err.message || 'Failed to load performance reviews.';
    }
  }

  async function loadHrmTransfers() {
    if (!adminKey || !currentFormationId) return;

    const container = document.getElementById('hrmTransfersTable');
    if (!container) return;

    container.textContent = 'Loading transfer requests...';

    try {
      // Note: We'll need to add a listTransfers action, or reuse existing data
      container.textContent = 'Transfer management coming soon.';
    } catch (err) {
      container.textContent = err.message || 'Failed to load transfers.';
    }
  }

  /**
   * Client-side file validation
   * Validates file type, size, and count before upload
   */
  function validateFileUpload(file, currentFileCount = 0) {
    const MAX_FILE_SIZE = 2 * 1024 * 1024; // 2MB in bytes
    const MAX_FILES_PER_STAFF = 10;
    const ALLOWED_MIME_TYPES = ['image/jpeg', 'image/jpg'];
    const ALLOWED_EXTENSIONS = ['.jpg', '.jpeg', '.JPG', '.JPEG'];

    // Explicitly reject PDFs
    const normalizedMimeType = file.type ? file.type.toLowerCase() : '';
    const normalizedFileName = file.name ? file.name.toLowerCase() : '';

    if (normalizedMimeType === 'application/pdf' || normalizedFileName.endsWith('.pdf')) {
      return {
        valid: false,
        error: 'PDF_NOT_ALLOWED',
        message: 'PDF files are not allowed. Only JPG/JPEG image files are permitted.'
      };
    }

    // Validate file type (only JPG/JPEG)
    const isValidMimeType = ALLOWED_MIME_TYPES.includes(normalizedMimeType);
    const isValidExtension = normalizedFileName.endsWith('.jpg') || normalizedFileName.endsWith('.jpeg');

    if (!isValidMimeType || !isValidExtension) {
      return {
        valid: false,
        error: 'INVALID_FILE_TYPE',
        message: 'Only JPG/JPEG files are allowed. PDFs and other file types are not permitted. File type: ' + (file.type || 'unknown')
      };
    }

    // Validate file size
    if (file.size > MAX_FILE_SIZE) {
      const fileSizeMB = (file.size / (1024 * 1024)).toFixed(2);
      return {
        valid: false,
        error: 'FILE_SIZE_EXCEEDED',
        message: `File size (${fileSizeMB}MB) exceeds maximum allowed size of 2MB.`
      };
    }

    // Validate file count
    if (currentFileCount >= MAX_FILES_PER_STAFF) {
      return {
        valid: false,
        error: 'MAX_FILES_EXCEEDED',
        message: `Maximum file limit reached. Employee already has ${currentFileCount} files (max: ${MAX_FILES_PER_STAFF}). Please delete existing files before uploading new ones.`
      };
    }

    return { valid: true };
  }

  /**
   * Convert file to base64
   */
  function fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  /**
   * Upload document for an employee
   */
  async function uploadEmployeeDocument(employeeId, formationId, subUnitId) {
    if (!adminKey || !employeeId) return;

    // Check permissions - only HRM_ADMIN and SUPER_ADMIN can upload
    if (adminRole !== 'HRM_ADMIN' && adminRole !== 'SUPER_ADMIN') {
      await UI.showError('Access Denied', 'Only HRM_ADMIN and SUPER_ADMIN can upload documents.');
      return;
    }

    // formationId and subUnitId are optional - backend will get from staff record if not provided
    if (!formationId) {
      formationId = currentFormationId || adminFormationId || '';
    }
    if (!subUnitId) {
      subUnitId = adminDepartmentId || '';
    }

    if (!window.Swal) {
      await UI.showError('Unavailable', 'SweetAlert2 is not loaded.');
      return;
    }

    // Get current file count
    let currentFileCount = 0;
    try {
      const docRes = await Api.call('listStaffDocuments', {
        key: adminKey,
        employeeId: employeeId
      });
      if (docRes && docRes.success && docRes.data && docRes.data.documents) {
        currentFileCount = docRes.data.documents.length;
      }
    } catch (err) {
      console.warn('Could not fetch current file count:', err);
    }

    // Show upload modal with instructions
    const uploadConfirm = await Swal.fire({
      title: 'Upload Staff Document',
      html: `
        <div style="text-align: left; padding: 1rem 0;">
          <p style="margin-bottom: 1rem; color: #374151;">
            <strong>Upload Requirements:</strong>
          </p>
          <ul style="margin-left: 1.5rem; margin-bottom: 1rem; line-height: 1.8; color: #6b7280;">
            <li>File type: JPG/JPEG images only</li>
            <li>Maximum file size: 2MB</li>
            <li>Maximum files per staff: 10 (Current: ${currentFileCount}/10)</li>
          </ul>
          <p style="margin-top: 1rem; padding: 0.75rem; background: #fef3c7; border-left: 4px solid #f59e0b; border-radius: 4px; color: #92400e;">
            <strong>Note:</strong> PDF files and other file types are not allowed.
          </p>
        </div>
      `,
      showCancelButton: true,
      confirmButtonText: 'Select Image',
      cancelButtonText: 'Cancel',
      confirmButtonColor: '#059669',
      icon: 'info'
    });

    if (!uploadConfirm.isConfirmed) {
      return;
    }

    // Create file input
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/jpeg,image/jpg,.jpg,.jpeg'; // Explicitly exclude PDFs
    input.style.display = 'none';
    document.body.appendChild(input);

    // Trigger file selection
    input.click();

    input.onchange = async (e) => {
      const file = e.target.files[0];
      if (!file) {
        document.body.removeChild(input);
        return;
      }

      // Client-side validation
      const validation = validateFileUpload(file, currentFileCount);
      if (!validation.valid) {
        document.body.removeChild(input);
        await UI.showError('Upload Rejected', validation.message);
        return;
      }

      // Show loading with file name
      Swal.fire({
        title: 'Uploading Image...',
        html: `
          <div style="text-align: center;">
            <p style="margin-bottom: 0.5rem; font-weight: 600;">${file.name}</p>
            <p style="color: #6b7280; font-size: 0.9rem;">${(file.size / 1024).toFixed(2)} KB</p>
          </div>
        `,
        allowOutsideClick: false,
        allowEscapeKey: false,
        showConfirmButton: false,
        didOpen: () => {
          Swal.showLoading();
        }
      });

      try {
        // Convert to base64
        const base64Data = await fileToBase64(file);

        // Upload to server
        // formationId and subUnitId are optional - backend will get from staff record if not provided
        const res = await Api.call('uploadDocumentMetadata', {
          key: adminKey,
          employeeId: employeeId,
          formationId: formationId || '', // Can be empty - backend will get from staff record
          subUnitId: subUnitId || '', // Can be empty - backend will get from staff record
          fileName: file.name,
          mimeType: file.type,
          fileData: base64Data,
          fileSize: file.size
        });

        if (res && res.success) {
          Swal.close();

          // Show success message with VIEW DOCUMENTS button
          const uploadResult = await Swal.fire({
            title: 'Upload Successful!',
            html: `
              <div style="text-align: center; padding: 1rem 0;">
                <div style="font-size: 3rem; margin-bottom: 1rem;">✅</div>
                <p style="font-size: 1.1rem; margin-bottom: 0.5rem; color: #374151;">
                  <strong>Image "${file.name}"</strong>
                </p>
                <p style="color: #6b7280; margin-bottom: 1.5rem;">
                  has been uploaded successfully!
                </p>
              </div>
            `,
            showConfirmButton: true,
            confirmButtonText: 'VIEW DOCUMENTS',
            confirmButtonColor: '#059669',
            showCancelButton: true,
            cancelButtonText: 'Close',
            icon: 'success',
            width: '500px'
          });

          // If user clicks VIEW DOCUMENTS, open the staff profile with documents
          if (uploadResult.isConfirmed) {
            await viewStaffProfile(employeeId);
          }
        } else {
          throw new Error(res.message || 'Upload failed');
        }
      } catch (err) {
        // Check if it's a Drive permission error
        const errorMsg = err.message || '';
        if (errorMsg.includes('Drive API permissions') ||
          errorMsg.includes('DRIVE_PERMISSION_REQUIRED') ||
          errorMsg.includes('DRIVE_UPLOAD_ERROR')) {
          // Show detailed instructions for Drive permission setup
          await Swal.fire({
            icon: 'warning',
            title: 'Drive Permissions Required',
            html: `
              <div style="text-align: left; padding: 1rem 0;">
                <p style="margin-bottom: 1rem; font-weight: 600; color: #dc2626;">
                  The system needs Google Drive permissions to upload staff documents.
                </p>
                <p style="margin-bottom: 1rem;"><strong>This is a one-time setup that must be done by the system administrator:</strong></p>
                <ol style="margin-left: 1.5rem; margin-bottom: 1rem; line-height: 2;">
                  <li>Open <strong>Google Apps Script Editor</strong> (script.google.com)</li>
                  <li>Select your project</li>
                  <li>In the function dropdown, select <code style="background: #f3f4f6; padding: 2px 6px; border-radius: 3px;">testDrivePermissions</code></li>
                  <li>Click <strong>"Run"</strong> button</li>
                  <li>When prompted, click <strong>"Review permissions"</strong></li>
                  <li>Select your Google account</li>
                  <li>Click <strong>"Advanced"</strong> → <strong>"Go to [Project Name] (unsafe)"</strong></li>
                  <li>Click <strong>"Allow"</strong> to grant Drive API access</li>
                  <li>Run function: <code style="background: #f3f4f6; padding: 2px 6px; border-radius: 3px;">initializeHrmDriveFolders()</code> to create folder structure</li>
                  <li>Try uploading again</li>
                </ol>
                <p style="margin-top: 1rem; padding: 0.75rem; background: #fef3c7; border-left: 4px solid #f59e0b; border-radius: 4px;">
                  <strong>Note:</strong> This setup only needs to be done once. After authorization, all HRM admins will be able to upload documents.
                </p>
              </div>
            `,
            width: '700px',
            confirmButtonText: 'I Understand',
            confirmButtonColor: '#059669'
          });
        } else {
          await UI.showError('Upload Failed', err.message || 'Failed to upload document. Please try again.');
        }
      } finally {
        document.body.removeChild(input);
      }
    };
  }

  async function viewEmployeeProfile(employeeId) {
    if (!adminKey || !currentFormationId) return;

    try {
      const res = await Api.call('getEmployeeProfile', {
        key: adminKey,
        formationId: currentFormationId,
        employeeId: employeeId,
      });

      // Get documents
      let documents = [];
      try {
        const docRes = await Api.call('listStaffDocuments', {
          key: adminKey,
          employeeId: employeeId
        });
        if (docRes && docRes.success && docRes.data && docRes.data.documents) {
          documents = docRes.data.documents;
        }
      } catch (docErr) {
        console.warn('Could not load documents:', docErr);
      }

      if (res && res.success && res.data) {
        const profile = res.data;

        // Build documents HTML
        let documentsHtml = '';
        if (documents.length > 0) {
          documentsHtml = '<div style="margin-top: 1rem;"><strong>Documents (' + documents.length + '/10):</strong><ul style="text-align: left; margin-top: 0.5rem;">';
          for (const doc of documents) {
            const fileSizeKB = doc.fileSize ? (doc.fileSize / 1024).toFixed(2) + 'KB' : 'N/A';
            documentsHtml += `<li><a href="${doc.fileUrl || doc.downloadUrl || '#'}" target="_blank">${doc.fileName}</a> (${fileSizeKB})</li>`;
          }
          documentsHtml += '</ul></div>';
        } else {
          documentsHtml = '<div style="margin-top: 1rem;"><em>No documents uploaded yet.</em></div>';
        }

        const result = await Swal.fire({
          title: 'Employee Profile',
          html: `
            <div style="text-align: left;">
              <p><strong>Name:</strong> ${profile.name || 'N/A'}</p>
              <p><strong>Department:</strong> ${profile.departmentName || profile.departmentId || 'N/A'}</p>
              <p><strong>Status:</strong> ${profile.profile?.status || 'ACTIVE'}</p>
              <p><strong>Phone:</strong> ${profile.profile?.phone || 'N/A'}</p>
              <p><strong>Email:</strong> ${profile.profile?.email || 'N/A'}</p>
              ${documentsHtml}
            </div>
          `,
          showCancelButton: true,
          confirmButtonText: 'Upload Document',
          cancelButtonText: 'Close',
          confirmButtonColor: '#059669',
          showDenyButton: documents.length > 0,
          denyButtonText: 'View All Documents',
          denyButtonColor: '#3b82f6'
        });

        if (result.isConfirmed) {
          await uploadEmployeeDocument(employeeId, currentFormationId, adminDepartmentId || '');
        } else if (result.isDenied) {
          // Show full documents list
          await showEmployeeDocuments(employeeId);
        }
      }
    } catch (err) {
      await UI.showError('Error', err.message || 'Failed to load profile.');
    }
  }

  /**
   * Show all documents for an employee
   */
  async function showEmployeeDocuments(employeeId) {
    if (!adminKey) return;

    try {
      const res = await Api.call('listStaffDocuments', {
        key: adminKey,
        employeeId: employeeId
      });

      if (res && res.success && res.data && res.data.documents) {
        const documents = res.data.documents;

        if (documents.length === 0) {
          await UI.showInfo('Documents', 'No documents found for this employee.');
          return;
        }

        // Use placeholder for doc images (Drive fileUrl/thumbnail in img src causes ERR_TOO_MANY_REDIRECTS for private files)
        const docPlaceholderSvg = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='180' height='150' viewBox='0 0 180 150'%3E%3Crect fill='%23e5e7eb' width='180' height='150'/%3E%3Ctext x='50%25' y='45%25' fill='%236b7280' font-size='14' text-anchor='middle'%3EDocument%3C/text%3E%3Ctext x='50%25' y='55%25' fill='%239ca3af' font-size='11' text-anchor='middle'%3EClick to open%3C/text%3E%3C/svg%3E";
        let html = '<div style="text-align: left; max-height: 500px; overflow-y: auto;">';
        html += '<div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(180px, 1fr)); gap: 1rem; margin-top: 0.5rem;">';

        for (const doc of documents) {
          const fileSizeKB = doc.fileSize ? (doc.fileSize / 1024).toFixed(2) + 'KB' : 'N/A';
          const uploadDate = doc.uploadedAt ? new Date(doc.uploadedAt).toLocaleDateString() : 'N/A';
          const isImage = doc.mimeType && doc.mimeType.startsWith('image/');
          const downloadUrl = doc.driveFileId ? ('https://drive.google.com/uc?id=' + doc.driveFileId + '&export=download') : (doc.downloadUrl || '#');
          const fileName = doc.fileName || 'Document';

          if (isImage) {
            html += `
              <div style="border: 1px solid #e5e7eb; border-radius: 0.5rem; padding: 0.5rem; text-align: center; background: #f9fafb; cursor: pointer; transition: transform 0.2s; box-shadow 0.2s;" 
                   onclick="window.open('${downloadUrl.replace(/'/g, "\\'")}', '_blank')"
                   onmouseover="this.style.transform='scale(1.02)'; this.style.boxShadow='0 4px 8px rgba(0,0,0,0.1)';"
                   onmouseout="this.style.transform='scale(1)'; this.style.boxShadow='none';">
                <img src="${docPlaceholderSvg}" alt="${(doc.fileName || 'Document').replace(/"/g, '&quot;')}" style="width: 100%; height: 150px; object-fit: cover; border-radius: 0.25rem; margin-bottom: 0.5rem; cursor: pointer;" />
                <div style="font-size: 0.75rem; color: #374151; word-break: break-word; font-weight: 600; margin-bottom: 0.25rem;" title="${fileName.replace(/"/g, '&quot;')}">${fileName.length > 20 ? fileName.substring(0, 20) + '...' : fileName}</div>
                <div style="font-size: 0.7rem; color: #6b7280; margin-top: 0.25rem;">${fileSizeKB}</div>
                <div style="font-size: 0.65rem; color: #9ca3af; margin-top: 0.25rem;">${uploadDate}</div>
                <div style="font-size: 0.65rem; color: #9ca3af;">By: ${doc.uploadedBy || 'N/A'}</div>
                <div style="font-size: 0.6rem; color: #059669; margin-top: 0.25rem; font-weight: 500;">Click to view larger</div>
              </div>
            `;
          } else {
            html += `
              <div style="border: 1px solid #e5e7eb; border-radius: 0.5rem; padding: 0.75rem; text-align: center; background: #f9fafb;">
                <a href="${downloadUrl}" target="_blank" rel="noopener" style="display: block; text-decoration: none; color: #059669;">
                  <div style="font-size: 2.5rem; margin-bottom: 0.5rem;">📄</div>
                  <div style="font-size: 0.75rem; color: #374151; word-break: break-word; font-weight: 600;">${fileName.length > 20 ? fileName.substring(0, 20) + '...' : fileName}</div>
                  <div style="font-size: 0.7rem; color: #6b7280; margin-top: 0.25rem;">${fileSizeKB}</div>
                  <div style="font-size: 0.65rem; color: #9ca3af; margin-top: 0.25rem;">${uploadDate}</div>
                  <div style="font-size: 0.65rem; color: #9ca3af;">By: ${doc.uploadedBy || 'N/A'}</div>
                </a>
              </div>
            `;
          }
        }
        html += '</div></div>';

        await Swal.fire({
          title: 'Employee Documents',
          html: html,
          width: '900px',
          confirmButtonText: 'Close',
          confirmButtonColor: '#059669',
          customClass: {
            popup: 'swal2-wide-modal',
            htmlContainer: 'swal2-html-container-wide'
          }
        });
      } else {
        await UI.showInfo('Documents', 'No documents found for this employee.');
      }
    } catch (err) {
      await UI.showError('Error', err.message || 'Failed to load documents.');
    }
  }

  async function approveLeave(requestId) {
    if (!adminKey) return;

    try {
      UI.showLoading('Processing', 'Approving leave request...');
      await Api.call('approveLeave', {
        key: adminKey,
        requestId: requestId,
      });
      UI.closeLoading();
      await UI.showSuccess('Success', 'Leave request approved.');
      await loadHrmLeaves();
      await loadHrmStats();
    } catch (err) {
      UI.closeLoading();
      await UI.showError('Error', err.message || 'Failed to approve leave.');
    }
  }

  async function rejectLeave(requestId) {
    if (!adminKey) return;

    const reason = await Swal.fire({
      title: 'Reject Leave Request',
      input: 'text',
      inputLabel: 'Rejection Reason',
      inputPlaceholder: 'Enter reason for rejection...',
      showCancelButton: true,
      confirmButtonText: 'Reject',
      confirmButtonColor: '#dc2626',
    });

    if (!reason.isConfirmed) return;

    try {
      UI.showLoading('Processing', 'Rejecting leave request...');
      await Api.call('rejectLeave', {
        key: adminKey,
        requestId: requestId,
        rejectionReason: reason.value || 'No reason provided.',
      });
      UI.closeLoading();
      await UI.showSuccess('Success', 'Leave request rejected.');
      await loadHrmLeaves();
      await loadHrmStats();
    } catch (err) {
      UI.closeLoading();
      await UI.showError('Error', err.message || 'Failed to reject leave.');
    }
  }

  async function showCreatePerformanceModal() {
    if (!window.Swal) {
      await UI.showError('Unavailable', 'SweetAlert2 is not loaded.');
      return;
    }

    const result = await Swal.fire({
      title: 'Create Performance Review',
      html: `
        <input id="perfEmployeeId" class="swal2-input" placeholder="Employee ID *">
        <input id="perfReviewPeriod" class="swal2-input" placeholder="Review Period (e.g., 2024-Q1) *">
        <input id="perfRating" class="swal2-input" type="number" min="1" max="5" placeholder="Rating (1-5) *">
        <textarea id="perfComments" class="swal2-textarea" placeholder="Comments"></textarea>
      `,
      focusConfirm: false,
      showCancelButton: true,
      confirmButtonText: 'Create',
      cancelButtonText: 'Cancel',
      confirmButtonColor: '#059669',
      preConfirm: async () => {
        const employeeId = document.getElementById('perfEmployeeId').value.trim();
        const reviewPeriod = document.getElementById('perfReviewPeriod').value.trim();
        const rating = parseInt(document.getElementById('perfRating').value, 10);
        const comments = document.getElementById('perfComments').value.trim();

        if (!employeeId || !reviewPeriod || !rating) {
          Swal.showValidationMessage('Employee ID, Review Period, and Rating are required.');
          return false;
        }

        if (rating < 1 || rating > 5) {
          Swal.showValidationMessage('Rating must be between 1 and 5.');
          return false;
        }

        try {
          await Api.call('createPerformanceReview', {
            key: adminKey,
            formationId: currentFormationId,
            employeeId: employeeId,
            reviewPeriod: reviewPeriod,
            rating: rating,
            comments: comments,
          });
          return true;
        } catch (err) {
          Swal.showValidationMessage(err.message || 'Failed to create review.');
          return false;
        }
      },
    });

    if (result.isConfirmed) {
      await loadHrmPerformance();
      await UI.showSuccess('Success', 'Performance review created successfully.');
    }
  }

  // Expose HRM functions globally (assign to window, not const AdminPage)
  if (!window.AdminPage) {
    window.AdminPage = {};
  }
  window.AdminPage.viewEmployeeProfile = viewEmployeeProfile;
  window.AdminPage.approveLeave = approveLeave;
  window.AdminPage.rejectLeave = rejectLeave;
  window.AdminPage.uploadEmployeeDocument = uploadEmployeeDocument;
  window.AdminPage.showEmployeeDocuments = showEmployeeDocuments;

  // ============================================================================
  // HRM STAFF MANAGEMENT FUNCTIONS
  // ============================================================================

  let currentStaffPage = 1;
  let currentStaffLimit = 20;

  /**
   * Load HRM Staff Dashboard statistics
   */
  async function loadHrmStaffStats() {
    if (!adminKey) return;

    if (adminRole !== 'HRM_ADMIN' && adminRole !== 'SUPER_ADMIN') return;

    try {
      // For HRM_ADMIN and SUPER_ADMIN, formationId is optional (can search all)
      const formationIdForStats = currentFormationId || adminFormationId || '';

      UI.showLoading('Loading', 'Fetching staff statistics...');
      const searchRes = await Api.call('searchStaff', {
        key: adminKey,
        formationId: formationIdForStats, // Can be empty for HRM_ADMIN/SUPER_ADMIN
        query: '',
        page: 1,
        limit: 1000,
        includeArchived: true
      });
      UI.closeLoading();

      if (searchRes && searchRes.success && searchRes.data) {
        const allStaff = searchRes.data.results || [];
        const active = allStaff.filter(s => s.status === 'ACTIVE').length;
        const archived = allStaff.filter(s => s.status === 'ARCHIVED').length;

        let withDocs = 0;
        for (const staff of allStaff.slice(0, 50)) { // Limit to first 50 for performance
          try {
            const docRes = await Api.call('listStaffDocuments', {
              key: adminKey,
              employeeId: staff.employeeId
            });
            if (docRes && docRes.success && docRes.data && docRes.data.documents && docRes.data.documents.length > 0) {
              withDocs++;
            }
          } catch (e) {
            // Ignore errors
          }
        }

        const totalEl = document.getElementById('hrmStaffTotal');
        const activeEl = document.getElementById('hrmStaffActive');
        const archivedEl = document.getElementById('hrmStaffArchived');
        const withDocsEl = document.getElementById('hrmStaffWithDocs');

        if (totalEl) totalEl.textContent = allStaff.length;
        if (activeEl) activeEl.textContent = active;
        if (archivedEl) archivedEl.textContent = archived;
        if (withDocsEl) withDocsEl.textContent = withDocs;
      }
    } catch (err) {
      UI.closeLoading();
      console.error('Failed to load HRM staff stats:', err);
    }
  }

  /**
   * Show Add Staff modal
   */
  async function showAddStaffModal() {
    if (!adminKey) {
      await UI.showError('Error', 'Please log in first.');
      return;
    }

    if (!window.Swal) {
      await UI.showError('Unavailable', 'SweetAlert2 is not loaded.');
      return;
    }

    // Load formations and departments for selection
    let formations = [];
    let departments = [];

    try {
      UI.showLoading('Loading', 'Fetching formations and departments...');
      // Load formations (for SUPER_ADMIN and HRM_ADMIN)
      if (adminRole === 'SUPER_ADMIN' || adminRole === 'HRM_ADMIN') {
        const formRes = await Api.call('listFormations', { key: adminKey });
        if (formRes && formRes.success && formRes.data && formRes.data.formations) {
          formations = formRes.data.formations.filter(f => f.active !== false);
        }
      }

      // Load departments (if formation is selected or admin has one)
      const formIdForDept = currentFormationId || adminFormationId || (formations.length > 0 ? formations[0].formationId : '');
      if (formIdForDept) {
        const deptRes = await Api.call('listDepartments', {
          key: adminKey,
          formationId: formIdForDept
        });
        if (deptRes && deptRes.success && deptRes.data && deptRes.data.departments) {
          departments = deptRes.data.departments;
        }
      }
      UI.closeLoading();
    } catch (err) {
      UI.closeLoading();
      console.warn('Could not load formations/departments:', err);
    }

    // Store step 1 data in a closure variable
    let step1Data = {};

    // Step 1: Basic Information
    const step1 = await Swal.fire({
      title: 'Add Staff Record - Step 1 of 2',
      width: '800px',
      html: `
        <p style="text-align: left; color: #666; font-size: 0.9rem; margin-bottom: 1rem;">
          <strong>Note:</strong> Employee ID will be automatically generated. Formation and Department can be added later.
        </p>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-bottom: 1rem;">
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Surname *</label>
            <input id="swalStaffSurname" class="swal2-input" placeholder="Surname" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Other Names</label>
            <input id="swalStaffOtherNames" class="swal2-input" placeholder="Other Names" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Email (optional)</label>
            <input id="swalStaffEmail" class="swal2-input" type="email" placeholder="Email" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Telephone Number</label>
            <input id="swalStaffTelephone" class="swal2-input" type="tel" placeholder="Telephone Number" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          ${(adminRole === 'SUPER_ADMIN' || adminRole === 'HRM_ADMIN') && formations.length > 0 ? `
          <div style="grid-column: 1 / -1;">
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Formation (optional - can be added later)</label>
            <select id="swalStaffFormation" class="swal2-select" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
              <option value="">-- Select Formation (Optional) --</option>
              ${formations.map(form => `<option value="${form.formationId}" ${currentFormationId && form.formationId === currentFormationId ? 'selected' : ''}>${form.name || form.formationId}</option>`).join('')}
            </select>
          </div>
          ` : ''}
          <div style="grid-column: 1 / -1;">
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Department/Sub-Unit (optional - can be added later)</label>
            <select id="swalStaffSubUnit" class="swal2-select" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
              <option value="">-- Select Department (Optional) --</option>
              ${departments.map(dept => `<option value="${dept.departmentId || dept.subUnitId}" ${adminDepartmentId && dept.departmentId === adminDepartmentId ? 'selected' : ''}>${dept.name || dept.departmentId || dept.subUnitId}</option>`).join('')}
            </select>
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">File Number</label>
            <input id="swalStaffFileNumber" class="swal2-input" placeholder="File Number" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">IPPIS Number</label>
            <input id="swalStaffIppis" class="swal2-input" placeholder="IPPIS Number" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Date of Birth</label>
            <input id="swalStaffDob" class="swal2-input" type="date" placeholder="Date of Birth" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Date of First Appointment</label>
            <input id="swalStaffFirstAppointment" class="swal2-input" type="date" placeholder="Date of First Appointment" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Date of Present Appointment</label>
            <input id="swalStaffPresentAppointment" class="swal2-input" type="date" placeholder="Date of Present Appointment" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Confirmation Date</label>
            <input id="swalStaffConfirmationDate" class="swal2-input" type="date" placeholder="Confirmation Date" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Cadre</label>
            <input id="swalStaffCadre" class="swal2-input" placeholder="Cadre" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Rank</label>
            <input id="swalStaffRank" class="swal2-input" placeholder="Rank" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Grade Level</label>
            <input id="swalStaffGradeLevel" class="swal2-input" placeholder="Grade Level" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">State of Origin</label>
            <input id="swalStaffStateOfOrigin" class="swal2-input" placeholder="State of Origin" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">LGA</label>
            <input id="swalStaffLga" class="swal2-input" placeholder="Local Government Area" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div style="grid-column: 1 / -1;">
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Qualification</label>
            <input id="swalStaffQualification" class="swal2-input" placeholder="Educational Qualification" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
        </div>
      `,
      focusConfirm: false,
      showConfirmButton: true,
      showCancelButton: true,
      confirmButtonText: 'Next Page',
      cancelButtonText: 'Cancel',
      confirmButtonColor: '#059669',
      allowOutsideClick: false,
      allowEscapeKey: true,
      allowEnterKey: true,
      buttonsStyling: true,
      customClass: {
        popup: 'swal2-wide-modal',
        htmlContainer: 'swal2-html-container-wide',
        confirmButton: 'swal2-confirm'
      },
      didOpen: () => {
        // Ensure confirm button is visible
        setTimeout(() => {
          const confirmBtn = document.querySelector('.swal2-confirm');
          if (confirmBtn) {
            confirmBtn.style.display = '';
            confirmBtn.disabled = false;
          }
        }, 100);

        // Load departments when formation is selected
        const formationSelect = document.getElementById('swalStaffFormation');
        const subUnitSelect = document.getElementById('swalStaffSubUnit');

        if (formationSelect && subUnitSelect) {
          formationSelect.addEventListener('change', async function () {
            const selectedFormationId = this.value;
            if (selectedFormationId) {
              // Show loading indicator in the select itself instead of blocking modal
              subUnitSelect.disabled = true;
              subUnitSelect.innerHTML = '<option value="">Loading departments...</option>';
              
              try {
                const deptRes = await Api.call('listDepartments', {
                  key: adminKey,
                  formationId: selectedFormationId
                });
                
                if (deptRes && deptRes.success && deptRes.data && deptRes.data.departments) {
                  subUnitSelect.innerHTML = '<option value="">-- Select Department (Optional) --</option>';
                  deptRes.data.departments.forEach(dept => {
                    const option = document.createElement('option');
                    option.value = dept.departmentId || dept.subUnitId;
                    option.textContent = dept.name || dept.departmentId || dept.subUnitId;
                    subUnitSelect.appendChild(option);
                  });
                } else {
                  subUnitSelect.innerHTML = '<option value="">-- Select Department (Optional) --</option>';
                }
              } catch (err) {
                console.warn('Could not load departments:', err);
                subUnitSelect.innerHTML = '<option value="">-- Select Department (Optional) --</option>';
              } finally {
                subUnitSelect.disabled = false;
              }
            } else {
              subUnitSelect.innerHTML = '<option value="">-- Select Department (Optional) --</option>';
            }
          });
        }
      }
    });

    // Validate step 1
    if (!step1.isConfirmed) return;

    // Collect step 1 data - use small delay to ensure DOM is ready
    await new Promise(resolve => setTimeout(resolve, 50));
    
    try {
      // Try to get elements - they might still be in DOM briefly
      let surnameEl = document.getElementById('swalStaffSurname');
      let surname = '';
      
      // If element not found, try accessing through Swal container
      if (!surnameEl && Swal && Swal.getContainer) {
        const container = Swal.getContainer();
        if (container) {
          surnameEl = container.querySelector('#swalStaffSurname');
        }
      }
      
      surname = surnameEl ? surnameEl.value.trim() : '';
      
      if (!surname) {
        await UI.showError('Validation Error', 'Surname is required to proceed.');
        return;
      }

      const otherNamesEl = document.getElementById('swalStaffOtherNames');
      const emailEl = document.getElementById('swalStaffEmail');
      const telephoneEl = document.getElementById('swalStaffTelephone');
      const formationEl = document.getElementById('swalStaffFormation');
      const subUnitEl = document.getElementById('swalStaffSubUnit');
      const fileNumberEl = document.getElementById('swalStaffFileNumber');
      const ippisEl = document.getElementById('swalStaffIppis');
      const dobEl = document.getElementById('swalStaffDob');
      const firstAppointmentEl = document.getElementById('swalStaffFirstAppointment');
      const presentAppointmentEl = document.getElementById('swalStaffPresentAppointment');
      const confirmationDateEl = document.getElementById('swalStaffConfirmationDate');
      const cadreEl = document.getElementById('swalStaffCadre');
      const rankEl = document.getElementById('swalStaffRank');
      const gradeLevelEl = document.getElementById('swalStaffGradeLevel');
      const stateOfOriginEl = document.getElementById('swalStaffStateOfOrigin');
      const lgaEl = document.getElementById('swalStaffLga');
      const qualificationEl = document.getElementById('swalStaffQualification');

      step1Data = {
        surname: surname,
        otherNames: otherNamesEl ? otherNamesEl.value.trim() : '',
        email: emailEl ? emailEl.value.trim() : '',
        telephone: telephoneEl ? telephoneEl.value.trim() : '',
        formationId: formationEl ? (formationEl.value.trim() || currentFormationId || adminFormationId || '') : (currentFormationId || adminFormationId || ''),
        subUnitId: subUnitEl ? (subUnitEl.value.trim() || adminDepartmentId || '') : (adminDepartmentId || ''),
        fileNumber: fileNumberEl ? fileNumberEl.value.trim() : '',
        ippisNumber: ippisEl ? ippisEl.value.trim() : '',
        dob: dobEl ? dobEl.value.trim() : '',
        firstAppointment: firstAppointmentEl ? firstAppointmentEl.value.trim() : '',
        presentAppointment: presentAppointmentEl ? presentAppointmentEl.value.trim() : '',
        confirmationDate: confirmationDateEl ? confirmationDateEl.value.trim() : '',
        cadre: cadreEl ? cadreEl.value.trim() : '',
        rank: rankEl ? rankEl.value.trim() : '',
        gradeLevel: gradeLevelEl ? gradeLevelEl.value.trim() : '',
        stateOfOrigin: stateOfOriginEl ? stateOfOriginEl.value.trim() : '',
        lga: lgaEl ? lgaEl.value.trim() : '',
        qualification: qualificationEl ? qualificationEl.value.trim() : ''
      };
    } catch (err) {
      console.error('Error collecting step 1 data:', err);
      await UI.showError('Error', 'Failed to collect form data. Please try again.');
      return;
    }

    // Step 2: Personal Information
    const step2 = await Swal.fire({
      title: 'Add Staff Record - Step 2 of 2',
      width: '800px',
      html: `
        <p style="text-align: left; color: #666; font-size: 0.9rem; margin-bottom: 1rem;">
          <strong>Personal Information:</strong> Fill in the following details (all fields are optional).
        </p>
        ${(adminRole === 'SUPER_ADMIN' || adminRole === 'HRM_ADMIN') ? `
        <div style="grid-column: 1 / -1; margin-bottom: 1rem; padding: 1rem; border: 1px dashed #059669; border-radius: 8px; background: #f0fdf4;">
          <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #059669;">Passport / Profile Picture (optional)</label>
          <p style="font-size: 0.85rem; color: #666; margin-bottom: 0.5rem;">JPEG, PNG or WebP. Max 2MB. Stored in staff profile folder on Drive.</p>
          <input type="file" id="swalStaffProfilePhoto" accept="image/jpeg,image/png,image/webp,image/jpg" style="width: 100%; padding: 0.5rem;">
        </div>
        ` : ''}
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-bottom: 1rem;">
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Marital Status</label>
            <select id="swalStaffMaritalStatus" class="swal2-select" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
              <option value="">-- Select Marital Status --</option>
              <option value="Single">Single</option>
              <option value="Married">Married</option>
              <option value="Divorced">Divorced</option>
              <option value="Widowed">Widowed</option>
            </select>
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Spouse Name</label>
            <input id="swalStaffSpouseName" class="swal2-input" placeholder="Spouse Name" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div style="grid-column: 1 / -1;">
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Home Address</label>
            <textarea id="swalStaffHomeAddress" class="swal2-textarea" placeholder="Home Address" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem; min-height: 80px;"></textarea>
          </div>
          <div style="grid-column: 1 / -1;">
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Permanent Home Address</label>
            <textarea id="swalStaffPermanentAddress" class="swal2-textarea" placeholder="Permanent Home Address" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem; min-height: 80px;"></textarea>
          </div>
          <div style="grid-column: 1 / -1;">
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Spouse Address</label>
            <textarea id="swalStaffSpouseAddress" class="swal2-textarea" placeholder="Spouse Address" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem; min-height: 80px;"></textarea>
          </div>
          <div>
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Next of Kin</label>
            <input id="swalStaffNextOfKin" class="swal2-input" placeholder="Next of Kin Name" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
          </div>
          <div style="grid-column: 1 / -1;">
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Next of Kin Address</label>
            <textarea id="swalStaffNextOfKinAddress" class="swal2-textarea" placeholder="Next of Kin Address" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem; min-height: 80px;"></textarea>
          </div>
        </div>
      `,
      focusConfirm: false,
      showCancelButton: true,
      confirmButtonText: 'Create Staff Record',
      cancelButtonText: 'Back',
      confirmButtonColor: '#059669',
      customClass: {
        popup: 'swal2-wide-modal',
        htmlContainer: 'swal2-html-container-wide'
      },
      preConfirm: async () => {
        // Get values from step 2 (current modal)
        const maritalStatus = document.getElementById('swalStaffMaritalStatus')?.value.trim() || '';
        const spouseName = document.getElementById('swalStaffSpouseName')?.value.trim() || '';
        const homeAddress = document.getElementById('swalStaffHomeAddress')?.value.trim() || '';
        const permanentAddress = document.getElementById('swalStaffPermanentAddress')?.value.trim() || '';
        const spouseAddress = document.getElementById('swalStaffSpouseAddress')?.value.trim() || '';
        const nextOfKin = document.getElementById('swalStaffNextOfKin')?.value.trim() || '';
        const nextOfKinAddress = document.getElementById('swalStaffNextOfKinAddress')?.value.trim() || '';

        // Optional profile picture (HRM Admin / Super Admin only)
        let profilePictureData = null;
        const photoInput = document.getElementById('swalStaffProfilePhoto');
        if (photoInput && photoInput.files && photoInput.files[0]) {
          const file = photoInput.files[0];
          const maxSize = 2 * 1024 * 1024; // 2MB
          if (file.size > maxSize) {
            Swal.showValidationMessage('Profile picture must be 2MB or less.');
            return false;
          }
          const allowed = ['image/jpeg', 'image/jpg', 'image/png', 'image/webp'];
          if (!allowed.includes(file.type)) {
            Swal.showValidationMessage('Profile picture must be JPEG, PNG or WebP.');
            return false;
          }
          profilePictureData = await new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve({ base64: reader.result, fileName: file.name, mimeType: file.type, fileSize: file.size });
            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsDataURL(file);
          });
        }

        // Use step 1 data (stored when step 1 closed)
        const surname = step1Data.surname || '';
        if (!surname) {
          Swal.showValidationMessage('Surname is required.');
          return false;
        }

        try {
          UI.showLoading('Creating', 'Creating staff record...');
          const res = await Api.call('createStaff', {
            key: adminKey,
            staff: {
              formationId: step1Data.formationId || '',
              subUnitId: step1Data.subUnitId || '',
              email: step1Data.email || '',
              telephone: step1Data.telephone || '',
              fileNumber: step1Data.fileNumber || '',
              ippisNumber: step1Data.ippisNumber || '',
              surname: surname,
              otherNames: step1Data.otherNames || '',
              dob: step1Data.dob || '',
              dateOfFirstAppointment: step1Data.firstAppointment || '',
              dateOfPresentAppointment: step1Data.presentAppointment || '',
              confirmationDate: step1Data.confirmationDate || '',
              cadre: step1Data.cadre || '',
              rank: step1Data.rank || '',
              gradeLevel: step1Data.gradeLevel || '',
              stateOfOrigin: step1Data.stateOfOrigin || '',
              lga: step1Data.lga || '',
              qualification: step1Data.qualification || '',
              maritalStatus: maritalStatus || '',
              spouseName: spouseName || '',
              homeAddress: homeAddress || '',
              permanentHomeAddress: permanentAddress || '',
              spouseAddress: spouseAddress || '',
              nextOfKin: nextOfKin || '',
              nextOfKinAddress: nextOfKinAddress || ''
            }
          });
          if (!res || !res.success) {
            UI.closeLoading();
            throw new Error(res?.message || 'Failed to create staff record.');
          }
          const employeeId = res.data?.employeeId || 'N/A';
          const formationId = res.data?.formationId || step1Data.formationId || currentFormationId || adminFormationId || '';

          if (profilePictureData && employeeId && (adminRole === 'SUPER_ADMIN' || adminRole === 'HRM_ADMIN')) {
            UI.showLoading('Uploading', 'Uploading profile picture...');
            try {
              await Api.call('uploadStaffProfilePicture', {
                key: adminKey,
                employeeId,
                formationId: formationId || undefined,
                subUnitId: res.data?.subUnitId || step1Data.subUnitId || undefined,
                fileData: profilePictureData.base64,
                fileName: profilePictureData.fileName,
                mimeType: profilePictureData.mimeType,
                fileSize: profilePictureData.fileSize
              });
            } catch (picErr) {
              console.warn('Profile picture upload failed:', picErr);
              // Don't fail create; user can add photo later in edit
            }
          }
          UI.closeLoading();
          return { success: true, employeeId, formationId };
        } catch (err) {
          UI.closeLoading();
          Swal.showValidationMessage(err.message || 'Failed to create staff record.');
          return false;
        }
      }
    });

    if (!step2 || !step2.isConfirmed || !step2.value || !step2.value.success) return;

    const result = step2;
    if (result.isConfirmed && result.value && result.value.success) {
      const employeeId = result.value.employeeId;
      const formationId = result.value.formationId || currentFormationId || adminFormationId || '';

      await UI.showSuccess(
        'Success',
        `Staff record created successfully!`
      );

      // Ask if user wants to upload documents
      if (formationId) {
        const uploadResult = await Swal.fire({
          title: 'Upload Documents?',
          text: 'Would you like to upload documents/images for this staff member now?',
          icon: 'question',
          showCancelButton: true,
          confirmButtonText: 'Yes, Upload Documents',
          cancelButtonText: 'Skip for Now',
          confirmButtonColor: '#059669'
        });

        if (uploadResult.isConfirmed) {
          // Get staff record to get formationId and subUnitId
          try {
            const staffRes = await Api.call('getStaffById', {
              key: adminKey,
              employeeId: employeeId
            });
            const staffFormationId = (staffRes && staffRes.success && staffRes.data && staffRes.data.staff)
              ? staffRes.data.staff.formationId || formationId || ''
              : formationId || '';
            const staffSubUnitId = (staffRes && staffRes.success && staffRes.data && staffRes.data.staff)
              ? staffRes.data.staff.subUnitId || ''
              : '';
            await uploadEmployeeDocument(employeeId, staffFormationId, staffSubUnitId);
          } catch (err) {
            // Fallback to just employeeId and formationId
            await uploadEmployeeDocument(employeeId, formationId, '');
          }
        }
      }

      await loadHrmStaffStats();
      await loadStaffList(1);
    }
  }

  /**
   * Load staff list with pagination
   */
  async function loadStaffList(page = null) {
    if (!adminKey) return;

    if (page !== null) {
      currentStaffPage = page;
    }

    const query = document.getElementById('hrmStaffSearchInput')?.value.trim() || '';
    const includeArchived = document.getElementById('hrmIncludeArchivedCheck')?.checked || false;

    const container = document.getElementById('hrmStaffTable');
    if (!container) return;

    container.textContent = 'Loading staff records...';

    try {
      // For HRM_ADMIN and SUPER_ADMIN, formationId is optional (can search all)
      // For others, use currentFormationId or adminFormationId
      const formationIdForSearch = (adminRole === 'SUPER_ADMIN' || adminRole === 'HRM_ADMIN')
        ? (currentFormationId || adminFormationId || '')
        : (currentFormationId || adminFormationId || '');

      const res = await Api.call('searchStaff', {
        key: adminKey,
        formationId: formationIdForSearch, // Can be empty for HRM_ADMIN/SUPER_ADMIN
        query: query,
        page: currentStaffPage,
        limit: currentStaffLimit,
        includeArchived: includeArchived
      });

      if (!res || !res.success || !res.data) {
        container.textContent = 'No staff records found.';
        return;
      }

      const staff = res.data.results || [];
      const pagination = res.data.pagination || {};

      if (staff.length === 0) {
        container.textContent = 'No staff records found.';
        const paginationEl = document.getElementById('hrmStaffPagination');
        if (paginationEl) paginationEl.innerHTML = '';
        return;
      }

      // Load formations and departments to map IDs to names
      let formationMap = {};
      let departmentMap = {};
      try {
        if (adminRole === 'SUPER_ADMIN' || adminRole === 'HRM_ADMIN') {
          const formRes = await Api.call('listFormations', { key: adminKey });
          if (formRes && formRes.success && formRes.data && formRes.data.formations) {
            formRes.data.formations.forEach(f => {
              formationMap[f.formationId] = f.name || f.formationId;
            });
          }
        }
        
        // Load departments for all unique formation IDs in the staff list
        const uniqueFormationIds = [...new Set(staff.map(s => s.formationId).filter(Boolean))];
        for (const formId of uniqueFormationIds) {
          try {
            const deptRes = await Api.call('listDepartments', {
              key: adminKey,
              formationId: formId
            });
            if (deptRes && deptRes.success && deptRes.data && deptRes.data.departments) {
              deptRes.data.departments.forEach(d => {
                const deptId = d.departmentId || d.subUnitId;
                departmentMap[deptId] = d.name || deptId;
              });
            }
          } catch (err) {
            console.warn('Could not load departments for formation:', formId, err);
          }
        }
      } catch (err) {
        console.warn('Could not load formations/departments for mapping:', err);
      }

      // Calculate serial number offset based on pagination
      const serialOffset = (pagination.page - 1) * (pagination.limit || currentStaffLimit) || 0;

      let html = '<table class="data-table"><thead><tr><th>S/N</th><th>File Number</th><th>Name</th><th>Cadre</th><th>Rank</th><th>Grade Level</th><th>Formation</th><th>Department</th><th>Status</th><th>Actions</th></tr></thead><tbody>';

      for (let idx = 0; idx < staff.length; idx++) {
        const s = staff[idx];
        const serialNumber = serialOffset + idx + 1;
        // Format DOB for display - prevent ISO format
        let dobDisplay = 'N/A';
        if (s.dob) {
          try {
            const dobValue = s.dob;
            // If it's not an ISO date string, use as is
            if (typeof dobValue === 'string' && !dobValue.match(/^\d{4}-\d{2}-\d{2}/)) {
              dobDisplay = dobValue;
            } else {
              const dobDate = new Date(dobValue);
              if (!isNaN(dobDate.getTime())) {
                dobDisplay = dobDate.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
              } else {
                dobDisplay = String(dobValue);
              }
            }
          } catch (e) {
            dobDisplay = String(s.dob);
          }
        }

        // Format Grade Level - ensure it's treated as text, not date
        // Convert to string and escape any date-like formatting
        let gradeLevelDisplay = 'N/A';
        if (s.gradeLevel !== null && s.gradeLevel !== undefined && s.gradeLevel !== '') {
          const gradeStr = String(s.gradeLevel);
          // If it looks like a date (ISO format), it's probably a mistake - return as plain text
          gradeLevelDisplay = gradeStr;
        }

        // Get formation and department names
        const formationName = s.formationId ? (formationMap[s.formationId] || 'Not Assigned') : 'Not Assigned';
        const departmentName = s.subUnitId ? (departmentMap[s.subUnitId] || 'Not Assigned') : 'Not Assigned';

        html += `<tr>
          <td>${serialNumber}</td>
          <td>${s.fileNumber || 'N/A'}</td>
          <td><strong>${(s.surname || '') + (s.otherNames ? ' ' + s.otherNames : '') || s.name || 'N/A'}</strong></td>
          <td>${s.cadre || 'N/A'}</td>
          <td>${s.rank || 'N/A'}</td>
          <td>${gradeLevelDisplay}</td>
          <td>${formationName}</td>
          <td>${departmentName}</td>
          <td><span class="badge ${s.status === 'ACTIVE' ? 'badge-success' : 'badge-warning'}">${s.status || 'ACTIVE'}</span></td>
          <td style="white-space: nowrap;">
            <button class="btn btn-xs btn-primary" onclick="AdminPage.viewStaffProfile('${s.employeeId}')" title="View full details and documents">View</button>
            <button class="btn btn-xs btn-secondary" onclick="AdminPage.editStaff('${s.employeeId}')" title="Edit staff record">Edit</button>
            ${s.status !== 'ARCHIVED' ? `<button class="btn btn-xs btn-danger" onclick="AdminPage.archiveStaffConfirm('${s.employeeId}')" title="Archive/Delete staff record">Archive</button>` : `<button class="btn btn-xs btn-success" onclick="AdminPage.unarchiveStaff('${s.employeeId}')" title="Restore archived record">Restore</button>`}
          </td>
        </tr>`;
      }

      html += '</tbody></table>';
      container.innerHTML = html;

      const paginationEl = document.getElementById('hrmStaffPagination');
      if (paginationEl && pagination.totalPages > 1) {
        let paginationHtml = '';
        if (pagination.hasPrev) {
          paginationHtml += `<button class="btn btn-sm btn-secondary" onclick="AdminPage.loadStaffListPage(${pagination.page - 1})">Previous</button>`;
        }
        paginationHtml += `<span style="margin: 0 1rem;">Page ${pagination.page} of ${pagination.totalPages} (${pagination.total} total)</span>`;
        if (pagination.hasNext) {
          paginationHtml += `<button class="btn btn-sm btn-secondary" onclick="AdminPage.loadStaffListPage(${pagination.page + 1})">Next</button>`;
        }
        paginationEl.innerHTML = paginationHtml;
      } else if (paginationEl) {
        paginationEl.innerHTML = '';
      }
    } catch (err) {
      container.textContent = err.message || 'Failed to load staff records.';
    }
  }

  function loadStaffListPage(page) {
    loadStaffList(page);
  }

  /**
   * Export full staff records as CSV (opens in Excel). HRM Admin & Super Admin only.
   * Uses current filters: formationId, includeArchived, search query.
   */
  async function downloadStaffRecordsAsCsv() {
    if (!adminKey || (adminRole !== 'HRM_ADMIN' && adminRole !== 'SUPER_ADMIN')) return;
    const formationIdForExport = (adminRole === 'SUPER_ADMIN' || adminRole === 'HRM_ADMIN')
      ? (currentFormationId || adminFormationId || '')
      : (currentFormationId || adminFormationId || '');
    const includeArchived = document.getElementById('hrmIncludeArchivedCheck')?.checked || false;
    const query = document.getElementById('hrmStaffSearchInput')?.value.trim() || '';

    try {
      UI.showLoading('Exporting', 'Fetching full staff records...');
      const res = await Api.call('exportStaffRecords', {
        key: adminKey,
        formationId: formationIdForExport,
        includeArchived,
        query
      });
      UI.closeLoading();

      if (!res || !res.success || !res.data) {
        await UI.showError('Export failed', res?.message || 'No data returned.');
        return;
      }
      const results = res.data.results || [];
      const total = res.data.total || 0;
      if (total === 0) {
        await UI.showError('No data', 'No staff records to export.');
        return;
      }

      const escapeCsv = (val) => {
        if (val === null || val === undefined) return '';
        const s = String(val);
        if (s.includes(',') || s.includes('"') || s.includes('\n') || s.includes('\r')) {
          return '"' + s.replace(/"/g, '""') + '"';
        }
        return s;
      };

      const columns = [
        'employeeId', 'fileNumber', 'ippisNumber', 'surname', 'otherNames', 'email', 'telephone', 'dob',
        'dateOfFirstAppointment', 'dateOfPresentAppointment', 'confirmationDate',
        'cadre', 'rank', 'gradeLevel', 'stateOfOrigin', 'lga', 'qualification',
        'maritalStatus', 'spouseName', 'homeAddress', 'permanentHomeAddress', 'spouseAddress',
        'nextOfKin', 'nextOfKinAddress', 'formationId', 'subUnitId', 'status'
      ];
      const header = columns.join(',');
      const rows = results.map(r => columns.map(c => escapeCsv(r[c])).join(','));
      const csv = [header, ...rows].join('\r\n');
      const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'staff_records_' + new Date().toISOString().split('T')[0] + '.csv';
      a.click();
      URL.revokeObjectURL(url);

      await UI.showSuccess('Export complete', `Downloaded ${total} staff record(s) as CSV. You can open it in Excel.`);
    } catch (err) {
      UI.closeLoading();
      await UI.showError('Export failed', err.message || 'Could not export staff records.');
    }
  }

  async function viewStaffProfile(employeeId) {
    if (!adminKey) return;

    try {
      UI.showLoading('Loading', 'Fetching staff profile...');
      const res = await Api.call('getStaffById', {
        key: adminKey,
        employeeId: employeeId
        // formationId is optional - backend will find the record
      });

      if (!res || !res.success || !res.data || !res.data.staff) {
        UI.closeLoading();
        await UI.showError('Error', 'Staff record not found.');
        return;
      }

      const staff = res.data.staff;

      // Load formation and department names
      let formationName = 'Not Assigned';
      let departmentName = 'Not Assigned';
      try {
        if (staff.formationId) {
          const formRes = await Api.call('listFormations', { key: adminKey });
          if (formRes && formRes.success && formRes.data && formRes.data.formations) {
            const formation = formRes.data.formations.find(f => f.formationId === staff.formationId);
            if (formation) {
              formationName = formation.name || staff.formationId;
            }
          }
        }
        if (staff.subUnitId && staff.formationId) {
          const deptRes = await Api.call('listDepartments', {
            key: adminKey,
            formationId: staff.formationId
          });
          if (deptRes && deptRes.success && deptRes.data && deptRes.data.departments) {
            const dept = deptRes.data.departments.find(d => (d.departmentId || d.subUnitId) === staff.subUnitId);
            if (dept) {
              departmentName = dept.name || staff.subUnitId;
            }
          }
        }
      } catch (err) {
        console.warn('Could not load formation/department names:', err);
      }

      let documents = [];
      try {
        UI.showLoading('Loading', 'Fetching documents...');
        const docRes = await Api.call('listStaffDocuments', {
          key: adminKey,
          employeeId: employeeId
          // formationId is optional - backend will get from staff record if needed
        });
        UI.closeLoading();
        if (docRes && docRes.success && docRes.data && docRes.data.documents) {
          documents = docRes.data.documents;
        } else if (docRes && !docRes.success) {
          console.warn('Failed to load documents:', docRes.message || docRes.reason);
        }
      } catch (docErr) {
        UI.closeLoading();
        console.error('Error loading documents:', docErr);
        // Continue to display staff profile even if documents fail to load
      }

      // Use placeholder for doc images (Drive fileUrl/thumbnail in img src causes ERR_TOO_MANY_REDIRECTS for private files)
      const docPlaceholderSvg = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='160' height='140' viewBox='0 0 160 140'%3E%3Crect fill='%23e5e7eb' width='160' height='140'/%3E%3Ctext x='50%25' y='45%25' fill='%236b7280' font-size='14' text-anchor='middle'%3EDocument%3C/text%3E%3Ctext x='50%25' y='55%25' fill='%239ca3af' font-size='11' text-anchor='middle'%3EClick to open%3C/text%3E%3C/svg%3E";
      let documentsHtml = '';
      if (documents.length > 0) {
        documentsHtml = '<div style="margin-top: 1.5rem; border-top: 1px solid #e5e7eb; padding-top: 1rem;"><strong style="display: block; margin-bottom: 0.75rem; font-size: 1rem;">Attached Documents (' + documents.length + '/10):</strong>';
        documentsHtml += '<div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(160px, 1fr)); gap: 1rem; margin-top: 0.5rem;">';
        for (const doc of documents) {
          const fileSizeKB = doc.fileSize ? (doc.fileSize / 1024).toFixed(2) + 'KB' : 'N/A';
          const isImage = doc.mimeType && doc.mimeType.startsWith('image/');
          const downloadUrl = doc.driveFileId ? ('https://drive.google.com/uc?id=' + doc.driveFileId + '&export=download') : (doc.downloadUrl || '#');
          const fileName = doc.fileName || 'Document';
          const uploadDate = doc.uploadedAt ? new Date(doc.uploadedAt).toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' }) : 'N/A';
          const safeFileName = fileName.replace(/'/g, "\\'").replace(/"/g, '&quot;');

          if (isImage) {
            documentsHtml += `
              <div style="border: 1px solid #e5e7eb; border-radius: 0.5rem; padding: 0.5rem; text-align: center; background: #f9fafb; cursor: pointer; transition: transform 0.2s, box-shadow 0.2s;" 
                   onclick="window.open('${downloadUrl}', '_blank')"
                   onmouseover="this.style.transform='scale(1.02)'; this.style.boxShadow='0 4px 8px rgba(0,0,0,0.1)';"
                   onmouseout="this.style.transform='scale(1)'; this.style.boxShadow='none';">
                <img src="${docPlaceholderSvg}" alt="${safeFileName}" style="width: 100%; height: 140px; object-fit: cover; border-radius: 0.25rem; margin-bottom: 0.5rem; cursor: pointer;" />
                <div style="font-size: 0.75rem; color: #374151; word-break: break-word; font-weight: 500; margin-bottom: 0.25rem;" title="${safeFileName}">${fileName.length > 20 ? fileName.substring(0, 20) + '...' : fileName}</div>
                <div style="font-size: 0.65rem; color: #6b7280; margin-top: 0.25rem;">${fileSizeKB}</div>
                <div style="font-size: 0.6rem; color: #9ca3af; margin-top: 0.15rem;">${uploadDate}</div>
                <div style="font-size: 0.6rem; color: #9ca3af; margin-top: 0.1rem;">Click to view</div>
              </div>
            `;
          } else {
            documentsHtml += `
              <div style="border: 1px solid #e5e7eb; border-radius: 0.5rem; padding: 0.75rem; text-align: center; background: #f9fafb;">
                <a href="${downloadUrl}" target="_blank" rel="noopener" style="display: block; text-decoration: none; color: #059669;">
                  <div style="font-size: 2rem; margin-bottom: 0.5rem;">📄</div>
                  <div style="font-size: 0.75rem; color: #374151; word-break: break-word; font-weight: 500;">${fileName.length > 20 ? fileName.substring(0, 20) + '...' : fileName}</div>
                  <div style="font-size: 0.7rem; color: #6b7280; margin-top: 0.25rem;">${fileSizeKB}</div>
                  <div style="font-size: 0.65rem; color: #9ca3af; margin-top: 0.25rem;">${uploadDate}</div>
                </a>
              </div>
            `;
          }
        }
        documentsHtml += '</div></div>';
      } else {
        documentsHtml = '<div style="margin-top: 1.5rem; border-top: 1px solid #e5e7eb; padding-top: 1rem;"><em style="color: #6b7280;">No documents uploaded yet.</em></div>';
      }

      // Check if user can upload (HRM_ADMIN or SUPER_ADMIN)
      const canUpload = adminRole === 'HRM_ADMIN' || adminRole === 'SUPER_ADMIN';

      UI.closeLoading();
      const result = await Swal.fire({
        title: 'Staff Profile - Quick View',
        width: '600px',
        html: `
          <div style="text-align: left;">
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-bottom: 1.5rem;">
              <div><strong>File Number:</strong> ${staff.fileNumber || 'N/A'}</div>
              <div><strong>Surname:</strong> ${staff.surname || 'N/A'}</div>
              <div><strong>Other Names:</strong> ${staff.otherNames || 'N/A'}</div>
              <div><strong>Full Name:</strong> ${(staff.surname || '') + (staff.otherNames ? ' ' + staff.otherNames : '') || 'N/A'}</div>
              <div><strong>Email:</strong> ${staff.email || 'N/A'}</div>
              <div><strong>Telephone:</strong> ${staff.telephone ? String(staff.telephone) : 'N/A'}</div>
              <div><strong>Cadre:</strong> ${staff.cadre || 'N/A'}</div>
              <div><strong>Rank:</strong> ${staff.rank || 'N/A'}</div>
              <div><strong>Grade Level:</strong> ${(() => {
            const gradeLevel = staff.gradeLevel;
            if (!gradeLevel) return 'N/A';
            const gradeStr = String(gradeLevel);
            return gradeStr;
          })()}</div>
              <div><strong>Formation:</strong> ${formationName}</div>
              <div><strong>Department:</strong> ${departmentName}</div>
              <div><strong>Status:</strong> <span class="badge ${staff.status === 'ACTIVE' ? 'badge-success' : 'badge-warning'}">${staff.status || 'ACTIVE'}</span></div>
            </div>
            <div style="text-align: center; margin-top: 1.5rem; padding-top: 1.5rem; border-top: 1px solid #e5e7eb;">
              <button id="viewFullProfileBtn" class="btn btn-primary" style="padding: 0.75rem 2rem; font-size: 1rem;">
                View Full Profile
              </button>
            </div>
          </div>
        `,
        showConfirmButton: false,
        showCancelButton: true,
        cancelButtonText: 'Close',
        customClass: {
          popup: 'swal2-wide-modal',
          htmlContainer: 'swal2-html-container-wide'
        },
        didOpen: () => {
          const viewFullBtn = document.getElementById('viewFullProfileBtn');
          if (viewFullBtn) {
            viewFullBtn.addEventListener('click', () => {
              Swal.close();
              showFullStaffProfile(employeeId);
            });
          }
        }
      });
    } catch (err) {
      UI.closeLoading();
      await UI.showError('Error', err.message || 'Failed to load staff profile.');
    }
  }

  /** Fallback when no profile photo or image fails. Data URI only (no external file). */
  const PROFILE_PHOTO_FALLBACK = 'data:image/svg+xml,%3Csvg xmlns=%22http://www.w3.org/2000/svg%22 width=%22120%22 height=%22120%22 viewBox=%220 0 120 120%22%3E%3Crect fill=%22%23e5e7eb%22 width=%22120%22 height=%22120%22/%3E%3Ctext x=%2250%25%22 y=%2250%25%22 fill=%22%236b7280%22 font-size=%2214%22 text-anchor=%22middle%22 dy=%22.3em%22%3EProfile%3C/text%3E%3C/svg%3E';

  /**
   * Safe profile photo renderer. Only uses drive.google.com/thumbnail (no redirects).
   * All URLs in template literals. onerror fallback so UI never breaks.
   */
  function renderProfilePhoto(imgEl, fileId, sz) {
    if (!imgEl) return;
    const size = sz || 'w300';
    if (!fileId) {
      imgEl.src = PROFILE_PHOTO_FALLBACK;
      imgEl.onerror = null;
      return;
    }
    imgEl.src = `https://drive.google.com/thumbnail?id=${encodeURIComponent(fileId)}&sz=${size}`;
    imgEl.onerror = function () {
      imgEl.src = PROFILE_PHOTO_FALLBACK;
      imgEl.onerror = null;
    };
  }

  async function showFullStaffProfile(employeeId) {
    if (!adminKey) return;

    try {
      UI.showLoading('Loading', 'Fetching staff profile...');
      const res = await Api.call('getStaffProfile', {
        key: adminKey,
        employeeId: employeeId
      });

      if (!res || !res.success || !res.data || !res.data.staff) {
        UI.closeLoading();
        await UI.showError('Error', res?.message || 'Staff record not found.');
        return;
      }

      const staff = res.data.staff;
      const documents = res.data.documents || [];

      // Load formation and department names
      let formationName = 'Not Assigned';
      let departmentName = 'Not Assigned';
      try {
        if (staff.formationId) {
          const formRes = await Api.call('listFormations', { key: adminKey });
          if (formRes && formRes.success && formRes.data && formRes.data.formations) {
            const formation = formRes.data.formations.find(f => f.formationId === staff.formationId);
            if (formation) {
              formationName = formation.name || staff.formationId;
            }
          }
        }
        if (staff.subUnitId && staff.formationId) {
          const deptRes = await Api.call('listDepartments', {
            key: adminKey,
            formationId: staff.formationId
          });
          if (deptRes && deptRes.success && deptRes.data && deptRes.data.departments) {
            const dept = deptRes.data.departments.find(d => (d.departmentId || d.subUnitId) === staff.subUnitId);
            if (dept) {
              departmentName = dept.name || staff.subUnitId;
            }
          }
        }
      } catch (err) {
        console.warn('Could not load formation/department names:', err);
      }

      // Format date helper
      const formatDate = (dateValue) => {
        if (!dateValue) return 'N/A';
        try {
          if (typeof dateValue === 'string' && !dateValue.match(/^\d{4}-\d{2}-\d{2}/)) {
            return dateValue;
          }
          const date = new Date(dateValue);
          if (!isNaN(date.getTime())) {
            return date.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
          }
        } catch (e) { }
        return String(dateValue);
      };

      const canEdit = adminRole === 'HRM_ADMIN' || adminRole === 'SUPER_ADMIN';
      const canArchive = canEdit;

      UI.closeLoading();
      const result = await Swal.fire({
        title: 'Full Staff Profile',
        width: '1000px',
        html: `
          <div id="fullStaffProfileContent" style="text-align: left; max-height: 80vh; overflow-y: auto;">
            <div style="margin-bottom: 1.5rem; display: flex; align-items: center; gap: 1.5rem; padding: 1rem; background: #f9fafb; border-radius: 8px;">
              <div style="flex-shrink: 0; width: 120px; height: 120px; background: #e5e7eb; border-radius: 8px; overflow: hidden; display: flex; align-items: center; justify-content: center; color: #6b7280; font-size: 0.85rem;">
                <img id="fullStaffProfilePhotoImg" src="" alt="Profile" loading="lazy" style="width: 100%; height: 100%; object-fit: cover;" />
              </div>
              <div>
                <h3 style="margin: 0 0 0.25rem 0; color: #111;">${(staff.surname || '') + (staff.otherNames ? ' ' + staff.otherNames : '') || 'N/A'}</h3>
                <p style="margin: 0; color: #666; font-size: 0.9rem;">${staff.fileNumber ? 'File: ' + staff.fileNumber : ''} ${staff.cadre ? ' • ' + staff.cadre : ''}</p>
              </div>
            </div>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-bottom: 1.5rem;">
              <h3 style="grid-column: 1 / -1; margin-bottom: 0.5rem; color: #059669; border-bottom: 2px solid #059669; padding-bottom: 0.5rem;">Basic Information</h3>
              <div><strong>File Number:</strong> ${staff.fileNumber || 'N/A'}</div>
              <div><strong>IPPIS Number:</strong> ${staff.ippisNumber || 'N/A'}</div>
              <div><strong>Surname:</strong> ${staff.surname || 'N/A'}</div>
              <div><strong>Other Names:</strong> ${staff.otherNames || 'N/A'}</div>
              <div><strong>Full Name:</strong> ${(staff.surname || '') + (staff.otherNames ? ' ' + staff.otherNames : '') || 'N/A'}</div>
              <div><strong>Email:</strong> ${staff.email || 'N/A'}</div>
              <div><strong>Telephone:</strong> ${staff.telephone ? String(staff.telephone) : 'N/A'}</div>
              <div><strong>Date of Birth:</strong> ${formatDate(staff.dob)}</div>
              <div><strong>Cadre:</strong> ${staff.cadre || 'N/A'}</div>
              <div><strong>Rank:</strong> ${staff.rank || 'N/A'}</div>
              <div><strong>Grade Level:</strong> ${staff.gradeLevel ? String(staff.gradeLevel) : 'N/A'}</div>
              <div><strong>State of Origin:</strong> ${staff.stateOfOrigin || 'N/A'}</div>
              <div><strong>LGA:</strong> ${staff.lga || 'N/A'}</div>
              <div><strong>Qualification:</strong> ${staff.qualification || 'N/A'}</div>
              <div><strong>Formation:</strong> ${formationName}</div>
              <div><strong>Department/Sub-Unit:</strong> ${departmentName}</div>
              <div><strong>Status:</strong> <span class="badge ${staff.status === 'ACTIVE' ? 'badge-success' : 'badge-warning'}">${staff.status || 'ACTIVE'}</span></div>
            </div>
            
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-bottom: 1.5rem;">
              <h3 style="grid-column: 1 / -1; margin-bottom: 0.5rem; color: #059669; border-bottom: 2px solid #059669; padding-bottom: 0.5rem;">Appointment Information</h3>
              <div><strong>Date of First Appointment:</strong> ${formatDate(staff.dateOfFirstAppointment)}</div>
              <div><strong>Date of Present Appointment:</strong> ${formatDate(staff.dateOfPresentAppointment)}</div>
              <div><strong>Confirmation Date:</strong> ${formatDate(staff.confirmationDate)}</div>
            </div>

            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-bottom: 1.5rem;">
              <h3 style="grid-column: 1 / -1; margin-bottom: 0.5rem; color: #059669; border-bottom: 2px solid #059669; padding-bottom: 0.5rem;">Personal Information</h3>
              <div><strong>Marital Status:</strong> ${staff.maritalStatus || 'N/A'}</div>
              <div><strong>Spouse Name:</strong> ${staff.spouseName || 'N/A'}</div>
              <div style="grid-column: 1 / -1;"><strong>Home Address:</strong> ${staff.homeAddress || 'N/A'}</div>
              <div style="grid-column: 1 / -1;"><strong>Permanent Home Address:</strong> ${staff.permanentHomeAddress || 'N/A'}</div>
              <div style="grid-column: 1 / -1;"><strong>Spouse Address:</strong> ${staff.spouseAddress || 'N/A'}</div>
              <div><strong>Next of Kin:</strong> ${staff.nextOfKin || 'N/A'}</div>
              <div style="grid-column: 1 / -1;"><strong>Next of Kin Address:</strong> ${staff.nextOfKinAddress || 'N/A'}</div>
            </div>

            <div style="margin-bottom: 1.5rem;">
              <h3 style="margin-bottom: 0.5rem; color: #059669; border-bottom: 2px solid #059669; padding-bottom: 0.5rem;">Documents</h3>
              ${documents.length === 0
                ? '<p class="info info-muted">No documents uploaded.</p>'
                : `<div class="staff-doc-gallery" style="display: grid; grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); gap: 1rem; margin-top: 0.5rem;">
                ${documents.map(d => {
                  const downloadUrl = 'https://drive.google.com/uc?id=' + (d.driveFileId || '') + '&export=download';
                  const safeName = (d.fileName || 'Document').replace(/</g, '&lt;').replace(/"/g, '&quot;');
                  const docPlaceholderSvg = "data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='140' height='140' viewBox='0 0 140 140'%3E%3Crect fill='%23e5e7eb' width='140' height='140'/%3E%3Ctext x='50%25' y='45%25' fill='%236b7280' font-size='14' text-anchor='middle'%3EDocument%3C/text%3E%3Ctext x='50%25' y='55%25' fill='%239ca3af' font-size='11' text-anchor='middle'%3EClick to open%3C/text%3E%3C/svg%3E";
                  return `<a href="${downloadUrl}" target="_blank" rel="noopener" style="display: block; border: 1px solid #e5e7eb; border-radius: 8px; overflow: hidden; text-decoration: none; color: inherit;">
                  <div style="aspect-ratio: 1; background: #f3f4f6; overflow: hidden;">
                    <img src="${docPlaceholderSvg}" alt="${safeName}" style="width: 100%; height: 100%; object-fit: cover;" />
                  </div>
                  <div style="padding: 0.5rem; font-size: 0.8rem; color: #374151; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${safeName}</div>
                </a>`;
                }).join('')}
              </div>`}
            </div>

            <div style="text-align: center; margin-top: 2rem; padding-top: 1.5rem; border-top: 2px solid #e5e7eb;">
              <div style="display: flex; gap: 1rem; justify-content: center;">
                ${canEdit ? `<button id="editStaffBtn" class="btn btn-secondary" style="padding: 0.75rem 2rem;">Edit Record</button>` : ''}
                ${canArchive ? `<button id="archiveStaffBtn" class="btn ${staff.status === 'ACTIVE' ? 'btn-danger' : 'btn-success'}" style="padding: 0.75rem 2rem;">${staff.status === 'ACTIVE' ? 'Archive Record' : 'Restore Record'}</button>` : ''}
                ${canEdit ? `<button id="uploadDocBtn" class="btn btn-primary" style="padding: 0.75rem 2rem;">Upload Document</button>` : ''}
              </div>
            </div>
          </div>
        `,
        showConfirmButton: false,
        showCancelButton: true,
        cancelButtonText: 'Close',
        width: '1000px',
        customClass: {
          popup: 'swal2-wide-modal',
          htmlContainer: 'swal2-html-container-wide'
        },
        didOpen: () => {
          // Profile photo: safe renderer (thumbnail URL only, onerror fallback)
          const profileImg = document.getElementById('fullStaffProfilePhotoImg');
          if (profileImg) {
            renderProfilePhoto(profileImg, staff.profilePhotoFileId || null, 'w300');
          }
          // Scroll to top of content
          const contentDiv = document.getElementById('fullStaffProfileContent');
          if (contentDiv) {
            contentDiv.scrollTop = 0;
          }
          
          const editBtn = document.getElementById('editStaffBtn');
          const archiveBtn = document.getElementById('archiveStaffBtn');
          const uploadDocBtn = document.getElementById('uploadDocBtn');
          const viewDocsBtn = document.getElementById('viewDocsBtn');

          if (editBtn) {
            editBtn.addEventListener('click', () => {
              Swal.close();
              editStaff(employeeId);
            });
          }

          if (archiveBtn) {
            archiveBtn.addEventListener('click', async () => {
              Swal.close();
              await archiveStaffConfirm(employeeId);
              // Reload profile after archiving
              await showFullStaffProfile(employeeId);
            });
          }

          if (uploadDocBtn) {
            uploadDocBtn.addEventListener('click', async () => {
              Swal.close();
              await uploadEmployeeDocument(employeeId, staff.formationId || currentFormationId || '', staff.subUnitId || '');
              // Reload profile after upload
              await showFullStaffProfile(employeeId);
            });
          }

          if (viewDocsBtn) {
            viewDocsBtn.addEventListener('click', () => {
              Swal.close();
              showEmployeeDocuments(employeeId);
            });
          }
        }
      });
    } catch (err) {
      UI.closeLoading();
      await UI.showError('Error', err.message || 'Failed to load full staff profile.');
    }
  }

  async function editStaff(employeeId) {
    if (!adminKey) return;

    try {
      UI.showLoading('Loading', 'Fetching staff record...');
      const res = await Api.call('getStaffById', {
        key: adminKey,
        employeeId: employeeId
        // formationId is optional
      });

      if (!res || !res.success || !res.data || !res.data.staff) {
        UI.closeLoading();
        await UI.showError('Error', 'Staff record not found.');
        return;
      }

      const staff = res.data.staff;

      // Convert date to yyyy-mm-dd for type="date" inputs (accepts plain text like "25 Jan 2025" or ISO)
      function toDateInputValue(val) {
        if (!val) return '';
        if (/^\d{4}-\d{2}-\d{2}/.test(String(val))) return String(val).slice(0, 10);
        const d = new Date(val);
        return isNaN(d.getTime()) ? '' : d.toISOString().split('T')[0];
      }

      // Load formations and departments for selection
      let formations = [];
      let departments = [];

      try {
        UI.showLoading('Loading', 'Fetching formations and departments...');
        if (adminRole === 'SUPER_ADMIN' || adminRole === 'HRM_ADMIN') {
          const formRes = await Api.call('listFormations', { key: adminKey });
          if (formRes && formRes.success && formRes.data && formRes.data.formations) {
            formations = formRes.data.formations.filter(f => f.active !== false);
          }
        }

        // Load departments for current or selected formation
        const formIdForDept = staff.formationId || currentFormationId || adminFormationId || (formations.length > 0 ? formations[0].formationId : '');
        if (formIdForDept) {
          const deptRes = await Api.call('listDepartments', {
            key: adminKey,
            formationId: formIdForDept
          });
          if (deptRes && deptRes.success && deptRes.data && deptRes.data.departments) {
            departments = deptRes.data.departments;
          }
        }
        UI.closeLoading();
      } catch (err) {
        UI.closeLoading();
        console.warn('Could not load formations/departments:', err);
      }

      const result = await Swal.fire({
        title: 'Edit Staff Record',
        width: '800px',
        html: `
          ${(adminRole === 'SUPER_ADMIN' || adminRole === 'HRM_ADMIN') ? `
          <div style="margin-bottom: 1rem; padding: 1rem; border: 1px solid #e5e7eb; border-radius: 8px; background: #f9fafb;">
            <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Profile Picture</label>
            <div style="display: flex; align-items: center; gap: 1rem; flex-wrap: wrap;">
              <div style="width: 80px; height: 80px; background: #e5e7eb; border-radius: 8px; overflow: hidden; flex-shrink: 0; display: flex; align-items: center; justify-content: center;">
                <img id="swalEditProfilePhotoImg" src="" alt="Profile" loading="lazy" style="width: 100%; height: 100%; object-fit: cover;" />
              </div>
              <div>
                <input type="file" id="swalEditProfilePhoto" accept="image/jpeg,image/png,image/webp,image/jpg" style="font-size: 0.9rem;">
                <p style="margin: 0.25rem 0 0 0; font-size: 0.8rem; color: #666;">JPEG, PNG or WebP. Max 2MB.</p>
              </div>
            </div>
          </div>
          ` : ''}
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1rem; margin-bottom: 1rem;">
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">File Number</label>
              <input id="swalEditFileNumber" class="swal2-input" placeholder="File Number" value="${staff.fileNumber || ''}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">IPPIS Number</label>
              <input id="swalEditIppis" class="swal2-input" placeholder="IPPIS Number" value="${staff.ippisNumber || ''}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Surname *</label>
              <input id="swalEditSurname" class="swal2-input" placeholder="Surname *" value="${staff.surname || ''}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Other Names</label>
              <input id="swalEditOtherNames" class="swal2-input" placeholder="Other Names" value="${staff.otherNames || ''}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Email</label>
              <input id="swalEditEmail" class="swal2-input" type="email" placeholder="Email" value="${staff.email || ''}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Telephone Number</label>
              <input id="swalEditTelephone" class="swal2-input" type="tel" placeholder="Telephone Number" value="${String(staff.telephone || '')}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            ${(adminRole === 'SUPER_ADMIN' || adminRole === 'HRM_ADMIN') && formations.length > 0 ? `
            <div style="grid-column: 1 / -1;">
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Formation (can be changed)</label>
              <select id="swalEditFormation" class="swal2-select" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
                <option value="">-- No Formation --</option>
                ${formations.map(form => `<option value="${form.formationId}" ${staff.formationId && form.formationId === staff.formationId ? 'selected' : ''}>${form.name || form.formationId}</option>`).join('')}
              </select>
            </div>
            ` : ''}
            <div style="grid-column: 1 / -1;">
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Department/Sub-Unit (can be changed)</label>
              <select id="swalEditSubUnit" class="swal2-select" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
                <option value="">-- No Department --</option>
                ${departments.map(dept => `<option value="${dept.departmentId || dept.subUnitId}" ${staff.subUnitId && (dept.departmentId === staff.subUnitId || dept.subUnitId === staff.subUnitId) ? 'selected' : ''}>${dept.name || dept.departmentId || dept.subUnitId}</option>`).join('')}
              </select>
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Date of Birth</label>
              <input id="swalEditDob" class="swal2-input" type="date" placeholder="Date of Birth" value="${toDateInputValue(staff.dob)}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Date of First Appointment</label>
              <input id="swalEditFirstAppointment" class="swal2-input" type="date" placeholder="Date of First Appointment" value="${toDateInputValue(staff.dateOfFirstAppointment)}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Date of Present Appointment</label>
              <input id="swalEditPresentAppointment" class="swal2-input" type="date" placeholder="Date of Present Appointment" value="${toDateInputValue(staff.dateOfPresentAppointment)}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Confirmation Date</label>
              <input id="swalEditConfirmationDate" class="swal2-input" type="date" placeholder="Confirmation Date" value="${toDateInputValue(staff.confirmationDate)}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Cadre</label>
              <input id="swalEditCadre" class="swal2-input" placeholder="Cadre" value="${staff.cadre || ''}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Rank</label>
              <input id="swalEditRank" class="swal2-input" placeholder="Rank" value="${staff.rank || ''}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Grade Level</label>
              <input id="swalEditGradeLevel" class="swal2-input" placeholder="Grade Level" value="${staff.gradeLevel || ''}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">State of Origin</label>
              <input id="swalEditStateOfOrigin" class="swal2-input" placeholder="State of Origin" value="${staff.stateOfOrigin || ''}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">LGA</label>
              <input id="swalEditLga" class="swal2-input" placeholder="LGA" value="${staff.lga || ''}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div style="grid-column: 1 / -1;">
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Qualification</label>
              <input id="swalEditQualification" class="swal2-input" placeholder="Qualification" value="${staff.qualification || ''}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div style="grid-column: 1 / -1;">
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Marital Status</label>
              <select id="swalEditMaritalStatus" class="swal2-select" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
                <option value="">-- Select Marital Status --</option>
                <option value="Single" ${staff.maritalStatus === 'Single' ? 'selected' : ''}>Single</option>
                <option value="Married" ${staff.maritalStatus === 'Married' ? 'selected' : ''}>Married</option>
                <option value="Divorced" ${staff.maritalStatus === 'Divorced' ? 'selected' : ''}>Divorced</option>
                <option value="Widowed" ${staff.maritalStatus === 'Widowed' ? 'selected' : ''}>Widowed</option>
              </select>
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Spouse Name</label>
              <input id="swalEditSpouseName" class="swal2-input" placeholder="Spouse Name" value="${staff.spouseName || ''}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div style="grid-column: 1 / -1;">
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Home Address</label>
              <textarea id="swalEditHomeAddress" class="swal2-textarea" placeholder="Home Address" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem; min-height: 80px;">${staff.homeAddress || ''}</textarea>
            </div>
            <div style="grid-column: 1 / -1;">
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Permanent Home Address</label>
              <textarea id="swalEditPermanentAddress" class="swal2-textarea" placeholder="Permanent Home Address" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem; min-height: 80px;">${staff.permanentHomeAddress || ''}</textarea>
            </div>
            <div style="grid-column: 1 / -1;">
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Spouse Address</label>
              <textarea id="swalEditSpouseAddress" class="swal2-textarea" placeholder="Spouse Address" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem; min-height: 80px;">${staff.spouseAddress || ''}</textarea>
            </div>
            <div>
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Next of Kin</label>
              <input id="swalEditNextOfKin" class="swal2-input" placeholder="Next of Kin Name" value="${staff.nextOfKin || ''}" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
            </div>
            <div style="grid-column: 1 / -1;">
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Next of Kin Address</label>
              <textarea id="swalEditNextOfKinAddress" class="swal2-textarea" placeholder="Next of Kin Address" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem; min-height: 80px;">${staff.nextOfKinAddress || ''}</textarea>
            </div>
            <div style="grid-column: 1 / -1;">
              <label style="display: block; margin-bottom: 0.5rem; font-weight: 600; color: #333;">Status</label>
              <select id="swalEditStatus" class="swal2-select" style="width: 100%; padding: 0.75rem; border: 1px solid #ddd; border-radius: 4px; font-size: 1rem;">
                <option value="ACTIVE" ${staff.status === 'ACTIVE' ? 'selected' : ''}>Active</option>
                <option value="ARCHIVED" ${staff.status === 'ARCHIVED' ? 'selected' : ''}>Archived</option>
              </select>
            </div>
          </div>
        `,
        customClass: {
          popup: 'swal2-wide-modal',
          htmlContainer: 'swal2-html-container-wide'
        },
        didOpen: () => {
          // Load departments when formation is selected
          const formationSelect = document.getElementById('swalEditFormation');
          const subUnitSelect = document.getElementById('swalEditSubUnit');

          if (formationSelect && subUnitSelect) {
            formationSelect.addEventListener('change', async function () {
              const selectedFormationId = this.value;
              if (selectedFormationId) {
                try {
                  UI.showLoading('Loading', 'Fetching departments...');
                  const deptRes = await Api.call('listDepartments', {
                    key: adminKey,
                    formationId: selectedFormationId
                  });
                  UI.closeLoading();
                  if (deptRes && deptRes.success && deptRes.data && deptRes.data.departments) {
                    subUnitSelect.innerHTML = '<option value="">-- No Department --</option>';
                    deptRes.data.departments.forEach(dept => {
                      const option = document.createElement('option');
                      option.value = dept.departmentId || dept.subUnitId;
                      option.textContent = dept.name || dept.departmentId || dept.subUnitId;
                      subUnitSelect.appendChild(option);
                    });
                  }
                } catch (err) {
                  UI.closeLoading();
                  console.warn('Could not load departments:', err);
                }
              } else {
                subUnitSelect.innerHTML = '<option value="">-- No Department --</option>';
              }
            });
          }
        },
        focusConfirm: false,
        showCancelButton: true,
        showConfirmButton: false, // We'll add custom button
        cancelButtonText: 'Cancel',
        didOpen: () => {
          // Profile photo: safe renderer (thumbnail URL only, onerror fallback)
          const editProfileImg = document.getElementById('swalEditProfilePhotoImg');
          if (editProfileImg) {
            renderProfilePhoto(editProfileImg, staff.profilePhotoFileId || null, 'w80');
          }
          // Hide default confirm button if it exists
          const confirmButton = Swal.getConfirmButton();
          if (confirmButton) {
            confirmButton.style.display = 'none';
          }
          
          // Create custom Update Record button
          setTimeout(() => {
            const actionsDiv = Swal.getActions();
            if (actionsDiv) {
              // Check if custom button already exists
              let updateBtn = document.getElementById('swalCustomUpdateBtn');
              if (!updateBtn) {
                updateBtn = document.createElement('button');
                updateBtn.id = 'swalCustomUpdateBtn';
                updateBtn.className = 'swal2-confirm swal2-styled';
                updateBtn.style.backgroundColor = '#059669';
                updateBtn.style.marginRight = '0.5em';
                updateBtn.textContent = 'Update Record';
                
                updateBtn.addEventListener('click', async () => {
                  // Collect form data before closing modal
                  const fileNumberEl = document.getElementById('swalEditFileNumber');
                  const ippisNumberEl = document.getElementById('swalEditIppis');
                  const surnameEl = document.getElementById('swalEditSurname');
                  const otherNamesEl = document.getElementById('swalEditOtherNames');
                  const emailEl = document.getElementById('swalEditEmail');
                  const telephoneEl = document.getElementById('swalEditTelephone');
                  const formationIdEl = document.getElementById('swalEditFormation');
                  const subUnitIdEl = document.getElementById('swalEditSubUnit');
                  const dobEl = document.getElementById('swalEditDob');
                  const firstAppointmentEl = document.getElementById('swalEditFirstAppointment');
                  const presentAppointmentEl = document.getElementById('swalEditPresentAppointment');
                  const confirmationDateEl = document.getElementById('swalEditConfirmationDate');
                  const cadreEl = document.getElementById('swalEditCadre');
                  const rankEl = document.getElementById('swalEditRank');
                  const gradeLevelEl = document.getElementById('swalEditGradeLevel');
                  const stateOfOriginEl = document.getElementById('swalEditStateOfOrigin');
                  const lgaEl = document.getElementById('swalEditLga');
                  const qualificationEl = document.getElementById('swalEditQualification');
                  const maritalStatusEl = document.getElementById('swalEditMaritalStatus');
                  const spouseNameEl = document.getElementById('swalEditSpouseName');
                  const homeAddressEl = document.getElementById('swalEditHomeAddress');
                  const permanentAddressEl = document.getElementById('swalEditPermanentAddress');
                  const spouseAddressEl = document.getElementById('swalEditSpouseAddress');
                  const nextOfKinEl = document.getElementById('swalEditNextOfKin');
                  const nextOfKinAddressEl = document.getElementById('swalEditNextOfKinAddress');
                  const statusEl = document.getElementById('swalEditStatus');

                  if (!surnameEl || !surnameEl.value.trim()) {
                    Swal.showValidationMessage('Surname is required.');
                    return;
                  }

                  const fileNumber = fileNumberEl ? fileNumberEl.value.trim() : '';
                  const ippisNumber = ippisNumberEl ? ippisNumberEl.value.trim() : '';
                  const surname = surnameEl ? surnameEl.value.trim() : '';
                  const otherNames = otherNamesEl ? otherNamesEl.value.trim() : '';
                  const email = emailEl ? emailEl.value.trim() : '';
                  const telephone = telephoneEl ? String(telephoneEl.value.trim()) : ''; // Ensure string
                  const formationId = formationIdEl ? (formationIdEl.value.trim() || undefined) : undefined;
                  const subUnitId = subUnitIdEl ? (subUnitIdEl.value.trim() || undefined) : undefined;
                  const dob = dobEl ? dobEl.value.trim() : '';
                  const firstAppointment = firstAppointmentEl ? firstAppointmentEl.value.trim() : '';
                  const presentAppointment = presentAppointmentEl ? presentAppointmentEl.value.trim() : '';
                  const confirmationDate = confirmationDateEl ? confirmationDateEl.value.trim() : '';
                  const cadre = cadreEl ? cadreEl.value.trim() : '';
                  const rank = rankEl ? rankEl.value.trim() : '';
                  const gradeLevel = gradeLevelEl ? gradeLevelEl.value.trim() : '';
                  const stateOfOrigin = stateOfOriginEl ? stateOfOriginEl.value.trim() : '';
                  const lga = lgaEl ? lgaEl.value.trim() : '';
                  const qualification = qualificationEl ? qualificationEl.value.trim() : '';
                  const maritalStatus = maritalStatusEl ? maritalStatusEl.value.trim() : '';
                  const spouseName = spouseNameEl ? spouseNameEl.value.trim() : '';
                  const homeAddress = homeAddressEl ? homeAddressEl.value.trim() : '';
                  const permanentAddress = permanentAddressEl ? permanentAddressEl.value.trim() : '';
                  const spouseAddress = spouseAddressEl ? spouseAddressEl.value.trim() : '';
                  const nextOfKin = nextOfKinEl ? nextOfKinEl.value.trim() : '';
                  const nextOfKinAddress = nextOfKinAddressEl ? nextOfKinAddressEl.value.trim() : '';
                  const status = statusEl ? statusEl.value : 'ACTIVE';

                  try {
                    Swal.close();
                    UI.showLoading('Updating', 'Saving staff record...');
                    const updateRes = await Api.call('updateStaff', {
                      key: adminKey,
                      staff: {
                        employeeId: employeeId,
                        formationId: formationId,
                        subUnitId: subUnitId,
                        fileNumber: fileNumber || '',
                        ippisNumber: ippisNumber || '',
                        email: email || '',
                        telephone: telephone || '',
                        surname: surname,
                        otherNames: otherNames || '',
                        dob: dob || '',
                        dateOfFirstAppointment: firstAppointment || '',
                        dateOfPresentAppointment: presentAppointment || '',
                        confirmationDate: confirmationDate || '',
                        cadre: cadre || '',
                        rank: rank || '',
                        gradeLevel: gradeLevel || '',
                        stateOfOrigin: stateOfOrigin || '',
                        lga: lga || '',
                        qualification: qualification || '',
                        maritalStatus: maritalStatus || '',
                        spouseName: spouseName || '',
                        homeAddress: homeAddress || '',
                        permanentHomeAddress: permanentAddress || '',
                        spouseAddress: spouseAddress || '',
                        nextOfKin: nextOfKin || '',
                        nextOfKinAddress: nextOfKinAddress || '',
                        status
                      }
                    });

                    if (updateRes && updateRes.success) {
                      // Optional: upload new profile picture if selected (HRM Admin / Super Admin)
                      const photoInput = document.getElementById('swalEditProfilePhoto');
                      if (photoInput && photoInput.files && photoInput.files[0] && (adminRole === 'SUPER_ADMIN' || adminRole === 'HRM_ADMIN')) {
                        const file = photoInput.files[0];
                        if (file.size <= 2 * 1024 * 1024 && ['image/jpeg', 'image/jpg', 'image/png', 'image/webp'].includes(file.type)) {
                          try {
                            UI.showLoading('Uploading', 'Uploading profile picture...');
                            const base64 = await new Promise((resolve, reject) => {
                              const reader = new FileReader();
                              reader.onload = () => resolve(reader.result);
                              reader.onerror = () => reject(new Error('Failed to read file'));
                              reader.readAsDataURL(file);
                            });
                            await Api.call('uploadStaffProfilePicture', {
                              key: adminKey,
                              employeeId,
                              formationId: formationId || staff.formationId || undefined,
                              subUnitId: subUnitId || staff.subUnitId || undefined,
                              fileData: base64,
                              fileName: file.name,
                              mimeType: file.type,
                              fileSize: file.size
                            });
                          } catch (picErr) {
                            console.warn('Profile picture upload failed:', picErr);
                          }
                        }
                      }
                      UI.closeLoading();
                      await UI.showSuccess('Success', 'Staff record updated successfully!');
                      await loadHrmStaffStats();
                      await loadStaffList(currentStaffPage);
                    } else {
                      UI.closeLoading();
                      throw new Error(updateRes.message || 'Failed to update staff record.');
                    }
                  } catch (err) {
                    UI.closeLoading();
                    await UI.showError('Error', err.message || 'Failed to update staff record.');
                  }
                });
                
                // Insert before cancel button
                const cancelBtn = Swal.getCancelButton();
                if (cancelBtn && cancelBtn.parentNode) {
                  cancelBtn.parentNode.insertBefore(updateBtn, cancelBtn);
                } else if (actionsDiv) {
                  actionsDiv.appendChild(updateBtn);
                }
              }
            }
          }, 100);
        }
      });
    } catch (err) {
      UI.closeLoading();
      await UI.showError('Error', err.message || 'Failed to load staff record.');
    }
  }

  async function archiveStaffConfirm(employeeId) {
    if (!adminKey) return;

    const confirmed = await UI.confirmAction(
      'Archive Staff Record',
      'Are you sure you want to archive this staff record? This action can be reversed by editing the record.',
      'Archive',
      'Cancel'
    );

    if (!confirmed) return;

    try {
      const res = await Api.call('archiveStaff', {
        key: adminKey,
        employeeId: employeeId
        // formationId is optional - backend will find the record
      });

      if (res && res.success) {
        await UI.showSuccess('Success', 'Staff record archived successfully!');
        await loadHrmStaffStats();
        await loadStaffList(currentStaffPage);
      } else {
        throw new Error(res.message || 'Failed to archive staff record.');
      }
    } catch (err) {
      await UI.showError('Error', err.message || 'Failed to archive staff record.');
    }
  }

  // Export functions (assign to window.AdminPage, not const AdminPage)
  if (!window.AdminPage) {
    window.AdminPage = {};
  }
  window.AdminPage.loadStaffList = loadStaffList;
  window.AdminPage.loadStaffListPage = loadStaffListPage;
  window.AdminPage.viewStaffProfile = viewStaffProfile;
  window.AdminPage.showFullStaffProfile = showFullStaffProfile;
  window.AdminPage.editStaff = editStaff;
  window.AdminPage.archiveStaffConfirm = archiveStaffConfirm;
  window.AdminPage.loadHrmStaffStats = loadHrmStaffStats;
  window.AdminPage.showAddStaffModal = showAddStaffModal;

  // ============================================================================
  // HRM ADMIN MANAGEMENT FUNCTIONS (SUPER_ADMIN ONLY)
  // ============================================================================

  /**
   * Load HRM Admins table
   */
  async function loadHrmAdminsTable() {
    const container = document.getElementById('hrmAdminsTable');
    if (!container || !adminKey || adminRole !== 'SUPER_ADMIN') return;

    container.textContent = 'Loading HRM admins...';

    try {
      const res = await Api.call('listHrmAdmins', { key: adminKey });

      if (!res || !res.success || !res.data) {
        container.textContent = 'No HRM admins found.';
        return;
      }

      const admins = res.data.admins || [];

      if (admins.length === 0) {
        container.textContent = 'No HRM admins found.';
        return;
      }

      let html = '<table><thead><tr><th>S/N</th><th>Email</th><th>Role</th><th>Status</th><th>Actions</th></tr></thead><tbody>';

      for (let idx = 0; idx < admins.length; idx++) {
        const admin = admins[idx];
        html += `<tr>
          <td>${idx + 1}</td>
          <td>${admin.email || ''}</td>
          <td><span class="badge badge-info">${admin.role || ''}</span></td>
          <td><span class="badge ${admin.active ? 'badge-success' : 'badge-warning'}">${admin.active ? 'Active' : 'Inactive'}</span></td>
          <td>
            <button class="btn btn-xs btn-secondary" onclick="AdminPage.editHrmAdminRole('${admin.adminId}', '${admin.role || ''}')">Change Role</button>
            ${admin.active ?
            `<button class="btn btn-xs btn-danger" onclick="AdminPage.deactivateHrmAdmin('${admin.adminId}')">Deactivate</button>` :
            `<button class="btn btn-xs btn-success" onclick="AdminPage.activateHrmAdmin('${admin.adminId}')">Activate</button>`
          }
          </td>
        </tr>`;
      }

      html += '</tbody></table>';
      container.innerHTML = html;
    } catch (err) {
      container.textContent = err.message || 'Failed to load HRM admins.';
    }
  }

  /**
   * Show Create HRM Admin modal
   */
  async function showCreateHrmAdminModal() {
    if (!adminKey || adminRole !== 'SUPER_ADMIN') {
      await UI.showError('Access Denied', 'Only SUPER_ADMIN can create HRM admins.');
      return;
    }

    if (!window.Swal) {
      await UI.showError('Unavailable', 'SweetAlert2 is not loaded.');
      return;
    }

    const result = await Swal.fire({
      title: 'Create HRM Admin',
      html: `
        <input id="swalHrmAdminId" class="swal2-input" placeholder="Admin ID *">
        <input id="swalHrmAdminEmail" class="swal2-input" type="email" placeholder="Email Address *">
        <select id="swalHrmAdminRole" class="swal2-select" style="width: 100%;">
          <option value="">Select Role *</option>
          <option value="HRM_ADMIN">HRM Admin (Full CRUD)</option>
          <option value="HRM_VIEWER">HRM Viewer (Read-Only)</option>
        </select>
      `,
      focusConfirm: false,
      showCancelButton: true,
      confirmButtonText: 'Create HRM Admin',
      cancelButtonText: 'Cancel',
      confirmButtonColor: '#059669',
      preConfirm: async () => {
        const adminId = document.getElementById('swalHrmAdminId').value.trim();
        const email = document.getElementById('swalHrmAdminEmail').value.trim();
        const role = document.getElementById('swalHrmAdminRole').value;

        if (!adminId) {
          Swal.showValidationMessage('Admin ID is required.');
          return false;
        }
        if (!email) {
          Swal.showValidationMessage('Email is required.');
          return false;
        }
        if (!role) {
          Swal.showValidationMessage('Role is required.');
          return false;
        }

        const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
        if (!emailRegex.test(email)) {
          Swal.showValidationMessage('Please enter a valid email address.');
          return false;
        }

        try {
          const res = await Api.call('createHrmAdmin', {
            key: adminKey,
            adminId: adminId,
            email: email,
            role: role
          });

          if (res && res.success) {
            return { success: true };
          } else {
            throw new Error(res.message || 'Failed to create HRM admin.');
          }
        } catch (err) {
          Swal.showValidationMessage(err.message || 'Failed to create HRM admin.');
          return false;
        }
      }
    });

    if (result.isConfirmed && result.value && result.value.success) {
      await UI.showSuccess('Success', 'HRM admin created successfully!');
      await loadHrmAdminsTable();
    }
  }

  /**
   * Edit HRM Admin Role
   */
  async function editHrmAdminRole(adminId, currentRole) {
    if (!adminKey || adminRole !== 'SUPER_ADMIN') {
      await UI.showError('Access Denied', 'Only SUPER_ADMIN can update HRM admin roles.');
      return;
    }

    if (!window.Swal) {
      await UI.showError('Unavailable', 'SweetAlert2 is not loaded.');
      return;
    }

    const result = await Swal.fire({
      title: 'Change HRM Admin Role',
      html: `
        <p style="text-align: left; margin-bottom: 1rem;"><strong>Admin ID:</strong> ${adminId}</p>
        <select id="swalHrmAdminRoleEdit" class="swal2-select" style="width: 100%;">
          <option value="HRM_ADMIN" ${currentRole === 'HRM_ADMIN' ? 'selected' : ''}>HRM Admin (Full CRUD)</option>
          <option value="HRM_VIEWER" ${currentRole === 'HRM_VIEWER' ? 'selected' : ''}>HRM Viewer (Read-Only)</option>
        </select>
      `,
      focusConfirm: false,
      showCancelButton: true,
      confirmButtonText: 'Update Role',
      cancelButtonText: 'Cancel',
      confirmButtonColor: '#059669',
      preConfirm: async () => {
        const role = document.getElementById('swalHrmAdminRoleEdit').value;

        if (!role) {
          Swal.showValidationMessage('Role is required.');
          return false;
        }

        try {
          const res = await Api.call('updateHrmAdmin', {
            key: adminKey,
            adminId: adminId,
            role: role
          });

          if (res && res.success) {
            return { success: true };
          } else {
            throw new Error(res.message || 'Failed to update HRM admin role.');
          }
        } catch (err) {
          Swal.showValidationMessage(err.message || 'Failed to update HRM admin role.');
          return false;
        }
      }
    });

    if (result.isConfirmed && result.value && result.value.success) {
      await UI.showSuccess('Success', 'HRM admin role updated successfully!');
      await loadHrmAdminsTable();
    }
  }

  /**
   * Deactivate HRM Admin
   */
  async function deactivateHrmAdmin(adminId) {
    if (!adminKey || adminRole !== 'SUPER_ADMIN') {
      await UI.showError('Access Denied', 'Only SUPER_ADMIN can deactivate HRM admins.');
      return;
    }

    const confirmed = await UI.confirmAction(
      'Deactivate HRM Admin',
      `Are you sure you want to deactivate HRM admin "${adminId}"? They will lose access to HRM functions.`,
      'Deactivate',
      'Cancel'
    );

    if (!confirmed) return;

    try {
      const res = await Api.call('deactivateHrmAdmin', {
        key: adminKey,
        adminId: adminId
      });

      if (res && res.success) {
        await UI.showSuccess('Success', 'HRM admin deactivated successfully!');
        await loadHrmAdminsTable();
      } else {
        throw new Error(res.message || 'Failed to deactivate HRM admin.');
      }
    } catch (err) {
      await UI.showError('Error', err.message || 'Failed to deactivate HRM admin.');
    }
  }

  /**
   * Activate HRM Admin
   */
  async function activateHrmAdmin(adminId) {
    if (!adminKey || adminRole !== 'SUPER_ADMIN') {
      await UI.showError('Access Denied', 'Only SUPER_ADMIN can activate HRM admins.');
      return;
    }

    const confirmed = await UI.confirmAction(
      'Activate HRM Admin',
      `Are you sure you want to activate HRM admin "${adminId}"?`,
      'Activate',
      'Cancel'
    );

    if (!confirmed) return;

    try {
      const res = await Api.call('activateHrmAdmin', {
        key: adminKey,
        adminId: adminId
      });

      if (res && res.success) {
        await UI.showSuccess('Success', 'HRM admin activated successfully!');
        await loadHrmAdminsTable();
      } else {
        throw new Error(res.message || 'Failed to activate HRM admin.');
      }
    } catch (err) {
      await UI.showError('Error', err.message || 'Failed to activate HRM admin.');
    }
  }

  // Export HRM Admin functions (assign to window.AdminPage, not const AdminPage)
  if (!window.AdminPage) {
    window.AdminPage = {};
  }
  window.AdminPage.loadHrmAdminsTable = loadHrmAdminsTable;
  window.AdminPage.showCreateHrmAdminModal = showCreateHrmAdminModal;
  window.AdminPage.editHrmAdminRole = editHrmAdminRole;
  window.AdminPage.deactivateHrmAdmin = deactivateHrmAdmin;
  window.AdminPage.activateHrmAdmin = activateHrmAdmin;

  // ============================================================================
  // HRM AUDIT LOGS FUNCTIONS (SUPER_ADMIN ONLY)
  // ============================================================================

  let currentHrmLogsPage = 1;
  let currentHrmLogsLimit = 50;

  /**
   * Load HRM Audit Logs
   */
  async function loadHrmAuditLogs(page = null) {
    if (!adminKey || adminRole !== 'SUPER_ADMIN') return;

    if (page !== null) {
      currentHrmLogsPage = page;
    }

    const container = document.getElementById('hrmAuditLogsTable');
    if (!container) return;

    container.textContent = 'Loading HRM audit logs...';

    const searchQuery = document.getElementById('hrmLogsSearchInput')?.value.trim() || '';
    const actionFilter = document.getElementById('hrmLogsActionFilter')?.value || '';
    const statusFilter = document.getElementById('hrmLogsStatusFilter')?.value || '';

    try {
      const res = await Api.call('getHrmAuditLogs', {
        key: adminKey,
        limit: currentHrmLogsLimit,
        actor: searchQuery || undefined,
        actionType: actionFilter || undefined,
        status: statusFilter || undefined
      });

      if (!res || !res.success || !res.data) {
        container.textContent = 'No HRM audit logs found.';
        return;
      }

      const logs = res.data.logs || [];

      if (logs.length === 0) {
        container.textContent = 'No HRM audit logs found matching your filters.';
        const paginationEl = document.getElementById('hrmLogsPagination');
        if (paginationEl) paginationEl.innerHTML = '';
        return;
      }

      // Build table
      let html = '<table><thead><tr><th>Timestamp</th><th>Action</th><th>Actor</th><th>Role</th><th>Formation</th><th>Status</th><th>Details</th></tr></thead><tbody>';

      for (const log of logs) {
        const timestamp = log.timestamp ? new Date(log.timestamp).toLocaleString() : 'N/A';
        const statusBadge = log.status === 'SUCCESS' ? 'badge-success' :
          log.status === 'FAILED' ? 'badge-danger' :
            log.status === 'REJECTED' ? 'badge-warning' : 'badge-info';

        // Format extra details
        let extraText = '';
        if (log.extra) {
          if (typeof log.extra === 'object') {
            extraText = log.extra.extra || JSON.stringify(log.extra);
          } else {
            extraText = log.extra;
          }
        }
        if (extraText.length > 100) {
          extraText = extraText.substring(0, 100) + '...';
        }

        html += `<tr>
          <td>${timestamp}</td>
          <td><span class="badge badge-info">${log.actionType || 'N/A'}</span></td>
          <td>${log.actor || 'N/A'}</td>
          <td>${log.actorRole || 'N/A'}</td>
          <td>${log.formationId || 'N/A'}</td>
          <td><span class="badge ${statusBadge}">${log.status || 'N/A'}</span></td>
          <td style="max-width: 300px; overflow: hidden; text-overflow: ellipsis;" title="${extraText}">${extraText || '-'}</td>
        </tr>`;
      }

      html += '</tbody></table>';
      container.innerHTML = html;

      // Pagination (if needed)
      const total = logs.length;
      const totalPages = Math.ceil(total / currentHrmLogsLimit);
      const paginationEl = document.getElementById('hrmLogsPagination');

      if (paginationEl && totalPages > 1) {
        let paginationHtml = '';
        if (currentHrmLogsPage > 1) {
          paginationHtml += `<button class="btn btn-sm btn-secondary" onclick="AdminPage.loadHrmLogsPage(${currentHrmLogsPage - 1})">Previous</button>`;
        }
        paginationHtml += `<span style="margin: 0 1rem;">Page ${currentHrmLogsPage} of ${totalPages} (${total} logs)</span>`;
        if (currentHrmLogsPage < totalPages) {
          paginationHtml += `<button class="btn btn-sm btn-secondary" onclick="AdminPage.loadHrmLogsPage(${currentHrmLogsPage + 1})">Next</button>`;
        }
        paginationEl.innerHTML = paginationHtml;
      } else if (paginationEl) {
        paginationEl.innerHTML = `<span>Showing ${total} logs</span>`;
      }
    } catch (err) {
      container.textContent = err.message || 'Failed to load HRM audit logs.';
    }
  }

  function loadHrmLogsPage(page) {
    loadHrmAuditLogs(page);
  }

  /**
   * Export HRM Audit Logs
   */
  async function exportHrmAuditLogs() {
    if (!adminKey || adminRole !== 'SUPER_ADMIN') {
      await UI.showError('Access Denied', 'Only SUPER_ADMIN can export HRM audit logs.');
      return;
    }

    try {
      const res = await Api.call('getHrmAuditLogs', {
        key: adminKey,
        limit: 10000 // Get more for export
      });

      if (!res || !res.success || !res.data || !res.data.logs) {
        await UI.showError('Error', 'No logs to export.');
        return;
      }

      const logs = res.data.logs;

      // Create CSV content
      let csv = 'Timestamp,Action Type,Actor,Role,Formation ID,Status,Reason,Extra Details\n';

      for (const log of logs) {
        const timestamp = log.timestamp ? new Date(log.timestamp).toISOString() : '';
        const actionType = (log.actionType || '').replace(/,/g, ';');
        const actor = (log.actor || '').replace(/,/g, ';');
        const role = (log.actorRole || '').replace(/,/g, ';');
        const formationId = (log.formationId || '').replace(/,/g, ';');
        const status = (log.status || '').replace(/,/g, ';');
        const reason = (log.reason || '').replace(/,/g, ';');
        let extra = '';
        if (log.extra) {
          if (typeof log.extra === 'object') {
            extra = (log.extra.extra || JSON.stringify(log.extra)).replace(/,/g, ';');
          } else {
            extra = String(log.extra).replace(/,/g, ';');
          }
        }

        csv += `${timestamp},${actionType},${actor},${role},${formationId},${status},${reason},${extra}\n`;
      }

      // Create download link
      const blob = new Blob([csv], { type: 'text/csv' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `hrm_audit_logs_${new Date().toISOString().split('T')[0]}.csv`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);

      await UI.showSuccess('Export Complete', `Exported ${logs.length} HRM audit logs.`);
    } catch (err) {
      await UI.showError('Export Failed', err.message || 'Failed to export HRM audit logs.');
    }
  }

  // Export HRM Audit Logs functions (assign to window.AdminPage, not const AdminPage)
  if (!window.AdminPage) {
    window.AdminPage = {};
  }
  window.AdminPage.loadHrmAuditLogs = loadHrmAuditLogs;
  window.AdminPage.loadHrmLogsPage = loadHrmLogsPage;
  window.AdminPage.exportHrmAuditLogs = exportHrmAuditLogs;

  // ============================================================================
  // MODULE MANAGEMENT FUNCTIONS
  // ============================================================================

  function setupModuleSelectors() {
    const moduleCards = document.querySelectorAll('.module-card');
    moduleCards.forEach(card => {
      card.addEventListener('click', () => {
        const module = card.getAttribute('data-module');
        if (module) {
          openModule(module);
        }
      });
    });
  }

  function showModuleSelector() {
    console.log('showModuleSelector called');
    const moduleSelector = document.getElementById('moduleSelector');
    const adminLayout = document.getElementById('adminLayout');
    const adminSidebar = document.getElementById('adminSidebar');
    const sidebarOverlay = document.getElementById('sidebarOverlay');

    // Hide all workspace contents
    document.querySelectorAll('.workspace-content').forEach(ws => {
      ws.classList.remove('active');
    });

    if (moduleSelector) {
      moduleSelector.classList.remove('hidden');
      console.log('Module selector shown');
    }
    if (adminLayout) {
      adminLayout.classList.add('hidden');
      console.log('Admin layout hidden');
    }

    // Close mobile sidebar if open
    if (adminSidebar) adminSidebar.classList.remove('mobile-open');
    if (sidebarOverlay) sidebarOverlay.classList.remove('active');

    currentModule = null;
  }

  function openModule(module) {
    currentModule = module;

    // Hide module selector, show admin layout
    const moduleSelector = document.getElementById('moduleSelector');
    const adminLayout = document.getElementById('adminLayout');

    if (moduleSelector) moduleSelector.classList.add('hidden');
    if (adminLayout) adminLayout.classList.remove('hidden');

    // Close mobile sidebar if open
    const adminSidebar = document.getElementById('adminSidebar');
    const sidebarOverlay = document.getElementById('sidebarOverlay');
    if (adminSidebar) adminSidebar.classList.remove('mobile-open');
    if (sidebarOverlay) sidebarOverlay.classList.remove('active');

    // Update active module card
    document.querySelectorAll('.module-card').forEach(card => {
      card.classList.remove('active');
      if (card.getAttribute('data-module') === module) {
        card.classList.add('active');
      }
    });

    // Setup back button handlers for the current workspace
    // This ensures buttons work even if they were hidden when page loaded
    setTimeout(() => {
      const backButtons = document.querySelectorAll('.back-to-modules-header');
      backButtons.forEach(btn => {
        // Remove any existing listeners by replacing the button
        const newBtn = btn.cloneNode(true);
        btn.parentNode.replaceChild(newBtn, btn);

        // Add fresh event listener
        newBtn.addEventListener('click', (e) => {
          e.preventDefault();
          e.stopPropagation();
          console.log('Back to modules clicked (direct handler)');
          showModuleSelector();
        });
      });
    }, 100);

    // Load module-specific sidebar (only for System Admin) and workspace
    loadModuleSidebar(module);
    loadModuleActions(module);
    loadModuleWorkspace(module);
  }

  function loadModuleSidebar(module) {
    const adminSidebar = document.getElementById('adminSidebar');
    const adminWorkspace = document.getElementById('adminWorkspace');

    // Only show sidebar for System Admin
    if (module === 'system') {
      if (adminSidebar) adminSidebar.classList.remove('module-hidden');
      if (adminWorkspace) adminWorkspace.classList.remove('full-width');

      const sidebarContent = document.getElementById('sidebarContent');
      if (!sidebarContent) return;

      const html = `
        <div class="sidebar-section">
          <div class="sidebar-section-title">Settings</div>
          <a href="#" class="sidebar-link active" data-view="dashboard">
            <span class="sidebar-link-icon">📊</span>
            <span>Overview</span>
          </a>
          <a href="#" class="sidebar-link" data-view="formations">
            <span class="sidebar-link-icon">🏢</span>
            <span>Formations</span>
          </a>
          <a href="#" class="sidebar-link" data-view="admins">
            <span class="sidebar-link-icon">👥</span>
            <span>Admins</span>
          </a>
          <a href="#" class="sidebar-link" data-view="hrm-admins">
            <span class="sidebar-link-icon">👔</span>
            <span>HRM Admins</span>
          </a>
          <a href="#" class="sidebar-link" data-view="activity-logs">
            <span class="sidebar-link-icon">📋</span>
            <span>Activity Logs</span>
          </a>
          <a href="#" class="sidebar-link" data-view="retention">
            <span class="sidebar-link-icon">🗄️</span>
            <span>Data Storage</span>
          </a>
        </div>
      `;

      sidebarContent.innerHTML = html;

      // Setup sidebar link click handlers
      sidebarContent.querySelectorAll('.sidebar-link').forEach(link => {
        link.addEventListener('click', (e) => {
          e.preventDefault();
          const view = link.getAttribute('data-view');
          if (view) {
            // Remove active class from all links
            sidebarContent.querySelectorAll('.sidebar-link').forEach(l => l.classList.remove('active'));
            link.classList.add('active');

            // Close mobile sidebar after navigation
            if (adminSidebar) adminSidebar.classList.remove('mobile-open');
            const sidebarOverlay = document.getElementById('sidebarOverlay');
            if (sidebarOverlay) sidebarOverlay.classList.remove('active');

            // Load the view
            loadModuleView(module, view);
          }
        });
      });
    } else {
      // Hide sidebar for other modules
      if (adminSidebar) adminSidebar.classList.add('module-hidden');
      if (adminWorkspace) adminWorkspace.classList.add('full-width');
    }
  }

  function loadModuleActions(module) {
    // Load contextual action buttons for each module
    if (module === 'attendance') {
      loadAttendanceActions();
    } else if (module === 'visitors') {
      loadVisitorsActions();
    } else if (module === 'hrm') {
      loadHrmActions();
    }
  }

  function loadAttendanceActions() {
    const actionsContainer = document.getElementById('attendanceActions');
    if (!actionsContainer) return;

    actionsContainer.innerHTML = `
      <button class="action-btn active" data-view="dashboard">
        <span class="action-btn-icon">📊</span>
        <span>Overview</span>
      </button>
      <button class="action-btn" data-view="employees">
        <span class="action-btn-icon">👤</span>
        <span>Staff</span>
      </button>
      <button class="action-btn" data-view="devices">
        <span class="action-btn-icon">📱</span>
        <span>Devices</span>
      </button>
      <button class="action-btn" data-view="logs">
        <span class="action-btn-icon">📋</span>
        <span>Records</span>
      </button>
      <button class="action-btn" data-view="tokens">
        <span class="action-btn-icon">🔑</span>
        <span>QR Codes</span>
      </button>
    `;

    // Setup click handlers
    actionsContainer.querySelectorAll('.action-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const view = btn.getAttribute('data-view');
        if (view) {
          actionsContainer.querySelectorAll('.action-btn').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          loadModuleView('attendance', view);
        }
      });
    });
  }

  function loadVisitorsActions() {
    const actionsContainer = document.getElementById('visitorsActions');
    if (!actionsContainer) return;

    actionsContainer.innerHTML = `
      <button class="action-btn active" data-view="dashboard">
        <span class="action-btn-icon">📊</span>
        <span>Overview</span>
      </button>
      <button class="action-btn" data-view="requests">
        <span class="action-btn-icon">📝</span>
        <span>Requests</span>
      </button>
      <button class="action-btn" data-view="logs">
        <span class="action-btn-icon">📋</span>
        <span>Records</span>
      </button>
    `;

    // Setup click handlers
    actionsContainer.querySelectorAll('.action-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const view = btn.getAttribute('data-view');
        if (view) {
          actionsContainer.querySelectorAll('.action-btn').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          loadModuleView('visitors', view);
        }
      });
    });
  }

  function loadHrmActions() {
    const actionsContainer = document.getElementById('hrmActions');
    if (!actionsContainer) return;

    let html = `
      <button class="action-btn active" data-view="dashboard">
        <span class="action-btn-icon">📊</span>
        <span>Overview</span>
      </button>
      <button class="action-btn" data-view="staff">
        <span class="action-btn-icon">👔</span>
        <span>Staff</span>
      </button>
      <button class="action-btn" data-view="profiles">
        <span class="action-btn-icon">👤</span>
        <span>Profiles</span>
      </button>
      <button class="action-btn" data-view="leaves">
        <span class="action-btn-icon">🏖️</span>
        <span>Leaves</span>
      </button>
      <button class="action-btn" data-view="performance">
        <span class="action-btn-icon">⭐</span>
        <span>Performance</span>
      </button>
      <button class="action-btn" data-view="transfers">
        <span class="action-btn-icon">🔄</span>
        <span>Transfers</span>
      </button>
    `;

    // Add admin actions for SUPER_ADMIN and HRM_ADMIN
    if (adminRole === 'SUPER_ADMIN' || adminRole === 'HRM_ADMIN') {
      html += `
        <button class="action-btn" data-view="admins">
          <span class="action-btn-icon">👥</span>
          <span>Admins</span>
        </button>
        <button class="action-btn" data-view="audit">
          <span class="action-btn-icon">📜</span>
          <span>Logs</span>
        </button>
      `;
    }

    // Add "Add Staff" button for SUPER_ADMIN and HRM_ADMIN (not a view, but an action)
    if (adminRole === 'SUPER_ADMIN' || adminRole === 'HRM_ADMIN') {
      html += `
        <button class="action-btn action-btn-primary" id="addStaffBtn" style="background: var(--nysc-green); color: white; border: none;">
          <span class="action-btn-icon">➕</span>
          <span>Add Staff</span>
        </button>
      `;
    }

    actionsContainer.innerHTML = html;

    // Setup click handlers for view buttons
    actionsContainer.querySelectorAll('.action-btn[data-view]').forEach(btn => {
      btn.addEventListener('click', () => {
        const view = btn.getAttribute('data-view');
        if (view) {
          actionsContainer.querySelectorAll('.action-btn').forEach(b => b.classList.remove('active'));
          btn.classList.add('active');
          loadModuleView('hrm', view);
        }
      });
    });

    // Setup click handler for "Add Staff" button
    const addStaffBtn = document.getElementById('addStaffBtn');
    if (addStaffBtn) {
      addStaffBtn.addEventListener('click', () => {
        showAddStaffModal();
      });
    }
  }

  function loadModuleWorkspace(module) {
    // Hide all workspaces
    document.querySelectorAll('.workspace-content').forEach(ws => {
      ws.classList.remove('active');
    });

    // Show selected workspace
    const workspace = document.getElementById(`${module}Workspace`);
    if (workspace) {
      workspace.classList.add('active');
    }

    // Load default view (dashboard)
    loadModuleView(module, 'dashboard');
  }

  async function loadModuleView(module, view) {
    const contentId = `${module}Content`;
    const content = document.getElementById(contentId);
    if (!content) return;

    // Update active action button for non-system modules
    if (module !== 'system') {
      const actionsContainer = document.getElementById(`${module}Actions`);
      if (actionsContainer) {
        actionsContainer.querySelectorAll('.action-btn').forEach(btn => {
          btn.classList.remove('active');
          if (btn.getAttribute('data-view') === view) {
            btn.classList.add('active');
          }
        });
      }
    }

    // Show loading
    content.innerHTML = '<div class="info info-muted">Loading...</div>';

    try {
      if (module === 'attendance') {
        await loadAttendanceView(view, content);
      } else if (module === 'visitors') {
        await loadVisitorsView(view, content);
      } else if (module === 'hrm') {
        await loadHrmView(view, content);
      } else if (module === 'system') {
        await loadSystemView(view, content);
      }
    } catch (err) {
      console.error(`Error loading ${module} view ${view}:`, err);
      content.innerHTML = `<div class="info info-error">Error loading view: ${err.message}</div>`;
    }
  }

  async function loadAttendanceView(view, container) {
    if (view === 'dashboard') {
      container.innerHTML = `
        <section class="card">
          <h2>Attendance Overview</h2>
          <div class="dashboard-stats">
            <div class="stat-card">
              <div class="stat-value" id="attendanceToday">-</div>
              <div class="stat-label">Today's Attendance</div>
            </div>
            <div class="stat-card">
              <div class="stat-value" id="attendanceThisWeek">-</div>
              <div class="stat-label">This Week</div>
            </div>
            <div class="stat-card">
              <div class="stat-value" id="attendanceThisMonth">-</div>
              <div class="stat-label">This Month</div>
            </div>
          </div>
        </section>
      `;
      await loadAttendanceLogs();
    } else if (view === 'employees') {
      container.innerHTML = `
        <section class="card">
          <h2>Employee Management</h2>
          <div id="employeesTable" class="table-wrapper"></div>
        </section>
      `;
      await loadEmployees();
    } else if (view === 'devices') {
      container.innerHTML = `
        <section class="card">
          <h2>Device Registration</h2>
          <label class="switch">
            <input type="checkbox" id="registrationModeToggle" />
            <span class="slider"></span>
            <span class="switch-label">Enable Registration Mode</span>
          </label>
        </section>
      `;
      // Re-attach event listener
      const toggle = document.getElementById('registrationModeToggle');
      if (toggle) {
        toggle.addEventListener('change', handleRegistrationToggle);
      }
    } else if (view === 'logs') {
      container.innerHTML = `
        <section class="card">
          <h2>Attendance Logs</h2>
          <button id="exportAttendanceBtn" class="btn btn-secondary btn-sm">Export Sheet</button>
          <div id="attendanceLogs" class="table-wrapper"></div>
        </section>
      `;
      const exportBtn = document.getElementById('exportAttendanceBtn');
      if (exportBtn) {
        exportBtn.addEventListener('click', () => adminAction('exportAttendance'));
      }
      await loadAttendanceLogs();
    } else if (view === 'tokens') {
      container.innerHTML = `
        <section class="card">
          <h2>Token Management</h2>
          <button id="forceNewTokenBtn" class="btn btn-secondary btn-sm">Generate New Daily Token</button>
          <div id="tokenStatus" class="status"></div>
        </section>
      `;
      const tokenBtn = document.getElementById('forceNewTokenBtn');
      if (tokenBtn) {
        tokenBtn.addEventListener('click', () => adminAction('forceNewToken'));
      }
    }
  }

  async function loadVisitorsView(view, container) {
    if (view === 'dashboard') {
      container.innerHTML = `
        <section class="card">
          <h2>Visitors Overview</h2>
          <div class="dashboard-stats">
            <div class="stat-card">
              <div class="stat-value" id="visitorsToday">-</div>
              <div class="stat-label">Today</div>
            </div>
            <div class="stat-card">
              <div class="stat-value" id="visitorsPending">-</div>
              <div class="stat-label">Pending</div>
            </div>
            <div class="stat-card">
              <div class="stat-value" id="visitorsThisWeek">-</div>
              <div class="stat-label">This Week</div>
            </div>
          </div>
        </section>
      `;
    } else if (view === 'requests') {
      container.innerHTML = `
        <section class="card">
          <h2>Visitor Requests</h2>
          <div id="visitorRequests" class="table-wrapper"></div>
        </section>
      `;
    } else if (view === 'logs') {
      container.innerHTML = `
        <section class="card">
          <h2>Visitor Logs</h2>
          <button id="exportVisitorsBtn" class="btn btn-secondary btn-sm">Export Sheet</button>
          <div id="visitorLogs" class="table-wrapper"></div>
        </section>
      `;
      const exportBtn = document.getElementById('exportVisitorsBtn');
      if (exportBtn) {
        exportBtn.addEventListener('click', () => adminAction('exportVisitors'));
      }
      await loadVisitorLogs();
    }
  }

  async function loadHrmView(view, container) {
    if (view === 'dashboard') {
      container.innerHTML = `
        <section class="card">
          <h2>HRM Overview</h2>
          <div class="dashboard-stats">
            <div class="stat-card">
              <div class="stat-value" id="hrmTotalStaff">-</div>
              <div class="stat-label">Total Staff</div>
            </div>
            <div class="stat-card">
              <div class="stat-value" id="hrmActiveStaff">-</div>
              <div class="stat-label">Active Staff</div>
            </div>
            <div class="stat-card">
              <div class="stat-value" id="hrmPendingLeaves">-</div>
              <div class="stat-label">Pending Leaves</div>
            </div>
          </div>
        </section>
      `;
      await loadHrmStaffStats();
    } else if (view === 'staff') {
      container.innerHTML = `
        <section class="card card-full-width">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 1rem; flex-wrap: wrap; gap: 1rem;">
            <h2>Staff Records</h2>
            ${(adminRole === 'SUPER_ADMIN' || adminRole === 'HRM_ADMIN') ? `
            <div style="display: flex; gap: 0.5rem; flex-wrap: wrap;">
              <button class="btn btn-secondary" id="downloadStaffRecordsBtn" title="Download full staff records as CSV or Excel">
                <span>📥 Download / Extract</span>
              </button>
              <button class="btn btn-primary" id="addStaffBtnInView" style="background: var(--nysc-green);">
                <span>➕ Add Staff</span>
              </button>
            </div>
            ` : ''}
          </div>
          <div style="display: flex; gap: 1rem; margin-bottom: 1rem; flex-wrap: wrap; align-items: center;">
            <input 
              type="text" 
              id="hrmStaffSearchInput" 
              class="swal2-input" 
              placeholder="Search by name, employee ID, or file number..." 
              style="flex: 1; min-width: 200px; margin: 0;"
            />
            <button class="btn btn-secondary" id="hrmStaffSearchBtn">Search</button>
            <button class="btn btn-secondary" id="hrmStaffClearSearchBtn">Clear</button>
            <label style="display: flex; align-items: center; gap: 0.5rem; cursor: pointer;">
              <input type="checkbox" id="hrmIncludeArchivedCheck" />
              <span>Include Archived</span>
            </label>
          </div>
          <div id="hrmStaffTable" class="table-wrapper"></div>
          <div id="hrmStaffPagination"></div>
        </section>
      `;

      // Setup Add Staff button in view
      const addStaffBtnInView = document.getElementById('addStaffBtnInView');
      if (addStaffBtnInView) {
        addStaffBtnInView.addEventListener('click', () => {
          showAddStaffModal();
        });
      }

      // Setup Download / Extract button (HRM Admin & Super Admin only)
      const downloadStaffRecordsBtn = document.getElementById('downloadStaffRecordsBtn');
      if (downloadStaffRecordsBtn) {
        downloadStaffRecordsBtn.addEventListener('click', () => {
          downloadStaffRecordsAsCsv();
        });
      }

      // Setup search and filter handlers
      const searchBtn = document.getElementById('hrmStaffSearchBtn');
      const clearBtn = document.getElementById('hrmStaffClearSearchBtn');
      const searchInput = document.getElementById('hrmStaffSearchInput');
      const includeArchivedCheck = document.getElementById('hrmIncludeArchivedCheck');

      if (searchBtn) {
        searchBtn.addEventListener('click', () => {
          loadStaffList(1);
        });
      }

      if (clearBtn) {
        clearBtn.addEventListener('click', () => {
          if (searchInput) searchInput.value = '';
          loadStaffList(1);
        });
      }

      if (searchInput) {
        searchInput.addEventListener('keypress', (e) => {
          if (e.key === 'Enter') {
            loadStaffList(1);
          }
        });
      }

      if (includeArchivedCheck) {
        includeArchivedCheck.addEventListener('change', () => {
          loadStaffList(1);
        });
      }

      await loadStaffList(1);
    } else if (view === 'profiles') {
      container.innerHTML = `
        <section class="card">
          <h2>Employee Profiles</h2>
          <div class="info info-muted">Profile management coming soon...</div>
        </section>
      `;
    } else if (view === 'leaves') {
      container.innerHTML = `
        <section class="card">
          <h2>Leave Management</h2>
          <div class="info info-muted">Leave management coming soon...</div>
        </section>
      `;
    } else if (view === 'performance') {
      container.innerHTML = `
        <section class="card">
          <h2>Performance Reviews</h2>
          <div class="info info-muted">Performance management coming soon...</div>
        </section>
      `;
    } else if (view === 'transfers') {
      container.innerHTML = `
        <section class="card">
          <h2>Department Transfers</h2>
          <div class="info info-muted">Transfer management coming soon...</div>
        </section>
      `;
    } else if (view === 'admins') {
      container.innerHTML = `
        <section class="card">
          <h2>HRM Admins</h2>
          <div id="hrmAdminsTable" class="table-wrapper"></div>
        </section>
      `;
      if (adminRole === 'SUPER_ADMIN') {
        await loadHrmAdminsTable();
      }
    } else if (view === 'audit') {
      container.innerHTML = `
        <section class="card">
          <h2>HRM Audit Logs</h2>
          <div id="hrmAuditLogs" class="table-wrapper"></div>
          <div id="hrmAuditPagination"></div>
        </section>
      `;
      await loadHrmAuditLogs(1);
    } else {
      container.innerHTML = `<div class="info info-muted">Loading HRM ${view}...</div>`;
    }
  }

  async function loadSystemView(view, container) {
    if (!adminKey || adminRole !== 'SUPER_ADMIN') {
      container.innerHTML = '<div class="info info-muted">Access denied. Super Admin only.</div>';
      return;
    }

    if (view === 'dashboard') {
      container.innerHTML = `
        <section class="card card-full-width">
          <h2>Nationwide Overview</h2>
          <div class="dashboard-stats">
            <div class="stat-card">
              <div class="stat-value" id="statTotalFormations">-</div>
              <div class="stat-label">Formations</div>
            </div>
            <div class="stat-card">
              <div class="stat-value" id="statTotalEmployees">-</div>
              <div class="stat-label">Total Staff</div>
            </div>
            <div class="stat-card">
              <div class="stat-value" id="statAttendanceToday">-</div>
              <div class="stat-label">Attendance Today</div>
            </div>
            <div class="stat-card">
              <div class="stat-value" id="statPendingVisits">-</div>
              <div class="stat-label">Pending Visits</div>
            </div>
          </div>
        </section>
        <section class="card card-full-width">
          <h2>Formations</h2>
          <div id="formationsTable" class="table-wrapper"></div>
        </section>
      `;
      await loadNationwideDashboard();
      await loadFormationsTable();
    } else if (view === 'formations') {
      container.innerHTML = `
        <section class="card card-full-width">
          <h2>Formations</h2>
          <div class="contextual-actions" style="margin-bottom: 1rem;">
            <button class="btn btn-primary" id="addFormationBtn">Add Formation</button>
          </div>
          <div id="formationsTable" class="table-wrapper"></div>
        </section>
      `;
      const addFormationBtn = document.getElementById('addFormationBtn');
      if (addFormationBtn) addFormationBtn.addEventListener('click', () => showFormationModal());
      await loadFormationsTable();
    } else if (view === 'admins') {
      container.innerHTML = `
        <section class="card card-full-width" id="adminAssignmentSection">
          <h2>Admins (All Modules)</h2>
          <p class="info info-muted" style="margin-bottom: 1rem;">Create, update roles, and manage admins across Attendance, Visitors, HRM, and System.</p>
          <div class="contextual-actions" style="margin-bottom: 1rem;">
            <button class="btn btn-primary" id="createAdminBtn">Create Admin</button>
            <button class="btn btn-secondary" id="refreshAdminsBtn">Refresh</button>
          </div>
          <div id="adminsTable" class="table-wrapper"></div>
        </section>
      `;
      const createAdminBtn = document.getElementById('createAdminBtn');
      const refreshAdminsBtn = document.getElementById('refreshAdminsBtn');
      if (createAdminBtn) createAdminBtn.addEventListener('click', () => showAdminModal());
      if (refreshAdminsBtn) refreshAdminsBtn.addEventListener('click', () => loadAdminsTable());
      await loadAdminsTable();
    } else if (view === 'hrm-admins') {
      container.innerHTML = `
        <section class="card card-full-width">
          <h2>HRM Admins</h2>
          <div class="contextual-actions" style="margin-bottom: 1rem;">
            <button class="btn btn-primary" id="createHrmAdminBtn">Create HRM Admin</button>
            <button class="btn btn-secondary" id="refreshHrmAdminsBtn">Refresh</button>
          </div>
          <div id="hrmAdminsTable" class="table-wrapper"></div>
        </section>
      `;
      const createHrmAdminBtn = document.getElementById('createHrmAdminBtn');
      const refreshHrmAdminsBtn = document.getElementById('refreshHrmAdminsBtn');
      if (createHrmAdminBtn) createHrmAdminBtn.addEventListener('click', () => showCreateHrmAdminModal());
      if (refreshHrmAdminsBtn) refreshHrmAdminsBtn.addEventListener('click', () => loadHrmAdminsTable());
      await loadHrmAdminsTable();
    } else if (view === 'activity-logs') {
      const todayStr = new Date().toISOString().split('T')[0];
      container.innerHTML = `
        <section class="card card-full-width">
          <h2>Activity Logs – Who Did What & When</h2>
          <p class="info info-muted" style="margin-bottom: 1rem;">Track all admin and system activity. Initially shows today's activities only. Use filters and Apply Filters to expand search.</p>
          <div class="contextual-actions" style="margin-bottom: 1rem; flex-wrap: wrap; gap: 0.5rem;">
            <input type="text" id="activityLogsActorFilter" placeholder="Filter by actor (key or name)" style="max-width: 220px; padding: 0.5rem; border: 1px solid #ddd; border-radius: 4px;">
            <input type="text" id="activityLogsActionFilter" placeholder="Filter by action type" style="max-width: 180px; padding: 0.5rem; border: 1px solid #ddd; border-radius: 4px;">
            <input type="text" id="activityLogsFormationFilter" placeholder="Filter by formation" style="max-width: 160px; padding: 0.5rem; border: 1px solid #ddd; border-radius: 4px;">
            <input type="date" id="activityLogsStartDate" title="From date" value="${todayStr}" style="padding: 0.5rem; border: 1px solid #ddd; border-radius: 4px;">
            <input type="date" id="activityLogsEndDate" title="To date" value="${todayStr}" style="padding: 0.5rem; border: 1px solid #ddd; border-radius: 4px;">
            <button class="btn btn-primary" id="activityLogsApplyBtn">Apply Filters</button>
            <button class="btn btn-secondary" id="activityLogsRefreshBtn">Refresh (Today)</button>
          </div>
          <div id="systemAuditLogsTable" class="table-wrapper" style="overflow-x: auto;"></div>
        </section>
      `;
      const applyBtn = document.getElementById('activityLogsApplyBtn');
      const refreshBtn = document.getElementById('activityLogsRefreshBtn');
      const getFilters = () => {
        const actor = document.getElementById('activityLogsActorFilter');
        const actionType = document.getElementById('activityLogsActionFilter');
        const formationId = document.getElementById('activityLogsFormationFilter');
        const startDate = document.getElementById('activityLogsStartDate');
        const endDate = document.getElementById('activityLogsEndDate');
        return {
          limit: 300,
          actor: actor && actor.value.trim() ? actor.value.trim() : undefined,
          actionType: actionType && actionType.value.trim() ? actionType.value.trim() : undefined,
          formationId: formationId && formationId.value.trim() ? formationId.value.trim() : undefined,
          startDate: startDate && startDate.value ? startDate.value : undefined,
          endDate: endDate && endDate.value ? endDate.value : undefined,
        };
      };
      const getTodayOnlyFilters = () => {
        const today = new Date().toISOString().split('T')[0];
        return { limit: 300, startDate: today, endDate: today };
      };
      if (applyBtn) applyBtn.addEventListener('click', () => loadSystemAuditLogs(getFilters()));
      if (refreshBtn) refreshBtn.addEventListener('click', () => {
        const today = new Date().toISOString().split('T')[0];
        const actorEl = document.getElementById('activityLogsActorFilter');
        const actionEl = document.getElementById('activityLogsActionFilter');
        const formationEl = document.getElementById('activityLogsFormationFilter');
        const startEl = document.getElementById('activityLogsStartDate');
        const endEl = document.getElementById('activityLogsEndDate');
        if (actorEl) actorEl.value = '';
        if (actionEl) actionEl.value = '';
        if (formationEl) formationEl.value = '';
        if (startEl) startEl.value = today;
        if (endEl) endEl.value = today;
        loadSystemAuditLogs(getTodayOnlyFilters());
      });
      await loadSystemAuditLogs(getTodayOnlyFilters());
    } else if (view === 'retention') {
      container.innerHTML = `
        <section class="card card-full-width">
          <h2>Data Storage & Retention</h2>
          <div id="retentionPolicy" class="info info-muted" style="margin-bottom: 1rem;">Loading retention policy...</div>
          <div style="margin-bottom: 1rem;">
            <button class="btn btn-secondary" id="runArchivalBtn">Run Archival Now</button>
            <button class="btn btn-secondary" id="refreshRetentionBtn">Refresh Policy</button>
          </div>
          <div id="archivalStatus" class="info"></div>
        </section>
      `;
      const runArchivalBtn = document.getElementById('runArchivalBtn');
      const refreshRetentionBtn = document.getElementById('refreshRetentionBtn');
      if (runArchivalBtn) runArchivalBtn.addEventListener('click', runArchival);
      if (refreshRetentionBtn) refreshRetentionBtn.addEventListener('click', loadRetentionPolicy);
      await loadRetentionPolicy();
    } else {
      container.innerHTML = `<div class="info info-muted">Loading System ${view}...</div>`;
    }
  }

  /**
   * View image in a larger modal
   */
  function viewImage(imageUrl, imageName) {
    if (!imageUrl || imageUrl === '#') {
      UI.showError('Error', 'Image URL is not available.');
      return;
    }

    Swal.fire({
      title: imageName || 'Image Viewer',
      html: `
        <div style="text-align: center;">
          <img src="${imageUrl}" alt="${imageName || 'Image'}" 
               style="max-width: 100%; max-height: 70vh; border-radius: 0.5rem; box-shadow: 0 4px 6px rgba(0,0,0,0.1);"
               onerror="this.src='data:image/svg+xml,%3Csvg xmlns=%27http://www.w3.org/2000/svg%27 width=%27400%27 height=%27300%27%3E%3Crect fill=%27%23f3f4f6%27 width=%27400%27 height=%27300%27/%3E%3Ctext x=%2750%25%27 y=%2750%25%27 text-anchor=%27middle%27 dy=%27.3em%27 fill=%27%239ca3af%27 font-size=%2718%27%3EImage not available%3C/text%3E%3C/svg%3E';">
        </div>
      `,
      width: '90%',
      maxWidth: '1200px',
      showCloseButton: true,
      showConfirmButton: true,
      confirmButtonText: 'Open in New Tab',
      confirmButtonColor: '#059669',
      showCancelButton: true,
      cancelButtonText: 'Close',
      footer: `
        <div style="margin-top: 1rem; padding-top: 1rem; border-top: 1px solid #e5e7eb;">
          <a href="${imageUrl}" target="_blank" style="color: #059669; text-decoration: none; font-weight: 500;">
            View Full Size in New Tab →
          </a>
        </div>
      `,
      didOpen: () => {
        // Add click handler to open in new tab
        const confirmBtn = document.querySelector('.swal2-confirm');
        if (confirmBtn) {
          confirmBtn.addEventListener('click', () => {
            window.open(imageUrl, '_blank');
          });
        }
      }
    });
  }

  // Merge return value with window.AdminPage (which already has functions assigned above)
  const returnValue = {
    init,
    refreshAll,
    loadHrmStats,
    viewEmployeeProfile,
    approveLeave,
    rejectLeave,
    uploadEmployeeDocument,
    showEmployeeDocuments,
    loadHrmStaffStats,
    loadStaffList,
    loadStaffListPage,
    viewStaffProfile,
    editStaff,
    archiveStaffConfirm,
    showAddStaffModal,
    loadAdminsTable,
    showAdminModal,
    loadSystemAuditLogs,
    loadHrmAdminsTable,
    showCreateHrmAdminModal,
    editHrmAdminRole,
    deactivateHrmAdmin,
    activateHrmAdmin,
    loadHrmAuditLogs,
    loadHrmLogsPage,
    exportHrmAuditLogs,
    openModule,
    showModuleSelector,
    viewImage,
  };

  // Merge with window.AdminPage (functions already assigned above)
  if (!window.AdminPage) {
    window.AdminPage = {};
  }
  Object.assign(window.AdminPage, returnValue);

  return window.AdminPage;
})();

