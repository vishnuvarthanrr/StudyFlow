/**
 * StudyFlow - Frontend Application Logic
 * Academic Task Manager & Syllabus-Progress Tracker
 */

(function () {
  'use strict';

  // State Management
  const state = {
    token: localStorage.getItem('studyflow_token') || null,
    user: JSON.parse(localStorage.getItem('studyflow_user') || 'null'),
    isDemo: false,
    tasks: [],
    courses: [],
    recommendation: null,
    filter: 'all', // 'all' | 'open' | 'completed'
    searchQuery: '',
    timer: {
      intervalId: null,
      totalSeconds: 25 * 60,
      remainingSeconds: 25 * 60,
      isRunning: false,
      activeTaskId: null
    }
  };

  // DOM Elements
  const elements = {
    // Views
    authView: document.getElementById('auth-view'),
    appView: document.getElementById('app-view'),
    toastContainer: document.getElementById('toast-container'),

    // Auth & Server Status
    serverStatusPill: document.getElementById('server-status-pill'),
    serverStatusText: document.getElementById('server-status-text'),
    fileProtocolWarning: document.getElementById('file-protocol-warning'),
    demoTriggerBtns: document.querySelectorAll('.demo-trigger-btn'),
    tabRegisterBtn: document.getElementById('tab-register-btn'),
    tabLoginBtn: document.getElementById('tab-login-btn'),
    registerForm: document.getElementById('register-form'),
    loginForm: document.getElementById('login-form'),
    switchToLogin: document.getElementById('switch-to-login'),
    switchToRegister: document.getElementById('switch-to-register'),
    authAlert: document.getElementById('auth-alert'),
    regSubmitBtn: document.getElementById('reg-submit-btn'),
    loginSubmitBtn: document.getElementById('login-submit-btn'),

    // App Header & User
    currentDateText: document.getElementById('current-date-text'),
    userAvatar: document.getElementById('user-avatar'),
    userDisplayName: document.getElementById('user-display-name'),
    userDisplayEmail: document.getElementById('user-display-email'),
    userGreeting: document.getElementById('user-greeting'),
    logoutBtn: document.getElementById('logout-btn'),

    // Summary Stats
    statOpenTasks: document.getElementById('stat-open-tasks'),
    statCompletedTasks: document.getElementById('stat-completed-tasks'),
    statCompletionRate: document.getElementById('stat-completion-rate'),
    statProgressBar: document.getElementById('stat-progress-bar'),
    statOverdueTasks: document.getElementById('stat-overdue-tasks'),
    cardOverdue: document.getElementById('card-overdue'),

    // Recommendation
    recommendationContainer: document.getElementById('recommendation-container'),

    // Tasks Section
    openAddTaskBtn: document.getElementById('open-add-task-btn'),
    quickAddTaskBtn: document.getElementById('quick-add-task-btn'),
    tasksListContainer: document.getElementById('tasks-list-container'),
    taskSearchInput: document.getElementById('task-search-input'),
    clearSearchBtn: document.getElementById('clear-search-btn'),
    filterPills: document.querySelectorAll('.filter-pill'),
    countFilterAll: document.getElementById('count-filter-all'),
    countFilterOpen: document.getElementById('count-filter-open'),
    countFilterCompleted: document.getElementById('count-filter-completed'),

    // Course Progress Section
    coursesListContainer: document.getElementById('courses-list-container'),

    // Add Task Modal
    addTaskModal: document.getElementById('add-task-modal'),
    closeModalBtn: document.getElementById('close-modal-btn'),
    cancelTaskBtn: document.getElementById('cancel-task-btn'),
    addTaskForm: document.getElementById('add-task-form'),
    saveTaskBtn: document.getElementById('save-task-btn'),
    taskFormAlert: document.getElementById('task-form-alert'),
    taskDueDate: document.getElementById('task-due-date'),
    existingCoursesDatalist: document.getElementById('existing-courses-datalist'),

    // Timer Modal
    timerModal: document.getElementById('timer-modal'),
    closeTimerBtn: document.getElementById('close-timer-btn'),
    timerCourseUnit: document.getElementById('timer-course-unit'),
    timerTaskTitle: document.getElementById('timer-task-title'),
    timerCountdown: document.getElementById('timer-countdown'),
    timerProgressRing: document.getElementById('timer-progress-ring'),
    timerToggleBtn: document.getElementById('timer-toggle-btn'),
    timerToggleIcon: document.getElementById('timer-toggle-icon'),
    timerToggleText: document.getElementById('timer-toggle-text'),
    timerResetBtn: document.getElementById('timer-reset-btn'),
    timerCompleteTaskBtn: document.getElementById('timer-complete-task-btn')
  };

  /* ==========================================================================
     API CLIENT HELPER & BASE URL RESOLVER
     ========================================================================== */

  // Dynamically resolve API base URL:
  // If the page is opened from file:// or a live reload port (e.g. 5500),
  // target the backend server running on port 3000.
  const API_BASE = (function () {
    if (window.location.protocol === 'file:') {
      return 'http://localhost:3000';
    }
    if (window.location.port && window.location.port !== '3000') {
      const host = window.location.hostname || 'localhost';
      return `http://${host}:3000`;
    }
    return '';
  })();

  async function apiRequest(endpoint, options = {}) {
    const fullUrl = endpoint.startsWith('http') ? endpoint : `${API_BASE}${endpoint}`;
    const headers = {
      'Content-Type': 'application/json',
      ...(options.headers || {})
    };

    if (state.token) {
      headers['Authorization'] = `Bearer ${state.token}`;
    }

    try {
      const response = await fetch(fullUrl, {
        ...options,
        headers
      });

      const data = await response.json().catch(() => ({}));

      if (!response.ok) {
        if (response.status === 401 && state.token && !state.isDemo) {
          // Token expired or invalid
          handleLogout(false);
          showToast('Session expired. Please log in again.', 'error');
        }
        throw new Error(data.error || `Request failed (${response.status})`);
      }

      return data;
    } catch (err) {
      // Clear, actionable error messaging instead of cryptic 'Failed to fetch'
      if (err.name === 'TypeError' || err.message === 'Failed to fetch' || (err.message && err.message.includes('fetch'))) {
        const targetServer = API_BASE || window.location.origin || 'http://localhost:3000';
        if (window.location.protocol === 'file:') {
          throw new Error(`Cannot connect to StudyFlow backend from file://. Please start the server (run 'npm start' in terminal) and open <a href="http://localhost:3000" target="_blank" style="color: inherit; text-decoration: underline;">http://localhost:3000</a> in your browser, or click "Explore in Instant Demo Mode" below.`);
        } else {
          throw new Error(`Cannot connect to backend server at ${targetServer}. Please ensure the server is running (run 'npm start' or 'node server.js' in terminal), or click "Explore in Instant Demo Mode".`);
        }
      }
      throw err;
    }
  }

  // Real-time server health check
  async function checkBackendHealth() {
    if (!elements.serverStatusPill || !elements.serverStatusText) return;

    if (window.location.protocol === 'file:' && elements.fileProtocolWarning) {
      elements.fileProtocolWarning.classList.remove('hidden');
    }

    const healthUrl = `${API_BASE || ''}/api/health`;
    try {
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 2800);
      const res = await fetch(healthUrl, { signal: controller.signal });
      clearTimeout(timeoutId);

      if (res.ok) {
        elements.serverStatusPill.className = 'server-status-pill online';
        elements.serverStatusText.textContent = 'Backend Online (http://localhost:3000)';
        return true;
      } else {
        throw new Error('Status ' + res.status);
      }
    } catch (err) {
      elements.serverStatusPill.className = 'server-status-pill offline';
      elements.serverStatusText.textContent = 'Backend Offline — Run `npm start` in terminal';
      return false;
    }
  }

  /* ==========================================================================
     NOTIFICATION TOAST SYSTEM
     ========================================================================== */

  function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    let icon = 'ℹ️';
    if (type === 'success') icon = '✓';
    if (type === 'error') icon = '⚠️';

    toast.innerHTML = `<span>${icon}</span> <span>${escapeHtml(message)}</span>`;
    elements.toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      setTimeout(() => {
        if (toast.parentNode) {
          toast.parentNode.removeChild(toast);
        }
      }, 250);
    }, 3800);
  }

  function escapeHtml(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  /* ==========================================================================
     DATE & TIME HELPERS
     ========================================================================== */

  function getGreetingName(fullName) {
    if (!fullName) return 'Student';
    const first = fullName.trim().split(' ')[0];
    return first.charAt(0).toUpperCase() + first.slice(1);
  }

  function getTimeGreeting() {
    const hour = new Date().getHours();
    if (hour < 12) return 'Good morning';
    if (hour < 17) return 'Good afternoon';
    return 'Good evening';
  }

  function formatFormattedCurrentDate() {
    const options = { weekday: 'long', month: 'short', day: 'numeric', year: 'numeric' };
    return new Date().toLocaleDateString('en-US', options);
  }

  function getTodayString() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  function formatDueDateLabel(dateString, isCompleted) {
    if (!dateString) return '';
    const today = getTodayString();
    
    if (dateString < today && !isCompleted) {
      const diffDays = Math.round((new Date(today) - new Date(dateString)) / (1000 * 60 * 60 * 24));
      return {
        text: `Overdue by ${diffDays}d`,
        className: 'overdue',
        isOverdue: true
      };
    } else if (dateString === today && !isCompleted) {
      return {
        text: 'Due Today',
        className: 'due-today',
        isOverdue: false
      };
    } else {
      const d = new Date(dateString + 'T00:00:00');
      const text = d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
      return {
        text: `Due ${text}`,
        className: '',
        isOverdue: false
      };
    }
  }

  function formatMinutes(minutes) {
    const mins = parseInt(minutes, 10) || 0;
    if (mins < 60) return `${mins}m`;
    const hrs = Math.floor(mins / 60);
    const remMins = mins % 60;
    return remMins > 0 ? `${hrs}h ${remMins}m` : `${hrs}h`;
  }

  /* ==========================================================================
     AUTH FLOW
     ========================================================================== */

  function showAuthAlert(message, type = 'error') {
    elements.authAlert.className = `auth-alert ${type}`;
    elements.authAlert.innerHTML = message;
    elements.authAlert.classList.remove('hidden');
  }

  function hideAuthAlert() {
    elements.authAlert.classList.add('hidden');
    elements.authAlert.textContent = '';
  }

  function clearFormErrors() {
    document.querySelectorAll('.field-error').forEach(el => { el.textContent = ''; });
    hideAuthAlert();
  }

  function switchAuthTab(target) {
    clearFormErrors();
    if (target === 'register') {
      elements.tabRegisterBtn.classList.add('active');
      elements.tabRegisterBtn.setAttribute('aria-selected', 'true');
      elements.tabLoginBtn.classList.remove('active');
      elements.tabLoginBtn.setAttribute('aria-selected', 'false');
      elements.registerForm.classList.remove('hidden');
      elements.loginForm.classList.add('hidden');
    } else {
      elements.tabLoginBtn.classList.add('active');
      elements.tabLoginBtn.setAttribute('aria-selected', 'true');
      elements.tabRegisterBtn.classList.remove('active');
      elements.tabRegisterBtn.setAttribute('aria-selected', 'false');
      elements.loginForm.classList.remove('hidden');
      elements.registerForm.classList.add('hidden');
    }
  }

  // Register Handler
  async function handleRegister(e) {
    e.preventDefault();
    clearFormErrors();

    const name = document.getElementById('reg-name').value.trim();
    const email = document.getElementById('reg-email').value.trim();
    const password = document.getElementById('reg-password').value;

    let hasError = false;

    if (!name) {
      document.getElementById('reg-name-error').textContent = 'Please enter your name.';
      hasError = true;
    }

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!email || !emailRegex.test(email)) {
      document.getElementById('reg-email-error').textContent = 'Please enter a valid university or personal email.';
      hasError = true;
    }

    if (!password || password.length < 6) {
      document.getElementById('reg-password-error').textContent = 'Password must be at least 6 characters.';
      hasError = true;
    }

    if (hasError) return;

    // Submit
    const btnText = elements.regSubmitBtn.querySelector('.btn-text');
    const spinner = elements.regSubmitBtn.querySelector('.btn-spinner');
    elements.regSubmitBtn.disabled = true;
    btnText.textContent = 'Creating Account...';
    spinner.classList.remove('hidden');

    try {
      const data = await apiRequest('/api/auth/register', {
        method: 'POST',
        body: JSON.stringify({ name, email, password })
      });

      // Save session
      state.token = data.token;
      state.user = data.user;
      localStorage.setItem('studyflow_token', data.token);
      localStorage.setItem('studyflow_user', JSON.stringify(data.user));

      showToast(`Welcome to StudyFlow, ${getGreetingName(data.user.name)}!`, 'success');
      elements.registerForm.reset();
      showAppView();
    } catch (err) {
      showAuthAlert(err.message, 'error');
    } finally {
      elements.regSubmitBtn.disabled = false;
      btnText.textContent = 'Create Account & Get Started';
      spinner.classList.add('hidden');
    }
  }

  // Login Handler
  async function handleLogin(e) {
    e.preventDefault();
    clearFormErrors();

    const email = document.getElementById('login-email').value.trim();
    const password = document.getElementById('login-password').value;

    let hasError = false;

    if (!email) {
      document.getElementById('login-email-error').textContent = 'Email address is required.';
      hasError = true;
    }

    if (!password) {
      document.getElementById('login-password-error').textContent = 'Password is required.';
      hasError = true;
    }

    if (hasError) return;

    // Submit
    const btnText = elements.loginSubmitBtn.querySelector('.btn-text');
    const spinner = elements.loginSubmitBtn.querySelector('.btn-spinner');
    elements.loginSubmitBtn.disabled = true;
    btnText.textContent = 'Signing In...';
    spinner.classList.remove('hidden');

    try {
      const data = await apiRequest('/api/auth/login', {
        method: 'POST',
        body: JSON.stringify({ email, password })
      });

      // Save session
      state.token = data.token;
      state.user = data.user;
      localStorage.setItem('studyflow_token', data.token);
      localStorage.setItem('studyflow_user', JSON.stringify(data.user));

      showToast(`Welcome back, ${getGreetingName(data.user.name)}!`, 'success');
      elements.loginForm.reset();
      showAppView();
    } catch (err) {
      showAuthAlert(err.message, 'error');
    } finally {
      elements.loginSubmitBtn.disabled = false;
      btnText.textContent = 'Log In to StudyFlow';
      spinner.classList.add('hidden');
    }
  }

  // Logout Handler
  async function handleLogout(callApi = true) {
    if (callApi && state.token && !state.isDemo) {
      try {
        await apiRequest('/api/auth/logout', { method: 'POST' });
      } catch (e) {
        // Ignore logout errors
      }
    }

    state.token = null;
    state.user = null;
    state.isDemo = false;
    state.tasks = [];
    state.courses = [];
    state.recommendation = null;

    localStorage.removeItem('studyflow_token');
    localStorage.removeItem('studyflow_user');

    // Stop active timer if running
    resetTimer();

    showAuthView();
    showToast('You have been logged out.', 'info');
    checkBackendHealth();
  }

  /* ==========================================================================
     VIEW SWITCHING
     ========================================================================== */

  function showAuthView() {
    elements.appView.classList.add('hidden');
    elements.authView.classList.remove('hidden');
    switchAuthTab('register');
  }

  function showAppView() {
    elements.authView.classList.add('hidden');
    elements.appView.classList.remove('hidden');

    // Render User Header Info
    renderUserProfile();

    // Fetch and render dashboard data
    fetchDashboardData();
  }

  function renderUserProfile() {
    if (!state.user) return;
    const firstName = getGreetingName(state.user.name);
    elements.userGreeting.textContent = `${getTimeGreeting()}, ${firstName} 👋`;
    elements.userDisplayName.textContent = state.user.name;
    elements.userDisplayEmail.textContent = state.user.email;
    elements.currentDateText.textContent = formatFormattedCurrentDate();

    // Initials Avatar
    const initials = state.user.name
      .split(' ')
      .map(part => part.charAt(0))
      .slice(0, 2)
      .join('')
      .toUpperCase() || 'S';
    elements.userAvatar.textContent = initials;
  }

  /* ==========================================================================
     DASHBOARD DATA & RENDERING
     ========================================================================== */

  /* ==========================================================================
     INSTANT DEMO MODE (OFFLINE / FALLBACK SUPPORT)
     ========================================================================== */

  function activateDemoMode() {
    state.isDemo = true;
    state.token = 'demo-token-' + Date.now();
    state.user = {
      id: 'demo-user-1',
      name: 'Vishnuvarthan R',
      email: '2516140@saec.ac.in'
    };

    const today = new Date();
    const fmt = d => d.toISOString().split('T')[0];
    const addDays = (n) => {
      const d = new Date(today);
      d.setDate(d.getDate() + n);
      return fmt(d);
    };

    state.tasks = [
      {
        id: 'demo-task-1',
        userId: 'demo-user-1',
        title: 'Review AVL Tree Rotations & Red-Black Balancing',
        course: 'CS201 Data Structures',
        unit: 'Unit 3: Balanced Binary Trees',
        difficulty: 'Hard',
        dueDate: fmt(today),
        estimatedMinutes: 60,
        completed: false,
        createdAt: new Date().toISOString()
      },
      {
        id: 'demo-task-2',
        userId: 'demo-user-1',
        title: 'Java OOPs: Inheritance, Polymorphism & Interface Design',
        course: 'NPTEL Java Programming',
        unit: 'Module 4: Object-Oriented Principles',
        difficulty: 'Hard',
        dueDate: addDays(1),
        estimatedMinutes: 90,
        completed: false,
        createdAt: new Date().toISOString()
      },
      {
        id: 'demo-task-3',
        userId: 'demo-user-1',
        title: 'Matrix Eigenvalues & Diagonalization Practice Problems',
        course: 'Math 301 Linear Algebra',
        unit: 'Unit 5: Eigenvectors & SVD',
        difficulty: 'Medium',
        dueDate: addDays(3),
        estimatedMinutes: 45,
        completed: false,
        createdAt: new Date().toISOString()
      },
      {
        id: 'demo-task-4',
        userId: 'demo-user-1',
        title: 'CPU Scheduling Algorithms Simulation Assignment',
        course: 'CS305 Operating Systems',
        unit: 'Unit 2: Process Management',
        difficulty: 'Easy',
        dueDate: addDays(5),
        estimatedMinutes: 30,
        completed: true,
        completedAt: new Date().toISOString(),
        createdAt: new Date().toISOString()
      }
    ];

    showAppView();
    showToast('Entered Instant Demo Mode! All features active.', 'success');
  }

  function syncDemoDashboard() {
    const todayStr = getTodayString();
    const today = new Date(todayStr);

    const totalTasks = state.tasks.length;
    const completedTasks = state.tasks.filter(t => t.completed).length;
    const openTasks = totalTasks - completedTasks;
    const completionPercentage = totalTasks > 0 ? Math.round((completedTasks / totalTasks) * 100) : 0;

    const overdueTasks = state.tasks.filter(t => {
      if (t.completed) return false;
      const diff = new Date(t.dueDate).getTime() - today.getTime();
      return Math.round(diff / (1000 * 60 * 60 * 24)) < 0;
    }).length;

    // Course progress
    const coursesMap = {};
    state.tasks.forEach(task => {
      const c = task.course.trim();
      if (!coursesMap[c]) {
        coursesMap[c] = { course: c, totalCount: 0, completedCount: 0, openCount: 0, percentage: 0 };
      }
      coursesMap[c].totalCount += 1;
      if (task.completed) coursesMap[c].completedCount += 1;
      else coursesMap[c].openCount += 1;
    });

    const courseProgress = Object.values(coursesMap).map(c => ({
      ...c,
      percentage: c.totalCount > 0 ? Math.round((c.completedCount / c.totalCount) * 100) : 0
    })).sort((a, b) => b.totalCount - a.totalCount);

    // Smart Recommendation
    const uncompleted = state.tasks.filter(t => !t.completed);
    let recommendation = null;
    if (uncompleted.length > 0) {
      const scored = uncompleted.map(task => {
        let score = 0;
        let reason = '';
        const diffDays = Math.round((new Date(task.dueDate).getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
        if (diffDays < 0) {
          score += 1000 + Math.abs(diffDays) * 120;
          reason = `Overdue by ${Math.abs(diffDays)} day(s) — urgent catch-up recommended for ${task.course}.`;
        } else if (diffDays === 0) {
          score += 650;
          reason = `Due today for ${task.course} — complete to maintain your syllabus schedule.`;
        } else if (diffDays <= 3) {
          score += 400 + (3 - diffDays) * 60;
          reason = `Due in ${diffDays} day(s) for ${task.course}.`;
        } else {
          score += 100;
          reason = `Upcoming syllabus topic for ${task.course}.`;
        }
        if (task.difficulty === 'Hard') score += 200;
        else if (task.difficulty === 'Medium') score += 100;
        return { task, score, reason };
      });
      scored.sort((a, b) => b.score - a.score);
      recommendation = {
        ...scored[0].task,
        recommendationScore: scored[0].score,
        reason: scored[0].reason
      };
    }

    state.courses = courseProgress;
    state.recommendation = recommendation;

    renderSummaryCards({ totalTasks, openTasks, completedTasks, completionPercentage, overdueTasks });
    renderSmartRecommendation(recommendation);
    renderTasksList();
    renderCourseProgress(courseProgress);
    updateExistingCoursesDatalist();
  }

  /* ==========================================================================
     DASHBOARD DATA & RENDERING
     ========================================================================== */

  async function fetchDashboardData() {
    if (state.isDemo) {
      syncDemoDashboard();
      return;
    }

    try {
      const [dashData, tasksData] = await Promise.all([
        apiRequest('/api/dashboard'),
        apiRequest('/api/tasks')
      ]);

      if (dashData.user) {
        state.user = dashData.user;
        localStorage.setItem('studyflow_user', JSON.stringify(dashData.user));
        renderUserProfile();
      }

      state.tasks = tasksData.tasks || [];
      state.courses = dashData.courseProgress || [];
      state.recommendation = dashData.recommendation;

      renderSummaryCards(dashData.summary);
      renderSmartRecommendation(dashData.recommendation);
      renderTasksList();
      renderCourseProgress(dashData.courseProgress);
      updateExistingCoursesDatalist();
    } catch (err) {
      console.error('Error fetching dashboard data:', err);
      showToast('Could not sync latest study data: ' + err.message, 'error');
    }
  }

  function renderSummaryCards(summary) {
    if (!summary) return;

    elements.statOpenTasks.textContent = summary.openTasks;
    elements.statCompletedTasks.textContent = summary.completedTasks;
    elements.statCompletionRate.textContent = `${summary.completionPercentage}%`;
    elements.statProgressBar.style.width = `${summary.completionPercentage}%`;

    elements.statOverdueTasks.textContent = summary.overdueTasks;
    if (summary.overdueTasks > 0) {
      elements.cardOverdue.classList.add('has-overdue');
      document.getElementById('stat-overdue-sub').textContent = 'Requires urgent catch-up';
    } else {
      elements.cardOverdue.classList.remove('has-overdue');
      document.getElementById('stat-overdue-sub').textContent = 'All deadlines on track';
    }
  }

  function renderSmartRecommendation(recommendation) {
    const container = elements.recommendationContainer;

    if (!recommendation) {
      const hasAnyTasks = state.tasks.length > 0;
      if (hasAnyTasks) {
        // All completed
        container.innerHTML = `
          <div class="recommendation-empty-card">
            <div class="empty-star-icon">🎉</div>
            <h3>All caught up for now!</h3>
            <p>You have finished all active syllabus tasks. Take a well-deserved break or add your next course unit to keep steady momentum.</p>
            <div style="margin-top: 1rem;">
              <button class="btn btn-secondary-sm" id="empty-rec-add-btn">+ Add Next Syllabus Task</button>
            </div>
          </div>
        `;
        const addBtn = document.getElementById('empty-rec-add-btn');
        if (addBtn) addBtn.addEventListener('click', openAddTaskModal);
      } else {
        // Zero tasks created yet
        container.innerHTML = `
          <div class="recommendation-empty-card">
            <div class="empty-star-icon">📖</div>
            <h3>Start Your Syllabus Roadmap</h3>
            <p>Your study space is currently clean and empty. Add your first university course task to unlock tailored priority recommendations and progress tracking.</p>
            <div style="margin-top: 1rem;">
              <button class="btn btn-primary" id="empty-rec-add-btn">+ Add Your First Task</button>
            </div>
          </div>
        `;
        const addBtn = document.getElementById('empty-rec-add-btn');
        if (addBtn) addBtn.addEventListener('click', openAddTaskModal);
      }
      return;
    }

    const dueMeta = formatDueDateLabel(recommendation.dueDate, recommendation.completed);
    const difficultyClass = recommendation.difficulty.toLowerCase();

    container.innerHTML = `
      <div class="recommendation-card">
        <div class="recommendation-header">
          <div class="recommendation-badge">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polygon points="12 2 15.09 8.26 22 9.27 17 14.14 18.18 21.02 12 17.77 5.82 21.02 7 14.14 2 9.27 8.91 8.26 12 2"/></svg>
            Start Here Next
          </div>
          <span class="badge-difficulty ${difficultyClass}">${recommendation.difficulty} Difficulty</span>
        </div>

        <div class="recommendation-body">
          <div class="recommendation-info">
            <div class="recommendation-meta-row">
              <span class="course-tag">${escapeHtml(recommendation.course)}</span>
              <span class="unit-tag">${escapeHtml(recommendation.unit)}</span>
              <span class="duration-tag">⏱️ ${formatMinutes(recommendation.estimatedMinutes)}</span>
              <span class="due-badge ${dueMeta.className}">${dueMeta.text}</span>
            </div>

            <h3 class="recommendation-title">${escapeHtml(recommendation.title)}</h3>

            <div class="recommendation-reason-box">
              <span>💡</span>
              <div>
                <strong>Priority Insight:</strong> ${escapeHtml(recommendation.reason)}
              </div>
            </div>
          </div>

          <div class="recommendation-actions">
            <button class="btn btn-primary" id="rec-start-btn" data-task-id="${recommendation.id}">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polygon points="10 8 16 12 10 16 10 8"/></svg>
              <span>Start Task</span>
            </button>
            <button class="btn btn-secondary" id="rec-complete-btn" data-task-id="${recommendation.id}">
              ✓ Mark Complete
            </button>
          </div>
        </div>
      </div>
    `;

    document.getElementById('rec-start-btn').addEventListener('click', () => {
      openTimerModal(recommendation);
    });

    document.getElementById('rec-complete-btn').addEventListener('click', () => {
      toggleTaskCompletion(recommendation.id, true);
    });
  }

  /* ==========================================================================
     TASK MANAGER & LIST RENDERING
     ========================================================================== */

  function renderTasksList() {
    const container = elements.tasksListContainer;
    const query = state.searchQuery.toLowerCase().trim();

    // Filter counts
    const totalAll = state.tasks.length;
    const totalOpen = state.tasks.filter(t => !t.completed).length;
    const totalCompleted = state.tasks.filter(t => t.completed).length;

    elements.countFilterAll.textContent = totalAll;
    elements.countFilterOpen.textContent = totalOpen;
    elements.countFilterCompleted.textContent = totalCompleted;

    // Filter by tab
    let filtered = state.tasks.filter(t => {
      if (state.filter === 'open') return !t.completed;
      if (state.filter === 'completed') return t.completed;
      return true;
    });

    // Filter by search query
    if (query) {
      filtered = filtered.filter(t =>
        t.title.toLowerCase().includes(query) ||
        t.course.toLowerCase().includes(query) ||
        t.unit.toLowerCase().includes(query)
      );
    }

    if (filtered.length === 0) {
      if (state.tasks.length === 0) {
        container.innerHTML = `
          <div class="empty-state">
            <div class="empty-state-icon">📝</div>
            <h3>No Tasks Yet</h3>
            <p>Break your semester syllabus into actionable units. Add your first study task above.</p>
            <button class="btn btn-primary" id="empty-state-add-btn">+ Add Your First Task</button>
          </div>
        `;
        document.getElementById('empty-state-add-btn').addEventListener('click', openAddTaskModal);
      } else {
        container.innerHTML = `
          <div class="empty-state">
            <div class="empty-state-icon">🔍</div>
            <h3>No Matching Tasks Found</h3>
            <p>Try searching for a different course keyword or clearing your active filters.</p>
          </div>
        `;
      }
      return;
    }

    container.innerHTML = filtered.map(task => {
      const dueMeta = formatDueDateLabel(task.dueDate, task.completed);
      const diffClass = task.difficulty.toLowerCase();

      return `
        <article class="task-item ${task.completed ? 'completed' : ''}" data-task-id="${task.id}">
          <div class="task-checkbox-wrap">
            <input type="checkbox" class="task-checkbox" ${task.completed ? 'checked' : ''} data-task-id="${task.id}" aria-label="Toggle ${escapeHtml(task.title)}">
          </div>

          <div class="task-item-main">
            <div class="task-item-title-row">
              <span class="task-item-title">${escapeHtml(task.title)}</span>
            </div>

            <div class="task-item-meta">
              <span class="course-tag">${escapeHtml(task.course)}</span>
              <span class="unit-tag">${escapeHtml(task.unit)}</span>
              <span class="badge-difficulty ${diffClass}">${task.difficulty}</span>
              <span class="duration-tag">⏱️ ${formatMinutes(task.estimatedMinutes)}</span>
              <span class="due-badge ${dueMeta.className}">${dueMeta.text}</span>
            </div>
          </div>

          <div class="task-actions">
            ${!task.completed ? `
              <button class="task-action-btn start-task-btn" title="Focus Study Timer" data-task-id="${task.id}" aria-label="Start study focus on ${escapeHtml(task.title)}">
                <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polygon points="10 8 16 12 10 16 10 8"/></svg>
              </button>
            ` : ''}
            <button class="task-action-btn delete-btn" title="Delete task" data-task-id="${task.id}" aria-label="Delete ${escapeHtml(task.title)}">
              <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="3 6 5 6 21 6"/><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><line x1="10" y1="11" x2="10" y2="17"/><line x1="14" y1="11" x2="14" y2="17"/></svg>
            </button>
          </div>
        </article>
      `;
    }).join('');

    // Attach listeners
    container.querySelectorAll('.task-checkbox').forEach(box => {
      box.addEventListener('change', (e) => {
        const taskId = e.target.dataset.taskId;
        toggleTaskCompletion(taskId, e.target.checked);
      });
    });

    container.querySelectorAll('.delete-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const taskId = e.currentTarget.dataset.taskId;
        deleteTask(taskId);
      });
    });

    container.querySelectorAll('.start-task-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const taskId = e.currentTarget.dataset.taskId;
        const task = state.tasks.find(t => t.id === taskId);
        if (task) openTimerModal(task);
      });
    });
  }

  /* ==========================================================================
     COURSE PROGRESS TRACKER RENDERING
     ========================================================================== */

  function renderCourseProgress(courses) {
    const container = elements.coursesListContainer;

    if (!courses || courses.length === 0) {
      container.innerHTML = `
        <div class="empty-state">
          <div class="empty-state-icon">📚</div>
          <h3>No Course Progress Yet</h3>
          <p>Tasks are automatically categorized by subject to reflect your syllabus completion rate.</p>
        </div>
      `;
      return;
    }

    container.innerHTML = courses.map(course => {
      return `
        <div class="course-progress-card">
          <div class="course-card-top">
            <span class="course-card-title">${escapeHtml(course.course)}</span>
            <span class="course-card-percent">${course.percentage}%</span>
          </div>

          <div class="course-progress-bar-wrap">
            <div class="course-progress-bar-fill" style="width: ${course.percentage}%"></div>
          </div>

          <div class="course-card-stats">
            <span><strong>${course.completedCount}</strong> of ${course.totalCount} tasks completed</span>
            <span>${course.openCount} open</span>
          </div>
        </div>
      `;
    }).join('');
  }

  function updateExistingCoursesDatalist() {
    const datalist = elements.existingCoursesDatalist;
    if (!datalist) return;
    const uniqueCourses = [...new Set(state.tasks.map(t => t.course.trim()))];
    datalist.innerHTML = uniqueCourses.map(c => `<option value="${escapeHtml(c)}">`).join('');
  }

  /* ==========================================================================
     TASK ACTIONS (TOGGLE, DELETE, CREATE)
     ========================================================================== */

  async function toggleTaskCompletion(taskId, isCompleted) {
    // Optimistic update in UI
    const target = state.tasks.find(t => t.id === taskId);
    if (target) {
      target.completed = isCompleted;
      target.completedAt = isCompleted ? new Date().toISOString() : null;
      renderTasksList();
    }

    if (state.isDemo) {
      showToast(isCompleted ? 'Topic mastered! Marked complete.' : 'Task reopened.', 'success');
      syncDemoDashboard();
      return;
    }

    try {
      await apiRequest(`/api/tasks/${taskId}`, {
        method: 'PATCH',
        body: JSON.stringify({ completed: isCompleted })
      });

      showToast(isCompleted ? 'Topic mastered! Marked complete.' : 'Task reopened.', 'success');
      // Sync fresh dashboard data (percentages, recommendation)
      fetchDashboardData();
    } catch (err) {
      showToast('Could not update task. ' + err.message, 'error');
      // Revert on failure
      if (target) {
        target.completed = !isCompleted;
        renderTasksList();
      }
    }
  }

  async function deleteTask(taskId) {
    const task = state.tasks.find(t => t.id === taskId);
    const title = task ? `"${task.title}"` : 'this task';
    if (!confirm(`Are you sure you want to delete ${title}?`)) {
      return;
    }

    // Optimistic delete
    state.tasks = state.tasks.filter(t => t.id !== taskId);
    renderTasksList();

    if (state.isDemo) {
      showToast('Task removed from syllabus.', 'info');
      syncDemoDashboard();
      return;
    }

    try {
      await apiRequest(`/api/tasks/${taskId}`, { method: 'DELETE' });
      showToast('Task removed from syllabus.', 'info');
      fetchDashboardData();
    } catch (err) {
      showToast('Could not delete task: ' + err.message, 'error');
      fetchDashboardData();
    }
  }

  // Create Task Modal & Form
  function openAddTaskModal() {
    elements.addTaskModal.classList.remove('hidden');
    elements.taskFormAlert.classList.add('hidden');
    elements.addTaskForm.reset();
    document.querySelectorAll('#add-task-form .field-error').forEach(el => { el.textContent = ''; });

    // Set default due date to today
    elements.taskDueDate.value = getTodayString();
    document.getElementById('task-minutes').value = '45';
    document.getElementById('task-title').focus();
  }

  function closeAddTaskModal() {
    elements.addTaskModal.classList.add('hidden');
  }

  async function handleCreateTask(e) {
    e.preventDefault();
    document.querySelectorAll('#add-task-form .field-error').forEach(el => { el.textContent = ''; });
    elements.taskFormAlert.classList.add('hidden');

    const form = elements.addTaskForm;
    const title = form.title.value.trim();
    const course = form.course.value.trim();
    const unit = form.unit.value.trim();
    const difficulty = form.difficulty.value;
    const dueDate = form.dueDate.value;
    const estimatedMinutes = parseInt(form.estimatedMinutes.value, 10);

    let hasError = false;

    if (!title) {
      document.getElementById('task-title-error').textContent = 'Please enter a task title.';
      hasError = true;
    }
    if (!course) {
      document.getElementById('task-course-error').textContent = 'Please enter the course/subject name.';
      hasError = true;
    }
    if (!unit) {
      document.getElementById('task-unit-error').textContent = 'Please enter unit or module details.';
      hasError = true;
    }
    if (!dueDate) {
      document.getElementById('task-due-date-error').textContent = 'Please choose a due date.';
      hasError = true;
    }
    if (isNaN(estimatedMinutes) || estimatedMinutes <= 0 || estimatedMinutes > 720) {
      document.getElementById('task-minutes-error').textContent = 'Estimated study time must be between 1 and 720 minutes.';
      hasError = true;
    }

    if (hasError) return;

    const btnText = elements.saveTaskBtn.querySelector('.btn-text');
    const spinner = elements.saveTaskBtn.querySelector('.btn-spinner');
    elements.saveTaskBtn.disabled = true;
    btnText.textContent = 'Saving...';
    spinner.classList.remove('hidden');

    if (state.isDemo) {
      const newTask = {
        id: 'demo-task-' + Date.now(),
        userId: state.user.id,
        title,
        course,
        unit,
        difficulty,
        dueDate,
        estimatedMinutes,
        completed: false,
        createdAt: new Date().toISOString()
      };
      state.tasks.unshift(newTask);
      closeAddTaskModal();
      showToast('Syllabus task added successfully!', 'success');
      syncDemoDashboard();
      elements.saveTaskBtn.disabled = false;
      btnText.textContent = 'Save Task';
      spinner.classList.add('hidden');
      return;
    }

    try {
      await apiRequest('/api/tasks', {
        method: 'POST',
        body: JSON.stringify({
          title,
          course,
          unit,
          difficulty,
          dueDate,
          estimatedMinutes
        })
      });

      closeAddTaskModal();
      showToast('Syllabus task added successfully!', 'success');
      fetchDashboardData();
    } catch (err) {
      elements.taskFormAlert.className = 'auth-alert error';
      elements.taskFormAlert.textContent = err.message;
      elements.taskFormAlert.classList.remove('hidden');
    } finally {
      elements.saveTaskBtn.disabled = false;
      btnText.textContent = 'Save Task';
      spinner.classList.add('hidden');
    }
  }

  /* ==========================================================================
     FOCUS STUDY TIMER MODAL
     ========================================================================== */

  function openTimerModal(task) {
    state.timer.activeTaskId = task.id;
    const mins = task.estimatedMinutes || 25;
    state.timer.totalSeconds = mins * 60;
    state.timer.remainingSeconds = mins * 60;
    state.timer.isRunning = false;

    elements.timerTaskTitle.textContent = task.title;
    elements.timerCourseUnit.textContent = `${task.course} • ${task.unit}`;
    updateTimerDisplay();

    elements.timerToggleText.textContent = 'Start Focus';
    elements.timerToggleIcon.textContent = '▶';
    elements.timerModal.classList.remove('hidden');
  }

  function closeTimerModal() {
    resetTimer();
    elements.timerModal.classList.add('hidden');
  }

  function updateTimerDisplay() {
    const mins = Math.floor(state.timer.remainingSeconds / 60);
    const secs = state.timer.remainingSeconds % 60;
    elements.timerCountdown.textContent = `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;

    // Update circular ring (circumference is 2 * PI * 95 approx 597)
    const circumference = 597;
    const progress = (state.timer.totalSeconds - state.timer.remainingSeconds) / state.timer.totalSeconds;
    const offset = circumference * (1 - progress);
    elements.timerProgressRing.style.strokeDashoffset = offset;
  }

  function toggleTimer() {
    if (state.timer.isRunning) {
      // Pause
      clearInterval(state.timer.intervalId);
      state.timer.isRunning = false;
      elements.timerToggleText.textContent = 'Resume Focus';
      elements.timerToggleIcon.textContent = '\u25B6';
    } else {
      // Start
      state.timer.isRunning = true;
      elements.timerToggleText.textContent = 'Pause Focus';
      elements.timerToggleIcon.textContent = '\u23F8';

      state.timer.intervalId = setInterval(() => {
        if (state.timer.remainingSeconds > 0) {
          state.timer.remainingSeconds -= 1;
          updateTimerDisplay();
        } else {
          clearInterval(state.timer.intervalId);
          state.timer.isRunning = false;
          elements.timerToggleText.textContent = 'Session Complete!';
          elements.timerToggleIcon.textContent = '\u2713';
          showToast('Focus session complete! Outstanding work.', 'success');
        }
      }, 1000);
    }
  }

  function resetTimer() {
    if (state.timer.intervalId) {
      clearInterval(state.timer.intervalId);
    }
    state.timer.isRunning = false;
    state.timer.remainingSeconds = state.timer.totalSeconds;
    updateTimerDisplay();
    elements.timerToggleText.textContent = 'Start Focus';
    elements.timerToggleIcon.textContent = '\u25B6';
  }

  /* ==========================================================================
     EVENT LISTENERS INITIALIZATION
     ========================================================================== */

  function initEventListeners() {
    // Tab switching
    elements.tabRegisterBtn.addEventListener('click', () => switchAuthTab('register'));
    elements.tabLoginBtn.addEventListener('click', () => switchAuthTab('login'));
    elements.switchToLogin.addEventListener('click', () => switchAuthTab('login'));
    elements.switchToRegister.addEventListener('click', () => switchAuthTab('register'));

    // Auth Form submissions
    elements.registerForm.addEventListener('submit', handleRegister);
    elements.loginForm.addEventListener('submit', handleLogin);
    elements.logoutBtn.addEventListener('click', () => handleLogout(true));

    // Password show/hide toggles
    document.querySelectorAll('.toggle-password-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const targetId = btn.dataset.target;
        const input = document.getElementById(targetId);
        if (input) {
          const isPassword = input.type === 'password';
          input.type = isPassword ? 'text' : 'password';
          btn.style.color = isPassword ? 'var(--coral-main)' : 'var(--text-muted)';
        }
      });
    });

    // Demo Mode Trigger Buttons
    if (elements.demoTriggerBtns) {
      elements.demoTriggerBtns.forEach(btn => {
        btn.addEventListener('click', activateDemoMode);
      });
    }

    // Task Filter Pills
    elements.filterPills.forEach(pill => {
      pill.addEventListener('click', () => {
        elements.filterPills.forEach(p => p.classList.remove('active'));
        pill.classList.add('active');
        state.filter = pill.dataset.filter;
        renderTasksList();
      });
    });

    // Task Search
    elements.taskSearchInput.addEventListener('input', (e) => {
      state.searchQuery = e.target.value;
      if (state.searchQuery) {
        elements.clearSearchBtn.classList.remove('hidden');
      } else {
        elements.clearSearchBtn.classList.add('hidden');
      }
      renderTasksList();
    });

    elements.clearSearchBtn.addEventListener('click', () => {
      elements.taskSearchInput.value = '';
      state.searchQuery = '';
      elements.clearSearchBtn.classList.add('hidden');
      renderTasksList();
    });

    // Add Task Modal triggers
    elements.openAddTaskBtn.addEventListener('click', openAddTaskModal);
    elements.quickAddTaskBtn.addEventListener('click', openAddTaskModal);
    elements.closeModalBtn.addEventListener('click', closeAddTaskModal);
    elements.cancelTaskBtn.addEventListener('click', closeAddTaskModal);
    elements.addTaskForm.addEventListener('submit', handleCreateTask);

    // Quick minute chips in modal
    document.querySelectorAll('.quick-chip').forEach(chip => {
      chip.addEventListener('click', () => {
        document.getElementById('task-minutes').value = chip.dataset.mins;
      });
    });

    // Timer modal triggers
    elements.closeTimerBtn.addEventListener('click', closeTimerModal);
    elements.timerToggleBtn.addEventListener('click', toggleTimer);
    elements.timerResetBtn.addEventListener('click', resetTimer);
    elements.timerCompleteTaskBtn.addEventListener('click', () => {
      if (state.timer.activeTaskId) {
        toggleTaskCompletion(state.timer.activeTaskId, true);
        closeTimerModal();
      }
    });

    // Close modals on Escape key or backdrop click
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        if (!elements.addTaskModal.classList.contains('hidden')) closeAddTaskModal();
        if (!elements.timerModal.classList.contains('hidden')) closeTimerModal();
      }
    });

    [elements.addTaskModal, elements.timerModal].forEach(modal => {
      modal.addEventListener('click', (e) => {
        if (e.target === modal) {
          if (modal === elements.addTaskModal) closeAddTaskModal();
          if (modal === elements.timerModal) closeTimerModal();
        }
      });
    });
  }

  /* ==========================================================================
     APP BOOTSTRAP
     ========================================================================== */

  async function initApp() {
    initEventListeners();
    checkBackendHealth();

    // Check if user session already exists in localStorage
    if (state.token) {
      try {
        const dashData = await apiRequest('/api/dashboard');
        state.user = dashData.user;
        localStorage.setItem('studyflow_user', JSON.stringify(dashData.user));
        showAppView();
      } catch (err) {
        console.warn('Session verification failed:', err);
        handleLogout(false);
      }
    } else {
      showAuthView();
    }
  }

  // Start app when DOM is ready
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initApp);
  } else {
    initApp();
  }

})();
