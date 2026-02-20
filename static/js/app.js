// ===== GLOBAL STATE =====
let allSheetsData = window.AppConfig?.allSheetsData || {};
let currentDataVersion = window.AppConfig?.dataVersion || 0;
let currentSheet = window.AppConfig?.initialSheet || '';
let buttonData = allSheetsData[currentSheet];
let selectedTask = null;
let allTasks = [];
let filterBarOpen = false;
let originalDetails = {};
let addTaskKnownColumns = [];

const buttonColors = [
  'bg-blue-500 hover:bg-blue-600',
  'bg-emerald-500 hover:bg-emerald-600',
  'bg-rose-500 hover:bg-rose-600',
  'bg-amber-500 hover:bg-amber-600',
  'bg-violet-500 hover:bg-violet-600',
  'bg-pink-500 hover:bg-pink-600',
  'bg-indigo-500 hover:bg-indigo-600',
  'bg-orange-500 hover:bg-orange-600',
  'bg-teal-500 hover:bg-teal-600',
  'bg-cyan-500 hover:bg-cyan-600'
];

// ===== DOM ELEMENTS =====
const filterBar = document.getElementById('filterBar');
const filterContent = document.getElementById('filterContent');
const expandIcon = document.getElementById('expandIcon');
const activeFilterBadge = document.getElementById('activeFilterBadge');
const scrollContainer = document.getElementById('scrollContainer');
const scrollLeftBtn = document.getElementById('scrollLeft');
const scrollRightBtn = document.getElementById('scrollRight');
const leftShadow = document.getElementById('leftShadow');
const rightShadow = document.getElementById('rightShadow');

// ===== FILTER BAR (click-to-toggle) =====
function toggleFilterBar() {
  filterBarOpen = !filterBarOpen;
  if (filterBarOpen) {
    filterContent.classList.remove('hidden');
    filterBar.classList.add('filter-expanded');
    expandIcon.style.transform = 'rotate(180deg)';
  } else {
    filterContent.classList.add('hidden');
    filterBar.classList.remove('filter-expanded');
    expandIcon.style.transform = 'rotate(0deg)';
  }
}

window.toggleFilterBar = toggleFilterBar;

function updateActiveFilterBadge() {
  const searchTerm = document.getElementById('searchBox').value;
  const statusFilter = document.getElementById('statusFilter').value;
  const priorityFilter = document.getElementById('priorityFilter').value;
  const hideCompleted = document.getElementById('hideCompleted').checked;

  let activeCount = 0;
  if (searchTerm) activeCount++;
  if (statusFilter !== 'All') activeCount++;
  if (priorityFilter !== 'All') activeCount++;
  if (hideCompleted) activeCount++;

  if (activeCount > 0) {
    activeFilterBadge.textContent = `${activeCount} active`;
    activeFilterBadge.classList.remove('hidden');
    document.getElementById('clearFilters').classList.remove('hidden');
  } else {
    activeFilterBadge.classList.add('hidden');
    document.getElementById('clearFilters').classList.add('hidden');
  }
}

// ===== TASK PARSING =====
function parseTaskDetails(details) {
  const lines = details.split('\n');
  let description = '';
  let status = '';
  let priority = '';
  let assignedTo = '';
  let deadline = '';

  lines.forEach(line => {
    if (line.includes('Description:')) {
      description = line.split('Description:')[1].trim();
    }
    if (line.includes('Status:')) {
      status = line.split('Status:')[1].trim();
    }
    if (line.includes('Priority:')) {
      priority = line.split('Priority:')[1].trim();
    }
    if (line.includes('Assigned To:')) {
      assignedTo = line.split('Assigned To:')[1].trim();
    }
    if (line.includes('Deadline:')) {
      const rawD = line.split('Deadline:')[1].trim();
      const dd = new Date(rawD);
      deadline = isNaN(dd) ? rawD : `${String(dd.getDate()).padStart(2,'0')}/${String(dd.getMonth()+1).padStart(2,'0')}/${String(dd.getFullYear()).slice(-2)}`;
    }
  });

  return { description, status, priority, assignedTo, deadline };
}

function truncateDescription(description, maxWords = 5) {
  if (!description) return '';
  const words = description.split(' ');
  if (words.length <= maxWords) return description;
  return words.slice(0, maxWords).join(' ') + '...';
}

function getStatusBadge(status) {
  const statusConfig = {
    'Not Started': 'bg-gray-500/20 text-gray-400',
    'In Progress': 'bg-amber-500/20 text-amber-400',
    'Completed': 'bg-emerald-500/20 text-emerald-400',
    'Blocked': 'bg-rose-500/20 text-rose-400'
  };
  const color = statusConfig[status] || 'bg-gray-500/20 text-gray-400';
  return `<span class="inline-flex items-center px-2 py-0.5 rounded-md text-xs font-medium ${color}">${status}</span>`;
}

function getPriorityIndicator(priority) {
  const config = {
    'High': '<span class="text-rose-400 text-xs font-semibold">High</span>',
    'Medium': '<span class="text-amber-400 text-xs font-semibold">Med</span>',
    'Low': '<span class="text-blue-400 text-xs font-semibold">Low</span>'
  };
  return config[priority] || '';
}

// ===== FILTERING =====
function getFilteredTasks() {
  const searchTerm = document.getElementById('searchBox').value.toLowerCase();
  const statusFilter = document.getElementById('statusFilter').value;
  const priorityFilter = document.getElementById('priorityFilter').value;
  const hideCompleted = document.getElementById('hideCompleted').checked;

  return allTasks.filter(task => {
    if (searchTerm && !task.name.toLowerCase().includes(searchTerm)) return false;
    if (hideCompleted && task.status === 'Completed') return false;
    if (statusFilter !== 'All' && task.status !== statusFilter) return false;
    if (priorityFilter !== 'All' && task.priority !== priorityFilter) return false;
    return true;
  });
}

function updateTaskCounter(filteredCount, totalCount) {
  const counter = document.getElementById('taskCounter');
  if (filteredCount === totalCount) {
    counter.textContent = `Showing ${totalCount} task${totalCount !== 1 ? 's' : ''}`;
  } else {
    counter.textContent = `Showing ${filteredCount} of ${totalCount} task${totalCount !== 1 ? 's' : ''}`;
  }
}

function applyFilters() {
  createButtons(currentSheet, true);
  updateActiveFilterBadge();
  if (selectedTask) {
    showTaskDetails(selectedTask);
  }
}

// ===== BUTTON CREATION =====
function createButtons(sheetName, shouldApplyFilters = false) {
  const container = document.getElementById('buttonContainer');
  container.innerHTML = '';

  allTasks = [];
  const taskNames = Object.keys(allSheetsData[sheetName] || {});

  taskNames.forEach(taskName => {
    const instances = allSheetsData[sheetName][taskName];

    instances.forEach((instance, index) => {
      const details = typeof instance === 'string' ? instance : instance.details;
      const metadata = typeof instance === 'object' && instance.metadata ? instance.metadata : null;
      const { description, status, priority, assignedTo, deadline } = parseTaskDetails(details);

      allTasks.push({
        name: taskName,
        instanceIndex: index,
        description, status, priority, assignedTo, deadline,
        details, metadata
      });
    });
  });

  const tasksToDisplay = shouldApplyFilters ? getFilteredTasks() : allTasks;
  updateTaskCounter(tasksToDisplay.length, allTasks.length);

  const uniqueTaskNames = [...new Set(tasksToDisplay.map(t => t.name))];

  uniqueTaskNames.forEach((taskName, index) => {
    const button = document.createElement('button');
    button.textContent = taskName;
    button.onclick = () => showTaskDetails(taskName);
    button.id = `task-${index}`;

    const colorClass = buttonColors[index % buttonColors.length];
    button.className = `task-button ${colorClass} text-white`;

    container.appendChild(button);
  });

  if (selectedTask && !uniqueTaskNames.includes(selectedTask)) {
    selectedTask = null;
    document.getElementById('textDisplay').innerHTML = `
      <div class="flex flex-col items-center justify-center h-32 text-gray-600">
        <svg class="w-10 h-10 mb-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"></path>
        </svg>
        <p class="text-sm">Click a task to view details</p>
      </div>`;
  }

  updateScrollIndicators();
}

// ===== TASK INSTANCE TOGGLE =====
function toggleTaskInstance(instanceId) {
  const instance = document.getElementById(`instance-${instanceId}`);
  const details = document.getElementById(`details-${instanceId}`);
  const icon = document.getElementById(`icon-${instanceId}`);

  if (instance.classList.contains('task-instance-collapsed')) {
    instance.classList.remove('task-instance-collapsed');
    instance.classList.add('task-instance-expanded');
    details.classList.remove('task-details-hidden');
    details.classList.add('task-details-visible');
    icon.classList.add('toggle-icon-expanded');
  } else {
    instance.classList.remove('task-instance-expanded');
    instance.classList.add('task-instance-collapsed');
    details.classList.remove('task-details-visible');
    details.classList.add('task-details-hidden');
    icon.classList.remove('toggle-icon-expanded');
  }
}

// ===== TASK DETAILS =====
function showTaskDetails(taskName) {
  selectedTask = taskName;

  const display = document.getElementById('textDisplay');
  const filteredTasks = getFilteredTasks();
  const instances = filteredTasks.filter(t => t.name === taskName);

  if (instances.length === 0) return;

  let html = '';

  const totalInstances = allTasks.filter(t => t.name === taskName).length;
  if (totalInstances > 1) {
    html += `<div class="mb-4 text-xs text-gray-500">
      <span class="font-semibold text-gray-400">${taskName}</span> &mdash; ${instances.length} ${instances.length !== totalInstances ? `of ${totalInstances}` : ''} instance${instances.length !== 1 ? 's' : ''}
    </div>`;
  }

  instances.forEach((task, index) => {
    const shortDesc = truncateDescription(task.description, 4);
    const statusBadge = getStatusBadge(task.status);
    const priorityIndicator = getPriorityIndicator(task.priority);
    const deadlineText = task.deadline ? `Due: ${task.deadline}` : '';
    const hasMetadata = task.metadata !== null;
    const taskIndex = allTasks.indexOf(task);

    const borderColor = index % 2 === 0 ? 'border-l-violet-500/50' : 'border-l-blue-500/50';

    const collapsedClass = 'task-instance-expanded';
    const detailsClass = 'task-details-visible';
    const iconClass = 'toggle-icon-expanded';

    const sourceFile = hasMetadata ? task.metadata.file_path.split(/[\\\/]/).pop() : '';

    html += `
      <div id="instance-${taskName}-${index}" class="task-instance ${collapsedClass} bg-white/[0.02] border border-white/5 border-l-2 ${borderColor} rounded-lg mb-3 overflow-hidden">
        <!-- Header -->
        <div class="p-4 cursor-pointer hover:bg-white/[0.03] transition flex items-center justify-between" onclick="toggleTaskInstance('${taskName}-${index}')">
          <div class="flex items-center gap-3 flex-1 min-w-0">
            <svg id="icon-${taskName}-${index}" class="toggle-icon ${iconClass} w-4 h-4 text-gray-500 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"></path>
            </svg>
            <div class="flex-1 flex flex-wrap items-center gap-2 min-w-0">
              <span class="font-semibold text-sm text-white">${taskName}</span>
              ${shortDesc ? `<span class="text-gray-500 text-xs truncate">${shortDesc}</span>` : ''}
              ${statusBadge}
              ${priorityIndicator}
              ${deadlineText ? `<span class="text-xs text-orange-400/80">${deadlineText}</span>` : ''}
            </div>
          </div>
          ${hasMetadata ? `
            <button onclick="event.stopPropagation(); startEdit('${taskName}-${index}', ${taskIndex})"
                    class="edit-btn bg-accent-500/20 hover:bg-accent-500/30 text-accent-400 px-3 py-1 rounded-md text-xs font-medium ml-2 transition">
              Edit
            </button>
          ` : ''}
        </div>

        <!-- Details -->
        <div id="details-${taskName}-${index}" class="task-details ${detailsClass} px-4 pb-4 pt-0">
          <!-- View Mode -->
          <div id="view-${taskName}-${index}" class="pl-7 border-l border-white/5 ml-2 space-y-2">
            <div class="whitespace-pre-line text-gray-400 text-sm leading-relaxed">${task.details.replace(/Deadline:\s*(\d{4}-\d{2}-\d{2})\s*\d{2}:\d{2}:\d{2}/g, (m, d) => { const p = d.split('-'); return 'Deadline: ' + p[2] + '/' + p[1] + '/' + p[0].slice(-2); })}</div>
          </div>

          <!-- Edit Mode -->
          <div id="edit-${taskName}-${index}" class="hidden pl-7 border-l border-accent-500/30 ml-2 space-y-4">
            <textarea id="textarea-${taskName}-${index}"
                      class="edit-textarea"
                      data-task-index="${taskIndex}">${task.details}</textarea>

            <div class="flex gap-2">
              <button onclick="saveEdit('${taskName}-${index}')"
                      id="save-btn-${taskName}-${index}"
                      class="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-1.5 rounded-lg text-sm font-medium transition">
                Save Changes
              </button>
              <button onclick="cancelEdit('${taskName}-${index}')"
                      class="bg-white/5 hover:bg-white/10 text-gray-300 px-4 py-1.5 rounded-lg text-sm font-medium transition">
                Cancel
              </button>
            </div>

            <div class="text-xs text-gray-600">
              <p>Format: <code class="bg-white/5 px-1.5 py-0.5 rounded text-gray-400">ColumnName: value</code></p>
              <p>Add new columns by adding new lines in this format.</p>
              ${hasMetadata ? `<p class="mt-1 text-gray-600">Source: ${sourceFile}</p>` : ''}
            </div>
          </div>
        </div>
      </div>
    `;
  });

  display.innerHTML = html;

  document.querySelectorAll('.task-button').forEach(btn => {
    btn.classList.remove('task-button-selected');
    if (btn.textContent === taskName) {
      btn.classList.add('task-button-selected');
    }
  });
}

// ===== EDIT FUNCTIONALITY =====
function startEdit(instanceId, taskIndex) {
  const viewDiv = document.getElementById(`view-${instanceId}`);
  const editDiv = document.getElementById(`edit-${instanceId}`);
  const textarea = document.getElementById(`textarea-${instanceId}`);

  originalDetails[instanceId] = textarea.value;

  viewDiv.classList.add('hidden');
  editDiv.classList.remove('hidden');
  textarea.focus();
}

function cancelEdit(instanceId) {
  const viewDiv = document.getElementById(`view-${instanceId}`);
  const editDiv = document.getElementById(`edit-${instanceId}`);
  const textarea = document.getElementById(`textarea-${instanceId}`);

  if (originalDetails[instanceId]) {
    textarea.value = originalDetails[instanceId];
  }

  editDiv.classList.add('hidden');
  viewDiv.classList.remove('hidden');
}

async function saveEdit(instanceId) {
  const textarea = document.getElementById(`textarea-${instanceId}`);
  const saveBtn = document.getElementById(`save-btn-${instanceId}`);
  const taskIndex = parseInt(textarea.dataset.taskIndex);
  const task = allTasks[taskIndex];

  if (!task || !task.metadata) {
    showNotification('Error: Cannot save - missing metadata', 'error');
    return;
  }

  const newDetails = textarea.value;
  const updates = {};
  const newColumns = {};
  const existingColumns = task.metadata.columns;
  const newColumnValues = {};

  newDetails.split('\n').forEach(line => {
    const colonIndex = line.indexOf(':');
    if (colonIndex > 0) {
      const key = line.substring(0, colonIndex).trim();
      const value = line.substring(colonIndex + 1).trim();
      newColumnValues[key] = value;

      if (existingColumns.includes(key)) {
        updates[key] = value;
      } else {
        newColumns[key] = value;
      }
    }
  });

  existingColumns.forEach(col => {
    if (col !== existingColumns[0]) {
      const oldValue = task.metadata.raw_values[col];
      const hasNewValue = newColumnValues.hasOwnProperty(col);
      if (oldValue && !hasNewValue) {
        updates[col] = '';
      }
    }
  });

  saveBtn.disabled = true;
  const originalBtnText = saveBtn.innerHTML;
  saveBtn.innerHTML = '<span class="saving-spinner">&#x21bb;</span> Saving...';

  try {
    const response = await fetch('/api/save-task', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file_path: task.metadata.file_path,
        sheet_name: task.metadata.sheet_name,
        row_index: task.metadata.row_index,
        task_name: task.metadata.task_name,
        updates: updates,
        new_columns: Object.keys(newColumns).length > 0 ? newColumns : null
      })
    });

    const result = await response.json();

    if (response.ok) {
      showNotification('Changes saved successfully!', 'success');
      cancelEdit(instanceId);
    } else {
      throw new Error(result.detail || 'Failed to save');
    }
  } catch (error) {
    showNotification(`Error: ${error.message}`, 'error');
  } finally {
    saveBtn.disabled = false;
    saveBtn.innerHTML = originalBtnText;
  }
}

function showNotification(message, type = 'success') {
  let notification = document.getElementById('appNotification');
  if (!notification) {
    notification = document.createElement('div');
    notification.id = 'appNotification';
    notification.className = 'notification-toast';
    document.body.appendChild(notification);
  }

  notification.classList.remove('bg-emerald-600', 'bg-rose-600');
  notification.classList.add(type === 'success' ? 'bg-emerald-600' : 'bg-rose-600');
  notification.textContent = message;

  notification.classList.add('visible');

  setTimeout(() => {
    notification.classList.remove('visible');
  }, 3000);
}

// Keep old name for backward compat with due-soon code
const showEditNotification = showNotification;

// ===== SHEET SWITCHING =====
function switchSheet(sheetName) {
  currentSheet = sheetName;
  buttonData = allSheetsData[sheetName];

  document.getElementById('pageTitle').textContent = sheetName;

  document.getElementById('searchBox').value = '';
  document.getElementById('statusFilter').value = 'All';
  document.getElementById('priorityFilter').value = 'All';
  document.getElementById('hideCompleted').checked = true;

  createButtons(sheetName, true);
  updateActiveFilterBadge();

  document.getElementById('textDisplay').innerHTML = `
    <div class="flex flex-col items-center justify-center h-32 text-gray-600">
      <svg class="w-10 h-10 mb-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"></path>
      </svg>
      <p class="text-sm">Click a task to view details</p>
    </div>`;

  document.querySelectorAll('.sheet-btn').forEach(btn => {
    btn.classList.remove('sheet-btn-active');
    if (btn.dataset.sheet === sheetName) {
      btn.classList.add('sheet-btn-active');
    }
  });

  updateExcelButton();
  updateAddTaskButton();
}

// ===== HORIZONTAL SCROLL =====
function updateScrollIndicators() {
  const container = scrollContainer;
  const isScrollable = container.scrollWidth > container.clientWidth;
  const isAtStart = container.scrollLeft <= 10;
  const isAtEnd = container.scrollLeft >= container.scrollWidth - container.clientWidth - 10;

  if (isScrollable) {
    scrollLeftBtn.style.opacity = isAtStart ? '0' : '1';
    scrollRightBtn.style.opacity = isAtEnd ? '0' : '1';
    leftShadow.style.opacity = isAtStart ? '0' : '1';
    rightShadow.style.opacity = isAtEnd ? '0' : '1';
  } else {
    scrollLeftBtn.style.opacity = '0';
    scrollRightBtn.style.opacity = '0';
    leftShadow.style.opacity = '0';
    rightShadow.style.opacity = '0';
  }
}

// ===== SSE (Real-time Updates) =====
let eventSource = null;
let reconnectAttempts = 0;
const maxReconnectAttempts = 10;

function connectSSE() {
  if (eventSource) eventSource.close();

  eventSource = new EventSource('/events');

  eventSource.onopen = function() {
    reconnectAttempts = 0;
  };

  eventSource.onmessage = function(event) {
    const newVersion = parseInt(event.data);
    if (newVersion > currentDataVersion) {
      fetchLatestData();
    }
  };

  eventSource.onerror = function() {
    eventSource.close();
    if (reconnectAttempts < maxReconnectAttempts) {
      reconnectAttempts++;
      setTimeout(connectSSE, 2000 * reconnectAttempts);
    }
  };
}

async function fetchLatestData(showToast = true) {
  try {
    const response = await fetch('/api/data');
    const data = await response.json();

    if (data.version > currentDataVersion) {
      allSheetsData = data.all_sheets_data;
      currentDataVersion = data.version;

      const newSheetNames = data.sheet_names;
      updateSheetButtons(newSheetNames);

      if (allSheetsData[currentSheet]) {
        buttonData = allSheetsData[currentSheet];
        createButtons(currentSheet, true);

        if (selectedTask && allSheetsData[currentSheet][selectedTask]) {
          showTaskDetails(selectedTask);
        } else if (selectedTask) {
          selectedTask = null;
          document.getElementById('textDisplay').innerHTML = `
            <div class="flex flex-col items-center justify-center h-32 text-gray-600">
              <svg class="w-10 h-10 mb-3" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2"></path>
              </svg>
              <p class="text-sm">Click a task to view details</p>
            </div>`;
        }
      } else {
        if (newSheetNames.length > 0) {
          switchSheet(newSheetNames[0]);
        }
      }

      if (showToast) {
        showNotification('Data updated', 'success');
      }
      updateDueSoonBadgeOnLoad();
      updateExcelButton();
      updateAddTaskButton();
    }
  } catch (err) {
    console.error('Failed to fetch latest data:', err);
  }
}

function updateSheetButtons(newSheetNames) {
  const container = document.getElementById('sheetButtons');
  const existingSheets = Array.from(container.querySelectorAll('.sheet-btn')).map(btn => btn.dataset.sheet);

  const sheetsChanged = newSheetNames.length !== existingSheets.length ||
                        newSheetNames.some((name, i) => name !== existingSheets[i]);

  if (sheetsChanged) {
    container.innerHTML = '';
    newSheetNames.forEach(sheetName => {
      const button = document.createElement('button');
      button.onclick = () => switchSheet(sheetName);
      button.className = 'sheet-btn w-full text-left px-4 py-2.5 rounded-lg text-sm font-medium transition-all duration-200 truncate';
      button.dataset.sheet = sheetName;
      button.textContent = sheetName;

      if (sheetName === currentSheet) {
        button.classList.add('sheet-btn-active');
      }

      container.appendChild(button);
    });
  }
}

// ===== DUE SOON =====
let dueSoonOriginalDetails = {};

function updateBodyScrollLock() {
  const dueSoonModal = document.getElementById('dueSoonModal');
  const addTaskModal = document.getElementById('addTaskModal');
  const dueSoonOpen = dueSoonModal && !dueSoonModal.classList.contains('hidden');
  const addTaskOpen = addTaskModal && !addTaskModal.classList.contains('hidden');
  document.body.style.overflow = (dueSoonOpen || addTaskOpen) ? 'hidden' : '';
}

function openDueSoonPopup() {
  const modal = document.getElementById('dueSoonModal');
  modal.classList.remove('hidden');
  updateBodyScrollLock();
  filterDueSoonTasks();
}

function closeDueSoonPopup() {
  const modal = document.getElementById('dueSoonModal');
  modal.classList.add('hidden');
  updateBodyScrollLock();
}

document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    closeDueSoonPopup();
    closeAddTaskModal();
  }
});

function getAllTasksAcrossProjects() {
  const allProjectTasks = [];

  Object.keys(allSheetsData).forEach(sheetName => {
    const taskNames = Object.keys(allSheetsData[sheetName] || {});

    taskNames.forEach(taskName => {
      const instances = allSheetsData[sheetName][taskName];

      instances.forEach((instance, index) => {
        const details = typeof instance === 'string' ? instance : instance.details;
        const metadata = typeof instance === 'object' && instance.metadata ? instance.metadata : null;
        const { description, status, priority, assignedTo, deadline } = parseTaskDetails(details);

        allProjectTasks.push({
          name: taskName, project: sheetName, instanceIndex: index,
          description, status, priority, assignedTo, deadline,
          deadlineDate: parseDeadlineToDate(deadline),
          details, metadata
        });
      });
    });
  });

  return allProjectTasks;
}

function parseDeadlineToDate(deadline) {
  if (!deadline) return null;

  const ddmmyy = deadline.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (ddmmyy) {
    let year = parseInt(ddmmyy[3]);
    if (year < 100) year += 2000;
    return new Date(year, parseInt(ddmmyy[2]) - 1, parseInt(ddmmyy[1]));
  }

  const yyyymmdd = deadline.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (yyyymmdd) {
    return new Date(parseInt(yyyymmdd[1]), parseInt(yyyymmdd[2]) - 1, parseInt(yyyymmdd[3]));
  }

  const parsed = new Date(deadline);
  return isNaN(parsed) ? null : parsed;
}

function getDaysUntilDeadline(deadlineDate) {
  if (!deadlineDate) return Infinity;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const deadline = new Date(deadlineDate);
  deadline.setHours(0, 0, 0, 0);
  return Math.ceil((deadline - today) / (1000 * 60 * 60 * 24));
}

function getDeadlineClass(daysUntil) {
  if (daysUntil < 0) return 'deadline-overdue';
  if (daysUntil === 0) return 'deadline-today';
  if (daysUntil <= 3) return 'deadline-soon';
  return 'deadline-later';
}

function getDeadlineText(daysUntil, deadline) {
  if (daysUntil === Infinity) return '';
  if (daysUntil < 0) return `Overdue (${deadline})`;
  if (daysUntil === 0) return 'Due Today';
  if (daysUntil === 1) return 'Due Tomorrow';
  return `Due in ${daysUntil} days`;
}

function filterDueSoonTasks() {
  const daysFilter = document.getElementById('dueSoonDaysFilter').value;
  const groupBy = document.getElementById('dueSoonGroupBy').value;
  const hideCompleted = document.getElementById('dueSoonHideCompleted').checked;

  let tasks = getAllTasksAcrossProjects();

  if (hideCompleted) {
    tasks = tasks.filter(t => t.status !== 'Completed');
  }

  if (daysFilter !== 'all') {
    const days = parseInt(daysFilter);
    tasks = tasks.filter(t => {
      const daysUntil = getDaysUntilDeadline(t.deadlineDate);
      return daysUntil <= days;
    });
  }

  const priorityOrder = { 'High': 0, 'Medium': 1, 'Low': 2, '': 3 };
  tasks.sort((a, b) => {
    const daysA = getDaysUntilDeadline(a.deadlineDate);
    const daysB = getDaysUntilDeadline(b.deadlineDate);
    if (daysA !== daysB) return daysA - daysB;
    return (priorityOrder[a.priority] || 3) - (priorityOrder[b.priority] || 3);
  });

  document.getElementById('dueSoonCounter').textContent = `${tasks.length} task${tasks.length !== 1 ? 's' : ''}`;
  updateDueSoonBadge(tasks.length);
  renderDueSoonTasks(tasks, groupBy);
}

function updateDueSoonBadge(count) {
  const badge = document.getElementById('dueSoonBadge');
  if (count > 0) {
    badge.textContent = count > 99 ? '99+' : count;
    badge.classList.remove('hidden');
  } else {
    badge.classList.add('hidden');
  }
}

function groupTasks(tasks, groupBy) {
  const groups = {};

  tasks.forEach(task => {
    let key;
    switch (groupBy) {
      case 'priority': key = task.priority || 'No Priority'; break;
      case 'status': key = task.status || 'No Status'; break;
      case 'project': key = task.project; break;
      case 'deadline':
        const daysUntil = getDaysUntilDeadline(task.deadlineDate);
        if (daysUntil < 0) key = 'Overdue';
        else if (daysUntil === 0) key = 'Due Today';
        else if (daysUntil <= 3) key = 'Due in 1-3 Days';
        else if (daysUntil <= 7) key = 'Due in 4-7 Days';
        else key = 'Due Later';
        break;
      default: key = 'All Tasks';
    }
    if (!groups[key]) groups[key] = [];
    groups[key].push(task);
  });

  return groups;
}

function getGroupOrder(groupBy) {
  switch (groupBy) {
    case 'priority': return ['High', 'Medium', 'Low', 'No Priority'];
    case 'status': return ['Blocked', 'In Progress', 'Not Started', 'Completed', 'No Status'];
    case 'deadline': return ['Overdue', 'Due Today', 'Due in 1-3 Days', 'Due in 4-7 Days', 'Due Later'];
    default: return null;
  }
}

function getGroupIndicator(groupBy, key) {
  switch (groupBy) {
    case 'priority':
      if (key === 'High') return '<div class="priority-high-indicator"></div>';
      if (key === 'Medium') return '<div class="priority-medium-indicator"></div>';
      if (key === 'Low') return '<div class="priority-low-indicator"></div>';
      return '';
    case 'status':
      if (key === 'Not Started') return '<div class="status-not-started-indicator"></div>';
      if (key === 'In Progress') return '<div class="status-in-progress-indicator"></div>';
      if (key === 'Completed') return '<div class="status-completed-indicator"></div>';
      if (key === 'Blocked') return '<div class="status-blocked-indicator"></div>';
      return '';
    default: return '';
  }
}

function renderDueSoonTasks(tasks, groupBy) {
  const container = document.getElementById('dueSoonTaskList');

  if (tasks.length === 0) {
    container.innerHTML = `
      <div class="due-soon-empty">
        <svg fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path>
        </svg>
        <h3>No tasks due soon</h3>
        <p>All caught up! No urgent tasks require your attention.</p>
      </div>`;
    return;
  }

  const groups = groupTasks(tasks, groupBy);
  const groupOrder = getGroupOrder(groupBy);

  let sortedKeys;
  if (groupOrder) {
    sortedKeys = groupOrder.filter(k => groups[k]);
    Object.keys(groups).forEach(k => {
      if (!sortedKeys.includes(k)) sortedKeys.push(k);
    });
  } else {
    sortedKeys = Object.keys(groups).sort();
  }

  let html = '';

  sortedKeys.forEach(groupKey => {
    const groupTaskList = groups[groupKey];
    const indicator = getGroupIndicator(groupBy, groupKey);

    html += `
      <div class="due-soon-group">
        <div class="due-soon-group-header">
          ${indicator}
          <h3>${groupKey}</h3>
          <span class="due-soon-group-badge">${groupTaskList.length}</span>
        </div>
        <div class="due-soon-group-tasks">
    `;

    groupTaskList.forEach((task) => {
      const taskId = `duesoon-${task.project}-${task.name}-${task.instanceIndex}`.replace(/[^a-zA-Z0-9-]/g, '_');
      const daysUntil = getDaysUntilDeadline(task.deadlineDate);
      const deadlineClass = getDeadlineClass(daysUntil);
      const deadlineText = getDeadlineText(daysUntil, task.deadline);
      const statusBadge = getStatusBadge(task.status);
      const priorityIndicator = getPriorityIndicator(task.priority);
      const hasMetadata = task.metadata !== null;

      html += `
        <div class="due-soon-task" id="task-${taskId}">
          <div class="due-soon-task-header" onclick="toggleDueSoonTask('${taskId}')">
            <svg class="due-soon-toggle-icon" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"></path>
            </svg>
            <div class="due-soon-task-name">
              <span title="${task.name}">${task.name}</span>
              <div class="due-soon-task-project">${task.project}</div>
            </div>
            <div class="due-soon-task-meta">
              ${statusBadge}
              ${priorityIndicator}
              ${deadlineText ? `<span class="due-soon-deadline ${deadlineClass}">${deadlineText}</span>` : ''}
            </div>
          </div>
          <div class="due-soon-task-details">
            <div class="due-soon-task-details-content">
              <div id="view-${taskId}">
                <pre>${task.details.replace(/Deadline:\s*(\d{4}-\d{2}-\d{2})\s*\d{2}:\d{2}:\d{2}/g, (m, d) => { const p = d.split('-'); return 'Deadline: ' + p[2] + '/' + p[1] + '/' + p[0].slice(-2); })}</pre>
                <div class="mt-3 flex gap-2">
                  ${hasMetadata ? `
                    <button onclick="event.stopPropagation(); startDueSoonEdit('${taskId}', '${task.project}', '${task.name}', ${task.instanceIndex})"
                            class="bg-accent-500/20 hover:bg-accent-500/30 text-accent-400 px-3 py-1.5 rounded-md text-xs font-medium transition">
                      Edit Task
                    </button>
                  ` : ''}
                  <button onclick="event.stopPropagation(); goToTask('${task.project}', '${task.name}')"
                          class="bg-white/5 hover:bg-white/10 text-gray-400 px-3 py-1.5 rounded-md text-xs font-medium transition">
                    Go to Project
                  </button>
                </div>
              </div>
              <div id="edit-${taskId}" class="hidden">
                <div class="due-soon-edit-area">
                  <textarea id="textarea-${taskId}" class="due-soon-edit-textarea"
                            data-project="${task.project}"
                            data-task-name="${task.name}"
                            data-instance-index="${task.instanceIndex}">${task.details}</textarea>
                  <div class="due-soon-edit-buttons">
                    <button onclick="saveDueSoonEdit('${taskId}')"
                            id="save-btn-${taskId}"
                            class="bg-emerald-600 hover:bg-emerald-700 text-white px-4 py-1.5 rounded-lg text-sm font-medium transition">
                      Save Changes
                    </button>
                    <button onclick="cancelDueSoonEdit('${taskId}')"
                            class="bg-white/5 hover:bg-white/10 text-gray-300 px-4 py-1.5 rounded-lg text-sm font-medium transition">
                      Cancel
                    </button>
                  </div>
                  <p class="text-xs text-gray-600 mt-2">Format: <code class="bg-white/5 px-1 rounded text-gray-400">ColumnName: value</code></p>
                </div>
              </div>
            </div>
          </div>
        </div>
      `;
    });

    html += `
        </div>
      </div>
    `;
  });

  container.innerHTML = html;
}

function toggleDueSoonTask(taskId) {
  const task = document.getElementById(`task-${taskId}`);
  if (task) task.classList.toggle('expanded');
}

function goToTask(project, taskName) {
  closeDueSoonPopup();
  switchSheet(project);
  setTimeout(() => {
    showTaskDetails(taskName);
    const buttons = document.querySelectorAll('.task-button');
    buttons.forEach(btn => {
      if (btn.textContent === taskName) {
        btn.scrollIntoView({ behavior: 'smooth', inline: 'center' });
      }
    });
  }, 100);
}

function startDueSoonEdit(taskId, project, taskName, instanceIndex) {
  const viewDiv = document.getElementById(`view-${taskId}`);
  const editDiv = document.getElementById(`edit-${taskId}`);
  const textarea = document.getElementById(`textarea-${taskId}`);

  dueSoonOriginalDetails[taskId] = textarea.value;
  viewDiv.classList.add('hidden');
  editDiv.classList.remove('hidden');
  textarea.focus();
}

function cancelDueSoonEdit(taskId) {
  const viewDiv = document.getElementById(`view-${taskId}`);
  const editDiv = document.getElementById(`edit-${taskId}`);
  const textarea = document.getElementById(`textarea-${taskId}`);

  if (dueSoonOriginalDetails[taskId]) {
    textarea.value = dueSoonOriginalDetails[taskId];
  }

  editDiv.classList.add('hidden');
  viewDiv.classList.remove('hidden');
}

async function saveDueSoonEdit(taskId) {
  const textarea = document.getElementById(`textarea-${taskId}`);
  const saveBtn = document.getElementById(`save-btn-${taskId}`);
  const project = textarea.dataset.project;
  const taskName = textarea.dataset.taskName;
  const instanceIndex = parseInt(textarea.dataset.instanceIndex);

  const instances = allSheetsData[project]?.[taskName];
  if (!instances || !instances[instanceIndex]) {
    showNotification('Error: Task not found', 'error');
    return;
  }

  const instance = instances[instanceIndex];
  const metadata = typeof instance === 'object' && instance.metadata ? instance.metadata : null;

  if (!metadata) {
    showNotification('Error: Cannot save - missing metadata', 'error');
    return;
  }

  const newDetails = textarea.value;
  const updates = {};
  const newColumns = {};
  const existingColumns = metadata.columns;
  const newColumnValues = {};

  newDetails.split('\n').forEach(line => {
    const colonIndex = line.indexOf(':');
    if (colonIndex > 0) {
      const key = line.substring(0, colonIndex).trim();
      const value = line.substring(colonIndex + 1).trim();
      newColumnValues[key] = value;

      if (existingColumns.includes(key)) {
        updates[key] = value;
      } else {
        newColumns[key] = value;
      }
    }
  });

  existingColumns.forEach(col => {
    if (col !== existingColumns[0]) {
      const oldValue = metadata.raw_values[col];
      const hasNewValue = newColumnValues.hasOwnProperty(col);
      if (oldValue && !hasNewValue) {
        updates[col] = '';
      }
    }
  });

  saveBtn.disabled = true;
  const originalBtnText = saveBtn.innerHTML;
  saveBtn.innerHTML = '<span class="saving-spinner">&#x21bb;</span> Saving...';

  try {
    const response = await fetch('/api/save-task', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file_path: metadata.file_path,
        sheet_name: metadata.sheet_name,
        row_index: metadata.row_index,
        task_name: metadata.task_name,
        updates: updates,
        new_columns: Object.keys(newColumns).length > 0 ? newColumns : null
      })
    });

    const result = await response.json();

    if (response.ok) {
      showNotification('Changes saved successfully!', 'success');
      cancelDueSoonEdit(taskId);
      setTimeout(() => filterDueSoonTasks(), 500);
    } else {
      throw new Error(result.detail || 'Failed to save');
    }
  } catch (error) {
    showNotification(`Error: ${error.message}`, 'error');
  } finally {
    saveBtn.disabled = false;
    saveBtn.innerHTML = originalBtnText;
  }
}

function updateDueSoonBadgeOnLoad() {
  let tasks = getAllTasksAcrossProjects();
  tasks = tasks.filter(t => t.status !== 'Completed');
  tasks = tasks.filter(t => getDaysUntilDeadline(t.deadlineDate) <= 7);
  updateDueSoonBadge(tasks.length);
}

// Make popup functions globally available
window.openDueSoonPopup = openDueSoonPopup;
window.closeDueSoonPopup = closeDueSoonPopup;
window.filterDueSoonTasks = filterDueSoonTasks;
window.toggleDueSoonTask = toggleDueSoonTask;
window.goToTask = goToTask;
window.startDueSoonEdit = startDueSoonEdit;
window.cancelDueSoonEdit = cancelDueSoonEdit;
window.saveDueSoonEdit = saveDueSoonEdit;

// ===== EXCEL FILE OPEN/CLOSE =====
function getExcelFilesForSheet(sheetName) {
  const tasks = allSheetsData[sheetName];
  if (!tasks) return [];

  const filePaths = new Set();
  for (const taskName in tasks) {
    for (const instance of tasks[taskName]) {
      if (instance.metadata && instance.metadata.file_path) {
        filePaths.add(instance.metadata.file_path);
      }
    }
  }
  return [...filePaths];
}

function getColumnsForSheetFile(sheetName, filePath) {
  const tasks = allSheetsData[sheetName];
  if (!tasks) return [];

  const columns = [];
  const seen = new Set();

  for (const taskName in tasks) {
    for (const instance of tasks[taskName]) {
      if (!instance.metadata || !Array.isArray(instance.metadata.columns)) {
        continue;
      }
      if (filePath && instance.metadata.file_path !== filePath) {
        continue;
      }
      for (const col of instance.metadata.columns) {
        if (!seen.has(col)) {
          seen.add(col);
          columns.push(col);
        }
      }
    }
  }

  return columns;
}

function updateAddTaskButton() {
  const btn = document.getElementById('addTaskBtn');
  if (!btn) return;

  const files = getExcelFilesForSheet(currentSheet);
  if (files.length === 0) {
    btn.classList.add('hidden');
  } else {
    btn.classList.remove('hidden');
  }
}

function updateAddTaskColumnHint() {
  const fileSelect = document.getElementById('addTaskFileSelect');
  const hint = document.getElementById('addTaskColumnHint');
  if (!fileSelect || !hint) return;

  const selectedFile = fileSelect.value;
  addTaskKnownColumns = getColumnsForSheetFile(currentSheet, selectedFile);
  const editableColumns = addTaskKnownColumns.slice(1);

  if (editableColumns.length === 0) {
    hint.innerHTML = 'Format each line as <code class="bg-white/5 px-1.5 py-0.5 rounded text-gray-300">ColumnName: value</code>. Unknown columns will be added as new columns.';
    return;
  }

  const preview = editableColumns.slice(0, 8).join(', ');
  const suffix = editableColumns.length > 8 ? ', ...' : '';
  hint.innerHTML = `Known columns: <span class="text-gray-400">${preview}${suffix}</span>. New column names are allowed.`;
}

function populateAddTaskFieldsTemplate() {
  const detailsInput = document.getElementById('addTaskDetailsInput');
  if (!detailsInput) return;

  const editableColumns = addTaskKnownColumns.slice(1);
  if (editableColumns.length === 0) {
    detailsInput.value = '';
    return;
  }

  detailsInput.value = editableColumns.map(col => `${col}: `).join('\n');
}

function onAddTaskFileChanged() {
  updateAddTaskColumnHint();
  populateAddTaskFieldsTemplate();
}

function openAddTaskModal() {
  const modal = document.getElementById('addTaskModal');
  const fileSelect = document.getElementById('addTaskFileSelect');
  const taskNameInput = document.getElementById('addTaskNameInput');
  const detailsInput = document.getElementById('addTaskDetailsInput');
  const sheetName = document.getElementById('addTaskSheetName');

  const files = getExcelFilesForSheet(currentSheet);
  if (files.length === 0) {
    showNotification('No source Excel file found for this sheet', 'error');
    return;
  }

  fileSelect.innerHTML = '';
  files.forEach(filePath => {
    const option = document.createElement('option');
    option.value = filePath;
    option.textContent = filePath.split(/[\\\/]/).pop();
    fileSelect.appendChild(option);
  });
  fileSelect.disabled = files.length <= 1;

  sheetName.textContent = currentSheet;
  taskNameInput.value = '';
  onAddTaskFileChanged();

  modal.classList.remove('hidden');
  updateBodyScrollLock();
  taskNameInput.focus();
}

function closeAddTaskModal() {
  const modal = document.getElementById('addTaskModal');
  if (!modal) return;
  modal.classList.add('hidden');
  updateBodyScrollLock();
}

async function submitAddTask() {
  const fileSelect = document.getElementById('addTaskFileSelect');
  const taskNameInput = document.getElementById('addTaskNameInput');
  const detailsInput = document.getElementById('addTaskDetailsInput');
  const saveBtn = document.getElementById('addTaskSaveBtn');

  const filePath = fileSelect.value;
  const taskName = taskNameInput.value.trim();
  const detailLines = detailsInput.value.split('\n');

  if (!filePath) {
    showNotification('Select a source file first', 'error');
    return;
  }
  if (!taskName) {
    showNotification('Task name is required', 'error');
    taskNameInput.focus();
    return;
  }

  const values = {};
  const newColumns = {};
  const knownColumns = new Set(addTaskKnownColumns.slice(1));

  for (const rawLine of detailLines) {
    const line = rawLine.trim();
    if (!line) continue;

    const colonIndex = line.indexOf(':');
    if (colonIndex <= 0) {
      showNotification(`Invalid line: "${line}". Use ColumnName: value`, 'error');
      return;
    }

    const key = line.substring(0, colonIndex).trim();
    const value = line.substring(colonIndex + 1).trim();

    if (knownColumns.has(key)) {
      values[key] = value;
    } else {
      newColumns[key] = value;
    }
  }

  saveBtn.disabled = true;
  const originalText = saveBtn.textContent;
  saveBtn.textContent = 'Adding...';

  try {
    const response = await fetch('/api/add-task', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        file_path: filePath,
        sheet_name: currentSheet,
        task_name: taskName,
        values,
        new_columns: Object.keys(newColumns).length > 0 ? newColumns : null
      })
    });

    const result = await response.json();
    if (!response.ok) {
      throw new Error(result.detail || 'Failed to add task');
    }

    closeAddTaskModal();
    showNotification('Task added successfully!', 'success');
    await fetchLatestData(false);
  } catch (error) {
    showNotification(`Error: ${error.message}`, 'error');
  } finally {
    saveBtn.disabled = false;
    saveBtn.textContent = originalText;
  }
}

function updateExcelButton() {
  const btn = document.getElementById('excelOpenBtn');
  const btnFile = document.getElementById('excelBtnFile');
  const files = getExcelFilesForSheet(currentSheet);

  if (!btn) return;

  if (files.length === 0) {
    btn.classList.add('hidden');
    return;
  }

  btn.classList.remove('hidden');

  const fileNames = files.map(f => f.split(/[\\\/]/).pop());
  if (fileNames.length > 0) {
    btnFile.textContent = fileNames.length === 1 ? fileNames[0] : `${fileNames.length} files`;
    btnFile.classList.remove('hidden');
  }
}

function openExcelFile() {
  const files = getExcelFilesForSheet(currentSheet);
  if (files.length === 0) return;

  // If only one file, open it directly
  if (files.length === 1) {
    openSingleExcelFile(files[0]);
    return;
  }

  // Multiple files: show selection popup
  toggleExcelFilePopup(files);
}

function toggleExcelFilePopup(files) {
  const popup = document.getElementById('excelFilePopup');
  const overlay = document.getElementById('excelPopupOverlay');

  if (!popup.classList.contains('hidden')) {
    closeExcelFilePopup();
    return;
  }

  const list = document.getElementById('excelFileList');
  list.innerHTML = '';

  for (const filePath of files) {
    const fileName = filePath.split(/[\\\/]/).pop();
    const item = document.createElement('div');
    item.className = 'excel-file-popup-item';
    item.innerHTML = `
      <svg fill="currentColor" viewBox="0 0 24 24">
        <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8l-6-6zM6 20V4h7v5h5v11H6z"/>
        <path d="M8 12.5L10.5 17H9l-1.5-3-1.5 3H4.5L7 12.5 4.5 8H6l1.5 3L9 8h1.5L8 12.5z"/>
      </svg>
      <span title="${fileName}">${fileName}</span>
    `;
    item.addEventListener('click', () => {
      closeExcelFilePopup();
      openSingleExcelFile(filePath);
    });
    list.appendChild(item);
  }

  popup.classList.remove('hidden');
  overlay.classList.remove('hidden');
}

function closeExcelFilePopup() {
  const popup = document.getElementById('excelFilePopup');
  const overlay = document.getElementById('excelPopupOverlay');
  popup.classList.add('hidden');
  overlay.classList.add('hidden');
}

async function openSingleExcelFile(filePath) {
  const btn = document.getElementById('excelOpenBtn');
  const btnText = document.getElementById('excelBtnText');

  btn.disabled = true;
  btnText.textContent = 'Opening...';

  try {
    const response = await fetch('/api/open-excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ file_path: filePath })
    });

    const result = await response.json();
    if (!response.ok) {
      throw new Error(result.detail || 'Failed to open');
    }

    const fileName = filePath.split(/[\\\/]/).pop();
    showNotification(`Opened: ${fileName}`, 'success');
  } catch (error) {
    showNotification(`Error: ${error.message}`, 'error');
  } finally {
    btn.disabled = false;
    btnText.textContent = 'Open Excel';
  }
}

window.openExcelFile = openExcelFile;
window.closeExcelFilePopup = closeExcelFilePopup;
window.openAddTaskModal = openAddTaskModal;
window.closeAddTaskModal = closeAddTaskModal;
window.submitAddTask = submitAddTask;

// ===== INITIALIZATION =====
function init() {
  // Scroll buttons
  scrollLeftBtn.addEventListener('click', () => {
    scrollContainer.scrollBy({ left: -300, behavior: 'smooth' });
  });

  scrollRightBtn.addEventListener('click', () => {
    scrollContainer.scrollBy({ left: 300, behavior: 'smooth' });
  });

  scrollContainer.addEventListener('scroll', updateScrollIndicators);
  window.addEventListener('resize', updateScrollIndicators);

  // Filter listeners
  document.getElementById('searchBox').addEventListener('input', applyFilters);
  document.getElementById('statusFilter').addEventListener('change', applyFilters);
  document.getElementById('priorityFilter').addEventListener('change', applyFilters);
  document.getElementById('hideCompleted').addEventListener('change', applyFilters);

  // Clear filters
  document.getElementById('clearFilters').addEventListener('click', () => {
    document.getElementById('searchBox').value = '';
    document.getElementById('statusFilter').value = 'All';
    document.getElementById('priorityFilter').value = 'All';
    document.getElementById('hideCompleted').checked = false;
    applyFilters();
  });

  // Initialize view
  switchSheet(currentSheet);

  // Due soon badge
  updateDueSoonBadgeOnLoad();

  // Excel button
  updateExcelButton();
  updateAddTaskButton();

  const addTaskFileSelect = document.getElementById('addTaskFileSelect');
  if (addTaskFileSelect) {
    addTaskFileSelect.addEventListener('change', onAddTaskFileChanged);
  }

  // Close excel file popup when clicking overlay
  const excelOverlay = document.getElementById('excelPopupOverlay');
  if (excelOverlay) {
    excelOverlay.addEventListener('click', closeExcelFilePopup);
  }

  // SSE
  connectSSE();

  // Cleanup
  window.addEventListener('beforeunload', () => {
    if (eventSource) eventSource.close();
  });
}

// Global function bindings
window.switchSheet = switchSheet;
window.showTaskDetails = showTaskDetails;
window.toggleTaskInstance = toggleTaskInstance;
window.startEdit = startEdit;
window.cancelEdit = cancelEdit;
window.saveEdit = saveEdit;

document.addEventListener('DOMContentLoaded', init);
