// ===== GLOBAL STATE =====
// These are initialized from the HTML template via window.AppConfig
let allSheetsData = window.AppConfig?.allSheetsData || {};
let currentDataVersion = window.AppConfig?.dataVersion || 0;
let currentSheet = window.AppConfig?.initialSheet || '';
let buttonData = allSheetsData[currentSheet];
let selectedTask = null;
let allTasks = [];
let filterBarLocked = false;
let collapseTimeout = null;
let originalDetails = {}; // Store original details for cancel functionality

const buttonColors = [
  'bg-blue-500 hover:bg-blue-700',
  'bg-green-500 hover:bg-green-700',
  'bg-red-500 hover:bg-red-700',
  'bg-yellow-500 hover:bg-yellow-700',
  'bg-purple-500 hover:bg-purple-700',
  'bg-pink-500 hover:bg-pink-700',
  'bg-indigo-500 hover:bg-indigo-700',
  'bg-orange-500 hover:bg-orange-700',
  'bg-teal-500 hover:bg-teal-700',
  'bg-cyan-500 hover:bg-cyan-700'
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

// ===== FILTER BAR FUNCTIONALITY =====
function expandFilterBar() {
  clearTimeout(collapseTimeout);
  filterBar.classList.remove('filter-bar-collapsed');
  filterBar.classList.add('filter-bar-expanded');
  filterContent.classList.remove('filter-content-hidden');
  filterContent.classList.add('filter-content-visible');
  expandIcon.style.transform = 'rotate(180deg)';
}

function collapseFilterBar() {
  if (filterBarLocked) return;

  collapseTimeout = setTimeout(() => {
    if (!filterBarLocked) {
      filterBar.classList.remove('filter-bar-expanded');
      filterBar.classList.add('filter-bar-collapsed');
      filterContent.classList.remove('filter-content-visible');
      filterContent.classList.add('filter-content-hidden');
      expandIcon.style.transform = 'rotate(0deg)';
    }
  }, 500);
}

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
    filterBar.classList.add('pulse-active');
    document.getElementById('clearFilters').classList.remove('hidden');
  } else {
    activeFilterBadge.classList.add('hidden');
    filterBar.classList.remove('pulse-active');
    document.getElementById('clearFilters').classList.add('hidden');
  }
}

// ===== TASK PARSING UTILITIES =====
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
  const statusColors = {
    'Not Started': 'bg-gray-600 text-gray-200',
    'In Progress': 'bg-yellow-600 text-yellow-100',
    'Completed': 'bg-green-600 text-green-100',
    'Blocked': 'bg-red-600 text-red-100'
  };
  const color = statusColors[status] || 'bg-gray-600 text-gray-200';
  return `<span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium ${color}">${status}</span>`;
}

function getPriorityIndicator(priority) {
  const priorityIcons = {
    'High': '<span class="text-red-400 font-bold">⚠️ High</span>',
    'Medium': '<span class="text-yellow-400 font-bold">➡️ Medium</span>',
    'Low': '<span class="text-blue-400 font-bold">⬇️ Low</span>'
  };
  return priorityIcons[priority] || '';
}

// ===== FILTERING =====
function getFilteredTasks() {
  const searchTerm = document.getElementById('searchBox').value.toLowerCase();
  const statusFilter = document.getElementById('statusFilter').value;
  const priorityFilter = document.getElementById('priorityFilter').value;
  const hideCompleted = document.getElementById('hideCompleted').checked;

  return allTasks.filter(task => {
    if (searchTerm && !task.name.toLowerCase().includes(searchTerm)) {
      return false;
    }
    if (hideCompleted && task.status === 'Completed') {
      return false;
    }
    if (statusFilter !== 'All' && task.status !== statusFilter) {
      return false;
    }
    if (priorityFilter !== 'All' && task.priority !== priorityFilter) {
      return false;
    }
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
        description: description,
        status: status,
        priority: priority,
        assignedTo: assignedTo,
        deadline: deadline,
        details: details,
        metadata: metadata
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
    button.className = `task-button ${colorClass} text-white font-bold py-4 px-6 rounded-lg transition duration-300 shadow-lg whitespace-nowrap`;

    container.appendChild(button);
  });

  if (selectedTask && !uniqueTaskNames.includes(selectedTask)) {
    selectedTask = null;
    document.getElementById('textDisplay').innerHTML = '<p class="text-gray-500 italic">Click a task to view details</p>';
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

// ===== TASK DETAILS DISPLAY =====
function showTaskDetails(taskName) {
  selectedTask = taskName;
  console.log(`Task selected: ${taskName} in project: ${currentSheet}`);

  const display = document.getElementById('textDisplay');
  const filteredTasks = getFilteredTasks();
  const instances = filteredTasks.filter(t => t.name === taskName);

  if (instances.length === 0) {
    return;
  }

  let html = '';

  const totalInstances = allTasks.filter(t => t.name === taskName).length;
  if (totalInstances > 1) {
    html += `<div class="mb-4 text-sm text-gray-400">
      <span class="font-semibold">${taskName}</span> - Showing ${instances.length} ${instances.length !== totalInstances ? `of ${totalInstances}` : ''} instance${instances.length !== 1 ? 's' : ''}
    </div>`;
  }

  instances.forEach((task, index) => {
    const shortDesc = truncateDescription(task.description, 4);
    const statusBadge = getStatusBadge(task.status);
    const priorityIndicator = getPriorityIndicator(task.priority);
    const deadlineText = task.deadline ? `Due: ${task.deadline}` : '';
    const hasMetadata = task.metadata !== null;
    const taskIndex = allTasks.indexOf(task);

    const bgColor = index % 2 === 0 ? 'bg-gray-800' : 'bg-gray-750';
    const borderColor = index % 2 === 0 ? 'border-l-blue-500' : 'border-l-purple-500';

    const collapsedClass = 'task-instance-expanded';
    const detailsClass = 'task-details-visible';
    const iconClass = 'toggle-icon-expanded';

    const sourceFile = hasMetadata ? task.metadata.file_path.split(/[\\\/]/).pop() : '';

    html += `
      <div id="instance-${taskName}-${index}" class="task-instance ${collapsedClass} ${bgColor} border-l-4 ${borderColor} rounded-lg mb-3 overflow-hidden">
        <!-- Instance Header (Always Visible) -->
        <div class="p-4 cursor-pointer hover:bg-gray-700 transition flex items-center justify-between" onclick="toggleTaskInstance('${taskName}-${index}')">
          <div class="flex items-center gap-3 flex-1">
            <svg id="icon-${taskName}-${index}" class="toggle-icon ${iconClass} w-5 h-5 text-gray-400 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"></path>
            </svg>
            <div class="flex-1 flex flex-wrap items-center gap-2">
              <span class="font-semibold text-white">${taskName}</span>
              ${shortDesc ? `<span class="text-gray-400">- ${shortDesc}</span>` : ''}
              ${statusBadge}
              ${priorityIndicator ? `<span class="text-sm">${priorityIndicator}</span>` : ''}
              ${deadlineText ? `<span class="text-sm text-orange-400 font-medium">${deadlineText}</span>` : ''}
            </div>
          </div>
          ${hasMetadata ? `
            <button onclick="event.stopPropagation(); startEdit('${taskName}-${index}', ${taskIndex})"
                    class="edit-btn bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-1 rounded text-sm font-medium ml-2">
              Edit
            </button>
          ` : ''}
        </div>

        <!-- Instance Details (Collapsible) -->
        <div id="details-${taskName}-${index}" class="task-details ${detailsClass} px-4 pb-4 pt-0">
          <!-- View Mode -->
          <div id="view-${taskName}-${index}" class="pl-8 border-l-2 border-gray-600 ml-2 space-y-2">
            <div class="whitespace-pre-line text-gray-300 leading-relaxed">${task.details.replace(/Deadline:\s*(\d{4}-\d{2}-\d{2})\s*\d{2}:\d{2}:\d{2}/g, (m, d) => { const p = d.split('-'); return 'Deadline: ' + p[2] + '/' + p[1] + '/' + p[0].slice(-2); })}</div>
          </div>

          <!-- Edit Mode (hidden by default) -->
          <div id="edit-${taskName}-${index}" class="hidden pl-8 border-l-2 border-indigo-500 ml-2 space-y-4">
            <textarea id="textarea-${taskName}-${index}"
                      class="edit-textarea"
                      data-task-index="${taskIndex}">${task.details}</textarea>

            <div class="flex gap-3">
              <button onclick="saveEdit('${taskName}-${index}')"
                      id="save-btn-${taskName}-${index}"
                      class="bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded font-medium">
                Save Changes
              </button>
              <button onclick="cancelEdit('${taskName}-${index}')"
                      class="bg-gray-600 hover:bg-gray-700 text-white px-4 py-2 rounded font-medium">
                Cancel
              </button>
            </div>

            <div class="text-xs text-gray-500">
              <p>Format: <code class="bg-gray-700 px-1 rounded">ColumnName: value</code></p>
              <p>Add new columns by adding new lines in this format.</p>
              ${hasMetadata ? `<p class="mt-1">Source: ${sourceFile}</p>` : ''}
            </div>
          </div>
        </div>
      </div>
    `;
  });

  display.innerHTML = html;

  document.querySelectorAll('.task-button').forEach(btn => {
    btn.classList.remove('ring-4', 'ring-white');
    if (btn.textContent === taskName) {
      btn.classList.add('ring-4', 'ring-white');
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
    showEditNotification('Error: Cannot save - missing metadata', 'error');
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
  saveBtn.innerHTML = '<span class="saving-spinner">↻</span> Saving...';

  try {
    const response = await fetch('/api/save-task', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
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
      showEditNotification('Changes saved successfully!', 'success');
      cancelEdit(instanceId);
    } else {
      throw new Error(result.detail || 'Failed to save');
    }

  } catch (error) {
    showEditNotification(`Error: ${error.message}`, 'error');
  } finally {
    saveBtn.disabled = false;
    saveBtn.innerHTML = originalBtnText;
  }
}

function showEditNotification(message, type = 'success') {
  let notification = document.getElementById('editNotification');
  if (!notification) {
    notification = document.createElement('div');
    notification.id = 'editNotification';
    notification.className = 'fixed bottom-4 right-4 px-4 py-2 rounded-lg shadow-lg transform translate-y-20 opacity-0 transition-all duration-300 z-50 text-white';
    document.body.appendChild(notification);
  }

  notification.classList.remove('bg-green-600', 'bg-red-600');
  notification.classList.add(type === 'success' ? 'bg-green-600' : 'bg-red-600');
  notification.innerHTML = message;

  notification.classList.remove('translate-y-20', 'opacity-0');

  setTimeout(() => {
    notification.classList.add('translate-y-20', 'opacity-0');
  }, 3000);
}

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

  document.getElementById('textDisplay').innerHTML = '<p class="text-gray-500 italic">Click a task to view details</p>';

  document.querySelectorAll('.sheet-btn').forEach(btn => {
    if (btn.dataset.sheet === sheetName) {
      btn.classList.remove('bg-gray-700');
      btn.classList.add('bg-purple-600');
    } else {
      btn.classList.remove('bg-purple-600');
      btn.classList.add('bg-gray-700');
    }
  });

  console.log(`Switched to project: ${sheetName}`);
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
  if (eventSource) {
    eventSource.close();
  }

  eventSource = new EventSource('/events');

  eventSource.onopen = function() {
    console.log('SSE connected - listening for data updates');
    reconnectAttempts = 0;
  };

  eventSource.onmessage = function(event) {
    const newVersion = parseInt(event.data);
    if (newVersion > currentDataVersion) {
      console.log(`Data update detected (v${currentDataVersion} -> v${newVersion}), refreshing...`);
      fetchLatestData();
    }
  };

  eventSource.onerror = function(err) {
    console.log('SSE connection error, will attempt reconnect...');
    eventSource.close();

    if (reconnectAttempts < maxReconnectAttempts) {
      reconnectAttempts++;
      setTimeout(connectSSE, 2000 * reconnectAttempts);
    }
  };
}

async function fetchLatestData() {
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
          document.getElementById('textDisplay').innerHTML = '<p class="text-gray-500 italic">Click a task to view details</p>';
        }
      } else {
        if (newSheetNames.length > 0) {
          switchSheet(newSheetNames[0]);
        }
      }

      showUpdateNotification();
      console.log(`Data refreshed to version ${currentDataVersion}`);

      // Update due soon badge
      updateDueSoonBadgeOnLoad();
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
      button.className = 'sheet-btn text-left bg-gray-700 hover:bg-gray-600 text-white font-semibold py-3 px-4 rounded-lg transition duration-300 shadow-md';
      button.dataset.sheet = sheetName;
      button.textContent = sheetName;

      if (sheetName === currentSheet) {
        button.classList.remove('bg-gray-700');
        button.classList.add('bg-purple-600');
      }

      container.appendChild(button);
    });
  }
}

function showUpdateNotification() {
  let notification = document.getElementById('updateNotification');
  if (!notification) {
    notification = document.createElement('div');
    notification.id = 'updateNotification';
    notification.className = 'fixed bottom-4 right-4 bg-green-600 text-white px-4 py-2 rounded-lg shadow-lg transform translate-y-20 opacity-0 transition-all duration-300 z-50';
    notification.innerHTML = '<span class="mr-2">&#10003;</span> Data updated';
    document.body.appendChild(notification);
  }

  notification.classList.remove('translate-y-20', 'opacity-0');

  setTimeout(() => {
    notification.classList.add('translate-y-20', 'opacity-0');
  }, 2000);
}

// ===== DUE SOON POPUP FUNCTIONALITY =====
let dueSoonOriginalDetails = {};

function openDueSoonPopup() {
  const modal = document.getElementById('dueSoonModal');
  modal.classList.remove('hidden');
  document.body.style.overflow = 'hidden';
  filterDueSoonTasks();
}

function closeDueSoonPopup() {
  const modal = document.getElementById('dueSoonModal');
  modal.classList.add('hidden');
  document.body.style.overflow = '';
}

// Close modal on Escape key
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    closeDueSoonPopup();
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
          name: taskName,
          project: sheetName,
          instanceIndex: index,
          description: description,
          status: status,
          priority: priority,
          assignedTo: assignedTo,
          deadline: deadline,
          deadlineDate: parseDeadlineToDate(deadline),
          details: details,
          metadata: metadata
        });
      });
    });
  });

  return allProjectTasks;
}

function parseDeadlineToDate(deadline) {
  if (!deadline) return null;

  // Try parsing DD/MM/YY format
  const ddmmyy = deadline.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (ddmmyy) {
    let year = parseInt(ddmmyy[3]);
    if (year < 100) year += 2000;
    return new Date(year, parseInt(ddmmyy[2]) - 1, parseInt(ddmmyy[1]));
  }

  // Try parsing YYYY-MM-DD format
  const yyyymmdd = deadline.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (yyyymmdd) {
    return new Date(parseInt(yyyymmdd[1]), parseInt(yyyymmdd[2]) - 1, parseInt(yyyymmdd[3]));
  }

  // Try general Date parsing
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

  // Filter by completion status
  if (hideCompleted) {
    tasks = tasks.filter(t => t.status !== 'Completed');
  }

  // Filter by days
  if (daysFilter !== 'all') {
    const days = parseInt(daysFilter);
    tasks = tasks.filter(t => {
      const daysUntil = getDaysUntilDeadline(t.deadlineDate);
      return daysUntil <= days;
    });
  }

  // Sort by deadline (soonest first), then by priority
  const priorityOrder = { 'High': 0, 'Medium': 1, 'Low': 2, '': 3 };
  tasks.sort((a, b) => {
    const daysA = getDaysUntilDeadline(a.deadlineDate);
    const daysB = getDaysUntilDeadline(b.deadlineDate);
    if (daysA !== daysB) return daysA - daysB;
    return (priorityOrder[a.priority] || 3) - (priorityOrder[b.priority] || 3);
  });

  // Update counter
  document.getElementById('dueSoonCounter').textContent = `${tasks.length} task${tasks.length !== 1 ? 's' : ''}`;

  // Update badge on sidebar button
  updateDueSoonBadge(tasks.length);

  // Group and render tasks
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
      case 'priority':
        key = task.priority || 'No Priority';
        break;
      case 'status':
        key = task.status || 'No Status';
        break;
      case 'project':
        key = task.project;
        break;
      case 'deadline':
        const daysUntil = getDaysUntilDeadline(task.deadlineDate);
        if (daysUntil < 0) key = 'Overdue';
        else if (daysUntil === 0) key = 'Due Today';
        else if (daysUntil <= 3) key = 'Due in 1-3 Days';
        else if (daysUntil <= 7) key = 'Due in 4-7 Days';
        else key = 'Due Later';
        break;
      default:
        key = 'All Tasks';
    }

    if (!groups[key]) groups[key] = [];
    groups[key].push(task);
  });

  return groups;
}

function getGroupOrder(groupBy) {
  switch (groupBy) {
    case 'priority':
      return ['High', 'Medium', 'Low', 'No Priority'];
    case 'status':
      return ['Blocked', 'In Progress', 'Not Started', 'Completed', 'No Status'];
    case 'deadline':
      return ['Overdue', 'Due Today', 'Due in 1-3 Days', 'Due in 4-7 Days', 'Due Later'];
    default:
      return null;
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
    default:
      return '';
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
      </div>
    `;
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
    const groupTasks = groups[groupKey];
    const indicator = getGroupIndicator(groupBy, groupKey);

    html += `
      <div class="due-soon-group">
        <div class="due-soon-group-header">
          ${indicator}
          <h3>${groupKey}</h3>
          <span class="due-soon-group-badge">${groupTasks.length}</span>
        </div>
        <div class="due-soon-group-tasks">
    `;

    groupTasks.forEach((task, idx) => {
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
              <div class="due-soon-task-project">Project: ${task.project}</div>
            </div>
            <div class="due-soon-task-meta">
              ${statusBadge}
              ${priorityIndicator}
              ${deadlineText ? `<span class="due-soon-deadline ${deadlineClass}">${deadlineText}</span>` : ''}
            </div>
          </div>
          <div class="due-soon-task-details">
            <div class="due-soon-task-details-content">
              <!-- View Mode -->
              <div id="view-${taskId}">
                <pre>${task.details.replace(/Deadline:\s*(\d{4}-\d{2}-\d{2})\s*\d{2}:\d{2}:\d{2}/g, (m, d) => { const p = d.split('-'); return 'Deadline: ' + p[2] + '/' + p[1] + '/' + p[0].slice(-2); })}</pre>
                ${hasMetadata ? `
                  <div class="mt-3 flex gap-2">
                    <button onclick="event.stopPropagation(); startDueSoonEdit('${taskId}', '${task.project}', '${task.name}', ${task.instanceIndex})"
                            class="bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-1.5 rounded text-sm font-medium transition">
                      Edit Task
                    </button>
                    <button onclick="event.stopPropagation(); goToTask('${task.project}', '${task.name}')"
                            class="bg-gray-600 hover:bg-gray-700 text-white px-3 py-1.5 rounded text-sm font-medium transition">
                      Go to Project
                    </button>
                  </div>
                ` : `
                  <div class="mt-3">
                    <button onclick="event.stopPropagation(); goToTask('${task.project}', '${task.name}')"
                            class="bg-gray-600 hover:bg-gray-700 text-white px-3 py-1.5 rounded text-sm font-medium transition">
                      Go to Project
                    </button>
                  </div>
                `}
              </div>
              <!-- Edit Mode -->
              <div id="edit-${taskId}" class="hidden">
                <div class="due-soon-edit-area">
                  <textarea id="textarea-${taskId}" class="due-soon-edit-textarea"
                            data-project="${task.project}"
                            data-task-name="${task.name}"
                            data-instance-index="${task.instanceIndex}">${task.details}</textarea>
                  <div class="due-soon-edit-buttons">
                    <button onclick="saveDueSoonEdit('${taskId}')"
                            id="save-btn-${taskId}"
                            class="bg-green-600 hover:bg-green-700 text-white px-4 py-2 rounded font-medium text-sm">
                      Save Changes
                    </button>
                    <button onclick="cancelDueSoonEdit('${taskId}')"
                            class="bg-gray-600 hover:bg-gray-700 text-white px-4 py-2 rounded font-medium text-sm">
                      Cancel
                    </button>
                  </div>
                  <p class="text-xs text-gray-500 mt-2">Format: <code class="bg-gray-700 px-1 rounded">ColumnName: value</code></p>
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
  if (task) {
    task.classList.toggle('expanded');
  }
}

function goToTask(project, taskName) {
  closeDueSoonPopup();
  switchSheet(project);
  setTimeout(() => {
    showTaskDetails(taskName);
    // Scroll the task button into view
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

  // Find the task in allSheetsData
  const instances = allSheetsData[project]?.[taskName];
  if (!instances || !instances[instanceIndex]) {
    showEditNotification('Error: Task not found', 'error');
    return;
  }

  const instance = instances[instanceIndex];
  const metadata = typeof instance === 'object' && instance.metadata ? instance.metadata : null;

  if (!metadata) {
    showEditNotification('Error: Cannot save - missing metadata', 'error');
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
  saveBtn.innerHTML = '<span class="saving-spinner">↻</span> Saving...';

  try {
    const response = await fetch('/api/save-task', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
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
      showEditNotification('Changes saved successfully!', 'success');
      cancelDueSoonEdit(taskId);
      // Refresh the popup after a brief delay
      setTimeout(() => filterDueSoonTasks(), 500);
    } else {
      throw new Error(result.detail || 'Failed to save');
    }

  } catch (error) {
    showEditNotification(`Error: ${error.message}`, 'error');
  } finally {
    saveBtn.disabled = false;
    saveBtn.innerHTML = originalBtnText;
  }
}

function updateDueSoonBadgeOnLoad() {
  // Count tasks due within 7 days (default filter) excluding completed
  let tasks = getAllTasksAcrossProjects();
  tasks = tasks.filter(t => t.status !== 'Completed');
  tasks = tasks.filter(t => {
    const daysUntil = getDaysUntilDeadline(t.deadlineDate);
    return daysUntil <= 7;
  });
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

// ===== INITIALIZATION =====
function init() {
  // Filter bar event listeners
  filterBar.addEventListener('mouseenter', expandFilterBar);
  filterBar.addEventListener('mouseleave', collapseFilterBar);

  // Lock filter bar when interacting with inputs
  const filterInputs = [
    document.getElementById('searchBox'),
    document.getElementById('statusFilter'),
    document.getElementById('priorityFilter'),
    document.getElementById('hideCompleted')
  ];

  filterInputs.forEach(input => {
    input.addEventListener('focus', () => {
      filterBarLocked = true;
      expandFilterBar();
    });

    input.addEventListener('blur', () => {
      setTimeout(() => {
        const anyFocused = filterInputs.some(i => i === document.activeElement);
        if (!anyFocused) {
          filterBarLocked = false;
          collapseFilterBar();
        }
      }, 100);
    });
  });

  // Clear filters button
  document.getElementById('clearFilters').addEventListener('click', () => {
    document.getElementById('searchBox').value = '';
    document.getElementById('statusFilter').value = 'All';
    document.getElementById('priorityFilter').value = 'All';
    document.getElementById('hideCompleted').checked = false;
    applyFilters();
  });

  // Scroll buttons
  scrollLeftBtn.addEventListener('click', () => {
    scrollContainer.scrollBy({ left: -300, behavior: 'smooth' });
  });

  scrollRightBtn.addEventListener('click', () => {
    scrollContainer.scrollBy({ left: 300, behavior: 'smooth' });
  });

  scrollContainer.addEventListener('scroll', updateScrollIndicators);
  window.addEventListener('resize', updateScrollIndicators);

  // Filter event listeners
  document.getElementById('searchBox').addEventListener('input', applyFilters);
  document.getElementById('statusFilter').addEventListener('change', applyFilters);
  document.getElementById('priorityFilter').addEventListener('change', applyFilters);
  document.getElementById('hideCompleted').addEventListener('change', applyFilters);

  // Initialize view
  switchSheet(currentSheet);

  // Initialize due soon badge count
  updateDueSoonBadgeOnLoad();

  // Start SSE connection
  connectSSE();

  // Cleanup on page unload
  window.addEventListener('beforeunload', () => {
    if (eventSource) {
      eventSource.close();
    }
  });
}

// Make functions globally available for onclick handlers
window.switchSheet = switchSheet;
window.showTaskDetails = showTaskDetails;
window.toggleTaskInstance = toggleTaskInstance;
window.startEdit = startEdit;
window.cancelEdit = cancelEdit;
window.saveEdit = saveEdit;

// Run init when DOM is ready
document.addEventListener('DOMContentLoaded', init);
