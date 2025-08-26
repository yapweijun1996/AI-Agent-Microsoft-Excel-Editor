import { AppState } from '../core/state.js';
import { db } from '../db/indexeddb.js';
import { escapeHtml, extractFirstJson, getSampleDataFromSheet, log } from '../utils/index.js';
import { Modal } from '../ui/modal.js';
import { showToast } from '../ui/toast.js';
import { getWorksheet } from '../spreadsheet/workbook-manager.js';
import { runExecutor, runValidator, fetchOpenAI, fetchGemini } from '../services/ai-agents.js';
import { pickProvider, getSelectedModel } from '../services/api-keys.js';
/* global applyEditsOrDryRun */

export async function loadTasks() {
  try {
    AppState.tasks = await db.getTasksByWorkbook('current') || [];
  } catch {
    AppState.tasks = [];
  }
}

export async function saveTasks() {
  for (const task of AppState.tasks) {
    await db.saveTask({ ...task, workbookId: 'current' });
  }
}

function renderTask(task) {
  const statusColors = {
    pending: 'bg-gray-100 text-gray-800',
    in_progress: 'bg-blue-100 text-blue-800',
    done: 'bg-green-100 text-green-800',
    failed: 'bg-red-100 text-red-800',
    blocked: 'bg-yellow-100 text-yellow-800'
  };

  const statusIcons = {
    pending: '⏳',
    in_progress: '🔄',
    done: '✅',
    failed: '❌',
    blocked: '🚫'
  };

  const canExecute = task.status === 'pending' || task.status === 'failed' || task.status === 'blocked';
  const showRetry = task.status === 'failed' || task.status === 'blocked';
  
  const priorityColors = {
    1: 'bg-red-500',
    2: 'bg-orange-500', 
    3: 'bg-yellow-500',
    4: 'bg-blue-500',
    5: 'bg-gray-500'
  };

  const priorityLabels = {
    1: 'Urgent',
    2: 'High',
    3: 'Medium', 
    4: 'Low',
    5: 'Lowest'
  };
  
  const priority = task.priority || 3;
  const dependencies = task.dependencies || [];
  
  function renderErrorSummary(task) {
    if (!task.result) return '';
    
    let errorText = '';
    if (typeof task.result === 'object') {
      if (task.result.errors && Array.isArray(task.result.errors)) {
        errorText = task.result.errors.slice(0, 2).join(', ');
        if (task.result.errors.length > 2) errorText += '...';
      } else if (task.result.analysis) {
        errorText = task.result.analysis.substring(0, 100) + '...';
      } else {
        errorText = 'Task validation failed';
      }
    } else {
      errorText = String(task.result).substring(0, 100);
      if (String(task.result).length > 100) errorText += '...';
    }
    
    return errorText;
  }

  return `
    <div class="task-item flex items-start justify-between p-3 bg-white rounded-lg border border-gray-200 hover:border-gray-300 transition-all duration-200 hover:shadow-sm ${task.status === 'in_progress' ? 'animate-pulse border-blue-300 bg-blue-50' : ''}" data-task-id="${task.id}">
      <div class="flex items-start space-x-3 flex-1 min-w-0">
        <div class="flex flex-col items-center space-y-1 flex-shrink-0">
          <div class="w-3 h-3 rounded-full ${priorityColors[priority]} opacity-75" title="Priority: ${priorityLabels[priority]}"></div>
          ${dependencies.length > 0 ? `<div class="text-xs text-gray-400" title="${dependencies.length} dependencies">🔗</div>` : ''}
        </div>
        <div class="flex-1 min-w-0">
          <div class="flex items-center space-x-2 mb-1">
            <h4 class="text-sm font-medium text-gray-900 truncate">${escapeHtml(task.title)}</h4>
            <span class="text-xs">${statusIcons[task.status] || statusIcons.pending}</span>
          </div>
          ${task.description ? `<p class="text-xs text-gray-500 mb-2 line-clamp-2">${escapeHtml(task.description)}</p>` : ''}
          <div class="flex items-center justify-between mb-2">
            <div class="flex items-center space-x-2">
              <span class="inline-flex items-center px-2 py-1 rounded-full text-xs font-medium ${statusColors[task.status] || statusColors.pending}">${(task.status || 'pending').replace('_', ' ')}</span>
              ${priority <= 2 ? `<span class="inline-flex items-center px-1.5 py-0.5 rounded-full text-xs font-medium bg-red-100 text-red-700">${priorityLabels[priority]}</span>` : ''}
            </div>
            ${task.context?.sheet ? `<span class="text-xs text-gray-400">📊 ${escapeHtml(task.context.sheet)}</span>` : ''}
          </div>
          ${task.result && (task.status === 'blocked' || task.status === 'failed') ? `
            <div class="mt-2 p-2 rounded-md text-xs ${
              task.status === 'blocked' ? 'bg-yellow-50 border border-yellow-200' : 'bg-red-50 border border-red-200'
            }">
              <div class="flex items-start space-x-2">
                <span class="${task.status === 'blocked' ? 'text-yellow-600' : 'text-red-600'} font-medium">
                  ${task.status === 'blocked' ? '⚠️ Blocked:' : '❌ Error:'}
                </span>
                <div class="${task.status === 'blocked' ? 'text-yellow-800' : 'text-red-800'} flex-1">
                  ${escapeHtml(renderErrorSummary(task))}
                </div>
              </div>
              ${(task.result?.recommendations || task.result?.risks) ? `
                <div class="mt-1 text-xs ${task.status === 'blocked' ? 'text-yellow-700' : 'text-red-700'}">
                  <strong>Suggestions:</strong> ${task.result.recommendations?.[0] || task.result.risks?.[0]?.mitigation || 'Review task details'}
                </div>
              ` : ''}
            </div>
          ` : ''}
          ${task.duration ? `<div class="text-xs text-gray-400 mt-1">Completed in ${(task.duration/1000).toFixed(1)}s</div>` : ''}
          ${task.createdAt && !task.duration ? `<div class="text-xs text-gray-400 mt-1">${new Date(task.createdAt).toLocaleString()}</div>` : ''}
        </div>
      </div>
      <div class="flex flex-col items-center space-y-1 ml-3 flex-shrink-0">
        ${canExecute ? `<button onclick="executeTask('${task.id}')" class="p-2 ${showRetry ? 'text-orange-600 hover:text-orange-800 bg-orange-50 hover:bg-orange-100' : 'text-blue-600 hover:text-blue-800 bg-blue-50 hover:bg-blue-100'} rounded-full transition-all duration-200 hover:scale-105" title="${showRetry ? 'Retry Task' : 'Execute Task'}">
          ${showRetry ?
      '<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg>' :
      '<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M14.828 14.828a4 4 0 01-5.656 0M9 10h1m4 0h1m-6 4h1m4 0h1m-6-8h8a2 2 0 012 2v8a2 2 0 01-2 2H8a2 2 0 01-2-2V6a2 2 0 012-2z"></path></svg>'
    }</button>` : ''}
        ${task.status === 'done' ? `<button onclick="viewTaskResult('${task.id}')" class="p-2 text-green-600 hover:text-green-800 bg-green-50 hover:bg-green-100 rounded-full transition-all duration-200 hover:scale-105" title="View Result">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"/></svg>
          </button>` : ''}
        ${(task.status === 'blocked' || task.status === 'failed') ? `<button onclick="viewTaskResult('${task.id}')" class="p-2 text-gray-600 hover:text-gray-800 bg-gray-50 hover:bg-gray-100 rounded-full transition-all duration-200 hover:scale-105" title="View Error Details">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
        </button>` : ''}
        <button onclick="deleteTask('${task.id}')" class="p-2 text-red-600 hover:text-red-800 bg-red-50 hover:bg-red-100 rounded-full transition-all duration-200 hover:scale-105" title="Delete Task">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/></svg>
        </button>
      </div>
    </div>`;
}

export function drawTasks() {
  const list = document.getElementById('task-list');
  const summary = document.getElementById('task-summary');
  
  // Add task controls if not already present
  const taskControls = document.getElementById('task-controls');
  if (taskControls && !taskControls.querySelector('.auto-execute-toggle')) {
    const controlsHtml = `
      <div class="flex items-center justify-between mb-3 p-2 bg-gray-50 rounded-lg">
        <div class="flex items-center space-x-4">
          <label class="flex items-center space-x-2 text-sm">
            <input type="checkbox" id="auto-execute-toggle" class="auto-execute-toggle rounded border-gray-300" ${AppState.autoExecute ? 'checked' : ''}>
            <span class="font-medium text-gray-700">Auto-execute tasks</span>
          </label>
          <div class="text-xs text-gray-500">Automatically run tasks after planning</div>
        </div>
        <div class="flex items-center space-x-2">
          <select id="task-filter" class="text-xs border border-gray-300 rounded px-2 py-1">
            <option value="all">All Tasks</option>
            <option value="active">Active Only</option>
            <option value="completed">Completed Only</option>
            <option value="failed">Failed/Blocked</option>
          </select>
          <button onclick="executeTasks(AppState.tasks.filter(t => t.status === 'pending'))" class="px-3 py-1 bg-blue-500 text-white text-xs rounded hover:bg-blue-600 transition-colors" title="Execute all pending tasks">
            Execute All
          </button>
        </div>
      </div>
    `;
    taskControls.innerHTML = controlsHtml;
    
    // Add event listeners
    document.getElementById('auto-execute-toggle')?.addEventListener('change', (e) => {
      AppState.autoExecute = e.target.checked;
      localStorage.setItem('autoExecute', AppState.autoExecute);
    });
    
    document.getElementById('task-filter')?.addEventListener('change', (e) => {
      filterTasks(e.target.value);
    });
  }

  if (summary) {
    const pending = AppState.tasks.filter(t => t.status === 'pending').length;
    const inProgress = AppState.tasks.filter(t => t.status === 'in_progress').length;
    const completed = AppState.tasks.filter(t => t.status === 'done').length;

    if (AppState.tasks.length === 0) {
      summary.textContent = 'No active tasks';
    } else {
      summary.textContent = `${pending} pending, ${inProgress} running, ${completed} done`;
    }
  }

  if (AppState.tasks.length === 0) {
    list.innerHTML = '<div class="text-center text-gray-500 text-sm py-4">No tasks yet. Chat with AI to create tasks!</div>';
    return;
  }
  
  // Get current filter
  const currentFilter = document.getElementById('task-filter')?.value || 'all';
  let filteredTasks = AppState.tasks;
  
  switch (currentFilter) {
    case 'active':
      filteredTasks = AppState.tasks.filter(t => ['pending', 'in_progress', 'blocked'].includes(t.status));
      break;
    case 'completed':
      filteredTasks = AppState.tasks.filter(t => t.status === 'done');
      break;
    case 'failed':
      filteredTasks = AppState.tasks.filter(t => ['failed', 'blocked'].includes(t.status));
      break;
    default:
      filteredTasks = AppState.tasks;
  }

  const tasksByStatus = {
    in_progress: filteredTasks.filter(t => t.status === 'in_progress'),
    pending: filteredTasks.filter(t => t.status === 'pending'),
    blocked: filteredTasks.filter(t => t.status === 'blocked'),
    failed: filteredTasks.filter(t => t.status === 'failed'),
    done: filteredTasks.filter(t => t.status === 'done')
  };

  let html = '';

  const activeTasks = [...tasksByStatus.in_progress, ...tasksByStatus.pending, ...tasksByStatus.blocked, ...tasksByStatus.failed];
  if (activeTasks.length > 0) {
    html += '<div class="space-y-2">';
    html += activeTasks.map(renderTask).join('');
    html += '</div>';
  }

  if (tasksByStatus.done.length > 0) {
    html += `
      <div class="mt-4 pt-4 border-t border-gray-200">
        <button onclick="toggleCompletedTasks()" class="flex items-center space-x-2 text-sm text-gray-600 hover:text-gray-800 mb-2">
          <svg id="completed-toggle-icon" class="w-4 h-4 transition-transform" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 5l7 7-7 7"/>
          </svg>
          <span>Completed (${tasksByStatus.done.length})</span>
        </button>
        <div id="completed-tasks" class="hidden space-y-2">
          ${tasksByStatus.done.map(renderTask).join('')}
        </div>
      </div>`;
  }

  list.innerHTML = html;
}

function filterTasks(filterType) {
  drawTasks(); // Redraw with the new filter
}

window.toggleCompletedTasks = function () {
  const completedTasks = document.getElementById('completed-tasks');
  const toggleIcon = document.getElementById('completed-toggle-icon');

  if (completedTasks.classList.contains('hidden')) {
    completedTasks.classList.remove('hidden');
    toggleIcon.style.transform = 'rotate(90deg)';
  } else {
    completedTasks.classList.add('hidden');
    toggleIcon.style.transform = 'rotate(0deg)';
  }
};

window.viewTaskResult = function (id) {
  const task = AppState.tasks.find(t => t.id === id);
  if (!task || !task.result) return;

  const modal = new Modal();
  const result = task.result;
  const isError = task.status === 'failed' || task.status === 'blocked';
  
  // Enhanced result formatting
  let contentHtml = '';
  
  if (isError) {
    const errorIcon = task.status === 'blocked' ? '🛑' : '❌';
    const errorColor = task.status === 'blocked' ? 'bg-yellow-50 border-yellow-200' : 'bg-red-50 border-red-200';
    const textColor = task.status === 'blocked' ? 'text-yellow-800' : 'text-red-800';
    
    contentHtml = `
      <div class="space-y-4">
        <div class="flex items-start space-x-3 p-3 ${errorColor} border rounded-lg">
          <span class="text-lg">${errorIcon}</span>
          <div>
            <h4 class="font-medium ${textColor} mb-1">${task.status === 'blocked' ? 'Task Blocked' : 'Task Failed'}</h4>
            <p class="text-sm ${textColor}">This task could not be completed. Review the details below for troubleshooting.</p>
          </div>
        </div>
        
        <div class="grid grid-cols-2 gap-3 text-sm">
          <div><strong>Status:</strong> ${task.status}</div>
          <div><strong>Sheet:</strong> ${task.context?.sheet || 'Unknown'}</div>
          <div><strong>Priority:</strong> ${['', 'Urgent', 'High', 'Medium', 'Low', 'Lowest'][task.priority || 3]}</div>
          <div><strong>Retry Count:</strong> ${task.retryCount || 0}/${task.maxRetries || 3}</div>
        </div>`;
    
    if (typeof result === 'object') {
      if (result.analysis) {
        contentHtml += `
          <div>
            <h5 class="font-medium text-gray-900 mb-2">Analysis</h5>
            <div class="bg-gray-50 p-3 rounded-lg text-sm">${escapeHtml(result.analysis)}</div>
          </div>`;
      }
      
      if (result.errors && result.errors.length > 0) {
        contentHtml += `
          <div>
            <h5 class="font-medium text-gray-900 mb-2">Errors</h5>
            <ul class="bg-red-50 p-3 rounded-lg text-sm space-y-1">
              ${result.errors.map(error => `<li class="flex items-start space-x-2"><span class="text-red-500">•</span><span>${escapeHtml(error)}</span></li>`).join('')}
            </ul>
          </div>`;
      }
      
      if (result.recommendations && result.recommendations.length > 0) {
        contentHtml += `
          <div>
            <h5 class="font-medium text-gray-900 mb-2">Recommendations</h5>
            <ul class="bg-blue-50 p-3 rounded-lg text-sm space-y-1">
              ${result.recommendations.map(rec => `<li class="flex items-start space-x-2"><span class="text-blue-500">💡</span><span>${escapeHtml(rec)}</span></li>`).join('')}
            </ul>
          </div>`;
      }
      
      if (result.risks && result.risks.length > 0) {
        contentHtml += `
          <div>
            <h5 class="font-medium text-gray-900 mb-2">Risk Assessment</h5>
            <div class="space-y-2">
              ${result.risks.map(risk => `
                <div class="flex items-start space-x-2 p-2 bg-orange-50 rounded">
                  <span class="text-orange-500 font-bold text-xs px-1.5 py-0.5 bg-orange-200 rounded">${risk.level?.toUpperCase()}</span>
                  <div class="text-sm">
                    <p class="font-medium">${escapeHtml(risk.description)}</p>
                    ${risk.mitigation ? `<p class="text-orange-700 mt-1">💡 ${escapeHtml(risk.mitigation)}</p>` : ''}
                  </div>
                </div>
              `).join('')}
            </div>
          </div>`;
      }
    } else {
      contentHtml += `
        <div>
          <h5 class="font-medium text-gray-900 mb-2">Error Details</h5>
          <div class="bg-gray-50 p-3 rounded-lg text-sm font-mono">${escapeHtml(String(result))}</div>
        </div>`;
    }
    
    contentHtml += '</div>';
  } else {
    // Success case
    contentHtml = `
      <div class="space-y-4">
        <div class="flex items-start space-x-3 p-3 bg-green-50 border border-green-200 rounded-lg">
          <span class="text-lg">✅</span>
          <div>
            <h4 class="font-medium text-green-800 mb-1">Task Completed Successfully</h4>
            <p class="text-sm text-green-700">This task has been executed and validated.</p>
          </div>
        </div>
        
        <div class="grid grid-cols-2 gap-3 text-sm">
          <div><strong>Status:</strong> ${task.status}</div>
          <div><strong>Sheet:</strong> ${task.context?.sheet || 'Unknown'}</div>
          <div><strong>Duration:</strong> ${task.duration ? (task.duration/1000).toFixed(1) + 's' : 'Unknown'}</div>
          <div><strong>Completed:</strong> ${task.completedAt ? new Date(task.completedAt).toLocaleString() : 'Unknown'}</div>
        </div>`;
    
    if (typeof result === 'object' && result.message) {
      contentHtml += `
        <div>
          <h5 class="font-medium text-gray-900 mb-2">Result Summary</h5>
          <div class="bg-gray-50 p-3 rounded-lg text-sm">${escapeHtml(result.message)}</div>
        </div>`;
    }
    
    if (typeof result === 'object') {
      contentHtml += `
        <details class="bg-gray-50 rounded-lg">
          <summary class="p-3 cursor-pointer font-medium text-gray-700 hover:bg-gray-100 rounded-lg">View Technical Details</summary>
          <div class="p-3 pt-0">
            <pre class="text-xs text-gray-600 whitespace-pre-wrap overflow-auto">${escapeHtml(JSON.stringify(result, null, 2))}</pre>
          </div>
        </details>`;
    }
    
    contentHtml += '</div>';
  }

  modal.show({
    title: `${isError ? (task.status === 'blocked' ? '🛑' : '❌') : '✅'} ${task.title}`,
    content: contentHtml,
    buttons: [
      ...(isError && (task.status === 'failed' || task.status === 'blocked') ? [{
        text: 'Retry Task', 
        action: 'retry', 
        onClick: () => executeTask(task.id)
      }] : []),
      { text: 'Close', action: 'close', primary: true }
    ],
    size: 'xl'
  });
};

window.deleteTask = function (id) {
  AppState.tasks = AppState.tasks.filter(t => t.id !== id);
  saveTasks();
  drawTasks();
};

export async function runOrchestrator(tasks) {
  const provider = pickProvider();

  if (provider === 'mock') {
    return {
      executionPlan: tasks.map((t, i) => ({ taskId: t.id, order: i + 1, dependencies: [] })),
      estimatedTime: tasks.length * 2000,
      riskAssessment: 'low',
      recommendations: ['Execute tasks sequentially']
    };
  }

  const ws = getWorksheet();
  const sheetContext = ws['!ref'] ? `Sheet "${AppState.activeSheet}" range: ${ws['!ref']}` : `Empty sheet "${AppState.activeSheet}"`;
  const sampleData = ws['!ref'] ? getSampleDataFromSheet(ws) : 'No data';

  const system = `You are the Orchestrator Agent...`; // Content omitted for brevity

  const tasksSummary = tasks.map(t => ({
    id: t.id,
    title: t.title,
    description: t.description,
    dependencies: t.dependencies || [],
    priority: t.priority || 3,
    context: t.context || {}
  }));

  const user = `Orchestrate execution of ${tasks.length} tasks:\n${JSON.stringify(tasksSummary, null, 2)}`;
  const messages = [{ role: 'system', content: system }, { role: 'user', content: user }];

  try {
    let data;
    const selectedModel = getSelectedModel();
    if (provider === 'openai') {
      data = await fetchOpenAI(AppState.keys.openai, messages, selectedModel);
    } else {
      data = await fetchGemini(AppState.keys.gemini, messages, selectedModel);
    }

    let text = '';
    if (provider === 'openai') {
      text = data.choices?.[0]?.message?.content || '';
    } else {
      text = data.candidates?.[0]?.content?.parts?.map(p => p.text).join('') || '';
    }

    let result = null;
    try {
      result = JSON.parse(text);
    } catch {
      result = extractFirstJson(text);
    }

    return result || {
      executionPlan: tasks.map((t, i) => ({ taskId: t.id, order: i + 1, dependencies: t.dependencies || [] })),
      riskAssessment: 'unknown',
      recommendations: ['Execute with caution - orchestrator analysis failed']
    };

  } catch (error) {
    console.error('Orchestrator failed:', error);
    return {
      executionPlan: tasks.map((t, i) => ({ taskId: t.id, order: i + 1, dependencies: t.dependencies || [] })),
      riskAssessment: 'high',
      recommendations: ['Manual review recommended - orchestrator unavailable']
    };
  }
}

window.executeTask = async function (id) {
  const task = AppState.tasks.find(t => t.id === id);
  if (!task) return;

  const uncompletedDeps = task.dependencies?.filter(depId => {
    const depTask = AppState.tasks.find(t => t.id === depId);
    return !depTask || depTask.status !== 'done';
  }) || [];

  if (uncompletedDeps.length > 0) {
    showToast(`Cannot execute: waiting for dependencies (${uncompletedDeps.join(', ')})`, 'warning');
    return;
  }

  task.status = 'in_progress';
  task.startTime = Date.now();
  saveTasks();
  drawTasks();

  try {
    const result = await runExecutor(task);
    if (!result) throw new Error('No executor result');

    const validation = await runValidator(result, task);
    if (!validation.valid) {
      task.status = 'blocked';
      task.result = validation;
      task.retryCount = (task.retryCount || 0) + 1;
      saveTasks();
      drawTasks();

      if (task.retryCount < task.maxRetries) {
        showToast(`Task blocked - ${task.maxRetries - task.retryCount} retries remaining`, 'warning');
      } else {
        showToast('Task failed after maximum retries', 'error');
        task.status = 'failed';
      }
      return;
    }

    await applyEditsOrDryRun(result);
    task.status = 'done';
    task.result = result;
    task.completedAt = Date.now();
    task.duration = task.completedAt - task.startTime;

    saveTasks();
    drawTasks();
    
    // Add completion animation
    setTimeout(() => {
      const taskElement = document.querySelector(`[data-task-id="${task.id}"]`);
      if (taskElement) {
        taskElement.classList.add('task-complete-flash');
        setTimeout(() => taskElement.classList.remove('task-complete-flash'), 600);
      }
    }, 100);
    
    showToast(`Task completed: ${task.title}`, 'success');

    const enabledTasks = AppState.tasks.filter(t =>
      t.status === 'pending' &&
      t.dependencies?.includes(id) &&
      t.dependencies.every(depId => {
        const depTask = AppState.tasks.find(dt => dt.id === depId);
        return depTask?.status === 'done';
      })
    );

    if (enabledTasks.length > 0) {
      showToast(`${enabledTasks.length} task(s) now ready to execute`, 'info');
    }

  } catch (e) {
    console.error('Task execution failed:', e);
    task.status = 'failed';
    task.result = String(e);
    task.retryCount = (task.retryCount || 0) + 1;
    saveTasks();
    drawTasks();
    showToast(`Task failed: ${task.title}`, 'error');
  }
};

export async function executeTasks(tasks, orchestration = null) {
  if (!tasks || tasks.length === 0) return;

  const startTime = Date.now();
  let completedCount = 0;
  let failedCount = 0;
  let blockedCount = 0;
  const results = [];

  // If a precomputed orchestration is provided (from chat flow), use it; else compute here
  if (orchestration && Array.isArray(orchestration.executionPlan)) {
    showToast(`Executing ${tasks.length} task(s) with precomputed plan...`, 'info', 3000);
  } else {
    showToast(`Orchestrating execution of ${tasks.length} task(s)...`, 'info');
    try {
      orchestration = await runOrchestrator(tasks);
      log('Orchestration plan:', orchestration);
    } catch (error) {
      console.error('Task orchestration failed:', error);
      orchestration = null;
    }
  }

  // Optional informational toasts
  if (orchestration?.riskAssessment) {
    showToast(`Orchestration risk: ${orchestration.riskAssessment}`, 'info', 3000);
  }
  if (typeof orchestration?.estimatedTime === 'number') {
    showToast(`Estimated time: ${(orchestration.estimatedTime / 1000).toFixed(1)}s`, 'info', 3000);
  }

  // Determine ordered tasks: prefer orchestration plan when available
  let sortedTasks = tasks;
  if (Array.isArray(orchestration?.executionPlan) && orchestration.executionPlan.length) {
    sortedTasks = orchestration.executionPlan
      .slice()
      .sort((a, b) => (a.order || 0) - (b.order || 0))
      .map(plan => tasks.find(t => t.id === plan.taskId))
      .filter(Boolean);
  }

  // Execute sequentially respecting the determined order
  try {
    for (const task of sortedTasks) {
      const taskStartTime = Date.now();
      await executeTask(task.id);
      const taskEndTime = Date.now();
      
      // Refresh task status from AppState
      const updatedTask = AppState.tasks.find(t => t.id === task.id);
      if (updatedTask) {
        if (updatedTask.status === 'done') {
          completedCount++;
          results.push({
            title: updatedTask.title,
            status: 'completed',
            duration: taskEndTime - taskStartTime,
            result: updatedTask.result
          });
        } else if (updatedTask.status === 'failed') {
          failedCount++;
          results.push({
            title: updatedTask.title,
            status: 'failed',
            error: updatedTask.result,
            duration: taskEndTime - taskStartTime
          });
        } else if (updatedTask.status === 'blocked') {
          blockedCount++;
          results.push({
            title: updatedTask.title,
            status: 'blocked',
            issue: updatedTask.result,
            duration: taskEndTime - taskStartTime
          });
        }
      }
      
      await new Promise(resolve => setTimeout(resolve, 500));
    }
    
    const totalTime = Date.now() - startTime;
    showToast('Task orchestration completed', 'success');
    
    // Add execution summary to chat
    addExecutionSummaryToChat(results, totalTime, completedCount, failedCount, blockedCount);
    
  } catch (error) {
    console.error('Task execution loop failed:', error);
    showToast('Task execution encountered an error', 'error');
    
    // Add error summary to chat
    addExecutionSummaryToChat(results, Date.now() - startTime, completedCount, failedCount, blockedCount, error);
  }
};

export async function autoExecuteTasks() {
  if (!AppState.autoExecute) return;

  const pendingTasks = AppState.tasks.filter(t => t.status === 'pending');
  if (pendingTasks.length > 0) {
    showToast(`Auto-executing ${pendingTasks.length} task(s)...`, 'info');
    await executeTasks(pendingTasks);
  }
}

// Add execution summary to chat
function addExecutionSummaryToChat(results, totalTime, completedCount, failedCount, blockedCount, error = null) {
  // Only add to chat if we have chat messages (meaning it was initiated from chat)
  if (!AppState.messages || AppState.messages.length === 0) return;
  
  const totalTasks = results.length;
  let content = '';
  let buttons = [];
  let structuredData = null;
  
  if (error) {
    content = `❌ **Task Execution Failed**\n\nExecution was interrupted: ${error.message}\n\nCompleted: ${completedCount}, Failed: ${failedCount}, Blocked: ${blockedCount}`;
    
    buttons = [
      {
        label: 'Retry Failed Tasks',
        action: 'retryTasks',
        type: 'primary',
        icon: '🔄'
      },
      {
        label: 'View Tasks',
        action: 'viewTasks',
        type: 'secondary',
        icon: '📋'
      }
    ];
  } else {
    // Success summary
    const timeStr = (totalTime / 1000).toFixed(1);
    content = `✅ **Task Execution Complete**\n\nExecuted ${totalTasks} tasks in ${timeStr}s`;
    
    if (completedCount > 0) content += `\n• ✅ ${completedCount} completed successfully`;
    if (failedCount > 0) content += `\n• ❌ ${failedCount} failed`;
    if (blockedCount > 0) content += `\n• 🚫 ${blockedCount} blocked`;
    
    // Add buttons based on results
    buttons = [];
    if (failedCount > 0 || blockedCount > 0) {
      buttons.push({
        label: 'Retry Failed Tasks',
        action: 'retryTasks',
        type: 'danger',
        icon: '🔄'
      });
    }
    
    buttons.push({
      label: 'View Tasks',
      action: 'viewTasks',
      type: 'secondary',
      icon: '📋'
    });
    
    if (completedCount > 0) {
      buttons.push({
        label: 'Clear Completed',
        action: 'clearCompleted',
        type: 'secondary',
        icon: '🗑️'
      });
    }
  }
  
  // Create structured data for detailed results
  if (results.length > 0) {
    structuredData = {
      headers: ['Task', 'Status', 'Duration'],
      rows: results.map(r => [
        r.title.length > 30 ? r.title.substring(0, 30) + '...' : r.title,
        r.status === 'completed' ? '✅ Done' : 
        r.status === 'failed' ? '❌ Failed' : '🚫 Blocked',
        `${(r.duration / 1000).toFixed(1)}s`
      ])
    };
  }
  
  const summaryMsg = {
    role: 'assistant',
    content: content,
    timestamp: Date.now(),
    agentType: 'executor',
    buttons: buttons,
    dataType: structuredData ? 'table' : null,
    structuredData: structuredData,
    taskStatus: {
      total: AppState.tasks.length,
      completed: AppState.tasks.filter(t => t.status === 'done').length
    }
  };
  
  AppState.messages.push(summaryMsg);
  
  // Import drawChat dynamically to avoid circular dependency
  import('../chat/chat-ui.js').then(({ drawChat }) => {
    drawChat();
  }).catch(console.error);
}

// Also expose to window for HTML onclick handlers
window.executeTasks = executeTasks;