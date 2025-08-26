import { AppState } from '../core/state.js';
import { db } from '../db/indexeddb.js';
import { escapeHtml, extractFirstJson } from '../utils/index.js';
import { Modal } from '../ui/modal.js';
import { showToast } from '../ui/toast.js';
import { getWorksheet } from '../spreadsheet/workbook-manager.js';
import { runExecutor, runValidator, fetchOpenAI, fetchGemini } from '../services/ai-agents.js';
/* global getSampleDataFromSheet, applyEditsOrDryRun, log, pickProvider, getSelectedModel */

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

  return `
    <div class="task-item flex items-start justify-between p-3 bg-white rounded-lg border border-gray-200 hover:border-gray-300 transition-colors ${task.status === 'in_progress' ? 'animate-pulse' : ''}" data-task-id="${task.id}">
      <div class="flex-1 min-w-0">
        <div class="flex items-center space-x-2 mb-1">
          <h4 class="text-sm font-medium text-gray-900 truncate">${escapeHtml(task.title)}</h4>
          <span class="text-xs">${statusIcons[task.status] || statusIcons.pending}</span>
        </div>
        ${task.description ? `<p class="text-xs text-gray-500 mb-2 line-clamp-2">${escapeHtml(task.description)}</p>` : ''}
        <div class="flex items-center justify-between">
          <span class="inline-flex items-center px-2 py-1 rounded-full text-xs font-medium ${statusColors[task.status] || statusColors.pending}">${(task.status || 'pending').replace('_', ' ')}</span>
          ${task.context?.sheet ? `<span class="text-xs text-gray-400">📊 ${escapeHtml(task.context.sheet)}</span>` : ''}
        </div>
        ${task.result && task.status === 'blocked' ? `<div class="mt-2 p-2 bg-yellow-50 rounded text-xs text-yellow-800">${escapeHtml(typeof task.result === 'object' ? task.result.errors?.join(', ') || 'Task blocked' : task.result)}</div>` : ''}
        ${task.result && task.status === 'failed' ? `<div class="mt-2 p-2 bg-red-50 rounded text-xs text-red-800">${escapeHtml(typeof task.result === 'string' ? task.result : 'Task failed')}</div>` : ''}
        ${task.createdAt ? `<div class="text-xs text-gray-400 mt-1">${new Date(task.createdAt).toLocaleString()}</div>` : ''}
      </div>
      <div class="flex items-center space-x-1 ml-3 flex-shrink-0">
        ${canExecute ? `<button onclick="executeTask('${task.id}')" class="p-1 text-blue-600 hover:text-blue-800 transition-colors" title="${showRetry ? 'Retry' : 'Execute'}">
          ${showRetry ?
      '<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9m11 11v-5h-.581m0 0a8.003 8.003 0 01-15.357-2m15.357 2H15"/></svg>' :
      '<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M14.828 14.828a4 4 0 01-5.656 0M9 10h1m4 0h1m-6 4h1m4 0h1m-6-8h8a2 2 0 012 2v8a2 2 0 01-2 2H8a2 2 0 01-2-2V6a2 2 0 012-2z"></path></svg>'
    }</button>` : ''}
        ${task.status === 'done' ? `<button onclick="viewTaskResult('${task.id}')" class="p-1 text-green-600 hover:text-green-800 transition-colors" title="View Result">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"/><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"/></svg>
          </button>` : ''}
        <button onclick="deleteTask('${task.id}')" class="p-1 text-red-600 hover:text-red-800 transition-colors" title="Delete">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"/></svg>
        </button>
      </div>
    </div>`;
}

export function drawTasks() {
  const list = document.getElementById('task-list');
  const summary = document.getElementById('task-summary');

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

  const tasksByStatus = {
    in_progress: AppState.tasks.filter(t => t.status === 'in_progress'),
    pending: AppState.tasks.filter(t => t.status === 'pending'),
    blocked: AppState.tasks.filter(t => t.status === 'blocked'),
    failed: AppState.tasks.filter(t => t.status === 'failed'),
    done: AppState.tasks.filter(t => t.status === 'done')
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
  const result = typeof task.result === 'object' ? JSON.stringify(task.result, null, 2) : String(task.result);

  modal.show({
    title: `Task Result: ${task.title}`,
    content: `
      <div class="space-y-3">
        <div class="text-sm text-gray-600">
          <strong>Status:</strong> ${task.status} <br>
          <strong>Sheet:</strong> ${task.context?.sheet || 'Unknown'} <br>
          <strong>Completed:</strong> ${new Date(task.createdAt).toLocaleString()}
        </div>
        <div class="bg-gray-50 p-3 rounded-lg">
          <pre class="text-sm text-gray-800 whitespace-pre-wrap">${escapeHtml(result)}</pre>
        </div>
      </div>`,
    buttons: [{ text: 'Close', action: 'close', primary: true }],
    size: 'lg'
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

export async function executeTasks(taskIds) {
  const tasks = taskIds.map(id => AppState.tasks.find(t => t.id === id)).filter(Boolean);
  if (tasks.length === 0) return;

  showToast(`Orchestrating execution of ${tasks.length} tasks...`, 'info');

  try {
    const orchestration = await runOrchestrator(tasks);
    log('Orchestration plan:', orchestration);

    if (orchestration.executionPlan) {
      const sortedTasks = orchestration.executionPlan
        .sort((a, b) => a.order - b.order)
        .map(plan => tasks.find(t => t.id === plan.taskId))
        .filter(Boolean);

      for (const task of sortedTasks) {
        await executeTask(task.id);
        await new Promise(resolve => setTimeout(resolve, 500));
      }
    }

    showToast('Task orchestration completed', 'success');

  } catch (error) {
    console.error('Task orchestration failed:', error);
    showToast('Orchestration failed, executing tasks sequentially', 'warning');

    for (const task of tasks) {
      await executeTask(task.id);
    }
  }
};

export async function autoExecuteTasks() {
  if (!AppState.autoExecute) return;

  const pendingTasks = AppState.tasks.filter(t => t.status === 'pending');
  if (pendingTasks.length > 0) {
    showToast(`Auto-executing ${pendingTasks.length} task(s)...`, 'info');
    await executeTasks(pendingTasks.map(t => t.id));
  }
}

// Also expose to window for HTML onclick handlers
window.executeTasks = executeTasks;