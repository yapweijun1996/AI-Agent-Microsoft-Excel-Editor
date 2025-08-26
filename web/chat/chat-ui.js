import { AppState } from '../core/state.js';
import { escapeHtml } from '../utils/index.js';
import { runPlanner, runIntentAgent } from '../services/ai-agents.js';
import { saveTasks, drawTasks, executeTasks, runOrchestrator } from '../tasks/task-manager.js';
import { showToast } from '../ui/toast.js';
import { showModal, showSettingsModal } from '../ui/modals.js';
/* global executeTask, openSettingsModal */

// Agent transition animation helper
function getAgentTransitionClass(agentType, isTransition) {
  if (!isTransition || AppState.reducedMotion) return '';
  return 'transition-all duration-500 transform scale-110 animate-bounce';
}

// Global chat action handler
window.handleChatAction = function(action, data) {
  switch (action) {
    case 'executeAllTasks':
      const pendingTasks = AppState.tasks.filter(t => t.status === 'pending' || t.status === 'failed');
      if (pendingTasks.length > 0) {
        executeTasks(pendingTasks);
      } else {
        showToast('No tasks to execute', 'info');
      }
      break;
      
    case 'executeTask':
      if (data && window.executeTask) {
        window.executeTask(data);
      }
      break;
      
    case 'viewDiff':
      if (data) {
        showDiffModal(data);
      }
      break;
      
    case 'openApiKeys':
      if (window.openSettingsModal) {
        window.openSettingsModal();
      }
      break;
      
    case 'retryTasks':
      const failedTasks = AppState.tasks.filter(t => t.status === 'failed' || t.status === 'blocked');
      if (failedTasks.length > 0) {
        executeTasks(failedTasks);
      }
      break;
      
    case 'clearTasks':
      if (confirm('Clear all tasks? This cannot be undone.')) {
        AppState.tasks = [];
        saveTasks();
        drawTasks();
        showToast('All tasks cleared', 'success');
      }
      break;
      
    case 'stopExecution':
      if (window.abortController) {
        window.abortController.abort();
        showToast('Execution stopped', 'info');
      }
      break;
      
    default:
      console.warn('Unknown chat action:', action);
  }
};

// Show diff modal for spreadsheet changes
function showDiffModal(diffData) {
  try {
    const data = typeof diffData === 'string' ? JSON.parse(diffData) : diffData;
    const modal = showModal('Spreadsheet Changes Preview', 
      `<div class="space-y-4">
        <div class="bg-yellow-50 p-3 rounded-md">
          <h4 class="font-medium text-yellow-800">Proposed Changes:</h4>
          <div class="mt-2 text-sm text-yellow-700">
            ${data.changes.map(change => 
              `<div class="mb-1">Cell ${change.cell}: "${change.oldValue}" → "${change.newValue}"</div>`
            ).join('')}
          </div>
        </div>
        <div class="flex justify-end space-x-2">
          <button onclick="this.closest('.modal').remove()" 
                  class="px-3 py-2 text-sm bg-gray-200 hover:bg-gray-300 rounded-md">
            Cancel
          </button>
          <button onclick="applyDiffChanges('${btoa(JSON.stringify(data))}'); this.closest('.modal').remove()" 
                  class="px-3 py-2 text-sm bg-blue-500 hover:bg-blue-600 text-white rounded-md">
            Apply Changes
          </button>
        </div>
      </div>`, 
      { size: 'lg' }
    );
  } catch (e) {
    console.error('Error showing diff modal:', e);
    showToast('Error displaying changes preview', 'error');
  }
}

// Apply diff changes
window.applyDiffChanges = function(encodedData) {
  try {
    const data = JSON.parse(atob(encodedData));
    if (window.applyEditsOrDryRun) {
      window.applyEditsOrDryRun(data.changes, false);
      showToast('Changes applied successfully', 'success');
    }
  } catch (e) {
    console.error('Error applying diff changes:', e);
    showToast('Error applying changes', 'error');
  }
};

// Stop button management
function showStopButton() {
  const chatInput = document.getElementById('chat-input-container');
  if (chatInput) {
    let stopButton = document.getElementById('stop-button');
    if (!stopButton) {
      stopButton = document.createElement('button');
      stopButton.id = 'stop-button';
      stopButton.innerHTML = '⏹️ Stop';
      stopButton.className = 'px-3 py-1 bg-red-500 text-white text-sm rounded-md hover:bg-red-600 transition-colors';
      stopButton.onclick = () => window.handleChatAction('stopExecution');
      chatInput.appendChild(stopButton);
    }
    stopButton.style.display = 'inline-block';
  }
}

function hideStopButton() {
  const stopButton = document.getElementById('stop-button');
  if (stopButton) {
    stopButton.style.display = 'none';
  }
}

// Additional action handlers
window.handleChatAction = function(action, data) {
  switch (action) {
    case 'executeAllTasks':
      const pendingTasks = AppState.tasks.filter(t => t.status === 'pending' || t.status === 'failed');
      if (pendingTasks.length > 0) {
        executeTasks(pendingTasks);
        updateTaskStatusInChat();
      } else {
        showToast('No tasks to execute', 'info');
      }
      break;
      
    case 'executeTask':
      if (data && window.executeTask) {
        window.executeTask(data);
        updateTaskStatusInChat();
      }
      break;
      
    case 'viewDiff':
      if (data) {
        showDiffModal(data);
      }
      break;

    case 'viewTasks':
      // Scroll to task panel or highlight it
      const taskPanel = document.querySelector('[data-panel="tasks"]');
      if (taskPanel) {
        taskPanel.scrollIntoView({ behavior: 'smooth' });
        taskPanel.classList.add('animate-pulse');
        setTimeout(() => taskPanel.classList.remove('animate-pulse'), 1000);
      }
      break;
      
    case 'openApiKeys':
      showSettingsModal();
      break;

    case 'retryLastRequest':
      const lastUserMessage = [...AppState.messages].reverse().find(m => m.role === 'user');
      if (lastUserMessage) {
        const input = document.getElementById('message-input');
        if (input) {
          input.value = lastUserMessage.content;
        }
      }
      break;
      
    case 'retryTasks':
      const failedTasks = AppState.tasks.filter(t => t.status === 'failed' || t.status === 'blocked');
      if (failedTasks.length > 0) {
        executeTasks(failedTasks);
        updateTaskStatusInChat();
      }
      break;
      
    case 'clearTasks':
      if (confirm('Clear all tasks? This cannot be undone.')) {
        AppState.tasks = [];
        saveTasks();
        drawTasks();
        showToast('All tasks cleared', 'success');
        updateTaskStatusInChat();
      }
      break;

    case 'clearCompleted':
      const completedTasks = AppState.tasks.filter(t => t.status === 'done');
      if (completedTasks.length > 0) {
        if (confirm(`Clear ${completedTasks.length} completed tasks? This cannot be undone.`)) {
          AppState.tasks = AppState.tasks.filter(t => t.status !== 'done');
          saveTasks();
          drawTasks();
          showToast('Completed tasks cleared', 'success');
          updateTaskStatusInChat();
        }
      }
      break;
      
    case 'stopExecution':
      if (window.abortController) {
        window.abortController.abort();
        showToast('Execution stopped', 'info');
      }
      break;
      
    default:
      console.warn('Unknown chat action:', action);
  }
};

// Update task status in the most recent AI message
function updateTaskStatusInChat() {
  const lastAiMessage = [...AppState.messages].reverse().find(m => m.role === 'assistant' && m.taskStatus);
  if (lastAiMessage) {
    const completed = AppState.tasks.filter(t => t.status === 'done').length;
    const total = AppState.tasks.length;
    lastAiMessage.taskStatus = { completed, total };
    drawChat();
  }
}

function renderChatMessage(msg) {
  const isUser = msg.role === 'user';
  const isTyping = msg.isTyping || false;
  const agentType = msg.agentType || 'assistant';

  let content = escapeHtml(msg.content);
  if (!isUser) {
    content = content.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    content = content.replace(/\n/g, '<br>');
    // Enhanced markdown-like formatting
    content = content.replace(/`([^`]+)`/g, '<code class="bg-gray-100 px-1 py-0.5 rounded text-xs font-mono">$1</code>');
    content = content.replace(/^- (.+)$/gm, '<div class="flex items-start space-x-2"><span class="text-blue-500">•</span><span>$1</span></div>');
  }

  // Render interactive buttons if present
  let buttonsHtml = '';
  if (msg.buttons && msg.buttons.length > 0) {
    const buttonElements = msg.buttons.map(btn => {
      const buttonClass = btn.type === 'primary' ? 'bg-blue-500 hover:bg-blue-600 text-white' : 
                         btn.type === 'danger' ? 'bg-red-500 hover:bg-red-600 text-white' :
                         btn.type === 'success' ? 'bg-green-500 hover:bg-green-600 text-white' :
                         'bg-gray-200 hover:bg-gray-300 text-gray-800';
      return `<button onclick="handleChatAction('${btn.action}', '${btn.data || ''}')
                       class="px-3 py-1.5 text-xs rounded-md transition-colors ${buttonClass} ${btn.disabled ? 'opacity-50 cursor-not-allowed' : ''}" 
                       ${btn.disabled ? 'disabled' : ''}>
                ${btn.icon || ''} ${btn.label}
              </button>`;
    }).join('');
    buttonsHtml = `<div class="mt-2 flex flex-wrap gap-2">${buttonElements}</div>`;
  }

  // Render structured data if present
  let structuredDataHtml = '';
  if (msg.dataType && msg.structuredData) {
    if (msg.dataType === 'table' && msg.structuredData.headers && msg.structuredData.rows) {
      const { headers, rows } = msg.structuredData;
      structuredDataHtml = `
        <div class="mt-2 overflow-x-auto">
          <table class="min-w-full text-xs border border-gray-200 rounded-md">
            <thead class="bg-gray-50">
              <tr>${headers.map(h => `<th class="px-2 py-1 text-left font-medium text-gray-700 border-b">${escapeHtml(h)}</th>`).join('')}</tr>
            </thead>
            <tbody>
              ${rows.map(row => `<tr class="hover:bg-gray-50">${row.map(cell => `<td class="px-2 py-1 border-b">${escapeHtml(cell)}</td>`).join('')}</tr>`).join('')}
            </tbody>
          </table>
        </div>`;
    } else if (msg.dataType === 'list' && Array.isArray(msg.structuredData)) {
      structuredDataHtml = `
        <div class="mt-2">
          <ul class="space-y-1 text-xs">
            ${msg.structuredData.map(item => `<li class="flex items-start space-x-2"><span class="text-blue-500 mt-0.5">•</span><span>${escapeHtml(item)}</span></li>`).join('')}
          </ul>
        </div>`;
    }
  }

  const agentIcons = {
    intent: '🧠',
    planner: '📋',
    executor: '⚡',
    validator: '✅',
    orchestrator: '🎯',
    assistant: '🤖'
  };

  const agentNames = {
    intent: 'Intent Agent',
    planner: 'Planner Agent', 
    executor: 'Executor Agent',
    validator: 'Validator Agent',
    orchestrator: 'Orchestrator Agent',
    assistant: 'AI Assistant'
  };

  const agentIcon = agentIcons[agentType] || agentIcons.assistant;
  const agentName = agentNames[agentType] || agentNames.assistant;

  // Apply error styling if message type is error
  const isError = msg.type === 'error';
  const messageClasses = isUser ? 'bg-blue-500 text-white' : 
                        isError ? 'bg-red-50 text-red-800 border-l-4 border-red-400' :
                        isTyping ? 'bg-yellow-100 text-yellow-800 border-l-4 border-yellow-400' : 
                        'bg-gray-200 text-gray-900';

  const animationClasses = AppState.reducedMotion ? '' : 
    (isTyping ? 'animate-pulse' : '') + ' transition-all duration-200';

  return `
    <div class="flex ${isUser ? 'justify-end' : 'justify-start'} ${animationClasses} chat-message" 
         role="${isTyping ? 'status' : 'log'}" 
         aria-live="${isTyping ? 'polite' : 'off'}" 
         aria-label="${isUser ? 'User message' : `${agentName} message`}">
      <div class="max-w-xs lg:max-w-md px-4 py-2 rounded-lg ${messageClasses} ${AppState.reducedMotion ? '' : 'transition-all duration-200'}">
        ${isUser ? '' : `<div class="flex items-center space-x-2 text-xs font-medium ${isTyping ? 'text-yellow-600' : isError ? 'text-red-600' : 'text-gray-500'} mb-1">
          <span class="${getAgentTransitionClass(agentType, msg.isTransition)}">${agentIcon}</span>
          <span>${agentName}</span>
          ${msg.step ? `<span class="text-gray-400">• ${msg.step}</span>` : ''}
          ${msg.progress ? `<div class="ml-2 flex-1">
            <div class="bg-gray-300 rounded-full h-1.5 w-16">
              <div class="bg-blue-500 h-1.5 rounded-full ${AppState.reducedMotion ? '' : 'transition-all duration-300'}" style="width: ${msg.progress}%"></div>
            </div>
          </div>` : ''}
        </div>`}
        <div class="text-sm">${isTyping && !content.includes('...') ? content + `<span class="typing-indicator ml-2 ${AppState.reducedMotion ? '' : 'animate-pulse'}"><span>•</span><span>•</span><span>•</span></span>` : content}</div>
        ${structuredDataHtml}
        ${buttonsHtml}
        ${msg.taskStatus ? `<div class="mt-2 text-xs text-gray-600">Tasks: ${msg.taskStatus.completed}/${msg.taskStatus.total} completed</div>` : ''}
        <div class="text-xs ${isUser ? 'text-blue-100' : (isTyping ? 'text-yellow-600' : isError ? 'text-red-500' : 'text-gray-500')} mt-1">${new Date(msg.timestamp).toLocaleTimeString()}</div>
      </div>
    </div>`;
}

export function drawChat() {
  const el = document.getElementById('chat-messages');
  if (!el) return;
  el.innerHTML = AppState.messages.map(renderChatMessage).join('');
  el.scrollTop = el.scrollHeight;
}

export async function onSend() {
  const input = document.getElementById('message-input');
  const text = input.value.trim();
  if (!text) return;

  const userMsg = { role: 'user', content: text, timestamp: Date.now() };
  AppState.messages.push(userMsg);
  drawChat();
  input.value = '';

  // Create abort controller for cancellation
  window.abortController = new AbortController();
  const signal = window.abortController.signal;

  // Show stop button
  showStopButton();

  // Enhanced typing message with step indicators and progress
  const updateTypingMessage = (content, agentType, step, progress, isTransition = false) => {
    const lastMsg = AppState.messages[AppState.messages.length - 1];
    if (lastMsg && lastMsg.isTyping) {
      lastMsg.content = content;
      lastMsg.agentType = agentType;
      lastMsg.step = step;
      lastMsg.progress = progress;
      lastMsg.isTransition = isTransition;
      drawChat();
    }
  };
  
  const typingMsg = { 
    role: 'assistant', 
    content: 'Analyzing your request...', 
    timestamp: Date.now(), 
    isTyping: true, 
    agentType: 'intent', 
    step: '1/4 steps',
    progress: 0
  };
  AppState.messages.push(typingMsg);
  drawChat();

  try {
    // Step 1: Intent Analysis
    updateTypingMessage('Analyzing your intent...', 'intent', '1/4 steps', 25);
    const intentResult = await runIntentAgent(text);
    
    if (signal.aborted) throw new Error('Process cancelled by user');
    
    AppState.messages = AppState.messages.filter(m => !m.isTyping);

    // If no tasks needed, provide conversational response
    if (!intentResult.needsTasks) {
      const response = intentResult.response || 'I understand. Is there anything you\'d like me to help you with in your spreadsheet?';
      const aiMsg = { role: 'assistant', content: response, timestamp: Date.now() };
      AppState.messages.push(aiMsg);
      drawChat();
      hideStopButton();
      return;
    }

    // Step 2: Task Planning  
    updateTypingMessage('Breaking down your request into tasks...', 'planner', '2/4 steps', 50, true);
    const tasks = await runPlanner(text);
    
    if (signal.aborted) throw new Error('Process cancelled by user');
    
    AppState.messages = AppState.messages.filter(m => !m.isTyping);

    if (tasks && tasks.length) {
      AppState.tasks.push(...tasks);
      await saveTasks();
      drawTasks();

      let responseContent = `✅ I've analyzed your request and created ${tasks.length} task(s).`;
      let buttons = [];
      
      if (AppState.autoExecute) {
        responseContent += ` I will now execute them automatically.`;
      } else {
        responseContent += `\n\nYou can execute tasks individually or all at once using the buttons below.`;
        
        // Add interactive buttons
        buttons = [
          {
            label: 'Execute All Tasks',
            action: 'executeAllTasks',
            type: 'primary',
            icon: '▶️'
          },
          {
            label: 'View Tasks',
            action: 'viewTasks',
            type: 'secondary',
            icon: '📋'
          }
        ];

        // Add preview button if tasks modify data
        const hasDataModifications = tasks.some(t => 
          t.description?.includes('update') || 
          t.description?.includes('modify') || 
          t.description?.includes('change')
        );
        
        if (hasDataModifications) {
          buttons.push({
            label: 'Preview Changes',
            action: 'viewDiff',
            data: JSON.stringify({
              changes: tasks.map(t => ({ 
                cell: 'A1', 
                oldValue: 'Current', 
                newValue: 'Updated',
                task: t.title 
              }))
            }),
            type: 'secondary',
            icon: '👁️'
          });
        }
      }
      
      const aiMsg = { 
        role: 'assistant', 
        content: responseContent, 
        timestamp: Date.now(),
        buttons: buttons,
        taskStatus: {
          total: tasks.length,
          completed: 0,
          pending: tasks.length
        }
      };
      AppState.messages.push(aiMsg);
      drawChat();

      if (AppState.autoExecute) {
        // Step 3: Orchestration
        updateTypingMessage('Planning task execution order...', 'orchestrator', '3/4 steps', 75, true);
        let orchestration = null;
        try {
          orchestration = await runOrchestrator(tasks);
          if (signal.aborted) throw new Error('Process cancelled by user');
        } catch (e) {
          console.error('Orchestrator error:', e);
        }

        // Prepare tasks to execute: prefer orchestrated order if available
        let tasksToExecute = tasks;
        if (orchestration && Array.isArray(orchestration.executionPlan)) {
          const fromPlan = orchestration.executionPlan
            .sort((a, b) => (a.order || 0) - (b.order || 0))
            .map(p => AppState.tasks.find(t => t.id === p.taskId))
            .filter(Boolean);
          if (fromPlan.length) tasksToExecute = fromPlan;
        }

        // Step 4: Execution
        updateTypingMessage('Executing tasks...', 'executor', '4/4 steps', 100, true);
        setTimeout(() => {
          AppState.messages = AppState.messages.filter(m => !m.isTyping);
          // Pass orchestration so executeTasks can optionally skip re-orchestrating
          executeTasks(tasksToExecute, orchestration);
          hideStopButton();
        }, 300);
      } else {
        hideStopButton();
      }
    } else {
      // Fallback for simple commands that don't generate tasks
      const singleTask = {
        id: 'task-' + Date.now(),
        title: text,
        description: 'Single task execution',
        status: 'pending',
        createdAt: new Date().toISOString()
      };
      AppState.tasks.push(singleTask);
      await saveTasks();
      drawTasks();
      await executeTask(singleTask.id);
    }
  } catch (error) {
    AppState.messages = AppState.messages.filter(m => !m.isTyping);
    hideStopButton();
    
    let errorContent = `❌ I encountered an error processing your request: ${error.message}`;
    let errorButtons = [];
    
    // Add specific error recovery buttons based on error type
    if (error.message.includes('API key') || error.message.includes('401') || error.message.includes('403')) {
      errorContent += `\n\nThis appears to be an API authentication issue. Please check your API keys.`;
      errorButtons.push({
        label: 'Open API Settings',
        action: 'openApiKeys',
        type: 'primary',
        icon: '🔑'
      });
    } else if (error.message.includes('cancelled')) {
      errorContent = `⏹️ Process cancelled by user.`;
      errorButtons.push({
        label: 'Try Again',
        action: 'retryLastRequest',
        type: 'secondary',
        icon: '🔄'
      });
    } else {
      errorContent += `\n\nYou can try again or check the browser console for more details.`;
      errorButtons.push(
        {
          label: 'Retry',
          action: 'retryLastRequest', 
          type: 'primary',
          icon: '🔄'
        },
        {
          label: 'Check API Keys',
          action: 'openApiKeys',
          type: 'secondary',
          icon: '🔑'
        }
      );
    }
    
    const errorMsg = {
      role: 'assistant',
      content: errorContent,
      timestamp: Date.now(),
      type: 'error',
      buttons: errorButtons
    };
    AppState.messages.push(errorMsg);
    drawChat();
    console.error('Chat error:', error);
    showToast('Chat processing failed', 'error');
  }
}