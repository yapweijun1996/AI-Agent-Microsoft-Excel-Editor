import { AppState } from '../core/state.js';
import { escapeHtml } from '../utils/index.js';
import { runPlanner, runIntentAgent } from '../services/ai-agents.js';
import { saveTasks, drawTasks, executeTasks, runOrchestrator } from '../tasks/task-manager.js';
import { showToast } from '../ui/toast.js';
/* global executeTask */

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

  return `
    <div class="flex ${isUser ? 'justify-end' : 'justify-start'} ${isTyping ? 'animate-pulse' : ''} chat-message">
      <div class="max-w-xs lg:max-w-md px-4 py-2 rounded-lg ${isUser ? 'bg-blue-500 text-white' : (isTyping ? 'bg-yellow-100 text-yellow-800 border-l-4 border-yellow-400' : 'bg-gray-200 text-gray-900')} transition-all duration-200">
        ${isUser ? '' : `<div class="flex items-center space-x-2 text-xs font-medium ${isTyping ? 'text-yellow-600' : 'text-gray-500'} mb-1">
          <span>${agentIcon}</span>
          <span>${agentName}</span>
          ${msg.step ? `<span class="text-gray-400">• ${msg.step}</span>` : ''}
        </div>`}
        <div class="text-sm">${isTyping && !content.includes('...') ? content + '<span class="typing-indicator ml-2"><span>•</span><span>•</span><span>•</span></span>' : content}</div>
        <div class="text-xs ${isUser ? 'text-blue-100' : (isTyping ? 'text-yellow-600' : 'text-gray-500')} mt-1">${new Date(msg.timestamp).toLocaleTimeString()}</div>
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

  // Enhanced typing message with step indicators
  const updateTypingMessage = (content, agentType, step) => {
    const lastMsg = AppState.messages[AppState.messages.length - 1];
    if (lastMsg && lastMsg.isTyping) {
      lastMsg.content = content;
      lastMsg.agentType = agentType;
      lastMsg.step = step;
      drawChat();
    }
  };
  
  const typingMsg = { role: 'assistant', content: 'Analyzing your request...', timestamp: Date.now(), isTyping: true, agentType: 'intent', step: '1/4 steps' };
  AppState.messages.push(typingMsg);
  drawChat();

  try {
    // Step 1: Intent Analysis
    updateTypingMessage('Analyzing your intent...', 'intent', '1/4 steps');
    const intentResult = await runIntentAgent(text);
    AppState.messages = AppState.messages.filter(m => !m.isTyping);

    // If no tasks needed, provide conversational response
    if (!intentResult.needsTasks) {
      const response = intentResult.response || 'I understand. Is there anything you\'d like me to help you with in your spreadsheet?';
      const aiMsg = { role: 'assistant', content: response, timestamp: Date.now() };
      AppState.messages.push(aiMsg);
      drawChat();
      return;
    }

    // Step 2: Task Planning  
    updateTypingMessage('Breaking down your request into tasks...', 'planner', '2/4 steps');
    const tasks = await runPlanner(text);
    AppState.messages = AppState.messages.filter(m => !m.isTyping);

    if (tasks && tasks.length) {
      AppState.tasks.push(...tasks);
      await saveTasks();
      drawTasks();

      let responseContent = `✅ I've analyzed your request and created ${tasks.length} task(s).`;
      if (AppState.autoExecute) {
        responseContent += ` I will now execute them automatically.`;
      } else {
        responseContent += `\n\n- Click the ▶️ button on each task to run them individually\n- Use the "Execute All" button for orchestrated execution\n- Toggle auto-execute in the task panel for automatic processing`;
      }
      
      const aiMsg = { role: 'assistant', content: responseContent, timestamp: Date.now() };
      AppState.messages.push(aiMsg);
      drawChat();

      if (AppState.autoExecute) {
        // Step 3: Orchestration
        updateTypingMessage('Planning task execution order...', 'orchestrator', '3/4 steps');
        let orchestration = null;
        try {
          orchestration = await runOrchestrator(tasks);
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
        updateTypingMessage('Executing tasks...', 'executor', '4/4 steps');
        setTimeout(() => {
          AppState.messages = AppState.messages.filter(m => !m.isTyping);
          // Pass orchestration so executeTasks can optionally skip re-orchestrating
          executeTasks(tasksToExecute, orchestration);
        }, 300);
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
    const errorMsg = {
      role: 'assistant',
      content: `❌ I encountered an error processing your request: ${error.message}\n\nPlease check your API keys and try again.`,
      timestamp: Date.now()
    };
    AppState.messages.push(errorMsg);
    drawChat();
    console.error('Chat error:', error);
    showToast('Chat processing failed', 'error');
  }
}