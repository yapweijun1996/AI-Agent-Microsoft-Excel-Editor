import { AppState } from '../core/state.js';
import { escapeHtml } from '../utils/index.js';
import { runPlanner } from '../services/ai-agents.js';
import { saveTasks, drawTasks, executeTasks, runOrchestrator } from '../tasks/task-manager.js';
import { showToast } from '../ui/toast.js';
/* global executeTask */

function renderChatMessage(msg) {
  const isUser = msg.role === 'user';
  const isTyping = msg.isTyping || false;

  let content = escapeHtml(msg.content);
  if (!isUser) {
    content = content.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
    content = content.replace(/\n/g, '<br>');
  }

  return `
    <div class="flex ${isUser ? 'justify-end' : 'justify-start'} ${isTyping ? 'animate-pulse' : ''}">
      <div class="max-w-xs lg:max-w-md px-4 py-2 rounded-lg ${isUser ? 'bg-blue-500 text-white' : (isTyping ? 'bg-yellow-100 text-yellow-800' : 'bg-gray-200 text-gray-900')}">
        ${isUser ? '' : `<div class="text-xs font-medium ${isTyping ? 'text-yellow-600' : 'text-gray-500'} mb-1">${isTyping ? '🤖 AI Agents' : 'AI Assistant'}</div>`}
        <div class="text-sm">${content}</div>
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

  const typingMsg = { role: 'assistant', content: '🤔 AI agents are planning your request...', timestamp: Date.now(), isTyping: true };
  AppState.messages.push(typingMsg);
  drawChat();

  try {
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
        responseContent += `\n\n🎯 Click the execute button on each task to run them, or use "Execute All" for orchestrated execution.`;
      }
      
      const aiMsg = { role: 'assistant', content: responseContent, timestamp: Date.now() };
      AppState.messages.push(aiMsg);
      drawChat();

      if (AppState.autoExecute) {
        const orchestration = await runOrchestrator(tasks);
        executeTasks(orchestration.executionPlan.map(p => AppState.tasks.find(t => t.id === p.taskId)));
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