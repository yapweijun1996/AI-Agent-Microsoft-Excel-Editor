'use strict';

import { AppState } from '../core/state.js';
import { log, getSampleDataFromSheet, extractFirstJson, uuid } from '../utils/index.js';
import { pickProvider, getSelectedModel } from './api-keys.js';
import { getWorksheet } from '../spreadsheet/workbook-manager.js';
import { showToast } from '../ui/toast.js';

export async function fetchOpenAI(apiKey, messages, model = 'gpt-4o') {
  const res = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${apiKey}` },
    body: JSON.stringify({ model, messages })
  });
  return res.json();
}

export async function fetchGemini(apiKey, messages, model = 'gemini-2.5-flash') {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      contents: messages.map(m => ({ role: m.role === 'assistant' ? 'model' : 'user', parts: [{ text: m.content }] }))
    })
  });
  return res.json();
}

export async function runIntentAgent(userText) {
  const provider = pickProvider();
  
  try {

    const system = `You are the Intent Agent - an expert at analyzing user input to determine whether it requires spreadsheet task planning or is a conversational message.

ROLE: Classify user input and determine the appropriate response type.

CLASSIFICATION TYPES:
1. "spreadsheet_operation" - User wants to perform spreadsheet tasks (create, edit, analyze, calculate, format, etc.)
2. "greeting" - Simple greetings, pleasantries, or casual conversation
3. "question" - Questions about the application, spreadsheet features, or general help
4. "clarification" - User asking for clarification or providing additional context

CURRENT CONTEXT:
- Active sheet: "${AppState.activeSheet}"
- Total sheets: ${AppState.wb.SheetNames.length}
- This is a spreadsheet application with AI automation

ANALYSIS CRITERIA:
- Does the input contain action words related to spreadsheet operations? (add, create, calculate, format, delete, sort, filter, etc.)
- Is the user requesting data manipulation or analysis?
- Is it just a greeting or casual conversation?
- Is the user asking questions about functionality?

OUTPUT FORMAT (JSON):
{
  "needsTasks": true/false,
  "intent": "spreadsheet_operation|greeting|question|clarification", 
  "confidence": 0.95,
  "reasoning": "Brief explanation of classification",
  "response": "Optional conversational response for non-task intents"
}

EXAMPLES:
- "hi" → {"needsTasks": false, "intent": "greeting", "confidence": 0.99, "reasoning": "Simple greeting", "response": "Hello! How can I help you with your spreadsheet today?"}
- "add a sum formula" → {"needsTasks": true, "intent": "spreadsheet_operation", "confidence": 0.95, "reasoning": "Clear spreadsheet operation request"}
- "how do I save?" → {"needsTasks": false, "intent": "question", "confidence": 0.9, "reasoning": "Question about application functionality", "response": "You can save by pressing Ctrl+S or using the File menu."}`;

    const messages = [{ role: 'system', content: system }, { role: 'user', content: userText }];
    let data;

    const selectedModel = getSelectedModel();
    if (provider === 'openai') {
      data = await fetchOpenAI(AppState.keys.openai, messages, selectedModel);
    } else {
      data = await fetchGemini(AppState.keys.gemini, messages, selectedModel);
    }

    let text = '';
    try {
      if (provider === 'openai') {
        text = data.choices?.[0]?.message?.content || '';
        if (!text && data.error) {
          throw new Error(data.error.message || 'OpenAI API error');
        }
      } else {
        text = data.candidates?.[0]?.content?.parts?.map(p => p.text).join('') || '';
        if (!text && data.error) {
          throw new Error(data.error.message || 'Gemini API error');
        }
      }
    } catch (parseError) {
      console.error('Failed to parse Intent Agent response:', parseError);
      // Default to needing tasks on parse error
      return { needsTasks: true, intent: 'spreadsheet_operation', confidence: 0.5 };
    }

    if (!text) {
      return { needsTasks: true, intent: 'spreadsheet_operation', confidence: 0.5 };
    }

    let result = null;
    try {
      result = JSON.parse(text);
    } catch {
      result = extractFirstJson(text);
    }

    if (result && typeof result.needsTasks === 'boolean') {
      return result;
    } else {
      // If LLM response is not parseable, default to needing tasks with low confidence
      console.warn('Intent Agent: LLM response not parseable, defaulting to spreadsheet_operation');
      return { 
        needsTasks: true, 
        intent: 'spreadsheet_operation', 
        confidence: 0.3,
        reasoning: 'LLM response parsing failed - defaulting to task-based processing'
      };
    }
  } catch (error) {
    console.error('Intent Agent failed:', error);
    // Default to needing tasks on error
    return { needsTasks: true, intent: 'spreadsheet_operation', confidence: 0.5 };
  }
}

export async function runPlanner(userText) {
  const provider = pickProvider();
  const tasks = [];

  try {

    // Get current sheet context for better planning
    const ws = getWorksheet();
    const sheetContext = ws['!ref'] ? `Current sheet "${AppState.activeSheet}" range: ${ws['!ref']}` : `Empty sheet "${AppState.activeSheet}"`;
    const sampleData = ws['!ref'] ? getSampleDataFromSheet(ws) : 'No data';

    const system = `You are the Planner Agent - an expert at analyzing user requests and breaking them down into precise, executable tasks for spreadsheet automation.

ROLE: Decompose complex spreadsheet operations into logical, sequential tasks that can be executed by specialized agents.

CAPABILITIES:
- Analyze natural language requests for spreadsheet operations
- Understand data patterns, relationships, and structure
- Plan multi-step workflows with dependencies
- Consider data validation and error handling needs
- Optimize task sequencing for efficiency

CURRENT CONTEXT:
- Active sheet: "${AppState.activeSheet}"
- Sheet structure: ${sheetContext}
- Sample data preview:
${sampleData}
- Available sheets: [${AppState.wb.SheetNames.join(', ')}]
- Total sheets: ${AppState.wb.SheetNames.length}

TASK BREAKDOWN STRATEGY:
1. Analyze the user request for complexity and dependencies
2. Identify required data operations (create, read, update, delete)
3. Consider data validation and formatting requirements
4. Plan for potential errors or edge cases
5. Sequence tasks logically with clear dependencies

OUTPUT FORMAT: Return a JSON array of task objects. Each task must include:
- "id": unique identifier
- "title": brief descriptive title
- "description": detailed operation description
- "priority": number 1-5 (1=highest)
- "dependencies": array of task IDs that must complete first
- "context": {"range": "A1:C10", "sheet": "SheetName", "operation": "type"}
- "validation": expected outcome or validation criteria

EXAMPLES:
For "Add totals row with formulas":
[
  {"id":"task1", "title":"Detect data range", "description":"Find the extent of existing data", "priority":1, "dependencies":[], "context":{"range":"detect", "sheet":"${AppState.activeSheet}", "operation":"analyze"}, "validation":"Data range identified"},
  {"id":"task2", "title":"Insert totals row", "description":"Add row below data for totals", "priority":2, "dependencies":["task1"], "context":{"range":"below_data", "sheet":"${AppState.activeSheet}", "operation":"insertRow"}, "validation":"Row inserted successfully"},
  {"id":"task3", "title":"Add SUM formulas", "description":"Create SUM formulas for numeric columns", "priority":3, "dependencies":["task2"], "context":{"range":"totals_row", "sheet":"${AppState.activeSheet}", "operation":"setFormula"}, "validation":"Formulas calculate correctly"}
]

IMPORTANT: 
- Always consider data integrity and user intent
- Plan for edge cases (empty data, invalid formats, etc.)
- Keep tasks atomic and focused
- Ensure proper sequencing with dependencies
- Include validation criteria for each task`;

    const messages = [{ role: 'system', content: system }, { role: 'user', content: userText }];
    let data;

    try {
      const selectedModel = getSelectedModel();
      if (provider === 'openai') {
        data = await fetchOpenAI(AppState.keys.openai, messages, selectedModel);
      } else {
        data = await fetchGemini(AppState.keys.gemini, messages, selectedModel);
      }
    } catch (apiError) {
      console.error('API call failed:', apiError);
      showToast(`${provider} API call failed. Check your API key and internet connection.`, 'error');
      return [];
    }

    let text = '';
    try {
      if (provider === 'openai') {
        text = data.choices?.[0]?.message?.content || '';
        if (!text && data.error) {
          throw new Error(data.error.message || 'OpenAI API error');
        }
      } else {
        text = data.candidates?.[0]?.content?.parts?.map(p => p.text).join('') || '';
        if (!text && data.error) {
          throw new Error(data.error.message || 'Gemini API error');
        }
      }
    } catch (parseError) {
      console.error('Failed to parse API response:', parseError);
      showToast('Failed to parse AI response', 'error');
      return [];
    }

    if (!text) {
      showToast('AI returned empty response', 'warning');
      return [];
    }

    let arr = null;
    try {
      arr = JSON.parse(text);
    } catch {
      arr = extractFirstJson(text);
    }

    if (Array.isArray(arr)) {
      return arr.map(t => ({
        id: t.id || uuid(),
        title: t.title || (t.description || 'Task'),
        description: t.description || '',
        status: 'pending',
        priority: t.priority || 3,
        dependencies: t.dependencies || [],
        context: { ...t.context, sheet: t.context?.sheet || AppState.activeSheet },
        validation: t.validation || null,
        createdAt: new Date().toISOString(),
        estimatedDuration: t.estimatedDuration || null,
        retryCount: 0,
        maxRetries: 3
      }));
    } else {
      showToast('AI response was not in expected format', 'warning');
      return [];
    }
  } catch (error) {
    console.error('Planner failed:', error);
    showToast('Planning failed: ' + error.message, 'error');
    return [];
  }
}

export async function runExecutor(task) {
  const provider = pickProvider();

  // Get current sheet context
  const ws = getWorksheet();
  const sheetContext = ws['!ref'] ? `Sheet "${AppState.activeSheet}" range: ${ws['!ref']}` : `Empty sheet "${AppState.activeSheet}"`;
  const sampleData = ws['!ref'] ? getSampleDataFromSheet(ws) : 'No data';

  const system = `You are the Executor Agent - a specialist in translating planned tasks into precise spreadsheet operations with intelligent analysis and error handling.

ROLE: Execute planned tasks by analyzing current spreadsheet state and generating optimal operation sequences.

CAPABILITIES:
- Analyze spreadsheet data patterns and structure
- Generate precise SheetJS-compatible operations
- Handle complex data transformations and calculations
- Implement intelligent error handling and rollback strategies
- Optimize operations for performance and data integrity

CURRENT CONTEXT:
- Active sheet: "${AppState.activeSheet}"
- Sheet structure: ${sheetContext}
- Sample data preview:
${sampleData}
- Available sheets: [${AppState.wb.SheetNames.join(', ')}]

EXECUTION STRATEGY:
1. Analyze current data structure and patterns
2. Determine optimal operation sequence
3. Consider data types and formatting requirements
4. Plan for edge cases and error conditions
5. Generate atomic, reversible operations

OPERATION SCHEMA (REQUIRED OUTPUT FORMAT):
{
  "success": true,
  "analysis": "Brief analysis of current state and planned changes",
  "edits": [
    {"op":"setCell","sheet":"SheetName","cell":"A1","value":"Total","dataType":"string"},
    {"op":"setRange","sheet":"SheetName","range":"A2:C3","values":[["a",1,2],["b",3,4]],"preserveTypes":true},
    {"op":"setFormula","sheet":"SheetName","cell":"D1","formula":"=SUM(A:A)"},
    {"op":"insertRow","sheet":"SheetName","row":2,"count":1},
    {"op":"deleteRow","sheet":"SheetName","row":2,"count":1},
    {"op":"insertColumn","sheet":"SheetName","col":"B","count":1},
    {"op":"deleteColumn","sheet":"SheetName","col":"B","count":1},
    {"op":"formatCell","sheet":"SheetName","cell":"A1","format":"0.00"},
    {"op":"formatRange","sheet":"SheetName","range":"A1:C3","format":"General"}
  ],
  "validation": {
    "expectedChanges": ["description of expected changes"],
    "rollbackPlan": ["steps to undo if needed"],
    "dataIntegrityChecks": ["validation points to verify"]
  },
  "warnings": ["any potential issues or considerations"],
  "message": "Detailed description of what was accomplished"
}

INTELLIGENT FEATURES:
- Auto-detect data types (numbers, dates, text, formulas)
- Preserve existing formatting where appropriate
- Handle formula dependencies and references
- Optimize range operations for efficiency
- Provide detailed rollback plans for safety

IMPORTANT: 
- Always analyze before executing
- Preserve data integrity and user intent
- Generate atomic, reversible operations
- Include comprehensive validation plans
- Handle edge cases gracefully`;

  const user = `Task: ${task.title}\nDescription: ${task.description || ''}\nContext: ${JSON.stringify(task.context || {})}`;
  const messages = [{ role: 'system', content: system }, { role: 'user', content: user }];
  let data;
  const selectedModel = getSelectedModel();
  if (provider === 'openai') { data = await fetchOpenAI(AppState.keys.openai, messages, selectedModel); }
  else { data = await fetchGemini(AppState.keys.gemini, messages, selectedModel); }
  let text = '';
  try {
    if (provider === 'openai') { text = data.choices?.[0]?.message?.content || ''; }
    else { text = data.candidates?.[0]?.content?.parts?.map(p => p.text).join('') || ''; }
  } catch { text = ''; }
  let obj = null;
  try { obj = JSON.parse(text); } catch { obj = extractFirstJson(text); }
  log('Executor raw', text);
  return obj;
}

export async function runValidator(executorObj, task) {
  const provider = pickProvider();

  // Quick null check before LLM validation
  if (!executorObj) {
    return {
      valid: false,
      confidence: 1.0,
      analysis: 'Executor result is null or undefined',
      errors: ['Missing executor result'],
      warnings: []
    };
  }

  try {
    const ws = getWorksheet();
    const sheetContext = ws['!ref'] ? `Sheet "${AppState.activeSheet}" range: ${ws['!ref']}` : `Empty sheet "${AppState.activeSheet}"`;
    const sampleData = ws['!ref'] ? getSampleDataFromSheet(ws) : 'No data';

    const system = `You are the Validator Agent - an expert in data integrity, conflict detection, and intelligent validation of spreadsheet operations.

ROLE: Analyze planned operations for potential conflicts, data integrity issues, and optimization opportunities while ensuring user intent is preserved. You must perform ALL validation including schema validation, operation type validation, and advanced integrity checks.

CAPABILITIES:
- Complete schema and structure validation of executor results
- Deep data integrity analysis and conflict detection
- Formula dependency and reference validation  
- Performance impact assessment for large operations
- Data type consistency and format validation
- User intent preservation and goal alignment
- Risk assessment with confidence scoring

CURRENT CONTEXT:
- Active sheet: "${AppState.activeSheet}"
- Sheet structure: ${sheetContext}
- Sample data preview: ${sampleData}
- Available sheets: [${AppState.wb.SheetNames.join(', ')}]

VALIDATION STRATEGY:
1. FIRST: Validate the executor result structure and schema
   - Check that 'edits' array exists and is valid
   - Verify operation types are supported: ['setCell', 'setRange', 'setFormula', 'insertRow', 'deleteRow', 'insertColumn', 'deleteColumn', 'formatCell', 'formatRange']
   - Ensure required fields are present for each operation
2. THEN: Analyze data integrity and potential conflicts
3. Validate formula references and dependencies
4. Assess performance impact and optimization opportunities
5. Check data type consistency and formatting
6. Verify alignment with user intent and task goals
7. Identify potential risks and provide recommendations

REQUIRED OUTPUT FORMAT:
{
  "valid": true,
  "confidence": 0.95,
  "analysis": "Detailed analysis including schema validation and operation impact",
  "dataIntegrityScore": 0.9,
  "schemaValidation": {
    "editsArrayValid": true,
    "operationTypesValid": true,
    "requiredFieldsPresent": true,
    "errors": []
  },
  "risks": [
    {"level": "medium", "description": "Potential data overwrite", "mitigation": "Create backup"},
    {"level": "low", "description": "Performance impact on large dataset", "mitigation": "Use batch operations"}
  ],
  "conflicts": [
    {"type": "formula_reference", "description": "Formula may reference moved cells", "severity": "high"}
  ],
  "optimizations": [
    "Batch similar cell operations for better performance",
    "Use range operations instead of individual cell updates"
  ],
  "recommendations": [
    "Execute in dry-run mode first",
    "Consider creating a backup before major structural changes"
  ],
  "userIntentAlignment": 0.95,
  "expectedOutcome": "Operations will successfully add totals row with proper formulas",
  "rollbackComplexity": "low",
  "warnings": ["Large dataset may impact browser performance"],
  "errors": []
}

VALIDATION CRITERIA:
- CRITICAL: Schema validation of executor result structure
- CRITICAL: Operation type validation against supported operations  
- Data integrity and consistency preservation
- Formula reference validity and dependency management
- Performance impact on current dataset size
- Alignment with original user request and task goals
- Potential for data loss or corruption
- Reversibility and rollback complexity

INTELLIGENCE FEATURES:
- Complete schema and structure validation
- Context-aware conflict detection
- Performance impact prediction
- User intent analysis and preservation
- Advanced risk assessment with mitigation strategies
- Optimization recommendations for efficiency

IMPORTANT: If the executor result fails schema validation (missing edits array, unsupported operations, missing required fields), set valid: false and include detailed error information.`;

    const operations = {
      task: {
        id: task?.id,
        title: task?.title,
        description: task?.description,
        context: task?.context
      },
      executorResult: executorObj
    };

    const user = `Validate these planned operations:\n${JSON.stringify(operations, null, 2)}`;
    const messages = [{ role: 'system', content: system }, { role: 'user', content: user }];

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

    if (result && typeof result.valid === 'boolean') {
      return result;
    }

    // Fallback when LLM response is not parseable - default to valid with warnings
    console.warn('Validator Agent: LLM response not parseable, defaulting to valid with low confidence');
    return {
      valid: true,
      confidence: 0.3,
      analysis: 'LLM validation failed - response not parseable. Proceeding with caution.',
      warnings: ['LLM validation unavailable - proceeding without advanced validation'],
      errors: [],
      dataIntegrityScore: 0.5
    };

  } catch (error) {
    console.error('Validator Agent failed:', error);
    return {
      valid: true, // Don't block on validator failure - allow operation to proceed with warnings
      confidence: 0.2,
      analysis: `LLM validation failed due to error: ${error.message}. Proceeding without validation.`,
      warnings: ['LLM Validator Agent unavailable - operations proceeding without advanced validation'],
      errors: [],
      dataIntegrityScore: 0.3,
      risks: [{ level: 'high', description: 'Operating without validation due to LLM failure', mitigation: 'Manual review recommended' }]
    };
  }
}