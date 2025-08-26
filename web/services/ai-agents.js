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

export async function runPlanner(userText) {
  const provider = pickProvider();
  const tasks = [];

  try {
    if (provider === 'mock') {
      tasks.push({ id: uuid(), title: 'Insert header row', description: 'Add Name, Age, Email', status: 'pending', context: { range: 'A1:C1', sheet: AppState.activeSheet }, createdAt: new Date().toISOString() });
      return tasks;
    }

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
  if (provider === 'mock') {
    return {
      edits: [
        { op: 'setCell', sheet: AppState.activeSheet, cell: 'A1', value: 'Total' },
        { op: 'setRange', sheet: AppState.activeSheet, range: 'A2:C3', values: [['a', 1, 2], ['b', 3, 4]] }
      ],
      export: null,
      message: `Mock applied 2 edits for ${task.title}`
    };
  }

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

  // Basic schema validation first
  const basicResult = { valid: true, errors: [], warnings: [] };
  if (!executorObj || !Array.isArray(executorObj.edits)) {
    basicResult.valid = false;
    basicResult.errors.push('Missing edits array');
    return basicResult;
  }

  const supportedOps = ['setCell', 'setRange', 'setFormula', 'insertRow', 'deleteRow', 'insertColumn', 'deleteColumn', 'formatCell', 'formatRange'];
  for (const e of executorObj.edits) {
    if (!e.op) {
      basicResult.valid = false;
      basicResult.errors.push('Edit missing operation type');
      break;
    }
    if (supportedOps.indexOf(e.op) === -1) {
      basicResult.valid = false;
      basicResult.errors.push(`Unsupported operation: ${e.op}`);
      break;
    }
  }

  if (!basicResult.valid) return basicResult;

  // Advanced AI-powered validation
  if (provider === 'mock') {
    return {
      valid: true,
      confidence: 0.8,
      analysis: 'Mock validation - basic schema checks passed',
      risks: [],
      recommendations: [],
      dataIntegrityScore: 0.9
    };
  }

  try {
    const ws = getWorksheet();
    const sheetContext = ws['!ref'] ? `Sheet "${AppState.activeSheet}" range: ${ws['!ref']}` : `Empty sheet "${AppState.activeSheet}"`;
    const sampleData = ws['!ref'] ? getSampleDataFromSheet(ws) : 'No data';

    const system = `You are the Validator Agent - an expert in data integrity, conflict detection, and intelligent validation of spreadsheet operations.

ROLE: Analyze planned operations for potential conflicts, data integrity issues, and optimization opportunities while ensuring user intent is preserved.

CAPABILITIES:
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
1. Analyze data integrity and potential conflicts
2. Validate formula references and dependencies
3. Assess performance impact and optimization opportunities
4. Check data type consistency and formatting
5. Verify alignment with user intent and task goals
6. Identify potential risks and provide recommendations

REQUIRED OUTPUT FORMAT:
{
  "valid": true,
  "confidence": 0.95,
  "analysis": "Detailed analysis of the planned operations and their impact",
  "dataIntegrityScore": 0.9,
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
  "warnings": ["Large dataset may impact browser performance"]
}

VALIDATION CRITERIA:
- Data integrity and consistency preservation
- Formula reference validity and dependency management
- Performance impact on current dataset size
- Alignment with original user request and task goals
- Potential for data loss or corruption
- Reversibility and rollback complexity

INTELLIGENCE FEATURES:
- Context-aware conflict detection
- Performance impact prediction
- User intent analysis and preservation
- Advanced risk assessment with mitigation strategies
- Optimization recommendations for efficiency`;

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

    // Fallback to basic validation
    return {
      valid: true,
      confidence: 0.7,
      analysis: 'AI validation failed, using basic schema validation',
      warnings: ['Advanced validation unavailable']
    };

  } catch (error) {
    console.error('Validator failed:', error);
    return {
      valid: true, // Don't block on validator failure
      confidence: 0.5,
      analysis: `Validation error: ${error.message}`,
      warnings: ['Validator agent unavailable - proceeding with basic validation only']
    };
  }
}