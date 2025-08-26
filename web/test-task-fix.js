/* Test script to verify task processing fix */

// Create a simple test task that will demonstrate the fix
async function createTestTransaction() {
  console.log('🧪 Creating test transaction...');
  
  // Mock task that would be created by AI agents
  const testTask = {
    id: 'test-task-' + Date.now(),
    title: 'Sample Transaction: Add Basic Data',
    description: 'Test task to verify applyEditsOrDryRun function works',
    status: 'pending',
    createdAt: new Date().toISOString(),
    context: {
      sheet: 'Sheet1',
      operation: 'test'
    },
    maxRetries: 3
  };

  // Mock executor result that would be generated
  const mockExecutorResult = {
    edits: [
      {
        op: 'setCell',
        cell: 'A1',
        value: 'Test Item',
        dataType: 'string'
      },
      {
        op: 'setCell',
        cell: 'B1',
        value: '100',
        dataType: 'number'
      },
      {
        op: 'setFormula',
        cell: 'C1',
        formula: '=B1*2'
      }
    ],
    validation: {
      confidence: 0.95,
      dataIntegrityScore: 0.9
    }
  };

  // Mock validator result (should validate successfully)
  const mockValidatorResult = {
    valid: true,
    confidence: 0.95,
    analysis: 'Test operations are valid and safe to execute',
    errors: []
  };

  console.log('📝 Test task:', testTask);
  console.log('⚙️ Mock executor result:', mockExecutorResult);
  console.log('✅ Mock validator result:', mockValidatorResult);
  
  return {
    task: testTask,
    executorResult: mockExecutorResult,
    validatorResult: mockValidatorResult
  };
}

// Test the applyEditsOrDryRun function directly
async function testApplyEditsFunction() {
  console.log('🔧 Testing applyEditsOrDryRun function...');
  
  const testData = await createTestTransaction();
  
  try {
    // Test if function exists globally
    if (typeof window.applyEditsOrDryRun === 'function') {
      console.log('✅ applyEditsOrDryRun function is available globally');
      
      // Test dry run mode
      console.log('🔍 Testing dry run mode...');
      window.AppState.dryRun = true;
      await window.applyEditsOrDryRun(testData.executorResult);
      console.log('✅ Dry run test completed');
      
      // Test actual execution
      console.log('⚡ Testing actual execution...');
      window.AppState.dryRun = false;
      await window.applyEditsOrDryRun(testData.executorResult);
      console.log('✅ Actual execution test completed');
      
      return { success: true, message: 'All tests passed!' };
    } else {
      console.error('❌ applyEditsOrDryRun function not found');
      return { success: false, message: 'Function not available' };
    }
  } catch (error) {
    console.error('❌ Test failed:', error);
    return { success: false, message: error.message };
  }
}

// Test complete task execution flow
async function testTaskExecution() {
  console.log('🚀 Testing complete task execution flow...');
  
  const testData = await createTestTransaction();
  
  // Add task to AppState
  if (window.AppState && Array.isArray(window.AppState.tasks)) {
    window.AppState.tasks.push(testData.task);
    console.log('📋 Test task added to AppState');
    
    // Try to execute the task using the window.executeTask function
    if (typeof window.executeTask === 'function') {
      console.log('⚡ Attempting to execute test task...');
      try {
        await window.executeTask(testData.task.id);
        console.log('✅ Task execution completed successfully!');
        return { success: true, message: 'Task execution successful' };
      } catch (error) {
        console.error('❌ Task execution failed:', error);
        return { success: false, message: `Task execution failed: ${error.message}` };
      }
    } else {
      console.error('❌ executeTask function not found');
      return { success: false, message: 'executeTask function not available' };
    }
  } else {
    console.error('❌ AppState.tasks not available');
    return { success: false, message: 'AppState.tasks not available' };
  }
}

// Main test function
async function runAllTests() {
  console.log('🎯 Starting comprehensive task processing tests...');
  
  const results = [];
  
  // Test 1: Function availability
  results.push(await testApplyEditsFunction());
  
  // Test 2: Complete task execution
  results.push(await testTaskExecution());
  
  // Summary
  const passed = results.filter(r => r.success).length;
  const total = results.length;
  
  console.log(`📊 Test Summary: ${passed}/${total} tests passed`);
  
  if (passed === total) {
    console.log('🎉 All tests passed! Task processing fix is working correctly.');
  } else {
    console.log('⚠️ Some tests failed. Review the issues above.');
  }
  
  return { passed, total, results };
}

// Expose functions to global scope for manual testing
window.testTaskProcessing = {
  createTestTransaction,
  testApplyEditsFunction, 
  testTaskExecution,
  runAllTests
};

console.log('🧪 Task processing test utilities loaded. Use window.testTaskProcessing.runAllTests() to run tests.');