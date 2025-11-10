import React, { useState } from 'react';
import { TEST_CASES, validateDiffs, generateTestReport } from './tests';
import { computeWordLevelDiff } from './diff';

/**
 * Test Runner Component
 * Runs automated tests on the diff algorithm
 */
export default function TestRunner() {
  const [testResults, setTestResults] = useState(null);
  const [isRunning, setIsRunning] = useState(false);
  const [selectedPriority, setSelectedPriority] = useState('all');

  const runTests = () => {
    setIsRunning(true);
    setTestResults(null);

    // Small delay to let UI update
    setTimeout(() => {
      const results = [];
      const filteredTests = selectedPriority === 'all' 
        ? TEST_CASES 
        : TEST_CASES.filter(tc => tc.priority === selectedPriority);

      for (const testCase of filteredTests) {
        try {
          const actualDiffs = computeWordLevelDiff(testCase.oldText, testCase.newText);
          const validation = validateDiffs(actualDiffs, testCase.expectedDiffs);

          results.push({
            id: testCase.id,
            name: testCase.name,
            priority: testCase.priority,
            passed: validation.passed,
            message: validation.message,
            notes: testCase.notes,
            oldText: testCase.oldText,
            newText: testCase.newText,
            actualDiffs,
            expectedDiffs: testCase.expectedDiffs
          });
        } catch (error) {
          results.push({
            id: testCase.id,
            name: testCase.name,
            priority: testCase.priority,
            passed: false,
            message: `Error: ${error.message}`,
            notes: testCase.notes,
            error: true
          });
        }
      }

      setTestResults(results);
      setIsRunning(false);
    }, 100);
  };

  const getStats = () => {
    if (!testResults) return null;

    const total = testResults.length;
    const passed = testResults.filter(r => r.passed).length;
    const failed = testResults.filter(r => !r.passed).length;
    const highPriorityFailed = testResults.filter(r => !r.passed && r.priority === 'high').length;

    return { total, passed, failed, highPriorityFailed };
  };

  const stats = getStats();

  return (
    <div style={{ padding: 12 }}>
      <h2 style={{ margin: 0 }}>Test Suite</h2>
      <p style={{ color: '#555', marginTop: 4 }}>
        Automated tests for diff algorithm (Unicode, duplicates, single chars, etc.)
      </p>

      <div style={{ display: 'flex', gap: 8, marginTop: 12, marginBottom: 12, alignItems: 'center' }}>
        <label style={{ fontWeight: 600 }}>Priority Filter:</label>
        <select 
          value={selectedPriority} 
          onChange={e => setSelectedPriority(e.target.value)}
          style={{ padding: '4px 8px' }}
        >
          <option value="all">All Tests</option>
          <option value="high">High Priority</option>
          <option value="medium">Medium Priority</option>
          <option value="low">Low Priority</option>
        </select>

        <button 
          onClick={runTests} 
          disabled={isRunning}
          style={{ 
            padding: '6px 12px', 
            fontWeight: 600,
            background: '#0078d4',
            color: 'white',
            border: 'none',
            borderRadius: 4,
            cursor: isRunning ? 'not-allowed' : 'pointer'
          }}
        >
          {isRunning ? 'Running Tests...' : 'Run Tests'}
        </button>
      </div>

      {stats && (
        <div style={{ 
          padding: 12, 
          background: stats.failed === 0 ? '#e3fcef' : '#ffebe6',
          borderRadius: 6,
          marginBottom: 12
        }}>
          <div style={{ fontWeight: 600, marginBottom: 8 }}>Test Results</div>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: 8 }}>
            <div>Total: {stats.total}</div>
            <div style={{ color: '#006644' }}>Passed: {stats.passed} ?</div>
            <div style={{ color: '#bf2600' }}>Failed: {stats.failed} ?</div>
            <div style={{ color: '#bf2600' }}>High Priority Failed: {stats.highPriorityFailed}</div>
          </div>
        </div>
      )}

      {testResults && (
        <div style={{ 
          maxHeight: '400px', 
          overflowY: 'auto',
          border: '1px solid #ddd',
          borderRadius: 6
        }}>
          {testResults.map((result, idx) => (
            <div 
              key={result.id} 
              style={{ 
                padding: 12, 
                borderBottom: idx < testResults.length - 1 ? '1px solid #eee' : 'none',
                background: result.passed ? '#f9f9f9' : '#fff5f5'
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, marginBottom: 4 }}>
                <span style={{ fontSize: 18 }}>
                  {result.passed ? '?' : '?'}
                </span>
                <span style={{ fontWeight: 600 }}>
                  {result.name}
                </span>
                <span style={{ 
                  fontSize: 11, 
                  padding: '2px 6px', 
                  background: result.priority === 'high' ? '#ffebe6' : result.priority === 'medium' ? '#fff4e6' : '#f0f0f0',
                  borderRadius: 3
                }}>
                  {result.priority.toUpperCase()}
                </span>
              </div>

              <div style={{ fontSize: 12, color: '#666', marginLeft: 26 }}>
                {result.id}
              </div>

              {!result.passed && (
                <div style={{ 
                  marginTop: 8, 
                  marginLeft: 26,
                  padding: 8,
                  background: '#fff',
                  border: '1px solid #ffccc7',
                  borderRadius: 4,
                  fontSize: 12
                }}>
                  <div style={{ fontWeight: 600, color: '#bf2600', marginBottom: 4 }}>
                    Error:
                  </div>
                  <div style={{ whiteSpace: 'pre-wrap', fontFamily: 'monospace' }}>
                    {result.message}
                  </div>
                </div>
              )}

              <details style={{ marginTop: 8, marginLeft: 26, fontSize: 12 }}>
                <summary style={{ cursor: 'pointer', color: '#0078d4' }}>
                  Show Details
                </summary>
                <div style={{ marginTop: 8, padding: 8, background: '#f9f9f9', borderRadius: 4 }}>
                  <div><strong>Notes:</strong> {result.notes}</div>
                  {result.oldText && (
                    <>
                      <div style={{ marginTop: 8 }}>
                        <strong>Old Text:</strong>
                        <pre style={{ 
                          background: '#fff', 
                          padding: 4, 
                          borderRadius: 3,
                          fontSize: 11,
                          overflow: 'auto'
                        }}>
                          {result.oldText}
                        </pre>
                      </div>
                      <div style={{ marginTop: 8 }}>
                        <strong>New Text:</strong>
                        <pre style={{ 
                          background: '#fff', 
                          padding: 4, 
                          borderRadius: 3,
                          fontSize: 11,
                          overflow: 'auto'
                        }}>
                          {result.newText}
                        </pre>
                      </div>
                    </>
                  )}
                  {result.actualDiffs && (
                    <div style={{ marginTop: 8 }}>
                      <strong>Actual Diffs:</strong>
                      <pre style={{ 
                        background: '#fff', 
                        padding: 4, 
                        borderRadius: 3,
                        fontSize: 11,
                        overflow: 'auto'
                      }}>
                        {JSON.stringify(result.actualDiffs, null, 2)}
                      </pre>
                    </div>
                  )}
                </div>
              </details>
            </div>
          ))}
        </div>
      )}

      {!testResults && !isRunning && (
        <div style={{ 
          padding: 24, 
          textAlign: 'center', 
          color: '#999',
          border: '2px dashed #ddd',
          borderRadius: 6
        }}>
          Click "Run Tests" to start automated testing
        </div>
      )}
    </div>
  );
}

