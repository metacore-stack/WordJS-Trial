/**
 * Comprehensive Test Suite for Word Diff Application
 * 
 * Based on specifications for handling:
 * - Duplicate/repeated text
 * - Single character operations
 * - Unicode normalization (NFC)
 * - Whitespace and newlines
 * - Grapheme clusters
 * - Boundary conditions
 * 
 * To run these tests manually in Word:
 * 1. Insert the old text into document
 * 2. Select it
 * 3. Run "Preview Changes" with the corresponding new text
 * 4. Verify the expected result
 */

export const TEST_CASES = [
  // ========== PRIORITY 1: DUPLICATE TEXT (Most Critical) ==========
  {
    id: 'duplicate-01',
    name: 'Repeated words - delete middle occurrence',
    priority: 'high',
    oldText: 'apple apple apple apple',
    newText: 'apple apple apple',
    expectedDiffs: [
      { op: 'equal', text: 'apple apple ' },
      { op: 'delete', text: 'apple ' },
      { op: 'equal', text: 'apple' }
    ],
    notes: 'Tests marker-based deletion to ensure correct instance is deleted'
  },
  {
    id: 'duplicate-02',
    name: 'Repeated words - delete first occurrence',
    priority: 'high',
    oldText: 'the the the document',
    newText: 'the the document',
    expectedDiffs: [
      { op: 'delete', text: 'the ' },
      { op: 'equal', text: 'the the document' }
    ],
    notes: 'Tests left-to-right deletion priority'
  },
  {
    id: 'duplicate-03',
    name: 'Repeated words - insert between duplicates',
    priority: 'high',
    oldText: 'hello hello hello',
    newText: 'hello NEW hello hello',
    expectedDiffs: [
      { op: 'equal', text: 'hello ' },
      { op: 'insert', text: 'NEW ' },
      { op: 'equal', text: 'hello hello' }
    ],
    notes: 'Tests insertion between repeated text'
  },
  {
    id: 'duplicate-04',
    name: 'Short repeated tokens',
    priority: 'high',
    oldText: 'a a a a a',
    newText: 'a a X a a',
    expectedDiffs: [
      { op: 'equal', text: 'a a ' },
      { op: 'insert', text: 'X ' },
      { op: 'equal', text: 'a a' }
    ],
    notes: 'Tests single-character repeated tokens with context'
  },

  // ========== PRIORITY 2: SINGLE CHARACTER OPERATIONS ==========
  {
    id: 'single-char-01',
    name: 'Single character deletion',
    priority: 'high',
    oldText: 'abcdef',
    newText: 'abdef',
    expectedDiffs: [
      { op: 'equal', text: 'ab' },
      { op: 'delete', text: 'c' },
      { op: 'equal', text: 'def' }
    ],
    notes: 'Tests single grapheme deletion with run splitting'
  },
  {
    id: 'single-char-02',
    name: 'Single character insertion',
    priority: 'high',
    oldText: 'abdef',
    newText: 'abcdef',
    expectedDiffs: [
      { op: 'equal', text: 'ab' },
      { op: 'insert', text: 'c' },
      { op: 'equal', text: 'def' }
    ],
    notes: 'Tests single grapheme insertion with xml:space="preserve"'
  },
  {
    id: 'single-char-03',
    name: 'Single space deletion',
    priority: 'high',
    oldText: 'hello  world',
    newText: 'hello world',
    expectedDiffs: [
      { op: 'equal', text: 'hello ' },
      { op: 'delete', text: ' ' },
      { op: 'equal', text: 'world' }
    ],
    notes: 'Tests single space handling with xml:space preservation'
  },
  {
    id: 'single-char-04',
    name: 'Single punctuation',
    priority: 'high',
    oldText: 'hello world',
    newText: 'hello, world',
    expectedDiffs: [
      { op: 'equal', text: 'hello' },
      { op: 'insert', text: ',' },
      { op: 'equal', text: ' world' }
    ],
    notes: 'Tests single punctuation insertion'
  },

  // ========== PRIORITY 3: UNICODE AND COMBINING CHARACTERS ==========
  {
    id: 'unicode-01',
    name: 'Combining character (é composed)',
    priority: 'high',
    oldText: 'café',
    newText: 'cafe',
    expectedDiffs: [
      { op: 'equal', text: 'caf' },
      { op: 'delete', text: 'é' },
      { op: 'insert', text: 'e' }
    ],
    notes: 'Tests NFC normalization with composed é'
  },
  {
    id: 'unicode-02',
    name: 'Combining character (é decomposed)',
    priority: 'high',
    oldText: 'cafe\u0301', // e + combining acute accent
    newText: 'cafe',
    expectedDiffs: [
      { op: 'equal', text: 'caf' },
      { op: 'delete', text: 'é' }
    ],
    notes: 'Tests NFC normalization with decomposed é (e + combining)'
  },
  {
    id: 'unicode-03',
    name: 'Full-width vs half-width',
    priority: 'medium',
    oldText: 'Hello', // Full-width
    newText: 'Hello',      // Half-width
    expectedDiffs: [
      { op: 'delete', text: 'Hello' },
      { op: 'insert', text: 'Hello' }
    ],
    notes: 'Tests full-width/half-width handling (may need normalization option)'
  },
  {
    id: 'unicode-04',
    name: 'Emoji (grapheme cluster)',
    priority: 'medium',
    oldText: 'Hello ??????????? world',
    newText: 'Hello ?? world',
    expectedDiffs: [
      { op: 'equal', text: 'Hello ' },
      { op: 'delete', text: '???????????' },
      { op: 'insert', text: '??' },
      { op: 'equal', text: ' world' }
    ],
    notes: 'Tests complex emoji as single grapheme clusters'
  },

  // ========== PRIORITY 4: WHITESPACE AND NEWLINES ==========
  {
    id: 'whitespace-01',
    name: 'Multiple spaces',
    priority: 'high',
    oldText: 'hello    world',
    newText: 'hello world',
    expectedDiffs: [
      { op: 'equal', text: 'hello ' },
      { op: 'delete', text: '   ' },
      { op: 'equal', text: 'world' }
    ],
    notes: 'Tests multiple consecutive spaces with xml:space'
  },
  {
    id: 'whitespace-02',
    name: 'Line breaks',
    priority: 'high',
    oldText: 'line1\n\nline2',
    newText: 'line1\nline2',
    expectedDiffs: [
      { op: 'equal', text: 'line1\n' },
      { op: 'delete', text: '\n' },
      { op: 'equal', text: 'line2' }
    ],
    notes: 'Tests newline normalization (\\r\\n -> \\n)'
  },
  {
    id: 'whitespace-03',
    name: 'Leading and trailing spaces',
    priority: 'high',
    oldText: ' hello world ',
    newText: 'hello world',
    expectedDiffs: [
      { op: 'delete', text: ' ' },
      { op: 'equal', text: 'hello world' },
      { op: 'delete', text: ' ' }
    ],
    notes: 'Tests leading/trailing space handling'
  },
  {
    id: 'whitespace-04',
    name: 'Tab characters',
    priority: 'medium',
    oldText: 'hello\tworld',
    newText: 'hello world',
    expectedDiffs: [
      { op: 'equal', text: 'hello' },
      { op: 'delete', text: '\t' },
      { op: 'insert', text: ' ' },
      { op: 'equal', text: 'world' }
    ],
    notes: 'Tests tab character handling'
  },

  // ========== PRIORITY 5: BOUNDARY CONDITIONS ==========
  {
    id: 'boundary-01',
    name: 'Insert at beginning',
    priority: 'medium',
    oldText: 'world',
    newText: 'Hello world',
    expectedDiffs: [
      { op: 'insert', text: 'Hello ' },
      { op: 'equal', text: 'world' }
    ],
    notes: 'Tests insertion at document start'
  },
  {
    id: 'boundary-02',
    name: 'Insert at end',
    priority: 'medium',
    oldText: 'Hello',
    newText: 'Hello world',
    expectedDiffs: [
      { op: 'equal', text: 'Hello' },
      { op: 'insert', text: ' world' }
    ],
    notes: 'Tests insertion at document end'
  },
  {
    id: 'boundary-03',
    name: 'Delete entire text',
    priority: 'medium',
    oldText: 'Delete me',
    newText: '',
    expectedDiffs: [
      { op: 'delete', text: 'Delete me' }
    ],
    notes: 'Tests complete text deletion'
  },
  {
    id: 'boundary-04',
    name: 'Insert into empty',
    priority: 'medium',
    oldText: '',
    newText: 'Hello world',
    expectedDiffs: [
      { op: 'insert', text: 'Hello world' }
    ],
    notes: 'Tests insertion into empty document'
  },
  {
    id: 'boundary-05',
    name: 'No change',
    priority: 'low',
    oldText: 'Same text',
    newText: 'Same text',
    expectedDiffs: [
      { op: 'equal', text: 'Same text' }
    ],
    notes: 'Tests no-op (identical text)'
  },

  // ========== PRIORITY 6: ADJACENT OPERATIONS ==========
  {
    id: 'adjacent-01',
    name: 'Adjacent delete and insert',
    priority: 'medium',
    oldText: 'abc',
    newText: 'aBc',
    expectedDiffs: [
      { op: 'equal', text: 'a' },
      { op: 'delete', text: 'b' },
      { op: 'insert', text: 'B' },
      { op: 'equal', text: 'c' }
    ],
    notes: 'Tests adjacent delete+insert (replacement)'
  },
  {
    id: 'adjacent-02',
    name: 'Multiple adjacent inserts',
    priority: 'medium',
    oldText: 'ac',
    newText: 'abc',
    expectedDiffs: [
      { op: 'equal', text: 'a' },
      { op: 'insert', text: 'b' },
      { op: 'equal', text: 'c' }
    ],
    notes: 'Tests insert in middle'
  },

  // ========== PRIORITY 7: LARGE/COMPLEX OPERATIONS ==========
  {
    id: 'large-01',
    name: 'Large text with multiple changes',
    priority: 'medium',
    oldText: 'The quick brown fox jumps over the lazy dog. The quick brown fox jumps over the lazy dog.',
    newText: 'The fast brown fox leaps over the sleepy dog. The fast brown fox leaps over the sleepy dog.',
    expectedDiffs: [
      { op: 'equal', text: 'The ' },
      { op: 'delete', text: 'quick' },
      { op: 'insert', text: 'fast' },
      { op: 'equal', text: ' brown fox ' },
      { op: 'delete', text: 'jumps' },
      { op: 'insert', text: 'leaps' },
      { op: 'equal', text: ' over the ' },
      { op: 'delete', text: 'lazy' },
      { op: 'insert', text: 'sleepy' },
      { op: 'equal', text: ' dog. The ' },
      { op: 'delete', text: 'quick' },
      { op: 'insert', text: 'fast' },
      { op: 'equal', text: ' brown fox ' },
      { op: 'delete', text: 'jumps' },
      { op: 'insert', text: 'leaps' },
      { op: 'equal', text: ' over the ' },
      { op: 'delete', text: 'lazy' },
      { op: 'insert', text: 'sleepy' },
      { op: 'equal', text: ' dog.' }
    ],
    notes: 'Tests handling of repeated pattern changes in longer text'
  },

  // ========== PRIORITY 8: SHORT TEXT (High Ambiguity) ==========
  {
    id: 'short-01',
    name: 'Very short repeated token',
    priority: 'high',
    oldText: 'I I I',
    newText: 'I X I',
    expectedDiffs: [
      { op: 'equal', text: 'I ' },
      { op: 'delete', text: 'I' },
      { op: 'insert', text: 'X' },
      { op: 'equal', text: ' I' }
    ],
    notes: 'Tests disambiguation with very short (1 char) tokens'
  },
  {
    id: 'short-02',
    name: 'Repeated short words',
    priority: 'high',
    oldText: 'to to to',
    newText: 'to for to',
    expectedDiffs: [
      { op: 'equal', text: 'to ' },
      { op: 'delete', text: 'to' },
      { op: 'insert', text: 'for' },
      { op: 'equal', text: ' to' }
    ],
    notes: 'Tests short word replacement'
  },

  // ========== REAL-WORLD EXAMPLES ==========
  {
    id: 'realworld-01',
    name: 'Patent text example (from spec)',
    priority: 'high',
    oldText: '[0075] FIG. 1 illustrates an illustration of a platform',
    newText: '[0075] FIG. 1 illustrates a platform',
    expectedDiffs: [
      { op: 'equal', text: '[0075] FIG. 1 illustrates ' },
      { op: 'delete', text: 'an illustration of ' },
      { op: 'equal', text: 'a platform' }
    ],
    notes: 'Real-world patent text deletion'
  },
  {
    id: 'realworld-02',
    name: 'Sentence restructuring',
    priority: 'medium',
    oldText: 'The platform (100-for-facilitating) may facilitate electronic signing.',
    newText: 'The platform (100) facilitates electronic signing.',
    expectedDiffs: [
      { op: 'equal', text: 'The platform (100' },
      { op: 'delete', text: '-for-facilitating' },
      { op: 'equal', text: ') ' },
      { op: 'delete', text: 'may facilitate' },
      { op: 'insert', text: 'facilitates' },
      { op: 'equal', text: ' electronic signing.' }
    ],
    notes: 'Real-world sentence simplification'
  },

  // ========== CRITICAL EDGE CASES FROM SPEC ==========
  
  {
    id: 'edge-formatted-01',
    name: 'Run-split deletion (formatted text)',
    priority: 'high',
    oldText: 'hello world',
    newText: 'helo world',
    expectedDiffs: [
      { op: 'equal', text: 'hel' },
      { op: 'delete', text: 'l' },
      { op: 'equal', text: 'o world' }
    ],
    notes: 'Delete single char that may be in middle of formatted run - requires run splitting'
  },
  
  {
    id: 'edge-repeated-delete-01',
    name: 'Delete first of many repeated words',
    priority: 'high',
    oldText: 'foo foo foo foo',
    newText: 'foo foo foo',
    expectedDiffs: [
      { op: 'delete', text: 'foo ' },
      { op: 'equal', text: 'foo foo foo' }
    ],
    notes: 'Test left-to-right consumption - delete FIRST occurrence'
  },
  
  {
    id: 'edge-repeated-delete-02',
    name: 'Delete last of many repeated words',
    priority: 'high',
    oldText: 'bar bar bar bar',
    newText: 'bar bar bar',
    expectedDiffs: [
      { op: 'equal', text: 'bar bar bar ' },
      { op: 'delete', text: 'bar' }
    ],
    notes: 'Test left-to-right consumption - delete LAST occurrence'
  },
  
  {
    id: 'edge-single-char-repeated',
    name: 'Single character repeated many times',
    priority: 'high',
    oldText: 'x x x x x',
    newText: 'x x y x x',
    expectedDiffs: [
      { op: 'equal', text: 'x x ' },
      { op: 'delete', text: 'x' },
      { op: 'insert', text: 'y' },
      { op: 'equal', text: ' x x' }
    ],
    notes: 'CRITICAL: Single char with high ambiguity - requires token mapping or markers'
  },
  
  {
    id: 'edge-overlapping-01',
    name: 'Adjacent delete and insert (replacement)',
    priority: 'high',
    oldText: 'quick',
    newText: 'fast',
    expectedDiffs: [
      { op: 'delete', text: 'quick' },
      { op: 'insert', text: 'fast' }
    ],
    notes: 'Tests handling of adjacent delete+insert to avoid overlapping revisions'
  },
  
  {
    id: 'edge-paragraph-01',
    name: 'Cross-paragraph deletion',
    priority: 'medium',
    oldText: 'Line one.\n\nLine two.',
    newText: 'Line one.\nLine two.',
    expectedDiffs: [
      { op: 'equal', text: 'Line one.\n' },
      { op: 'delete', text: '\n' },
      { op: 'equal', text: 'Line two.' }
    ],
    notes: 'Paragraph boundary - ensure proper paragraph handling'
  },
  
  {
    id: 'edge-whitespace-preserve',
    name: 'Leading and trailing space preservation',
    priority: 'high',
    oldText: '  text  ',
    newText: ' text ',
    expectedDiffs: [
      { op: 'delete', text: ' ' },
      { op: 'equal', text: ' text ' },
      { op: 'delete', text: ' ' }
    ],
    notes: 'Requires xml:space="preserve" - critical for spaces'
  },
  
  {
    id: 'edge-zero-width-01',
    name: 'Zero-width characters',
    priority: 'medium',
    oldText: 'hello\u200Bworld',  // Zero-width space
    newText: 'helloworld',
    expectedDiffs: [
      { op: 'equal', text: 'hello' },
      { op: 'delete', text: '\u200B' },
      { op: 'equal', text: 'world' }
    ],
    notes: 'Zero-width space handling - Unicode edge case'
  },
  
  {
    id: 'stress-rapid-changes',
    name: 'Many small rapid changes',
    priority: 'medium',
    oldText: 'a b c d e f g h i j',
    newText: 'A B C D E F G H I J',
    expectedDiffs: [
      { op: 'delete', text: 'a' }, { op: 'insert', text: 'A' }, { op: 'equal', text: ' ' },
      { op: 'delete', text: 'b' }, { op: 'insert', text: 'B' }, { op: 'equal', text: ' ' },
      { op: 'delete', text: 'c' }, { op: 'insert', text: 'C' }, { op: 'equal', text: ' ' },
      { op: 'delete', text: 'd' }, { op: 'insert', text: 'D' }, { op: 'equal', text: ' ' },
      { op: 'delete', text: 'e' }, { op: 'insert', text: 'E' }, { op: 'equal', text: ' ' },
      { op: 'delete', text: 'f' }, { op: 'insert', text: 'F' }, { op: 'equal', text: ' ' },
      { op: 'delete', text: 'g' }, { op: 'insert', text: 'G' }, { op: 'equal', text: ' ' },
      { op: 'delete', text: 'h' }, { op: 'insert', text: 'H' }, { op: 'equal', text: ' ' },
      { op: 'delete', text: 'i' }, { op: 'insert', text: 'I' }, { op: 'equal', text: ' ' },
      { op: 'delete', text: 'j' }, { op: 'insert', text: 'J' }
    ],
    notes: 'Stress test for race conditions and timing - many sequential operations'
  }
];

/**
 * Special test categories for systematic validation
 */
export const TEST_CATEGORIES = {
  DUPLICATE_TEXT: TEST_CASES.filter(t => t.id.startsWith('duplicate-')),
  SINGLE_CHAR: TEST_CASES.filter(t => t.id.startsWith('single-char-')),
  UNICODE: TEST_CASES.filter(t => t.id.startsWith('unicode-')),
  WHITESPACE: TEST_CASES.filter(t => t.id.startsWith('whitespace-')),
  BOUNDARY: TEST_CASES.filter(t => t.id.startsWith('boundary-')),
  EDGE_CASES: TEST_CASES.filter(t => t.id.startsWith('edge-')),
  STRESS: TEST_CASES.filter(t => t.id.startsWith('stress-')),
  REAL_WORLD: TEST_CASES.filter(t => t.id.startsWith('realworld-'))
};

/**
 * Validation helper to check if diffs match expected
 */
export function validateDiffs(actualDiffs, expectedDiffs) {
  if (actualDiffs.length !== expectedDiffs.length) {
    return {
      passed: false,
      message: `Expected ${expectedDiffs.length} diffs, got ${actualDiffs.length}`
    };
  }

  for (let i = 0; i < actualDiffs.length; i++) {
    const actual = actualDiffs[i];
    const expected = expectedDiffs[i];

    if (actual.op !== expected.op) {
      return {
        passed: false,
        message: `Diff ${i}: Expected op '${expected.op}', got '${actual.op}'`
      };
    }

    // Normalize text before comparison
    const normalizedActual = actual.text.normalize('NFC');
    const normalizedExpected = expected.text.normalize('NFC');

    if (normalizedActual !== normalizedExpected) {
      return {
        passed: false,
        message: `Diff ${i}: Text mismatch\nExpected: "${normalizedExpected}"\nGot: "${normalizedActual}"`
      };
    }
  }

  return {
    passed: true,
    message: 'All diffs match expected'
  };
}

/**
 * Run all tests
 */
export function runAllTests(diffFunction) {
  const results = [];
  
  for (const testCase of TEST_CASES) {
    const actualDiffs = diffFunction(testCase.oldText, testCase.newText);
    const validation = validateDiffs(actualDiffs, testCase.expectedDiffs);
    
    results.push({
      id: testCase.id,
      name: testCase.name,
      priority: testCase.priority,
      passed: validation.passed,
      message: validation.message,
      notes: testCase.notes
    });
  }
  
  return results;
}

/**
 * Generate test report
 */
export function generateTestReport(results) {
  const total = results.length;
  const passed = results.filter(r => r.passed).length;
  const failed = results.filter(r => !r.passed).length;
  const highPriorityFailed = results.filter(r => !r.passed && r.priority === 'high').length;

  let report = '\n========== TEST REPORT ==========\n';
  report += `Total: ${total}\n`;
  report += `Passed: ${passed} ?\n`;
  report += `Failed: ${failed} ?\n`;
  report += `High Priority Failed: ${highPriorityFailed}\n`;
  report += '================================\n\n';

  if (failed > 0) {
    report += 'FAILED TESTS:\n';
    results.filter(r => !r.passed).forEach(r => {
      report += `\n[${r.priority.toUpperCase()}] ${r.id}: ${r.name}\n`;
      report += `  ${r.message}\n`;
      report += `  Notes: ${r.notes}\n`;
    });
  }

  return report;
}

export default TEST_CASES;

