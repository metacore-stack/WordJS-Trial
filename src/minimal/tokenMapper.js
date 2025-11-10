/**
 * Token-based mapping system for precise diff application
 * Implements the specification for handling repeated tokens
 */

/**
 * Character-level tokenizer that matches diff-match-patch behavior
 * Returns array of tokens where each token is a single character or small chunk
 * This matches how diff-match-patch works (character-level diffing)
 */
export function tokenizeForMapping(text) {
  if (!text) return [];
  
  // For diff-match-patch compatibility, we tokenize character by character
  // but group consecutive characters that form "words" for efficiency
  const tokens = [];
  let currentToken = '';
  let i = 0;
  
  while (i < text.length) {
    const char = text[i];
    currentToken += char;
    i++;
    
    // Group characters into tokens (we can adjust chunk size)
    // For now, tokenize character by character to match diff-match-patch exactly
    tokens.push({
      text: char,
      startIndex: i - 1,
      endIndex: i,
      globalIndex: tokens.length
    });
  }
  
  return tokens;
}

/**
 * Build token stream from diffs
 * Each token in the stream represents a character or chunk from the diff
 */
export function buildDiffTokenStream(diffs) {
  const tokenStream = [];
  let globalTokenIndex = 0;
  
  for (const diff of diffs) {
    if (!diff || !diff.text) continue;
    
    const tokens = tokenizeForMapping(diff.text);
    for (const token of tokens) {
      tokenStream.push({
        op: diff.op, // 'equal', 'delete', or 'insert'
        text: token.text,
        startIndex: token.startIndex,
        endIndex: token.endIndex,
        diffGlobalIndex: globalTokenIndex++,
        mapped: null // Will be filled by mapping algorithm
      });
    }
  }
  
  return tokenStream;
}

/**
 * Parse OOXML to extract runs and their text
 * Returns array of run objects with text and position info
 */
export function parseOoxmlRuns(ooxml) {
  const runs = [];
  
  if (!ooxml || typeof ooxml !== 'string') {
    return runs;
  }
  
  try {
    // Simple XML parsing for runs
    // Match <w:r>...</w:r> blocks
    const runRegex = /<w:r[^>]*>(.*?)<\/w:r>/gs;
    let match;
    let runId = 0;
    let globalCharIndex = 0;
    
    while ((match = runRegex.exec(ooxml)) !== null) {
      const runXml = match[0];
      const runContent = match[1];
      
      // Extract text from <w:t> tags
      const textRegex = /<w:t[^>]*xml:space="preserve"[^>]*>(.*?)<\/w:t>/gs;
      const textMatches = [...runContent.matchAll(textRegex)];
      
      let runText = '';
      for (const textMatch of textMatches) {
        // Unescape XML entities
        const text = textMatch[1]
          .replace(/&lt;/g, '<')
          .replace(/&gt;/g, '>')
          .replace(/&amp;/g, '&')
          .replace(/&quot;/g, '"')
          .replace(/&apos;/g, "'");
        runText += text;
      }
      
      // Extract run properties
      const rPrMatch = runContent.match(/<w:rPr[^>]*>(.*?)<\/w:rPr>/s);
      const rPrXml = rPrMatch ? rPrMatch[0] : '';
      
      if (runText || runXml.includes('<w:br') || runXml.includes('<w:tab')) {
        runs.push({
          runId: runId++,
          rawXml: runXml,
          text: runText,
          rPrXml: rPrXml,
          startCharIndex: globalCharIndex,
          endCharIndex: globalCharIndex + runText.length,
          isComplex: runXml.includes('<w:hyperlink') || runXml.includes('<w:fldChar') || runXml.includes('<w:drawing')
        });
        
        globalCharIndex += runText.length;
      }
    }
  } catch (error) {
    console.warn('Error parsing OOXML runs:', error);
  }
  
  return runs;
}

/**
 * Build runTokenMap from parsed runs
 * Each token maps to a specific run and position
 */
export function buildRunTokenMap(runs) {
  const runTokenMap = [];
  let globalTokenIndex = 0;
  
  for (const run of runs) {
    const tokens = tokenizeForMapping(run.text);
    
    for (const token of tokens) {
      runTokenMap.push({
        token: token.text,
        runId: run.runId,
        runTokenIndex: token.startIndex,
        tokenStartOffsetInRun: token.startIndex,
        tokenEndOffsetInRun: token.endIndex,
        globalTokenIndex: globalTokenIndex++,
        run: run // Reference to full run object
      });
    }
  }
  
  return runTokenMap;
}

/**
 * Build occurrence map for each distinct token
 */
export function buildOccurrenceMap(runTokenMap) {
  const occurrenceMap = {};
  
  for (const tokenEntry of runTokenMap) {
    const token = tokenEntry.token;
    if (!occurrenceMap[token]) {
      occurrenceMap[token] = [];
    }
    occurrenceMap[token].push(tokenEntry.globalTokenIndex);
  }
  
  return occurrenceMap;
}

/**
 * Single-pass consume mapping algorithm
 * Maps each diff token to exact run/token position
 */
export function mapDiffTokensToRuns(diffTokenStream, runTokenMap) {
  let cursor = 0; // Points to next unconsumed token in runTokenMap
  const mappedDiffs = [];
  
  for (const diffToken of diffTokenStream) {
    if (diffToken.op === 'equal' || diffToken.op === 'delete') {
      // Must consume from runTokenMap
      if (cursor >= runTokenMap.length) {
        throw new Error(`Token mapping mismatch: cursor ${cursor} exceeds runTokenMap length ${runTokenMap.length}`);
      }
      
      const runToken = runTokenMap[cursor];
      
      // Validate token matches
      if (runToken.token !== diffToken.text) {
        // Allow some tolerance for whitespace/formatting differences
        // But log warning
        console.warn(`Token mismatch at cursor ${cursor}: expected "${runToken.token}", got "${diffToken.text}"`);
      }
      
      // Map the diff token to the run token
      diffToken.mapped = {
        runId: runToken.runId,
        runTokenIndex: runToken.runTokenIndex,
        tokenStartOffsetInRun: runToken.tokenStartOffsetInRun,
        tokenEndOffsetInRun: runToken.tokenEndOffsetInRun,
        globalTokenIndex: runToken.globalTokenIndex,
        run: runToken.run
      };
      
      cursor++;
    } else if (diffToken.op === 'insert') {
      // Insert tokens don't consume from runTokenMap
      // They map to insertion point between cursor-1 and cursor
      diffToken.insertionPoint = {
        beforeRunId: cursor < runTokenMap.length ? runTokenMap[cursor].runId : null,
        offsetInRun: cursor < runTokenMap.length ? runTokenMap[cursor].tokenStartOffsetInRun : null,
        afterGlobalIndex: cursor > 0 ? runTokenMap[cursor - 1].globalTokenIndex : -1
      };
    }
    
    mappedDiffs.push(diffToken);
  }
  
  // Validation: check if all original tokens were consumed
  if (cursor < runTokenMap.length) {
    const remaining = runTokenMap.length - cursor;
    console.warn(`Not all tokens consumed: ${remaining} tokens remaining`);
  }
  
  return mappedDiffs;
}

/**
 * Validate token alignment
 * Checks that original tokens from diffs match runTokenMap
 */
export function validateTokenAlignment(diffTokenStream, runTokenMap) {
  const originalTokens = diffTokenStream
    .filter(dt => dt.op === 'equal' || dt.op === 'delete')
    .map(dt => dt.text);
  
  const runTokens = runTokenMap.map(rt => rt.token);
  
  if (originalTokens.length !== runTokens.length) {
    return {
      valid: false,
      error: `Token count mismatch: diffs have ${originalTokens.length} original tokens, runs have ${runTokens.length}`
    };
  }
  
  for (let i = 0; i < originalTokens.length; i++) {
    if (originalTokens[i] !== runTokens[i]) {
      return {
        valid: false,
        error: `Token mismatch at index ${i}: diff has "${originalTokens[i]}", run has "${runTokens[i]}"`
      };
    }
  }
  
  return { valid: true };
}

