// Use the correct diff-match-patch package
const DiffMatchPatch = require('diff-match-patch');

// Create an instance of the diff_match_patch object
const dmp = new DiffMatchPatch();

function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/\"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeXml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

/**
 * Advanced normalization for consistent newline handling
 * Normalizes all newline variations (\r\n, \r, \n) to \n for diffing
 */
function normalizeNewlines(text) {
  if (!text) return '';
  // Replace \r\n with \n first (to avoid double replacement)
  // Then replace remaining \r with \n
  return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
}

/**
 * Denormalize newlines back to Word format (\r)
 */
function denormalizeNewlines(text) {
  if (!text) return '';
  // Convert \n back to \r for Word compatibility
  return text.replace(/\n/g, '\r');
}

/**
 * Tokenize text into words, preserving whitespace and punctuation
 * Returns an array of tokens with their text and type (word, space, punctuation)
 */
function tokenizeText(text) {
  const tokens = [];
  let currentToken = '';
  let currentType = 'word'; //'word', 'space', 'punctuation'

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const isSpace = /\s/.test(char);
    const isPunct = /[^\w\s]/.test(char);
    const isWord = /[\w]/.test(char);

    if (isSpace) {
      if (currentToken && currentType !== 'space') {
        tokens.push({ text: currentToken, type: currentType });
        currentToken = '';
      }
      currentToken += char;
      currentType = 'space';
    } else if (isWord) {
      if (currentToken && currentType !== 'word') {
        tokens.push({ text: currentToken, type: currentType });
        currentToken = '';
      }
      currentToken += char;
      currentType = 'word';
    } else if (isPunct) {
      if (currentToken && currentType !== 'punctuation') {
        tokens.push({ text: currentToken, type: currentType });
        currentToken = '';
      }
      currentToken += char;
      currentType = 'punctuation';
    } else {
      currentToken += char;
    }
  }

  if(currentToken) {
    tokens.push({text: currentToken, type: currentType});
  }

  return tokens;
}

/**
 * Refine diff to handle repeated patterns correctly
 * Splits delete/insert operations when they contain repeated patterns
 */
function refineDiffForRepeatedPatterns(diffs, oldText, newText) {
  const refined = [];
  
  for (let i = 0; i < diffs.length; i++) {
    const [op, text] = diffs[i];
    
    if (op === 0) {
      // Equal text - keep as is
      refined.push([op, text]);
    } else if (op === -1) {
      // Deletion - check if we can split it
      const nextDiff = i + 1 < diffs.length ? diffs[i + 1] : null;
      
      if (nextDiff && nextDiff[0] === 1) {
        // We have a delete followed by insert - check if they contain repeated patterns
        const deletedText = text;
        const insertedText = nextDiff[1];
        
        // Try to find the best alignment by looking for common patterns
        const splitResult = splitDeleteInsertPair(deletedText, insertedText, oldText, newText);
        
        if (splitResult.length > 0) {
          // We can split this into multiple operations
          for (const part of splitResult) {
            refined.push(part);
          }
          i++; // Skip the next insert since we've processed it
        } else {
          // Can't split - keep both delete and insert as original
          refined.push([op, text]);
          refined.push([nextDiff[0], nextDiff[1]]);
          i++; // Skip the next insert since we've added it
        }
      } else {
        // No matching insert - keep as is
        refined.push([op, text]);
      }
    } else if (op === 1) {
      // Insertion - only process if not already handled by delete
      const prevDiff = i - 1 >= 0 ? diffs[i - 1] : null;
      if (!prevDiff || prevDiff[0] !== -1) {
        // This insert wasn't part of a delete-insert pair, keep as is
        refined.push([op, text]);
      }
      // Otherwise, it was already processed with the delete
    }
  }
  
  return refined;
}

/**
 * Split a delete-insert pair into smaller operations when patterns repeat
 * Uses diff-match-patch on the pair itself to find granular changes
 */
function splitDeleteInsertPair(deletedText, insertedText, oldText, newText) {
  // Strategy: Apply diff-match-patch to the delete-insert pair itself
  // This will identify common substrings and split the operations correctly
  
  // First, try to find if there are repeated patterns by looking for
  // common substrings that appear multiple times
  const minPatternLength = 3; // Minimum length for a pattern to be considered
  
  // Find all common substrings of sufficient length
  const commonSubstrings = [];
  for (let len = Math.min(deletedText.length, insertedText.length); len >= minPatternLength; len--) {
    for (let i = 0; i <= deletedText.length - len; i++) {
      const pattern = deletedText.substring(i, i + len);
      // Check if pattern exists in inserted text
      if (insertedText.includes(pattern)) {
        // Count occurrences in both strings
        const delCount = (deletedText.match(new RegExp(pattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g')) || []).length;
        const insCount = (insertedText.match(new RegExp(pattern.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g')) || []).length;
        
        // Accept pattern if:
        // 1. It appears multiple times in both (repeated pattern), OR
        // 2. It appears once in both and helps create a split (at least 2 parts)
        if ((delCount > 1 && insCount > 1 && delCount === insCount) ||
            (delCount === 1 && insCount === 1 && delCount === insCount)) {
          // Check if splitting by this pattern creates meaningful parts
          const delParts = deletedText.split(pattern);
          const insParts = insertedText.split(pattern);
          if (delParts.length === insParts.length && delParts.length > 1) {
            commonSubstrings.push({ pattern, length: len, count: delCount, parts: delParts.length });
          }
        }
      }
    }
  }
  
  // If we found repeated patterns, use a more sophisticated approach
  if (commonSubstrings.length > 0) {
    // Sort by: parts count (more parts = better), then length (longer = better), then count
    commonSubstrings.sort((a, b) => {
      if (b.parts !== a.parts) return b.parts - a.parts;
      if (b.length !== a.length) return b.length - a.length;
      return b.count - a.count;
    });
    
    // Use the best pattern to guide the split
    const bestPattern = commonSubstrings[0];
    
    // Split both strings by this pattern
    const delParts = deletedText.split(bestPattern.pattern);
    const insParts = insertedText.split(bestPattern.pattern);
    
    if (delParts.length === insParts.length && delParts.length > 1) {
      // We can split by this pattern
      const result = [];
      for (let i = 0; i < delParts.length; i++) {
        if (i > 0) {
          // Add the pattern as equal
          result.push([0, bestPattern.pattern]);
        }
        
        // Process the parts before/after the pattern
        if (delParts[i] !== insParts[i]) {
          // Different parts - need to diff them
          const partDiffs = dmp.diff_main(delParts[i], insParts[i]);
          dmp.diff_cleanupSemantic(partDiffs);
          
          for (const [op, text] of partDiffs) {
            if (op === -1 && text.length > 0) result.push([-1, text]);
            else if (op === 1 && text.length > 0) result.push([1, text]);
            else if (op === 0 && text.length > 0) result.push([0, text]);
          }
        } else if (delParts[i].length > 0) {
          // Same parts - mark as equal
          result.push([0, delParts[i]]);
        }
      }
      
      if (result.length > 2) {
        return result;
      }
    }
  }
  
  // Fallback: Use diff-match-patch on the pair itself
  const pairDiffs = dmp.diff_main(deletedText, insertedText);
  dmp.diff_cleanupSemantic(pairDiffs);
  
  const result = [];
  
  for (const [op, text] of pairDiffs) {
    if (op === 0 && text.length > 0) {
      result.push([0, text]);
    } else if (op === -1 && text.length > 0) {
      result.push([-1, text]);
    } else if (op === 1 && text.length > 0) {
      result.push([1, text]);
    }
  }
  
  // Always return the result from diff-match-patch, even if it's just delete+insert
  // The diff-match-patch algorithm is sophisticated and should handle most cases
  // If it found any structure (equal parts or multiple operations), use it
  // Otherwise, still return it so we don't lose the insert operation
  if (result.length > 0) {
    return result;
  }
  
  // Fallback: if diff-match-patch produced nothing (shouldn't happen), return original
  return [[-1, deletedText], [1, insertedText]];
}

/**
 * ENHANCED: Split text by newlines into separate operations
 * This ensures newlines are handled as distinct operations with proper context
 */
function splitByNewlines(text) {
  const parts = [];
  let current = '';
  
  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    
    // After normalization, we only have \n
    if (char === '\n') {
      if (current) {
        parts.push({ text: current, hasNewline: false });
        current = '';
      }
      // Push newline as separate item (will be converted to \r for Word)
      parts.push({ text: '\r', hasNewline: true });
    }
    else {
      current += char;
    }
  }
  
  if (current) {
    parts.push({ text: current, hasNewline: false });
  }
  
  return parts;
}

/**
 * ENHANCED: Merge consecutive newline operations
 * If there are multiple consecutive newlines in old and new text,
 * they should be treated as a group to avoid fragmentation
 */
function mergeConsecutiveNewlines(wordDiffs) {
  const merged = [];
  let i = 0;

  while (i < wordDiffs.length) {
    const current = wordDiffs[i];

    // Check if this is the start of a newline sequence
    if (current.isNewline) {
      // NEW STRATEGY: Merge ALL consecutive newlines of same operation type
      // This groups multiple newline deletions into ONE tracked change
      // User accepts once → all newlines in the group are handled
      const sequence = [current];
      let j = i + 1;

      while (j < wordDiffs.length && wordDiffs[j].isNewline && wordDiffs[j].op === current.op) {
        sequence.push(wordDiffs[j]);
        j++;
      }

      // Merge consecutive newlines of same type
      if (sequence.length > 1) {
        merged.push({
          op: current.op,
          text: sequence.map(s => s.text).join(''),
          isNewline: true,
          count: sequence.length
        });
        i = j;
      } else {
        // Single newline - keep as is
        merged.push(current);
        i++;
      }
    } else {
      // Not a newline - keep as is
      merged.push(current);
      i++;
    }
  }

  return merged;
}

/**
 * ENHANCED: Compute word-level diff with superior newline handling
 * Returns array of { op: 'equal'|'delete'|'insert', text: string, isNewline?: boolean }
 */
export function computeWordLevelDiff(oldText, newText) {
  // Use diff-match-patch directly with word-level granularity
  const oldStr = oldText || '';
  const newStr = newText || '';
  
  console.log('?? Computing diff...');
  console.log(`   Old text: ${oldStr.length} chars`);
  console.log(`   New text: ${newStr.length} chars`);
  
  // Normalize newlines before diffing (treat \r, \n, and \r\n as equivalent)
  const normalizedOld = normalizeNewlines(oldStr);
  const normalizedNew = normalizeNewlines(newStr);
  
  console.log(`   Normalized old: ${normalizedOld.length} chars`);
  console.log(`   Normalized new: ${normalizedNew.length} chars`);
  
  // Compute diff on normalized text
  const diffs = dmp.diff_main(normalizedOld, normalizedNew);
  dmp.diff_cleanupSemantic(diffs);
  
  console.log(`   Raw diffs: ${diffs.length} operations`);
  
  // Refine diff to handle repeated patterns (using normalized text)
  const refinedDiffs = refineDiffForRepeatedPatterns(diffs, normalizedOld, normalizedNew);
  
  console.log(`   Refined diffs: ${refinedDiffs.length} operations`);
  
  // Convert diff-match-patch format to our format
  // Split newlines into separate operations for better handling
  const wordDiffs = [];
  
  for (const [op, text] of refinedDiffs) {
    // Split by newlines to handle them separately
    const parts = splitByNewlines(text);
    
    for (const part of parts) {
      if (part.hasNewline) {
        // Newline is a separate operation - mark it specially
        if (op === 0) {
          wordDiffs.push({ op: 'equal', text: part.text, isNewline: true });
        } else if (op === -1) {
          wordDiffs.push({ op: 'delete', text: part.text, isNewline: true });
        } else if (op === 1) {
          wordDiffs.push({ op: 'insert', text: part.text, isNewline: true });
        }
      } else if (part.text) {
        // Regular text (non-empty)
        if (op === 0) {
          wordDiffs.push({ op: 'equal', text: part.text });
        } else if (op === -1) {
          wordDiffs.push({ op: 'delete', text: part.text });
        } else if (op === 1) {
          wordDiffs.push({ op: 'insert', text: part.text });
        }
      }
    }
  }
  
  console.log(`   Split into ${wordDiffs.length} word-level operations`);
  
  // ENHANCEMENT: Merge consecutive newlines
  const mergedDiffs = mergeConsecutiveNewlines(wordDiffs);
  
  console.log(`   After merging consecutive newlines: ${mergedDiffs.length} operations`);
  
  // ENHANCEMENT: Smart newline consolidation
  // If we have patterns like: [delete newline, text changes, insert newline]
  // where the newlines are equivalent, convert to equal newline
  const consolidatedDiffs = consolidateEquivalentNewlines(mergedDiffs);
  
  console.log(`   After consolidating equivalent newlines: ${consolidatedDiffs.length} operations`);
  console.log('   ? Diff computation complete');
  
  return consolidatedDiffs;
}

/**
 * ENHANCEMENT: Consolidate equivalent newlines
 * If a newline is deleted and inserted in the same logical position,
 * treat it as equal (unchanged)
 */
function consolidateEquivalentNewlines(wordDiffs) {
  const consolidated = [];
  let i = 0;

  while (i < wordDiffs.length) {
    const current = wordDiffs[i];

    // Check for pattern: delete newline, [optional text changes], insert newline
    if (current.op === 'delete' && current.isNewline) {
      // Look ahead for insert newline
      let foundInsert = false;
      let textChangesOnly = true;
      let insertIndex = -1;

      for (let j = i + 1; j < Math.min(i + 10, wordDiffs.length); j++) {
        const candidate = wordDiffs[j];

        if (candidate.op === 'insert' && candidate.isNewline) {
          foundInsert = true;
          insertIndex = j;
          break;
        } else if (candidate.op === 'equal') {
          // Equal text breaks the pattern
          textChangesOnly = false;
          break;
        }
        // delete/insert text is OK
      }

      if (foundInsert && textChangesOnly) {
        console.log(`   Found equivalent newline pair at positions ${i} and ${insertIndex}`);
        
        // Convert to equal newline
        consolidated.push({ op: 'equal', text: '\r', isNewline: true });

        // Add all text changes between
        for (let k = i + 1; k < insertIndex; k++) {
          consolidated.push(wordDiffs[k]);
        }

        // Skip to after the insert
        i = insertIndex + 1;
        continue;
      }
    }

    // No consolidation - keep as is
    consolidated.push(current);
    i++;
  }

  return consolidated;
}

/**
 * Generate OOXML with w:del and w:ins tags for real tracked changes
 */
export function generateRevisionOoxml(diffs, author = 'Addin', rsid = '00A1B2C3') {
  const now = new Date().toISOString();
  const dateStr = now.substring(0, 19) + 'Z';
  
  // Build OOXML paragraph - Word requires valid structure
  let ooxml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">';
  
  for (const diff of diffs) {
    const escapedText = escapeXml(diff.text);
    
    // Skip empty diffs
    if (!diff.text || diff.text.length === 0) {
      continue;
    }
    
    if (diff.op === 'equal') {
      // Regular text run - split by line breaks to handle them properly
      const lines = escapedText.split('\n');
      for (let i = 0; i < lines.length; i++) {
        if (lines[i]) {
          ooxml += `<w:r><w:t xml:space="preserve">${lines[i]}</w:t></w:r>`;
        }
        if (i < lines.length - 1) {
          // Add line break
          ooxml += `<w:r><w:br/></w:r>`;
        }
      }
    } else if (diff.op === 'delete') {
      // Deletion - wrap in w:del
      ooxml += `<w:del w:author="${escapeXml(author)}" w:date="${dateStr}" w:rsidDel="${rsid}">`;
      ooxml += `<w:r><w:delText xml:space="preserve">${escapedText}</w:delText></w:r>`;
      ooxml += `</w:del>`;
    } else if (diff.op === 'insert') {
      // Insertion - wrap in w:ins
      ooxml += `<w:ins w:author="${escapeXml(author)}" w:date="${dateStr}" w:rsidIns="${rsid}">`;
      const lines = escapedText.split('\n');
      for (let i = 0; i < lines.length; i++) {
        if (lines[i]) {
          ooxml += `<w:r><w:t xml:space="preserve">${lines[i]}</w:t></w:r>`;
        }
        if (i < lines.length - 1) {
          ooxml += `<w:r><w:br/></w:r>`;
        }
      }
      ooxml += `</w:ins>`;
    }
  }
  
  ooxml += '</w:p>';
  
  // Validate basic XML structure
  if (!ooxml.includes('<w:p') || !ooxml.includes('</w:p>')) {
    console.warn('Generated invalid OOXML structure');
  }
  
  return ooxml;
}

export function trackChangesDiffPreview(oldText, newText) {
  const text1 = oldText || '';
  const text2 = newText || '';
  
  // Compute the diff
  const diffs = dmp.diff_main(text1, text2);
  
  // Clean up the diff for better semantic matching
  dmp.diff_cleanupSemantic(diffs);
  
  // Convert diff array to HTML
  const html = diffs
    .map((diff) => {
      const operation = diff[0]; // -1 = delete, 0 = equal, 1 = insert
      const text = diff[1];
      const escapedText = escapeHtml(text);
      
      if (operation === 1) {
        // Insertion - underline in blue
        return `<span style="color:#106ba3;text-decoration:underline;">${escapedText}</span>`;
      } else if (operation === -1) {
        // Deletion - strikethrough in red
        return `<span style="color:#c23030;text-decoration:line-through;">${escapedText}</span>`;
      } else {
        // Equal - unchanged text
        return escapedText;
      }
    })
    .join('');
    
  return `<div>${html}</div>`;
}

export default trackChangesDiffPreview;
