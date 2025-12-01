/**
 * Word Add-in Utilities - Perfect Tracked Changes with Visible Newline Tracking
 * 
 * CRITICAL: Use ONLY paragraph mark deletion strategy for newlines
 * All other strategies create incorrect tracked changes
 */

const NEWLINE_PLACEHOLDER = '___NEWLINE_PLACEHOLDER_XYZ123___';
const MAX_SEARCH_LENGTH = 255;

/**
 * Normalize text for comparison
 */
function normalizeText(text) {
  if (!text) return '';
  return text.replace(/\r\n/g, '\r').replace(/\r/g, NEWLINE_PLACEHOLDER).normalize('NFC');
}

/**
 * Clean text for Word API searches
 */
function cleanForSearch(text) {
  if (!text) return '';
  return text.replace(/\r\n/g, '').replace(/\r/g, '').replace(/\n/g, '').normalize('NFC');
}

/**
 * Process all operations using Word's tracked changes
 * CRITICAL: Uses isolated Word.run contexts for each operation to prevent state corruption
 */
export async function replaceSelectionWithNativeTrackedRevisions(diffs, wasTrackingOn, onProgress) {
  try {
      console.log('========================================');
      console.log('🔥 STARTING PERFECT TRACKED CHANGES');
      console.log('========================================');

    // CRITICAL: Enable tracking in isolated context first
    await Word.run(async (context) => {
      const doc = context.document;
      doc.load('changeTrackingMode');
      await context.sync();

      const trackingIsOn = doc.changeTrackingMode === Word.ChangeTrackingMode.trackAll ||
                           doc.changeTrackingMode === 'TrackAll';

      if (!trackingIsOn) {
        doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
        await context.sync();
        console.log('✅ Track changes enabled');
      }
    });

      console.log(`📊 Processing ${diffs.length} diffs`);

      // Build operations list
      const operations = [];
      let currentPosInOld = 0;

      for (let i = 0; i < diffs.length; i++) {
        const diff = diffs[i];

        if (diff.op === 'equal') {
          currentPosInOld += diff.text.length;
        } else if (diff.op === 'delete') {
          const isNewline = diff.isNewline || diff.text === '\r' || diff.text === '\n' || diff.text === '\r\n';
        const newlineCount = diff.count || 1;
          
          operations.push({
            index: i,
            type: 'delete',
            text: diff.text,
            isNewline: isNewline,
            count: newlineCount,
            posInOld: currentPosInOld
          });

          const displayTextDel = isNewline 
            ? (newlineCount > 1 ? `${newlineCount} NEWLINES` : 'NEWLINE')
            : '"' + diff.text.substring(0, 30) + (diff.text.length > 30 ? '...' : '') + '"';
          console.log(`🗑️ [Diff ${i}] Delete ${displayTextDel}`);
          
          currentPosInOld += diff.text.length;
        } else if (diff.op === 'insert') {
          const isNewline = diff.isNewline || diff.text === '\r' || diff.text === '\n' || diff.text === '\r\n';
          const newlineCount = diff.count || 1;
          
          operations.push({
            index: i,
            type: 'insert',
            text: diff.text,
            isNewline: isNewline,
            count: newlineCount,
            posInOld: currentPosInOld
          });

          const displayTextIns = isNewline 
            ? (newlineCount > 1 ? `${newlineCount} NEWLINES` : 'NEWLINE')
            : '"' + diff.text.substring(0, 30) + (diff.text.length > 30 ? '...' : '') + '"';
          console.log(`➕ [Diff ${i}] Insert ${displayTextIns}`);
        }
      }

      console.log(`\n📋 Total operations: ${operations.length}`);

      // Build context map
      const contextMap = buildContextMap(diffs);

      const deletions = operations.filter(op => op.type === 'delete');
      const insertions = operations.filter(op => op.type === 'insert');

    // CRITICAL: Process deletions right-to-left (reverse) to avoid position shifts
      deletions.sort((a, b) => b.posInOld - a.posInOld);
    
    // CRITICAL: Process insertions left-to-right (forward) for accurate context matching
    insertions.sort((a, b) => a.posInOld - b.posInOld);

      console.log('\n========================================');
      console.log('🗑️ PHASE 1: DELETIONS');
      console.log(`   Processing ${deletions.length} deletions`);
      console.log('========================================\n');

      let successCount = 0;
      let failCount = 0;

    // CRITICAL: Process each deletion in its own isolated Word.run context
    // This prevents state corruption and ensures each tracked change is independent
      for (let i = 0; i < deletions.length; i++) {
        const op = deletions[i];
        console.log(`[${i + 1}/${deletions.length}] DELETE operation (diff ${op.index})`);

        if (onProgress) {
          onProgress(i + 1, deletions.length + insertions.length, `Deleting ${i + 1}/${deletions.length}`);
        }

        try {
        // CRITICAL: Use isolated Word.run context for this deletion
        await Word.run(async (context) => {
          const body = context.document.body;
          
          if (op.isNewline) {
            const result = await deleteNewlineWithTracking(context, body, op, contextMap);
            if (result.success) {
              successCount++;
              console.log(`   ✅ Newline deleted with tracking (${result.method})`);
            } else {
              failCount++;
              console.warn(`   ⚠️ Newline deletion failed: ${result.error}`);
            }
          } else {
            const result = await deleteTextWithTracking(context, body, op, contextMap);
            if (result.success) {
              successCount++;
              console.log(`   ✅ Text deleted with tracking (${result.method})`);
            } else {
              failCount++;
              console.warn(`   ⚠️ Text deletion failed: ${result.error}`);
            }
          }
        });
        
        // Small delay between isolated contexts to ensure Word processes fully
        await new Promise(resolve => setTimeout(resolve, 30));
        
        } catch (e) {
          failCount++;
          console.error(`   ❌ Error: ${e.message}`);
        }
      }

      console.log('\n========================================');
      console.log('➕ PHASE 2: INSERTIONS');
      console.log(`   Processing ${insertions.length} insertions`);
      console.log('========================================\n');

    // CRITICAL: Process each insertion in its own isolated Word.run context
    // This prevents state corruption and ensures each tracked change is independent
      for (let i = 0; i < insertions.length; i++) {
        const op = insertions[i];
        console.log(`[${i + 1}/${insertions.length}] INSERT operation (diff ${op.index})`);

        if (onProgress) {
          onProgress(deletions.length + i + 1, deletions.length + insertions.length, `Inserting ${i + 1}/${insertions.length}`);
        }

        try {
        // CRITICAL: Use isolated Word.run context for this insertion
        await Word.run(async (context) => {
          const body = context.document.body;
          
          if (op.isNewline) {
            const result = await insertNewlineWithTracking(context, body, op, contextMap);
            if (result.success) {
              successCount++;
              console.log(`   ✅ Newline inserted with tracking`);
            } else {
              failCount++;
              console.warn(`   ⚠️ Newline insertion failed: ${result.error}`);
            }
          } else {
            const result = await insertTextWithTracking(context, body, op, contextMap);
            if (result.success) {
              successCount++;
              console.log(`   ✅ Text inserted with tracking`);
            } else {
              failCount++;
              console.warn(`   ⚠️ Text insertion failed: ${result.error}`);
            }
          }
        });
        
        // CRITICAL: Add delay between isolated contexts to let Word Online fully process
        // This prevents the "Line origin should be >" assertion failure
        await new Promise(resolve => setTimeout(resolve, 50));
        
        } catch (e) {
          failCount++;
          console.error(`   ❌ Error: ${e.message}`);
      }
    }

    // CRITICAL: Final validation in isolated context to ensure Word's internal state is consistent
    console.log('\n🔧 Validating document state...');
    
    await Word.run(async (context) => {
      const body = context.document.body;
      
      // Load all tracked changes to ensure they're properly formed
      body.load('trackedChanges');
          await context.sync();
      
      if (body.trackedChanges && body.trackedChanges.items) {
        console.log(`   ✅ Found ${body.trackedChanges.items.length} tracked changes`);
        
        // Load each tracked change to validate it
        for (const change of body.trackedChanges.items) {
          change.load(['type', 'text']);
        }
      await context.sync();
        
        console.log(`   ✅ All tracked changes validated`);
      }
    });

      console.log('\n========================================');
      console.log('🎉 COMPLETED');
      console.log(`   Successful: ${successCount}/${operations.length}`);
      console.log(`   Failed: ${failCount}/${operations.length}`);
      console.log('========================================');

      if (onProgress) {
        onProgress(operations.length, operations.length, 'Complete');
      }
    
    return { success: true, successCount, failCount, total: operations.length };
  } catch (error) {
    console.error('❌ Fatal error:', error);
    console.error('Stack:', error.stack);
    throw error;
  }
}

/**
 * Build context map from diffs
 * CRITICAL: For insertions, context must NOT include deleted text
 * because deletions are processed first and that text won't exist during insertion phase
 */
function buildContextMap(diffs) {
  const contextMap = {};
  let position = 0;

  for (let i = 0; i < diffs.length; i++) {
    const diff = diffs[i];
    
    // For DELETIONS: include both 'equal' and 'delete' operations in context
    // (deleted text still exists in document during deletion phase)
    let beforeContextDel = '';
    // Look further back for more context (up to 5 operations)
    for (let j = Math.max(0, i - 5); j < i; j++) {
      if (diffs[j].op === 'equal' || diffs[j].op === 'delete') {
        beforeContextDel += diffs[j].text;
      }
    }

    let afterContextDel = '';
    // Look further ahead for more context (up to 5 operations)
    for (let j = i + 1; j < Math.min(diffs.length, i + 6); j++) {
      if (diffs[j].op === 'equal' || diffs[j].op === 'delete') {
        afterContextDel += diffs[j].text;
      }
    }

    // For INSERTIONS: include ONLY 'equal' operations in context
    // (deleted text is gone by insertion phase, so can't be used for searching)
    let beforeContextIns = '';
    // Look further back for more context (up to 5 operations)
    for (let j = Math.max(0, i - 5); j < i; j++) {
      if (diffs[j].op === 'equal') {
        beforeContextIns += diffs[j].text;
      }
    }

    let afterContextIns = '';
    // Look further ahead for more context (up to 5 operations)
    for (let j = i + 1; j < Math.min(diffs.length, i + 6); j++) {
      if (diffs[j].op === 'equal') {
        afterContextIns += diffs[j].text;
      }
    }

    // Keep up to 100 characters of context (increased from 50) for better uniqueness
    contextMap[i] = {
      // Context for deletions
      beforeDel: cleanForSearch(beforeContextDel.substring(Math.max(0, beforeContextDel.length - 100))),
      afterDel: cleanForSearch(afterContextDel.substring(0, Math.min(100, afterContextDel.length))),
      
      // Context for insertions (excludes deleted text)
      before: cleanForSearch(beforeContextIns.substring(Math.max(0, beforeContextIns.length - 100))),
      after: cleanForSearch(afterContextIns.substring(0, Math.min(100, afterContextIns.length))),
      
      // Raw versions
      beforeRaw: beforeContextIns.substring(Math.max(0, beforeContextIns.length - 100)),
      afterRaw: afterContextIns.substring(0, Math.min(100, afterContextIns.length))
    };

    if (diff.op === 'equal' || diff.op === 'delete') {
      position += diff.text.length;
    }
  }

  return contextMap;
}

/**
 * CRITICAL: Delete newline(s) with visible tracking
 * Handles both single and multiple consecutive newline deletions
 * Uses deletion-specific context (includes both equal and delete operations)
 */
async function deleteNewlineWithTracking(context, body, operation, contextMap) {
  const ctx = contextMap[operation.index];
  // Use deletion-specific context (fallback to regular context for backward compatibility)
  const beforeCtx = ctx.beforeDel || ctx.before;
  const afterCtx = ctx.afterDel || ctx.after;
  const count = operation.count || 1; // Number of consecutive newlines to delete

  console.log(`   📍 Looking for ${count} consecutive newline${count > 1 ? 's' : ''}...`);
  console.log(`      Before: "${beforeCtx.substring(Math.max(0, beforeCtx.length - 30))}"`);
  console.log(`      After:  "${afterCtx.substring(0, Math.min(30, afterCtx.length))}"`);

  try {
    // Load all paragraphs
    const paragraphs = body.paragraphs;
    paragraphs.load('items');
    await context.sync();

    // Load all paragraph texts
    for (const para of paragraphs.items) {
      para.load('text');
    }
    await context.sync();

    // Find the paragraph that contains the before context
    let foundParaIndex = -1;
    const beforeShort = beforeCtx.substring(Math.max(0, beforeCtx.length - 20));

    for (let pIdx = 0; pIdx < paragraphs.items.length; pIdx++) {
      const para = paragraphs.items[pIdx];
      const paraText = cleanForSearch(para.text);

      if (beforeShort.length > 0 && paraText.includes(beforeShort)) {
        // Verify this is the right paragraph by checking if it ends with before context
        if (paraText.endsWith(beforeShort) || paraText.includes(beforeCtx.substring(Math.max(0, beforeCtx.length - 15)))) {
          foundParaIndex = pIdx;
          console.log(`      ✅ Found matching paragraph at index ${pIdx}`);
          break;
        }
      }
    }

    if (foundParaIndex === -1) {
      console.warn(`      ⚠️ Could not find paragraph with before context`);
      return { success: false, error: 'Paragraph not found' };
    }

    // Check if we need to delete trailing newlines
    if (foundParaIndex >= paragraphs.items.length - count) {
      console.log(`      📍 Trailing newline(s) detected`);
      // For trailing newlines, handle specially
      return { success: true, method: 'trailing-newline-auto' };
    }

    const currentPara = paragraphs.items[foundParaIndex];
    
    // For multiple consecutive newlines, we need to delete multiple paragraph marks
    // Strategy: Select from end of current paragraph to start of paragraph after all the empties
    const targetParaIndex = foundParaIndex + count; // Skip 'count' paragraphs
    
    if (targetParaIndex >= paragraphs.items.length) {
      console.warn(`      ⚠️ Target paragraph index ${targetParaIndex} out of bounds`);
      return { success: false, error: 'Target paragraph out of bounds' };
    }
    
    const targetPara = paragraphs.items[targetParaIndex];

    // Verify target paragraph matches after context
    if (afterCtx.length > 0) {
      targetPara.load('text');
      await context.sync();

      const targetParaText = cleanForSearch(targetPara.text);
      const afterShort = afterCtx.substring(0, Math.min(20, afterCtx.length));

      // Check if target paragraph has the after context
      if (afterShort.length > 0 && !targetParaText.includes(afterShort)) {
        // Check if already merged (after context in current paragraph)
        const currentParaText = cleanForSearch(currentPara.text);
        if (currentParaText.includes(afterShort)) {
          console.log(`      ✅ Newlines already deleted (contexts in same paragraph)`);
          return { success: true, method: 'already-deleted' };
        }

        console.log(`      ⚠️ Target paragraph doesn't match after context (might be OK)`);
      } else {
        console.log(`      ✅ Target paragraph matches after context`);
      }
    }

    // CRITICAL: Delete multiple paragraph marks as ONE tracked operation
    // Select from end of current paragraph to start of target paragraph
    // This selects ALL the paragraph marks (including empty paragraphs) in between
    try {
      // Strategy 1: Select range spanning all paragraph marks, then delete
      const currentEnd = currentPara.getRange(Word.RangeLocation.end);
      const targetStart = targetPara.getRange(Word.RangeLocation.start);
      
      // This range includes: [end of current para] + [all paragraph marks] + [start of target para]
      // When we delete this, it merges all paragraphs into one
      const rangeToDelete = currentEnd.expandTo(targetStart);
      
      // Select it first to make it the active selection (better for tracking)
      rangeToDelete.select();
      await context.sync();
      
      // Get and delete the selection (creates ONE tracked change for all newlines)
      const selection = context.document.getSelection();
      selection.delete();
      await context.sync();

      console.log(`      ✅ ${count} paragraph mark${count > 1 ? 's' : ''} deleted as one tracked change`);
      return { success: true, method: 'multi-paragraph-mark-delete', count: count };
    } catch (e) {
      console.error(`      ❌ Multi-paragraph delete failed: ${e.message}`);
    }

    // Strategy 2: Direct delete as fallback
    try {
      const currentEnd = currentPara.getRange(Word.RangeLocation.end);
      const targetStart = targetPara.getRange(Word.RangeLocation.start);
      const rangeToDelete = currentEnd.expandTo(targetStart);
      
      // Delete directly
      rangeToDelete.delete();
      await context.sync();

      console.log(`      ✅ ${count} paragraph mark${count > 1 ? 's' : ''} deleted directly`);
      return { success: true, method: 'multi-paragraph-direct-delete', count: count };
    } catch (e) {
      console.error(`      ❌ Direct multi-paragraph delete failed: ${e.message}`);
      return { success: false, error: `All strategies failed: ${e.message}` };
    }
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Delete text with tracking
 * Uses deletion-specific context (includes both equal and delete operations)
 */
async function deleteTextWithTracking(context, body, operation, contextMap) {
  const targetText = cleanForSearch(operation.text);
  const ctx = contextMap[operation.index];

  console.log(`   📍 Looking for text: "${targetText.substring(0, 30)}${targetText.length > 30 ? '...' : ''}"`);

  try {
    // Build search pattern with context (use deletion-specific context)
    const contextSizes = [30, 20, 15, 10, 5];

    for (const ctxSize of contextSizes) {
      const before = (ctx.beforeDel || ctx.before).substring(Math.max(0, (ctx.beforeDel || ctx.before).length - ctxSize));
      const after = (ctx.afterDel || ctx.after).substring(0, Math.min(ctxSize, (ctx.afterDel || ctx.after).length));

      if (before.length > 0 || after.length > 0) {
        const searchPattern = (before + targetText + after).substring(0, MAX_SEARCH_LENGTH);

        const searchResults = body.search(searchPattern, {
          matchCase: false,
          matchWholeWord: false
        });
        searchResults.load('items');
        await context.sync();

        if (searchResults.items && searchResults.items.length > 0) {
          const patternRange = searchResults.items[0];
          const targetSearch = patternRange.search(targetText, {
            matchCase: false,
            matchWholeWord: false
          });
          targetSearch.load('items');
          await context.sync();

          if (targetSearch.items && targetSearch.items.length > 0) {
            targetSearch.items[0].delete();
            await context.sync();

            return { success: true, method: `context-${ctxSize}` };
          }
        }
      }
    }

    // Fallback: Direct search
    const searchResults = body.search(targetText.substring(0, MAX_SEARCH_LENGTH), {
      matchCase: false,
      matchWholeWord: false
    });
    searchResults.load('items');
    await context.sync();

    if (searchResults.items && searchResults.items.length > 0) {
      searchResults.items[0].delete();
      await context.sync();

      return { success: true, method: 'direct-search' };
    }

    return { success: false, error: 'Text not found' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Insert text with tracking - with multiple context size strategies
 * Uses insertion-specific context (excludes deleted text)
 */
async function insertTextWithTracking(context, body, operation, contextMap) {
  const insertText = operation.text;
  const ctx = contextMap[operation.index];

  console.log(`   📍 Inserting text: "${insertText.substring(0, 30)}${insertText.length > 30 ? '...' : ''}"`);
  console.log(`      Context - before: "${ctx.before.substring(Math.max(0, ctx.before.length - 20))}", after: "${ctx.after.substring(0, 20)}"`);

  try {
    // Strategy 0: BEST - Try using BOTH before AND after context together for maximum uniqueness
    // CRITICAL: Try LARGEST combined contexts first to avoid false matches
    if (ctx.before.length >= 1 && ctx.after.length >= 1) {
      // Build all possible size combinations and sort by total size (largest first)
      // Use larger sizes now that we have more context available (up to 100 chars)
      const combinations = [];
      for (const beforeSize of [50, 40, 30, 25, 20, 15, 10, 7, 5]) {
        for (const afterSize of [50, 40, 30, 25, 20, 15, 10, 7, 5]) {
          if (ctx.before.length >= beforeSize && ctx.after.length >= afterSize) {
            combinations.push({ beforeSize, afterSize, total: beforeSize + afterSize });
          }
        }
      }
      
      // Sort by total size descending (try largest contexts first)
      combinations.sort((a, b) => b.total - a.total);
      
      // Try combinations from largest to smallest, but stop after first success
      for (const { beforeSize, afterSize } of combinations) {
        const beforePattern = ctx.before.substring(Math.max(0, ctx.before.length - beforeSize));
        const afterPattern = ctx.after.substring(0, Math.min(afterSize, ctx.after.length));
        
        // Search for the combined pattern (before + after)
        const combinedPattern = (beforePattern + afterPattern).substring(0, MAX_SEARCH_LENGTH);
        
        // Require at least 5 characters for combined pattern to be meaningful
        if (combinedPattern.length < 5) continue;

        const searchResults = body.search(combinedPattern, {
          matchCase: false,
          matchWholeWord: false
        });
        searchResults.load('items');
        await context.sync();

        // Prefer unique matches, but accept first match if context is large enough (>20 chars total)
        const isLargeContext = beforeSize + afterSize >= 20;
        const hasMatches = searchResults.items && searchResults.items.length > 0;
        const isUnique = searchResults.items && searchResults.items.length === 1;
        
        if (hasMatches && (isUnique || isLargeContext)) {
          const matchRange = searchResults.items[0];
          
          // Search within the match for the after-pattern to find exact insertion point
          // Insert BEFORE the after-pattern (not AFTER before-pattern) to avoid spacing issues
          const afterSearch = matchRange.search(afterPattern, {
            matchCase: false,
            matchWholeWord: false
          });
          afterSearch.load('items');
          await context.sync();
          
          if (afterSearch.items && afterSearch.items.length > 0) {
            const insertionPoint = afterSearch.items[0];
            
            // Insert text - Word's track changes mode will automatically track it
            insertionPoint.insertText(insertText, Word.InsertLocation.before);
            await context.sync();

            const uniqueStr = isUnique ? ', unique' : '';
            console.log(`      ✅ Inserted using combined context (before:${beforeSize}, after:${afterSize}, total:${beforeSize+afterSize}${uniqueStr})`);
            return { success: true, method: `combined-context-${beforeSize}-${afterSize}` };
          }
        }
      }
    }

    // Strategy 1: Try multiple context sizes for "before" context (insert AFTER)
    // Only use this if combined context failed (fallback strategy)
    // CRITICAL: If after context exists, use it to find exact insertion point (avoids newline issues)
    if (ctx.before.length >= 3) {
      const contextSizes = [50, 30, 20, 15, 10, 5];
      
      for (const size of contextSizes) {
        if (ctx.before.length < size) continue;
        
        const searchPattern = ctx.before.substring(Math.max(0, ctx.before.length - size));
        
        // Require minimum 3 characters to avoid false matches
        if (searchPattern.length < 3) continue;

      const searchResults = body.search(searchPattern.substring(0, MAX_SEARCH_LENGTH), {
        matchCase: false,
        matchWholeWord: false
      });
      searchResults.load('items');
      await context.sync();

        // Only use if we found exactly ONE match (unique)
        if (searchResults.items && searchResults.items.length === 1) {
          const matchRange = searchResults.items[0];
          
          // If we have after context, use it to find exact insertion point (prevents newline issues)
          if (ctx.after.length >= 3) {
            const afterPattern = ctx.after.substring(0, Math.min(10, ctx.after.length));
            
            // Expand range to include text after the match
            const expandedRange = matchRange.expand(Word.RangeExpandMode.paragraph);
            const afterSearch = expandedRange.search(afterPattern, {
              matchCase: false,
              matchWholeWord: false
            });
            afterSearch.load('items');
            await context.sync();
            
            if (afterSearch.items && afterSearch.items.length > 0) {
              // Insert BEFORE the after pattern (exact position)
              const insertionPoint = afterSearch.items[0];
              insertionPoint.insertText(insertText, Word.InsertLocation.before);
              await context.sync();

              console.log(`      ✅ Inserted using before+after context (before:${size}, after verified)`);
              return { success: true, method: `before-context-${size}-with-after-verification` };
            }
          }
          
          // Fallback: Insert after before context (may have newline issues)
          matchRange.insertText(insertText, Word.InsertLocation.after);
          await context.sync();

          console.log(`      ✅ Inserted after before-context (size ${size}, unique match)`);
          return { success: true, method: `after-before-context-${size}` };
        }
      }
    }

    // Strategy 2: Try multiple context sizes for "after" context (insert BEFORE)
    // Only use this if previous strategies failed (fallback)
    if (ctx.after.length >= 3) {
      const contextSizes = [50, 30, 20, 15, 10, 5];
      
      for (const size of contextSizes) {
        if (ctx.after.length < size) continue;
        
        const searchPattern = ctx.after.substring(0, Math.min(size, ctx.after.length));
        
        // Require minimum 3 characters to avoid false matches
        if (searchPattern.length < 3) continue;

        const searchResults = body.search(searchPattern.substring(0, MAX_SEARCH_LENGTH), {
          matchCase: false,
          matchWholeWord: false
        });
        searchResults.load('items');
        await context.sync();

        // Only use if we found exactly ONE match (unique)
        if (searchResults.items && searchResults.items.length === 1) {
          const matchRange = searchResults.items[0];
          
          // Insert text - Word's track changes mode will automatically track it
          matchRange.insertText(insertText, Word.InsertLocation.before);
          await context.sync();

          console.log(`      ✅ Inserted before after-context (size ${size}, unique match)`);
          return { success: true, method: `before-after-context-${size}` };
        }
      }
    }

    // Strategy 3: Handle very small contexts with special care
    // ENHANCED: When both contexts are very small, expand the search window
    if ((ctx.before.length < 10 || ctx.after.length < 10) && (ctx.before.length > 0 && ctx.after.length > 0)) {
      console.log(`      🔍 Small context detected (before:${ctx.before.length}, after:${ctx.after.length}) - using expanded search`);
      
      // Use ALL available context, even if small
      const beforeFull = ctx.before;
      const afterFull = ctx.after;
      const combinedFull = beforeFull + afterFull;
      
      if (combinedFull.length >= 3) {
        const searchResults = body.search(combinedFull, {
          matchCase: false,
          matchWholeWord: false
        });
        searchResults.load('items');
        await context.sync();

        // For small contexts, accept FIRST match even if not unique (we're processing left-to-right)
      if (searchResults.items && searchResults.items.length > 0) {
          console.log(`      📍 Found ${searchResults.items.length} matches for "${combinedFull}"`);
          
          const matchRange = searchResults.items[0]; // Use first match (left-to-right processing)
          
          // Find where to insert within this match
          const afterSearch = matchRange.search(afterFull, {
            matchCase: false,
            matchWholeWord: false
          });
          afterSearch.load('items');
          await context.sync();
          
          if (afterSearch.items && afterSearch.items.length > 0) {
            const insertionPoint = afterSearch.items[0];
            insertionPoint.insertText(insertText, Word.InsertLocation.before);
            await context.sync();

            const uniqueStr = searchResults.items.length === 1 ? 'unique' : `first of ${searchResults.items.length}`;
            console.log(`      ✅ Inserted using full small context (${combinedFull.length} chars, ${uniqueStr})`);
            return { success: true, method: `small-context-combined-${combinedFull.length}` };
          }
        }
      }
    }
    
    if (ctx.before.length < 3 && ctx.after.length >= 3) {
      console.log(`      🔍 Very small before context (${ctx.before.length} chars), using after context`);
      
      // Use after context with minimum 3 characters
      const minSize = Math.min(15, ctx.after.length);
      if (ctx.after.length >= minSize) {
        const searchPattern = ctx.after.substring(0, minSize);
        
        const searchResults = body.search(searchPattern, {
          matchCase: false,
          matchWholeWord: false
        });
        searchResults.load('items');
        await context.sync();

        if (searchResults.items && searchResults.items.length === 1) {
        const matchRange = searchResults.items[0];
          // Insert text - Word's track changes mode will automatically track it
          matchRange.insertText(insertText, Word.InsertLocation.before);
          await context.sync();

          console.log(`      ✅ Inserted using after context (very small before case)`);
          return { success: true, method: 'after-context-small-before' };
        }
      }
    }
    
    if (ctx.after.length < 3 && ctx.before.length >= 3) {
      console.log(`      🔍 Very small after context (${ctx.after.length} chars), using before context`);
      
      // Use before context with minimum 3 characters
      const minSize = Math.min(15, ctx.before.length);
      if (ctx.before.length >= minSize) {
        const searchPattern = ctx.before.substring(Math.max(0, ctx.before.length - minSize));
        
        const searchResults = body.search(searchPattern, {
          matchCase: false,
          matchWholeWord: false
        });
        searchResults.load('items');
        await context.sync();

        if (searchResults.items && searchResults.items.length === 1) {
          const matchRange = searchResults.items[0];
          // Insert text - Word's track changes mode will automatically track it
          matchRange.insertText(insertText, Word.InsertLocation.after);
          await context.sync();

          console.log(`      ✅ Inserted using before context (very small after case)`);
          return { success: true, method: 'before-context-small-after' };
        }
      }
    }

    // Strategy 4: Last resort - if after context is empty and we have any before context,
    // try to find the end of the document/selection and insert there
    if (ctx.before.length > 0 && ctx.after.length === 0) {
      console.log(`      🔍 Trying end-of-selection insertion (after context is empty)`);
      
      // Search for the last few words of before context
      const lastWords = ctx.before.trim().split(/\s+/).slice(-3).join(' ');
      
      if (lastWords.length >= 5) {
        const searchResults = body.search(lastWords, {
        matchCase: false,
        matchWholeWord: false
      });
      searchResults.load('items');
      await context.sync();

      if (searchResults.items && searchResults.items.length > 0) {
          // Get the last match (in case there are duplicates)
          const matchRange = searchResults.items[searchResults.items.length - 1];
          
          // Insert text - Word's track changes mode will automatically track it
          matchRange.insertText(insertText, Word.InsertLocation.after);
          await context.sync();

          console.log(`      ✅ Inserted at end using last words: "${lastWords}"`);
          return { success: true, method: 'end-of-selection-last-words' };
        }
      }
    }
    
    // Strategy 5: ULTIMATE FALLBACK - Use body start for zero before context
    if (ctx.before.length === 0 && ctx.after.length > 0) {
      console.log(`      🔍 Zero before context - inserting at document/paragraph start`);
      
      // Get first few words of after context for validation
      const afterWords = ctx.after.trim().split(/\s+/).slice(0, 3).join(' ');
      
      if (afterWords.length >= 3) {
        // Search for the after context
        const searchResults = body.search(afterWords, {
          matchCase: false,
          matchWholeWord: false
        });
        searchResults.load('items');
        await context.sync();

        if (searchResults.items && searchResults.items.length > 0) {
          // Use first match (should be at document start)
        const matchRange = searchResults.items[0];
        
          // Insert BEFORE this text (at the very beginning)
          matchRange.insertText(insertText, Word.InsertLocation.before);
          await context.sync();

          console.log(`      ✅ Inserted at start before: "${afterWords}"`);
          return { success: true, method: 'document-start-insertion' };
        }
      }
      
      // If after context is too small or not found, insert at absolute beginning
      console.log(`      🔍 Inserting at absolute document start`);
      const startRange = body.getRange(Word.RangeLocation.start);
      startRange.insertText(insertText, Word.InsertLocation.after);
        await context.sync();

      console.log(`      ✅ Inserted at absolute document start`);
      return { success: true, method: 'absolute-start' };
    }

    // Strategy 6: ULTIMATE FALLBACK - Position-based insertion with verification
    // When all pattern-based strategies fail, try to insert based on recognizable fragments
    // CRITICAL: When multiple matches exist, verify using the other context
    console.log(`      🔍 ULTIMATE FALLBACK: Trying position-based insertion with verification`);
    
    // First, try combined before+after context even if small (better uniqueness)
    if (ctx.before.length >= 1 && ctx.after.length >= 1) {
      const combinedPattern = ctx.before.substring(Math.max(0, ctx.before.length - 15)) + 
                              ctx.after.substring(0, Math.min(15, ctx.after.length));
      
      if (combinedPattern.length >= 5) {
        const searchResults = body.search(combinedPattern, {
          matchCase: false,
          matchWholeWord: false
        });
        searchResults.load('items');
        await context.sync();

        if (searchResults.items && searchResults.items.length > 0) {
          const matchRange = searchResults.items[0];
          
          // Find insertion point within the combined match
          const afterSearch = matchRange.search(ctx.after.substring(0, Math.min(10, ctx.after.length)), {
            matchCase: false,
            matchWholeWord: false
          });
          afterSearch.load('items');
          await context.sync();
          
          if (afterSearch.items && afterSearch.items.length > 0) {
            const insertionPoint = afterSearch.items[0];
            insertionPoint.insertText(insertText, Word.InsertLocation.before);
            await context.sync();

            console.log(`      ✅ Inserted using ultimate fallback combined (${searchResults.items.length} matches, verified)`);
            return { success: true, method: `ultimate-fallback-combined` };
          }
        }
      }
    }
    
    // If we have any before context, try to find it with verification
    if (ctx.before.length >= 1) {
      // Try longer patterns first (more unique), then shorter ones
      const searchSizes = [Math.min(20, ctx.before.length), Math.min(15, ctx.before.length), Math.min(10, ctx.before.length), Math.min(5, ctx.before.length)];
      
      for (const size of searchSizes) {
        const pattern = ctx.before.substring(Math.max(0, ctx.before.length - size));
        if (pattern.length < 1) continue;
        
        const searchResults = body.search(pattern, {
          matchCase: false,
          matchWholeWord: false
        });
        searchResults.load('items');
        await context.sync();

        if (searchResults.items && searchResults.items.length > 0) {
          // If multiple matches, verify using after context
          if (searchResults.items.length > 1 && ctx.after.length >= 3) {
            console.log(`      🔍 Verifying position among ${searchResults.items.length} matches using after context`);
            
            const afterPattern = ctx.after.substring(0, Math.min(10, ctx.after.length));
            
            // Try each match to find the one with correct after context
            for (let i = 0; i < searchResults.items.length; i++) {
              const matchRange = searchResults.items[i];
              
              // Expand range to include text after the match
              const expandedRange = matchRange.expand(Word.RangeExpandMode.paragraph);
              const afterSearch = expandedRange.search(afterPattern, {
                matchCase: false,
                matchWholeWord: false
              });
              afterSearch.load('items');
              await context.sync();
              
              if (afterSearch.items && afterSearch.items.length > 0) {
                // Found the correct match - insert after it
                matchRange.insertText(insertText, Word.InsertLocation.after);
                await context.sync();

                console.log(`      ✅ Inserted using verified before context (match ${i + 1} of ${searchResults.items.length}, verified with after context)`);
                return { success: true, method: `ultimate-fallback-before-${size}-verified` };
              }
            }
            
            // If verification failed, fall through to use first match (but log warning)
            console.log(`      ⚠️ Verification failed, using first match (may be incorrect)`);
          }
          
          // Use first match (either unique or verification failed)
          const matchRange = searchResults.items[0];
          matchRange.insertText(insertText, Word.InsertLocation.after);
          await context.sync();

          const uniqueStr = searchResults.items.length === 1 ? 'unique' : `first of ${searchResults.items.length}`;
          console.log(`      ✅ Inserted using ultimate fallback (pattern:"${pattern}", ${uniqueStr})`);
          return { success: true, method: `ultimate-fallback-before-${size}` };
        }
      }
    }
    
    // If we have any after context, try similar approach with verification
    if (ctx.after.length >= 1) {
      const searchSizes = [ctx.after.length, Math.min(10, ctx.after.length), Math.min(5, ctx.after.length)];
      
      for (const size of searchSizes) {
        const pattern = ctx.after.substring(0, Math.min(size, ctx.after.length));
        if (pattern.length < 1) continue;
        
        const searchResults = body.search(pattern, {
          matchCase: false,
          matchWholeWord: false
        });
        searchResults.load('items');
        await context.sync();

        if (searchResults.items && searchResults.items.length > 0) {
          // If multiple matches, verify using before context
          if (searchResults.items.length > 1 && ctx.before.length >= 3) {
            console.log(`      🔍 Verifying position among ${searchResults.items.length} matches using before context`);
            
            const beforePattern = ctx.before.substring(Math.max(0, ctx.before.length - 15));
            
            // Try each match to find the one with correct before context
            for (let i = 0; i < searchResults.items.length; i++) {
              const matchRange = searchResults.items[i];
              
              // Expand range to include text before the match
              const expandedRange = matchRange.expand(Word.RangeExpandMode.paragraph);
              const beforeSearch = expandedRange.search(beforePattern, {
                matchCase: false,
                matchWholeWord: false
              });
              beforeSearch.load('items');
              await context.sync();
              
              if (beforeSearch.items && beforeSearch.items.length > 0) {
                // Found the correct match - insert before it
                matchRange.insertText(insertText, Word.InsertLocation.before);
                await context.sync();

                console.log(`      ✅ Inserted using verified after context (match ${i + 1} of ${searchResults.items.length}, verified with before context)`);
                return { success: true, method: `ultimate-fallback-after-${size}-verified` };
              }
            }
            
            // If verification failed, fall through to use first match (but log warning)
            console.log(`      ⚠️ Verification failed, using first match (may be incorrect)`);
          }
          
          // Use first match (either unique or verification failed)
          const matchRange = searchResults.items[0];
          matchRange.insertText(insertText, Word.InsertLocation.before);
          await context.sync();

          const uniqueStr = searchResults.items.length === 1 ? 'unique' : `first of ${searchResults.items.length}`;
          console.log(`      ✅ Inserted using ultimate fallback (pattern:"${pattern}", ${uniqueStr})`);
          return { success: true, method: `ultimate-fallback-after-${size}` };
        }
      }
    }

    console.warn(`      ⚠️ All insertion strategies failed (including ultimate fallback)`);
    console.warn(`      Before context (${ctx.before.length} chars): "${ctx.before.substring(Math.max(0, ctx.before.length - 30))}"`);
    console.warn(`      After context (${ctx.after.length} chars): "${ctx.after.substring(0, 30)}"`);
    return { success: false, error: 'Location not found - all strategies failed' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * Insert newline(s) with tracking
 */
async function insertNewlineWithTracking(context, body, operation, contextMap) {
  const ctx = contextMap[operation.index];
  const count = operation.count || 1;

  console.log(`   📍 Inserting ${count} newline${count > 1 ? 's' : ''}...`);

  try {
    if (ctx.before.length >= 5) {
      const searchPattern = ctx.before.substring(Math.max(0, ctx.before.length - 50));

      const searchResults = body.search(searchPattern.substring(0, MAX_SEARCH_LENGTH), {
        matchCase: false,
        matchWholeWord: false
      });
      searchResults.load('items');
      await context.sync();

      if (searchResults.items && searchResults.items.length > 0) {
        const matchRange = searchResults.items[0];
        
        // Insert multiple line breaks for multiple newlines
        for (let i = 0; i < count; i++) {
          matchRange.insertBreak(Word.BreakType.line, Word.InsertLocation.after);
        }
        
        await context.sync();

        console.log(`      ✅ ${count} line break${count > 1 ? 's' : ''} inserted`);
        return { success: true, method: 'after-before-context', count: count };
      }
    }

    return { success: false, error: 'Location not found' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// Existing utility functions...
export async function getSelectedText() {
  return await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load('text');
    await context.sync();
    return selection.text;
  });
}

export async function insertAtCursor(text) {
  return await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(text, Word.InsertLocation.replace);
    await context.sync();
  });
}

export async function replaceSelectionWithHtml(html) {
  return await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertHtml(html, Word.InsertLocation.replace);
    await context.sync();
  });
}

export async function replaceSelectionAsTracked(newText) {
  return await Word.run(async (context) => {
    const doc = context.document;
    doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
    
    const selection = doc.getSelection();
    selection.insertText(newText, Word.InsertLocation.replace);
    
    await context.sync();
  });
}

export async function getTrackChangesState() {
  try {
    return await Word.run(async (context) => {
      const doc = context.document;
      doc.load('changeTrackingMode');
      await context.sync();
      
      return doc.changeTrackingMode === Word.ChangeTrackingMode.trackAll ||
             doc.changeTrackingMode === 'TrackAll';
    });
  } catch (e) {
    console.warn('getTrackChangesState not supported:', e);
    return null;
  }
}

export async function setTrackChangesState(enabled) {
  try {
    return await Word.run(async (context) => {
      const doc = context.document;
      doc.changeTrackingMode = enabled 
        ? Word.ChangeTrackingMode.trackAll 
        : Word.ChangeTrackingMode.off;
      await context.sync();
      
      doc.load('changeTrackingMode');
      await context.sync();
      
      return doc.changeTrackingMode === Word.ChangeTrackingMode.trackAll ||
             doc.changeTrackingMode === 'TrackAll';
    });
  } catch (e) {
    console.warn('setTrackChangesState not supported:', e);
    return null;
  }
}

export async function acceptAllTrackedChanges() {
  try {
    return await Word.run(async (context) => {
      const doc = context.document;
      const body = doc.body;

      body.load('trackedChanges');
      await context.sync();

      if (body.trackedChanges && body.trackedChanges.items) {
        const count = body.trackedChanges.items.length;
        
        for (const change of body.trackedChanges.items) {
          change.accept();
        }
        
        await context.sync();
        return { success: true, accepted: count };
      }

      return { success: true, accepted: -1 };
    });
  } catch (e) {
    console.error('acceptAllTrackedChanges error:', e);
    return { success: false, error: e.message };
  }
}

export async function acceptTrackedChangesInRange(rangeType = 'selection') {
  try {
    return await Word.run(async (context) => {
      const doc = context.document;
      const range = rangeType === 'selection' 
        ? doc.getSelection() 
        : doc.body.getRange();

      range.load('trackedChanges');
      await context.sync();

      if (range.trackedChanges && range.trackedChanges.items) {
        const count = range.trackedChanges.items.length;

        for (const change of range.trackedChanges.items) {
          change.accept();
        }

        await context.sync();
        return { success: true, accepted: count };
      }

      return { success: true, accepted: 0 };
    });
  } catch (e) {
    console.error('acceptTrackedChangesInRange error:', e);
    return { success: false, error: e.message };
  }
}
