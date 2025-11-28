/**
 * Word Add-in Utilities - Maximum Precision Tracked Changes
 * 
 * ENHANCED ALGORITHM for single-character accuracy:
 * - Adaptive context sizing (expands until unique)
 * - Position verification before applying
 * - Multiple search strategies with fallback
 * - Detailed logging for every operation
 */

/**
 * Placeholder token for newlines - must be unique and unlikely to appear in real text
 */
const NEWLINE_PLACEHOLDER = '___NEWLINE_PLACEHOLDER_XYZ123___';

/**
 * Normalize text for comparison
 * CRITICAL: Replace \r with placeholder so newlines can be handled as regular text
 * Word's search API can't match \r reliably, so we use a placeholder
 */
function normalizeText(text) {
  if (!text) return '';
  // Replace carriage returns (\r) with placeholder, then normalize Unicode
  return text.replace(/\r/g, NEWLINE_PLACEHOLDER).normalize('NFC');
}

/**
 * Convert placeholder back to newline character
 */
function denormalizeText(text) {
  if (!text) return '';
  return text.replace(new RegExp(NEWLINE_PLACEHOLDER.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g'), '\r');
}

/**
 * Find unique context for a position in text
 * Expands context window until pattern is GUARANTEED unique
 * For single characters, uses much larger context to prevent ambiguity
 */
function findUniqueContext(text, startPos, endPos, targetText) {
  const normalized = normalizeText(text);
  const target = normalizeText(targetText);
  
  // For single characters, use MUCH larger context sizes to ensure uniqueness
  const isSingleChar = target.length <= 3;
  const contextSizes = isSingleChar 
    ? [50, 100, 150, 200, 300, 500, 1000, 2000, normalized.length] // Much more aggressive for single chars
    : [20, 40, 60, 80, 100, 150, 200];
  
  for (const size of contextSizes) {
    const beforeStart = Math.max(0, startPos - size);
    const afterEnd = Math.min(normalized.length, endPos + size);
    
    const contextBefore = normalized.substring(beforeStart, startPos);
    const contextAfter = normalized.substring(endPos, afterEnd);
    const fullPattern = contextBefore + target + contextAfter;
    
    // Check if this pattern is unique in the text
    const firstMatch = normalized.indexOf(fullPattern);
    const lastMatch = normalized.lastIndexOf(fullPattern);
    
    if (firstMatch === lastMatch && firstMatch === beforeStart) {
      // Pattern is unique!
      console.log(`      ?? Found unique pattern with ${size} char context for "${target}"`);
      return {
        pattern: fullPattern,
        target: target,
        beforeContext: contextBefore,
        afterContext: contextAfter,
        isUnique: true,
        contextSize: size,
        positionInOriginal: startPos
      };
    }
  }
  
  // Could not find unique context even with maximum size - this is a critical error
  console.error(`      ? CRITICAL: Cannot find unique pattern for "${target}" at position ${startPos}`);
  console.error(`         This text appears in multiple identical contexts`);
  
  const size = Math.min(500, normalized.length);
  const beforeStart = Math.max(0, startPos - size);
  const afterEnd = Math.min(normalized.length, endPos + size);
  
  return {
    pattern: normalized.substring(beforeStart, afterEnd),
    target: target,
    beforeContext: normalized.substring(beforeStart, startPos),
    afterContext: normalized.substring(endPos, afterEnd),
    isUnique: false,
    contextSize: size,
    positionInOriginal: startPos
  };
}

/**
 * Apply diffs with maximum precision tracking
 * 
 * @param {Array} diffs - Array of {op: 'equal'|'delete'|'insert', text: string}
 * @param {boolean} wasTrackingOn - Whether tracking was on before
 * @param {Function} onProgress - Progress callback (current, total, message)
 */
export async function replaceSelectionWithNativeTrackedRevisions(diffs, wasTrackingOn, onProgress) {
  try {
    return await Word.run(async (context) => {
    const doc = context.document;
    const selection = doc.getSelection();
    
    // Load original text
    selection.load('text');
      await context.sync();
      
    const originalText = selection.text;
    
    // CRITICAL FIX: Replace newlines with placeholder so they can be handled as regular text
    // Build normalized text by replacing \r with placeholder
    // We'll work with this normalized version for all operations
    const normalized = normalizeText(originalText);
    
    console.log(`?? Text normalization: ${originalText.length} chars -> ${normalized.length} chars (newlines replaced with placeholder)`);
    
    console.log(`?? Starting maximum-precision tracking`);
    console.log(`   Processing ${diffs.length} diffs`);
    console.log(`   Original text: ${originalText.length} characters`);
    console.log(`   Normalized text: ${normalized.length} characters`);
    
    // Ensure tracking is ON
    doc.load('changeTrackingMode');
      await context.sync();
      
    const trackingIsOn = doc.changeTrackingMode === Word.ChangeTrackingMode.trackAll || 
                         doc.changeTrackingMode === 'TrackAll';
    
    if (!trackingIsOn) {
      doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      await context.sync();
      console.log('? Track changes enabled');
    }
    
    // Build position map with enhanced verification
    const positionMap = [];
    let currentPos = 0;
    let originalPos = 0; // Track position in original text (without placeholders)
    // normalized is already defined above (line 100)
    
    console.log('?? Building position map with uniqueness verification...');
    
    for (let i = 0; i < diffs.length; i++) {
      const diff = diffs[i];
      const normText = normalizeText(diff.text);
      
      if (diff.op === 'equal') {
        // CRITICAL: Don't process equal text - just track position
        // Equal text is UNCHANGED and should remain untouched
        const slice = normalized.slice(currentPos, currentPos + normText.length);
        if (slice === normText) {
          currentPos += normText.length;
          originalPos += diff.text.length; // Track original position
        } else {
          // Try to find nearby
          const searchIdx = normalized.indexOf(normText, Math.max(0, currentPos - 30));
          if (searchIdx >= 0 && searchIdx < currentPos + 150) {
            currentPos = searchIdx + normText.length;
            originalPos += diff.text.length; // Track original position
          } else {
            console.warn(`?? [Diff ${i}] Could not align equal text (skipping)`);
          }
        }
      } else if (diff.op === 'delete') {
        // CRITICAL: Don't skip newline deletions - they normalize to empty but must be handled specially
        // Check if this is a newline deletion (marked with isNewline flag or contains only newline chars)
        const isNewlineDelete = diff.isNewline || (diff.text === '\r' || diff.text === '\n' || diff.text === '\r\n');
        
        if (normText.length === 0 && !isNewlineDelete) {
          console.log(`? [Diff ${i}] Delete normalized to empty (non-newline) - auto-skipping`);
          continue; // Nothing to delete (but not a newline)
        }
        
        // For newline deletions, build context from ORIGINAL text (not normalized with placeholders)
        // because placeholders don't exist in the Word document
        if (isNewlineDelete) {
          // Build context from original text (without placeholders)
          // The newline is at originalPos in the original text
          const contextSize = 50;
          const beforeStart = Math.max(0, originalPos - contextSize);
          const afterEnd = Math.min(originalText.length, originalPos + diff.text.length + contextSize);
          
          // Get context from original text
          let beforeContext = originalText.substring(beforeStart, originalPos);
          let afterContext = originalText.substring(originalPos + diff.text.length, afterEnd);
          
          // Clean context: remove \r but don't add placeholders (just remove them)
          beforeContext = beforeContext.replace(/\r/g, '').normalize('NFC');
          afterContext = afterContext.replace(/\r/g, '').normalize('NFC');
          
          // For newlines, we don't need to verify uniqueness in normalized text
          // because we use paragraph-based matching, not text search
          // Just use the context we built from original text
          const uniqueContext = {
            pattern: beforeContext + NEWLINE_PLACEHOLDER + afterContext,
            target: NEWLINE_PLACEHOLDER,
            beforeContext: beforeContext,
            afterContext: afterContext,
            isUnique: false, // Not checked - we use paragraph matching instead
            contextSize: contextSize,
            positionInOriginal: originalPos
          };
          
          positionMap.push({
            index: i,
            op: 'delete',
            start: currentPos,
            end: currentPos + NEWLINE_PLACEHOLDER.length,
            text: NEWLINE_PLACEHOLDER,
            originalText: diff.text,
            context: uniqueContext,
            length: NEWLINE_PLACEHOLDER.length,
            isNewline: true
          });
          
          console.log(`? [Diff ${i}] Newline deletion with context (before: "${beforeContext.substring(Math.max(0, beforeContext.length - 20))}", after: "${afterContext.substring(0, Math.min(20, afterContext.length))}")`);
          currentPos += NEWLINE_PLACEHOLDER.length;
          originalPos += diff.text.length; // Track original position
          continue;
        }
        
        // Regular deletion - track original position
        originalPos += diff.text.length;
        
        // CRITICAL: Skip if this normalized to a placeholder (should have been caught above, but double-check)
        if (normText === NEWLINE_PLACEHOLDER) {
          console.warn(`?? [Diff ${i}] Placeholder detected in regular deletion path - this should not happen`);
          currentPos += NEWLINE_PLACEHOLDER.length;
          continue;
        }
        
        const slice = normalized.slice(currentPos, currentPos + normText.length);
        if (slice === normText) {
          // Find unique context for this deletion (use normalized text to avoid \r issues)
          const context = findUniqueContext(normalized, currentPos, currentPos + normText.length, diff.text);
          
          positionMap.push({
            index: i,
            op: 'delete',
            start: currentPos,
            end: currentPos + normText.length,
            text: diff.text,
            context: context,
            length: normText.length
          });
          
          if (!context.isUnique) {
            console.warn(`?? [Diff ${i}] Delete text not unique: "${diff.text}" (will try best match)`);
          } else {
            console.log(`? [Diff ${i}] Delete with unique context (size: ${context.contextSize})`);
          }
          
          currentPos += normText.length;
        } else {
          // Try to find nearby
          // CRITICAL: Skip if this is a placeholder (should have been caught above)
          if (normText === NEWLINE_PLACEHOLDER) {
            console.warn(`?? [Diff ${i}] Placeholder detected in fallback deletion path - this should not happen`);
            currentPos += NEWLINE_PLACEHOLDER.length;
            continue;
          }
          
          const searchIdx = normalized.indexOf(normText, Math.max(0, currentPos - 30));
          if (searchIdx >= 0 && searchIdx < currentPos + 150) {
            const context = findUniqueContext(normalized, searchIdx, searchIdx + normText.length, diff.text);
            
            positionMap.push({
              index: i,
              op: 'delete',
              start: searchIdx,
              end: searchIdx + normText.length,
              text: diff.text,
              context: context,
              length: normText.length
            });
            
            currentPos = searchIdx + normText.length;
          } else {
            console.error(`? [Diff ${i}] Cannot find delete text: "${diff.text.substring(0, 30)}..."`);
          }
        }
      } else if (diff.op === 'insert') {
        // Check if this is a newline insertion
        const isNewlineInsert = diff.isNewline || (diff.text === '\r' || diff.text === '\n' || diff.text === '\r\n');
        
        // For newline insertions, build context from ORIGINAL text (not normalized with placeholders)
        if (isNewlineInsert) {
          // Build context from original text at insertion point
          const contextSize = 50;
          const beforeStart = Math.max(0, originalPos - contextSize);
          const afterEnd = Math.min(originalText.length, originalPos + contextSize);
          
          // Get context from original text
          let beforeContext = originalText.substring(beforeStart, originalPos);
          let afterContext = originalText.substring(originalPos, afterEnd);
          
          // Clean context: remove \r but don't add placeholders
          beforeContext = beforeContext.replace(/\r/g, '').normalize('NFC');
          afterContext = afterContext.replace(/\r/g, '').normalize('NFC');
          
          const insertContext = {
            pattern: beforeContext + NEWLINE_PLACEHOLDER + afterContext,
            target: NEWLINE_PLACEHOLDER,
            beforeContext: beforeContext,
            afterContext: afterContext,
            isUnique: false, // Will be checked during insertion
            contextSize: contextSize,
            positionInOriginal: originalPos
          };
          
          positionMap.push({
            index: i,
            op: 'insert',
            position: currentPos,
            text: NEWLINE_PLACEHOLDER, // Use placeholder for newlines
            originalText: diff.text, // Keep original for reference
            context: insertContext,
            length: NEWLINE_PLACEHOLDER.length,
            isNewline: true // Mark as newline for final conversion
          });
          
          // Don't advance originalPos for insertions (they're not in the original text yet)
          currentPos += NEWLINE_PLACEHOLDER.length;
        } else {
          // Regular insertion - use normalized text for context
          const insertText = diff.text;
          const insertTextNormalized = normalizeText(insertText);
          
          // Find unique context for insertion point
          const insertContext = findUniqueContext(normalized, currentPos, currentPos, '');
          
          positionMap.push({
            index: i,
            op: 'insert',
            position: currentPos,
            text: insertText,
            originalText: diff.text,
            context: insertContext,
            length: insertTextNormalized.length,
            isNewline: false
          });
          
          // Don't advance originalPos for insertions
          currentPos += insertTextNormalized.length;
        }
        
        if (!insertContext.isUnique) {
          console.warn(`?? [Diff ${i}] Insert position not unique (will try best match)`);
              } else {
          console.log(`? [Diff ${i}] Insert with unique context (size: ${insertContext.contextSize})`);
        }
      }
    }
    
    console.log(`? Position map: ${positionMap.length} operations`);
    
    if (positionMap.length === 0) {
      console.log('?? No changes to apply');
      return;
    }
    
    // CRITICAL FIX: Split operations into deletions and insertions
    // Process deletions FIRST (reverse order), then insertions (reverse order)
    // Equal text is NOT processed - it remains unchanged
    
    const deletions = positionMap.filter(op => op.op === 'delete');
    const insertions = positionMap.filter(op => op.op === 'insert');
    
    // Sort deletions in REVERSE order (end to start)
    deletions.sort((a, b) => b.start - a.start);
    
    // Sort insertions in REVERSE order (end to start)  
    insertions.sort((a, b) => b.position - a.position);
    
    // Process deletions first, then insertions
    const orderedOps = [...deletions, ...insertions];
    
    console.log('?? Processing operations in two phases:');
    console.log(`   Phase 1: ${deletions.length} deletions (reverse order)`);
    console.log(`   Phase 2: ${insertions.length} insertions (reverse order)`);
    console.log(`   Note: Equal text (unchanged) is not processed`);
    
    let applied = 0;
    let skipped = 0;
    const totalOps = orderedOps.length;
    let currentPhase = null;
    
    for (const op of orderedOps) {
      try {
        const opNum = applied + skipped + 1;
        const phase = op.op === 'delete' ? 'PHASE 1' : 'PHASE 2';
        const prefix = `[${phase}] [${opNum}/${totalOps}]`;
        
        // Note: Tracking stays ON for insertions too
        // Insertions will be tracked (red underline) AND formatted blue
        if (phase === 'PHASE 2' && currentPhase !== 'PHASE 2') {
          console.log('');
          console.log('?? Entering PHASE 2: Insertions (tracking stays ON)');
          currentPhase = 'PHASE 2';
        }
        
        if (op.op === 'delete') {
            // Handle newline deletions as regular text (they're placeholders now)
            if (op.isNewline) {
              console.log(`${prefix} DELETE NEWLINE (as placeholder) at pos ${op.start}: "${op.text}"`);
              // Treat as regular text deletion - the placeholder is in the normalized text
              // We'll search for it using the context, but need to find the actual paragraph break location
              
              // Since placeholders aren't actually in the document, we need to find the paragraph break
              // by searching for the context and checking for paragraph boundaries
              let deletedSuccessfully = false;
              
              try {
                const body = context.document.body;
                // Context is already cleaned (no \r, no placeholders) - use it directly
                // Don't call normalizeText again as it might add placeholders
                const beforeCtx = op.context.beforeContext;
                const afterCtx = op.context.afterContext;
                
                // Search for before context + after context to find the location
                if (beforeCtx.length >= 5 && afterCtx.length >= 5) {
                  const searchText = beforeCtx.substring(Math.max(0, beforeCtx.length - 30)) + 
                                    afterCtx.substring(0, Math.min(30, afterCtx.length));
                  
                  // But wait - the search won't work because there's a paragraph break between them
                  // So search for before context, then check what comes after
                  const beforeSearch = beforeCtx.substring(Math.max(0, beforeCtx.length - 50));
                  const searchResults = body.search(beforeSearch, {
                    matchCase: false,
                    matchWholeWord: false
                  });
                  searchResults.load('items');
                  await context.sync();
                  
                  if (searchResults.items && searchResults.items.length > 0) {
                    const beforeRange = searchResults.items[0];
                    const paragraphs = body.paragraphs;
                    paragraphs.load('items');
                    await context.sync();
                    
                    // Find the paragraph containing beforeRange
                    for (let pIdx = 0; pIdx < paragraphs.items.length - 1; pIdx++) {
                      const para = paragraphs.items[pIdx];
                      para.load('text');
                      await context.sync();
                      
                      // Normalize paragraph text (remove \r but don't add placeholders)
                      const paraText = para.text.replace(/\r/g, '').normalize('NFC');
                      const beforeCtxEnd = beforeCtx.substring(Math.max(0, beforeCtx.length - 40));
                      
                      if (paraText.endsWith(beforeCtxEnd)) {
                        // Found the paragraph! Check next paragraph
                        if (pIdx < paragraphs.items.length - 1) {
                          const nextPara = paragraphs.items[pIdx + 1];
                          nextPara.load('text');
                          await context.sync();
                          
                          // Normalize next paragraph text (remove \r but don't add placeholders)
                          const nextParaText = nextPara.text.replace(/\r/g, '').normalize('NFC');
                          
                          // Check if next paragraph is empty (the newline)
                          if (nextParaText.trim().length === 0) {
                            nextPara.delete();
                            await context.sync();
                            deletedSuccessfully = true;
                            applied++;
                            console.log(`   ? Newline deleted (empty paragraph removed)`);
                            break;
                          }
                          
                          // Check if next paragraph matches after context
                          if (afterCtx.length >= 5) {
                            const afterCtxStart = afterCtx.substring(0, Math.min(40, afterCtx.length));
                            const nextParaTrimmed = nextParaText.trimStart();
                            const afterCtxTrimmed = afterCtxStart.trimStart();
                            
                            // More lenient matching - check if significant portion matches
                            const minMatchLen = Math.min(10, afterCtxTrimmed.length, nextParaTrimmed.length);
                            if (minMatchLen > 0 && nextParaTrimmed.substring(0, minMatchLen) === afterCtxTrimmed.substring(0, minMatchLen)) {
                              // Merge paragraphs to delete the newline
                              const nextParaTextContent = nextPara.text;
                              para.insertText(nextParaTextContent, Word.InsertLocation.end);
                              nextPara.delete();
                              await context.sync();
                              
                              deletedSuccessfully = true;
                              applied++;
                              console.log(`   ? Newline deleted successfully (paragraphs merged)`);
                              break;
                            } else {
                              console.log(`   ? Next paragraph doesn't match: "${nextParaTrimmed.substring(0, Math.min(20, nextParaTrimmed.length))}" vs "${afterCtxTrimmed.substring(0, Math.min(20, afterCtxTrimmed.length))}"`);
                            }
                          }
                        }
                      }
                    }
                  }
                }
                
                if (!deletedSuccessfully) {
                  skipped++;
                  console.warn(`   ?? Could not find newline to delete`);
                }
              } catch (e) {
                skipped++;
                console.error(`   ?? Error deleting newline: ${e.message}`);
              }
              
              continue; // Skip regular deletion processing
            }
            
            // Regular text deletion
            console.log(`${prefix} DELETE at pos ${op.start}: "${op.text.substring(0, 40)}..."`);
            console.log(`           Position in original: ${op.start}-${op.end}`);
            
            // Check if this is a placeholder (for newlines that are now regular text)
            const isPlaceholder = op.text === NEWLINE_PLACEHOLDER;
            const targetText = isPlaceholder ? NEWLINE_PLACEHOLDER : normalizeText(op.text);
            
            let deletedSuccessfully = false;
            
            // For placeholders, we need special handling since they're not actually in the document
            if (isPlaceholder) {
              // Placeholder represents a newline - find the paragraph break and merge paragraphs
              // Context is already cleaned (no \r, no placeholders) - use it directly
              const beforeCtx = op.context.beforeContext;
              const afterCtx = op.context.afterContext;
              
              console.log(`   ? Deleting placeholder (newline) - searching for paragraph break`);
              
              const body = context.document.body;
              
              // Search for before context to find the paragraph
              if (beforeCtx.length >= 5) {
                const beforeSearch = beforeCtx.substring(Math.max(0, beforeCtx.length - 50));
                const searchResults = body.search(beforeSearch, {
                  matchCase: false,
                  matchWholeWord: false
                });
                searchResults.load('items');
                await context.sync();
                
                if (searchResults.items && searchResults.items.length > 0) {
                  const paragraphs = body.paragraphs;
                  paragraphs.load('items');
                  await context.sync();
                  
                  // Find paragraph that ends with before context
                  for (let pIdx = 0; pIdx < paragraphs.items.length - 1; pIdx++) {
                    const para = paragraphs.items[pIdx];
                    para.load('text');
                    await context.sync();
                    
                    // Normalize paragraph text the same way as context (remove \r, no placeholders)
                    const paraText = para.text.replace(/\r/g, '').normalize('NFC');
                    const beforeCtxEnd = beforeCtx.substring(Math.max(0, beforeCtx.length - 40));
                    
                    if (paraText.endsWith(beforeCtxEnd)) {
                      // Found the paragraph! Check next paragraph
                      if (pIdx < paragraphs.items.length - 1) {
                        const nextPara = paragraphs.items[pIdx + 1];
                        nextPara.load('text');
                        await context.sync();
                        
                        // Normalize next paragraph text the same way (remove \r, no placeholders)
                        const nextParaText = nextPara.text.replace(/\r/g, '').normalize('NFC');
                        
                        // Check if next paragraph is empty (the newline)
                        if (nextParaText.trim().length === 0) {
                          nextPara.delete();
                          await context.sync();
                          deletedSuccessfully = true;
                          applied++;
                          console.log(`   ? Newline deleted (empty paragraph removed)`);
                          break;
                        }
                        
                        // Check if next paragraph matches after context
                        if (afterCtx.length >= 5) {
                          const afterCtxStart = afterCtx.substring(0, Math.min(40, afterCtx.length));
                          const nextParaTrimmed = nextParaText.trimStart();
                          const afterCtxTrimmed = afterCtxStart.trimStart();
                          
                          // More lenient matching - check if significant portion matches
                          const minMatchLen = Math.min(10, afterCtxTrimmed.length, nextParaTrimmed.length);
                          if (minMatchLen > 0 && nextParaTrimmed.substring(0, minMatchLen) === afterCtxTrimmed.substring(0, minMatchLen)) {
                            // Merge paragraphs
                            const nextParaTextContent = nextPara.text;
                            para.insertText(nextParaTextContent, Word.InsertLocation.end);
                            nextPara.delete();
                            await context.sync();
                            deletedSuccessfully = true;
                            applied++;
                            console.log(`   ? Newline deleted (paragraphs merged)`);
                            break;
                          }
                        }
                      } else {
                        // This is the last paragraph - check if after context is empty (trailing newline)
                        if (afterCtx.length === 0 || afterCtx.trim().length === 0) {
                          // Trailing newline - already handled (implicit paragraph break)
                          deletedSuccessfully = true;
                          applied++;
                          console.log(`   ? Trailing newline handled (implicit paragraph break)`);
                          break;
                        }
                      }
                      
                      // Fallback: Check if the paragraph already contains the after context
                      // (meaning paragraphs were already merged by a previous operation)
                      if (!deletedSuccessfully && afterCtx.length >= 5) {
                        const afterCtxStart = afterCtx.substring(0, Math.min(30, afterCtx.length));
                        const paraTextAfter = paraText.substring(paraText.length - Math.min(100, paraText.length));
                        if (paraTextAfter.includes(afterCtxStart)) {
                          // Paragraph already contains after context - newline already deleted
                          deletedSuccessfully = true;
                          applied++;
                          console.log(`   ? Newline already deleted (paragraphs already merged)`);
                          break;
                        }
                      }
                    }
                  }
                }
              }
              
              if (!deletedSuccessfully) {
                skipped++;
                console.warn(`   ?? Could not find newline/paragraph break to delete`);
              }
              
              continue; // Skip regular deletion processing for placeholders
            }
            
            // Regular text deletion logic
            // Strategy 1: Search with full unique pattern
            if (op.context.isUnique) {
              console.log(`   ? Using GUARANTEED unique pattern (${op.context.pattern.length} chars)`);
              
              // CRITICAL: For patterns longer than 255 chars, CENTER the window around the target
              let searchPattern = op.context.pattern;
              const targetText = normalizeText(op.text);
              
              if (searchPattern.length > 255) {
                // Find where the target is in the pattern
                const targetPosInPattern = searchPattern.indexOf(targetText);
                
                if (targetPosInPattern >= 0) {
                  // Center the 255-char window around the target
                  const targetCenter = targetPosInPattern + Math.floor(targetText.length / 2);
                  const windowStart = Math.max(0, targetCenter - 127); // 127 chars before center
                  const windowEnd = Math.min(searchPattern.length, windowStart + 255);
                  
                  searchPattern = searchPattern.substring(windowStart, windowEnd);
                  console.log(`      Pattern truncated to 255 chars centered on target (target at offset ${targetPosInPattern - windowStart})`);
                } else {
                  // Fallback: use last 255 chars
                  searchPattern = searchPattern.substring(searchPattern.length - 255);
                  console.log(`      Pattern truncated to last 255 chars (target not found in pattern!)`);
                }
              }
              
              // CRITICAL FIX: Search in document body, not selection
              // Selection text changes after each deletion!
              const body = context.document.body;
              const searchResults = body.search(searchPattern, {
            matchCase: false,
            matchWholeWord: false
          });
          searchResults.load('items');
            await context.sync();
      
              // Filter to only matches within original selection range
              let matchesInSelection = [];
          if (searchResults.items && searchResults.items.length > 0) {
                for (let i = 0; i < searchResults.items.length; i++) {
                  const item = searchResults.items[i];
              item.load('text');
                }
              await context.sync();
                matchesInSelection = searchResults.items; // Use first match for now
              }
              
              if (matchesInSelection.length > 0) {
                const matchRange = searchResults.items[0];
                
                // CRITICAL: We found the unique pattern!
                // The target MUST be within this range at a known offset
                // Calculate where the target is in the ORIGINAL searchPattern
                const targetOffsetInPattern = searchPattern.indexOf(targetText);
                
                console.log(`      Pattern match found, target at offset ${targetOffsetInPattern} in search pattern`);
                
                if (targetOffsetInPattern >= 0) {
                  let deletedSuccessfully = false;
                  
                  // Direct approach: Search for target in the document body with additional context
                  // Use the before and after context from the pattern to ensure correct match
                  // CRITICAL: Since we process deletions in reverse order, prefer after context
                  // as it's less likely to have been modified by previous deletions
                  const beforeCtx = normalizeText(op.context.beforeContext);
                  const afterCtx = normalizeText(op.context.afterContext);
                  
                  // Strategy 1: Build specific pattern with target and some context
                  const contextSizes = [30, 20, 15, 10, 5];
                  
                  for (const size of contextSizes) {
                    if (deletedSuccessfully) break;
                    
                    const beforeLen = Math.min(size, beforeCtx.length);
                    const afterLen = Math.min(size, afterCtx.length);
                    
                    // Build pattern: [before context] + [TARGET] + [after context]
                    const before = beforeCtx.slice(-beforeLen);
                    const after = afterCtx.slice(0, afterLen);
                    const fullPattern = before + targetText + after;
                    
                    if (fullPattern.length > targetText.length + 3) { // Need meaningful context
                      const patternPreview = fullPattern.length > 60 
                        ? `${fullPattern.substring(0, 25)}...${fullPattern.substring(fullPattern.length - 25)}`
                        : fullPattern;
                      console.log(`      Strategy 1 (ctx ${size}): "${patternPreview}"`);
                      
                      try {
                        const body = context.document.body;
                        const specificSearch = body.search(fullPattern.substring(0, 255), {matchCase: false, matchWholeWord: false});
                        specificSearch.load('items');
                        await context.sync();
                        
                        if (specificSearch.items && specificSearch.items.length > 0) {
                          // Found the pattern! Now we know EXACTLY where the target is
                          // The target is at offset beforeLen within this pattern
                          const patternRange = specificSearch.items[0];
                          patternRange.load('text');
                          await context.sync();
                          
                          const patternText = normalizeText(patternRange.text);
                          const targetOffset = patternText.indexOf(targetText);
                          
                          console.log(`         Pattern found, target at offset ${targetOffset} (expected ${beforeLen})`);
                          
                          // Verify the target is at the expected position within pattern
                          // More lenient tolerance for Word's text representation quirks
                          const positionTolerance = Math.max(5, Math.ceil(beforeLen * 0.2)); // 20% tolerance, min 5 chars
                          if (Math.abs(targetOffset - beforeLen) <= positionTolerance) {
                            // Search for target WITHIN this specific pattern range
                            const targetSearch = patternRange.search(targetText, {matchCase: false, matchWholeWord: false});
                            targetSearch.load('items');
                            await context.sync();
                            
                            // Find the match that's at the correct offset
                            if (targetSearch.items && targetSearch.items.length > 0) {
                              let correctItem = null;
                              
                              if (targetSearch.items.length === 1) {
                                // Only one match - use it
                                correctItem = targetSearch.items[0];
                              } else {
                                // Multiple matches within pattern - need to find the right one
                                console.log(`         Multiple target matches (${targetSearch.items.length}) within pattern, analyzing positions...`);
                                
                                // Load ranges to get relative positions
                                for (const item of targetSearch.items) {
                                  item.load('text');
                                }
                                await context.sync();
                                
                                // The correct match should be the one that appears after the "before context"
                                // We need to find which match in the pattern is at the right position
                                // Since targetOffset tells us where the FIRST occurrence is, we want that one if it's close to beforeLen
                                
                                if (Math.abs(targetOffset - beforeLen) <= 3) {
                                  // The first occurrence in the pattern text matches our expected position
                                  correctItem = targetSearch.items[0];
                                  console.log(`         Selected first match (at expected offset ${targetOffset})`);
                                } else {
                                  // Need to be more careful - search in pattern text to find all occurrences
                                  let allPositions = [];
                                  let searchPos = 0;
                                  while ((searchPos = patternText.indexOf(targetText, searchPos)) !== -1) {
                                    allPositions.push(searchPos);
                                    searchPos += 1;
                                  }
                                  
                                  console.log(`         Target appears at positions: ${allPositions.join(', ')} in pattern (expected near ${beforeLen})`);
                                  
                                  // Find the position closest to beforeLen
                                  let bestIdx = 0;
                                  let bestDistance = Math.abs(allPositions[0] - beforeLen);
                                  
                                  for (let i = 1; i < allPositions.length && i < targetSearch.items.length; i++) {
                                    const distance = Math.abs(allPositions[i] - beforeLen);
                                    if (distance < bestDistance) {
                                      bestDistance = distance;
                                      bestIdx = i;
                                    }
                                  }
                                  
                                  correctItem = targetSearch.items[Math.min(bestIdx, targetSearch.items.length - 1)];
                                  console.log(`         Selected match ${bestIdx} at position ${allPositions[bestIdx]} (closest to expected ${beforeLen})`);
                                }
                              }
                              
                              if (correctItem) {
                                correctItem.delete();
                                await context.sync();
                                applied++;
                                deletedSuccessfully = true;
                                console.log(`   ? Deleted (strategy 1, context ${size})`);
                              }
                            }
                          } else {
                            console.warn(`         Position mismatch (off by ${Math.abs(targetOffset - beforeLen)}), trying smaller context`);
                          }
                        }
                      } catch (e) {
                        console.warn(`      Strategy 1 (size ${size}) failed: ${e.message}`);
                      }
                    }
                  }
                  
                  if (deletedSuccessfully) {
                    continue;
                  }
                  
                  // Strategy 2: Use before+after context more aggressively
                  console.log(`      Strategy 2: Extended context search`);
                  
                  // Try to find a larger unique string that includes the target
                  const extendedBefore = beforeCtx.slice(-Math.min(50, beforeCtx.length));
                  const extendedAfter = afterCtx.slice(0, Math.min(50, afterCtx.length));
                  
                  // Build multiple search patterns to try
                  const searchPatterns = [];
                  
                  // Pattern 1: Large context on both sides
                  if (extendedBefore.length >= 10 && extendedAfter.length >= 10) {
                    searchPatterns.push({
                      pattern: extendedBefore + targetText + extendedAfter,
                      beforeLen: extendedBefore.length,
                      desc: 'large before+after context'
                    });
                  }
                  
                  // Pattern 2: Large before, small after
                  if (extendedBefore.length >= 15) {
                    const smallAfter = extendedAfter.slice(0, Math.min(5, extendedAfter.length));
                    searchPatterns.push({
                      pattern: extendedBefore + targetText + smallAfter,
                      beforeLen: extendedBefore.length,
                      desc: 'large before context'
                    });
                  }
                  
                  // Pattern 3: Small before, large after
                  if (extendedAfter.length >= 15) {
                    const smallBefore = extendedBefore.slice(-Math.min(5, extendedBefore.length));
                    searchPatterns.push({
                      pattern: smallBefore + targetText + extendedAfter,
                      beforeLen: smallBefore.length,
                      desc: 'large after context'
                    });
                  }
                  
                  try {
                    for (const {pattern, beforeLen, desc} of searchPatterns) {
                      if (deletedSuccessfully) break;
                      
                      console.log(`      Trying ${desc}: "${pattern.substring(0, 20)}...${pattern.substring(Math.max(0, pattern.length - 20))}"`);
                      
                      const body = context.document.body;
                      const patternSearch = body.search(pattern.substring(0, 255), {matchCase: false, matchWholeWord: false});
                      patternSearch.load('items');
                      await context.sync();
                      
                      if (patternSearch.items && patternSearch.items.length > 0) {
                        console.log(`      Found pattern, searching for target within it`);
                        
                        const patternRange = patternSearch.items[0];
                        const targetSearch = patternRange.search(targetText, {matchCase: false, matchWholeWord: false});
                        targetSearch.load('items');
                        await context.sync();
                        
                        if (targetSearch.items && targetSearch.items.length > 0) {
                          // Use first match (should be the correct one given the specific pattern)
                          targetSearch.items[0].delete();
                          await context.sync();
                          applied++;
                          deletedSuccessfully = true;
                          console.log(`   ? Deleted (strategy 2, ${desc})`);
                        }
                      }
                    }
                  } catch (e) {
                    console.warn(`      Strategy 2 failed: ${e.message}`);
                  }
                  
                  if (deletedSuccessfully) {
                    continue;
                  }
                  
                  // Strategy 3: Last resort - direct search with position-based selection
                  console.log(`      Strategy 3 (last resort): Direct search`);
                  try {
                    const body = context.document.body;
                    const directSearch = body.search(targetText, {matchCase: false, matchWholeWord: false});
                    directSearch.load('items');
                    await context.sync();
                    
                    if (directSearch.items && directSearch.items.length > 0) {
                      console.log(`      Found ${directSearch.items.length} matches, selecting based on deletion order`);
                      
                      let selectedMatch = null;
                      
                      // If multiple matches, try to use after context to narrow it down
                      if (directSearch.items.length > 1) {
                        const afterCtx = normalizeText(op.context.afterContext);
                        if (afterCtx.length >= 5) {
                          // Search for target + after context to find the correct match
                          const combinedPattern = targetText + afterCtx.substring(0, Math.min(20, afterCtx.length));
                          try {
                            const body = context.document.body;
                            const combinedSearch = body.search(combinedPattern.substring(0, 255), {
                              matchCase: false,
                              matchWholeWord: false
                            });
                            combinedSearch.load('items');
                            await context.sync();
                            
                            if (combinedSearch.items && combinedSearch.items.length > 0) {
                              // Found the combined pattern - search for target within it
                              const combinedRange = combinedSearch.items[0];
                              const targetInCombined = combinedRange.search(targetText, {
                                matchCase: false,
                                matchWholeWord: false
                              });
                              targetInCombined.load('items');
                              await context.sync();
                              
                              if (targetInCombined.items && targetInCombined.items.length > 0) {
                                selectedMatch = targetInCombined.items[0];
                                console.log(`      Selected match using target+after-context pattern`);
                              }
                            }
                          } catch (e) {
                            console.warn(`      Combined search failed: ${e.message}`);
                          }
                        }
                      }
                      
                      // Fallback: Since we process in REVERSE order (end to start), use the LAST match
                      // as it's most likely to be the one we haven't deleted yet
                      if (!selectedMatch) {
                        selectedMatch = directSearch.items[directSearch.items.length - 1];
                      }
                      
                      selectedMatch.delete();
                      await context.sync();
                      applied++;
                      deletedSuccessfully = true;
                      console.log(`   ? Deleted (strategy 3, match ${directSearch.items.length} of ${directSearch.items.length})`);
                    }
                  } catch (e) {
                    console.warn(`      Strategy 3 failed: ${e.message}`);
                  }
                  
                  if (deletedSuccessfully) {
                    continue;
                  }
                  
                  console.warn(`      ?? All strategies failed, skipping deletion`);
                  skipped++;
                  continue;
                } else {
                  skipped++;
                  console.warn(`   ?? Target not found in matched pattern`);
                }
              } else {
                // Pattern not found - try multiple fallback strategies
                console.log(`   ?? Full pattern not found, trying fallback strategies`);
                const afterCtx = normalizeText(op.context.afterContext);
                const beforeCtx = normalizeText(op.context.beforeContext);
                const targetText = normalizeText(op.text);
                let deletedSuccessfully = false;
                
                // Fallback 1: Try with after context (even if short)
                if (!deletedSuccessfully && afterCtx.length >= 3) {
                  console.log(`   Fallback 1: Trying after-context search (${afterCtx.length} chars)`);
                  try {
                    const afterOnlyPattern = targetText + afterCtx.substring(0, Math.min(50, afterCtx.length));
                    const body = context.document.body;
                    const afterSearch = body.search(afterOnlyPattern.substring(0, 255), {
                      matchCase: false,
                      matchWholeWord: false
                    });
                    afterSearch.load('items');
                    await context.sync();
                    
                    if (afterSearch.items && afterSearch.items.length > 0) {
                      const afterRange = afterSearch.items[0];
                      const targetSearch = afterRange.search(targetText, {matchCase: false, matchWholeWord: false});
                      targetSearch.load('items');
                      await context.sync();
                      
                      if (targetSearch.items && targetSearch.items.length > 0) {
                        targetSearch.items[0].delete();
                        await context.sync();
                        applied++;
                        deletedSuccessfully = true;
                        console.log(`   ? Deleted (fallback 1: after-context)`);
                      }
                    }
                  } catch (e) {
                    console.warn(`   ?? Fallback 1 failed: ${e.message}`);
                  }
                }
                
                // Fallback 2: Try with before context (if after failed)
                if (!deletedSuccessfully && beforeCtx.length >= 3) {
                  console.log(`   Fallback 2: Trying before-context search (${beforeCtx.length} chars)`);
                  try {
                    const beforeOnlyPattern = beforeCtx.substring(Math.max(0, beforeCtx.length - 50)) + targetText;
                    const body = context.document.body;
                    const beforeSearch = body.search(beforeOnlyPattern.substring(0, 255), {
                      matchCase: false,
                      matchWholeWord: false
                    });
                    beforeSearch.load('items');
                    await context.sync();
                    
                    if (beforeSearch.items && beforeSearch.items.length > 0) {
                      const beforeRange = beforeSearch.items[0];
                      const targetSearch = beforeRange.search(targetText, {matchCase: false, matchWholeWord: false});
                      targetSearch.load('items');
                      await context.sync();
                      
                      if (targetSearch.items && targetSearch.items.length > 0) {
                        // Use the last match (closest to end of before context)
                        const matchIndex = Math.min(targetSearch.items.length - 1, Math.max(0, targetSearch.items.length - 1));
                        targetSearch.items[matchIndex].delete();
                        await context.sync();
                        applied++;
                        deletedSuccessfully = true;
                        console.log(`   ? Deleted (fallback 2: before-context)`);
                      }
                    }
                  } catch (e) {
                    console.warn(`   ?? Fallback 2 failed: ${e.message}`);
                  }
                }
                
                // Fallback 3: Try searching for meaningful part (without leading whitespace)
                if (!deletedSuccessfully) {
                  console.log(`   Fallback 3: Trying meaningful-part search`);
                  try {
                    // Try searching for the meaningful part (trimmed or first word)
                    const trimmedTarget = targetText.trim();
                    const firstWord = targetText.match(/[^\s\r\n]+/)?.[0]; // First non-whitespace sequence
                    
                    const searchTargets = [];
                    if (trimmedTarget.length > 0 && trimmedTarget !== targetText) {
                      searchTargets.push({text: trimmedTarget, desc: 'trimmed'});
                    }
                    if (firstWord && firstWord !== targetText && firstWord.length >= 3) {
                      searchTargets.push({text: firstWord, desc: 'first-word'});
                    }
                    
                    for (const {text: searchText, desc} of searchTargets) {
                      if (deletedSuccessfully) break;
                      
                      const body = context.document.body;
                      const partSearch = body.search(searchText, {matchCase: false, matchWholeWord: false});
                      partSearch.load('items');
                      await context.sync();
                      
                      if (partSearch.items && partSearch.items.length > 0) {
                        // Use the last match (since we process in reverse order)
                        const match = partSearch.items[partSearch.items.length - 1];
                        
                        // Try to find the full target text near this match
                        // Search in document body with the meaningful part + target
                        try {
                          match.load('text');
                          await context.sync();
                          
                          if (!match.text) {
                            throw new Error('Match text is empty');
                          }
                          
                          // Check if target has leading whitespace that we need to delete
                          const leadingWhitespace = targetText.match(/^[\s\r\n]+/)?.[0] || '';
                          
                          if (leadingWhitespace.length > 0) {
                            // CRITICAL: Try to delete whitespace FIRST, then meaningful part
                            // This is more reliable because the document structure is still intact
                            
                            // Step 1: Try to find and delete the whitespace BEFORE deleting the meaningful part
                            const whitespaceOnly = leadingWhitespace.replace(/\r/g, ''); // Remove \r for search
                            
                            if (whitespaceOnly.length > 0) {
                              // Search for before context + whitespace + meaningful part
                              const beforeCtx = op.context.beforeContext.substring(Math.max(0, op.context.beforeContext.length - 20));
                              const fullPattern = beforeCtx + whitespaceOnly + searchText;
                              
                              try {
                                const fullPatternSearch = context.document.body.search(fullPattern.substring(0, 255), {
                                  matchCase: false,
                                  matchWholeWord: false
                                });
                                fullPatternSearch.load('items');
                                await context.sync();
                                
                                if (fullPatternSearch.items && fullPatternSearch.items.length > 0) {
                                  // Found the full pattern! Now search for the full target within it
                                  const fullRange = fullPatternSearch.items[0];
                                  
                                  // Search for the complete target (whitespace + meaningful part) in this range
                                  const completeTarget = whitespaceOnly + searchText;
                                  const targetInRange = fullRange.search(completeTarget, {
                                    matchCase: false,
                                    matchWholeWord: false
                                  });
                                  targetInRange.load('items');
                                  await context.sync();
                                  
                                  if (targetInRange.items && targetInRange.items.length > 0) {
                                    // Delete the complete target (whitespace + part)
                                    targetInRange.items[0].delete();
                                    await context.sync();
                                    applied++;
                                    deletedSuccessfully = true;
                                    console.log(`   ? Deleted (fallback 3: ${desc} search, full target with whitespace)`);
                                    break;
                                  }
                                }
                              } catch (fullPatternError) {
                                console.warn(`   ?? Full pattern search failed: ${getErrorMessage(fullPatternError)}`);
                              }
                              
                              // Step 2: If full pattern failed, try deleting whitespace separately BEFORE meaningful part
                              if (!deletedSuccessfully) {
                                // Search for whitespace using before context
                                const beforeCtx = op.context.beforeContext.substring(Math.max(0, op.context.beforeContext.length - 20));
                                const wsPattern = beforeCtx + whitespaceOnly;
                                
                                try {
                                  const wsSearch = context.document.body.search(wsPattern.substring(0, 255), {
                                    matchCase: false,
                                    matchWholeWord: false
                                  });
                                  wsSearch.load('items');
                                  await context.sync();
                                  
                                  if (wsSearch.items && wsSearch.items.length > 0) {
                                    const wsMatch = wsSearch.items[wsSearch.items.length - 1];
                                    
                                    // Try to find just the whitespace within this match
                                    const wsInMatch = wsMatch.search(whitespaceOnly, {
                                      matchCase: false,
                                      matchWholeWord: false
                                    });
                                    wsInMatch.load('items');
                                    await context.sync();
                                    
                                    if (wsInMatch.items && wsInMatch.items.length > 0) {
                                      // Delete whitespace first
                                      wsInMatch.items[0].delete();
                                      await context.sync();
                                      console.log(`   ? Deleted leading whitespace first (${whitespaceOnly.length} chars)`);
                                      
                                      // CRITICAL: After deleting whitespace, the 'match' range may be invalid
                                      // Re-search for the meaningful part to get a fresh, valid range
                                      try {
                                        const body = context.document.body;
                                        const refreshedSearch = body.search(searchText, {matchCase: false, matchWholeWord: false});
                                        refreshedSearch.load('items');
                                        await context.sync();
                                        
                                        if (refreshedSearch.items && refreshedSearch.items.length > 0) {
                                          // Use the last match (since we process in reverse order)
                                          const refreshedMatch = refreshedSearch.items[refreshedSearch.items.length - 1];
                                          refreshedMatch.delete();
                                          await context.sync();
                                          applied++;
                                          deletedSuccessfully = true;
                                          console.log(`   ? Deleted (fallback 3: ${desc} search, whitespace + part separately)`);
                                          break;
                                        } else {
                                          console.warn(`   ?? Could not find meaningful part after whitespace deletion`);
                                        }
                                      } catch (refreshError) {
                                        console.warn(`   ?? Failed to refresh search after whitespace deletion: ${getErrorMessage(refreshError)}`);
                                      }
                                    }
                                  }
                                } catch (wsError) {
                                  console.warn(`   ?? Whitespace deletion failed: ${getErrorMessage(wsError)}`);
                                }
                              }
                            }
                          }
                          
                          // If we still haven't deleted, delete just the meaningful part
                          if (!deletedSuccessfully) {
                            // Verify match is still valid before using it
                            try {
                              match.load('text');
                              await context.sync();
                              
                              if (match.text !== null && match.text !== undefined) {
                                match.delete();
                                await context.sync();
                                applied++;
                                deletedSuccessfully = true;
                                if (leadingWhitespace.length > 0) {
                                  console.log(`   ? Deleted (fallback 3: ${desc} search, partial match - whitespace cleanup attempted)`);
                                } else {
                                  console.log(`   ? Deleted (fallback 3: ${desc} search, partial match)`);
                                }
                                break;
                              } else {
                                // Match is invalid, re-search
                                throw new Error('Match range is invalid, need to re-search');
                              }
                            } catch (matchError) {
                              // Match might be invalid, re-search for it
                              console.warn(`   ?? Match may be invalid, re-searching: ${getErrorMessage(matchError)}`);
                              const body = context.document.body;
                              const refreshedSearch = body.search(searchText, {matchCase: false, matchWholeWord: false});
                              refreshedSearch.load('items');
                              await context.sync();
                              
                              if (refreshedSearch.items && refreshedSearch.items.length > 0) {
                                const refreshedMatch = refreshedSearch.items[refreshedSearch.items.length - 1];
                                refreshedMatch.delete();
                                await context.sync();
                                applied++;
                                deletedSuccessfully = true;
                                console.log(`   ? Deleted (fallback 3: ${desc} search, partial match - re-searched)`);
                                break;
                              }
                            }
                          }
                        } catch (expandError) {
                          // If expansion fails, try to delete the match, but re-search if needed
                          const expandErrorMsg = getErrorMessage(expandError);
                          console.warn(`   ?? Expansion failed: ${expandErrorMsg}`);
                          try {
                            // Try to verify match is valid first
                            match.load('text');
                            await context.sync();
                            
                            if (match.text !== null && match.text !== undefined) {
                              match.delete();
                              await context.sync();
                              applied++;
                              deletedSuccessfully = true;
                              console.log(`   ? Deleted (fallback 3: ${desc} search, partial match)`);
                              break;
                            } else {
                              throw new Error('Match range is invalid');
                            }
                          } catch (deleteError) {
                            // Match is invalid, try re-searching
                            try {
                              const body = context.document.body;
                              const refreshedSearch = body.search(searchText, {matchCase: false, matchWholeWord: false});
                              refreshedSearch.load('items');
                              await context.sync();
                              
                              if (refreshedSearch.items && refreshedSearch.items.length > 0) {
                                const refreshedMatch = refreshedSearch.items[refreshedSearch.items.length - 1];
                                refreshedMatch.delete();
                                await context.sync();
                                applied++;
                                deletedSuccessfully = true;
                                console.log(`   ? Deleted (fallback 3: ${desc} search, partial match - re-searched after error)`);
                                break;
                              } else {
                                const deleteErrorMsg = getErrorMessage(deleteError);
                                console.error(`   ❌ Failed to delete match: ${deleteErrorMsg}`);
                                throw deleteError;
                              }
                            } catch (refreshError) {
                              const deleteErrorMsg = getErrorMessage(deleteError);
                              console.error(`   ❌ Failed to delete match: ${deleteErrorMsg}`);
                              throw deleteError;
                            }
                          }
                        }
                      }
                    }
                  } catch (e) {
                    console.warn(`   ?? Fallback 3 (meaningful-part) failed: ${getErrorMessage(e)}`);
                  }
                }
                
                // Fallback 4: Direct search (last resort)
                if (!deletedSuccessfully) {
                  console.log(`   Fallback 4: Trying direct search`);
                  try {
                    const body = context.document.body;
                    const directSearch = body.search(targetText, {matchCase: false, matchWholeWord: false});
                    directSearch.load('items');
                    await context.sync();
                    
                    if (directSearch.items && directSearch.items.length > 0) {
                      // Since we process in reverse order, use the last match
                      const selectedMatch = directSearch.items[directSearch.items.length - 1];
                      selectedMatch.delete();
                      await context.sync();
                      applied++;
                      deletedSuccessfully = true;
                      console.log(`   ? Deleted (fallback 4: direct search, match ${directSearch.items.length} of ${directSearch.items.length})`);
                    }
                  } catch (e) {
                    console.warn(`   ?? Fallback 4 failed: ${getErrorMessage(e)}`);
                  }
                }
                
                if (!deletedSuccessfully) {
                skipped++;
                  console.warn(`   ?? All fallbacks failed, pattern not found in document`);
                }
              }
            } else {
              skipped++;
              console.warn(`   ?? No unique pattern found`);
            }
          
        } else if (op.op === 'insert') {
          // SPECIAL HANDLING: Newline insertions must use Word insertBreak() API
          if (op.isNewline) {
            console.log(`${prefix} INSERT NEWLINE at pos ${op.position}`);
            console.log(`           Using Word insertBreak() API for newline insertion`);
            
            let insertedSuccessfully = false;
            
            try {
              const body = context.document.body;
              const beforeCtx = normalizeText(op.context.beforeContext);
              const afterCtx = normalizeText(op.context.afterContext);
              
              // Try to find the insertion point using before context
              // CRITICAL: Since insertions happen in reverse order, other insertions may have already occurred
              // First try searching for before context + common insertion characters (like ":")
              if (beforeCtx.length >= 5) {
                const baseSearchText = beforeCtx.substring(Math.max(0, beforeCtx.length - 50));
                const commonInsertions = [':', ';', ',', '.', '!', '?'];
                let foundRange = null;
                let foundChar = null;
                
                // First, try to find before context + insertion character (most likely scenario)
                for (const insertChar of commonInsertions) {
                  const searchWithChar = baseSearchText + insertChar;
                  const searchResults = body.search(searchWithChar, {
                    matchCase: false,
                    matchWholeWord: false
                  });
                  searchResults.load('items');
                  await context.sync();
                  
                  if (searchResults.items && searchResults.items.length > 0) {
                    foundRange = searchResults.items[0];
                    foundChar = insertChar;
                    console.log(`      Found before context + "${insertChar}" - will insert newline after this`);
                    break;
                  }
                }
                
                // If not found with insertion character, try just the before context
                if (!foundRange) {
                  const searchResults = body.search(baseSearchText, {
                    matchCase: false,
                    matchWholeWord: false
                  });
                  searchResults.load('items');
                  await context.sync();
                  
                  if (searchResults.items && searchResults.items.length > 0) {
                    foundRange = searchResults.items[0];
                    console.log(`      Found before context only - will insert newline after it`);
                  }
                }
                
                if (foundRange) {
                  // Get the range after the found text
                  const insertPoint = foundRange.getRange(Word.RangeLocation.after);
                  
                  // Insert the newline (paragraph break)
                  try {
                    // Try insertBreak first (line break)
                    insertPoint.insertBreak(Word.BreakType.line, Word.InsertLocation.before);
                    await context.sync();
                    insertedSuccessfully = true;
                    applied++;
                    console.log(`   ? Newline inserted successfully using insertBreak(line)${foundChar ? ` after "${foundChar}"` : ''}`);
                  } catch (breakError) {
                    // insertBreak might not be available, use insertParagraph
                    console.log(`      insertBreak not available, using insertParagraph`);
                    const newPara = insertPoint.insertParagraph('', Word.InsertLocation.before);
                    newPara.font.color = '#0000FF'; // Blue for insertions
                    await context.sync();
                    insertedSuccessfully = true;
                    applied++;
                    console.log(`   ? Newline inserted successfully using insertParagraph${foundChar ? ` after "${foundChar}"` : ''}`);
                  }
                }
              }
              
              // Fallback: Try after context if before context failed
              if (!insertedSuccessfully && afterCtx.length >= 5) {
                const searchText = afterCtx.substring(0, Math.min(50, afterCtx.length));
                const searchResults = body.search(searchText, {
                  matchCase: false,
                  matchWholeWord: false
                });
                searchResults.load('items');
                await context.sync();
                
                if (searchResults.items && searchResults.items.length > 0) {
                  const afterRange = searchResults.items[0];
                  
                  // Get the range before the after context (where the newline should be inserted)
                  const insertPoint = afterRange.getRange(Word.RangeLocation.before);
                  
                  // Insert a paragraph break by inserting an empty paragraph
                  const newPara = insertPoint.insertParagraph('', Word.InsertLocation.before);
                  newPara.font.color = '#0000FF';
                  
                  await context.sync();
                  
                  insertedSuccessfully = true;
                  applied++;
                  console.log(`   ? Newline inserted successfully using insertBreak() (via after context)`);
                }
              }
              
              // Last resort: Try shorter context
              if (!insertedSuccessfully && beforeCtx.length >= 3) {
                const searchText = beforeCtx.substring(Math.max(0, beforeCtx.length - 10));
                const searchResults = body.search(searchText, {
                  matchCase: false,
                  matchWholeWord: false
                });
                searchResults.load('items');
                await context.sync();
                
                if (searchResults.items && searchResults.items.length > 0) {
                  const beforeRange = searchResults.items[searchResults.items.length - 1];
                  const insertPoint = beforeRange.getRange(Word.RangeLocation.after);
                  const newPara = insertPoint.insertParagraph('', Word.InsertLocation.before);
                  newPara.font.color = '#0000FF';
                  
                  await context.sync();
                  
                  insertedSuccessfully = true;
                  applied++;
                  console.log(`   ? Newline inserted successfully using insertBreak() (short context)`);
                }
              }
              
              if (!insertedSuccessfully) {
                skipped++;
                console.warn(`   ?? Could not find insertion point for newline`);
              }
            } catch (e) {
              skipped++;
              console.error(`   ?? Error inserting newline: ${e.message}`);
            }
            
            // Skip the rest of insertion processing for newlines
            continue;
          }
          
          console.log(`${prefix} INSERT at pos ${op.position}: "${op.text.substring(0, 40)}..."`);
          console.log(`           Position in original: ${op.position}`);
          
          // Note: Tracking is ON for insertions (they'll show as tracked)
          // We'll format them as BLUE so they're visually distinct
          
          // Find insertion point using context
          // After all deletions are done, the document is shorter, so we need to search carefully
          let insertedSuccessfully = false;
          
          if (op.context.beforeContext.length >= 10) {
            const searchText = op.context.beforeContext.substring(Math.max(0, op.context.beforeContext.length - 50));
            
            console.log(`   ? Searching for position after: "${searchText.substring(Math.max(0, searchText.length - 30))}..."`);
            
            // CRITICAL FIX: Search in document body, not selection
            const body = context.document.body;
            const searchResults = body.search(searchText, {
              matchCase: false,
              matchWholeWord: false
            });
            searchResults.load('items');
            await context.sync();
            
            if (searchResults.items && searchResults.items.length > 0) {
              const beforeRange = searchResults.items[0];
              
              // Insert after the before context
              const insertedRange = beforeRange.insertText(op.text, Word.InsertLocation.after);
              insertedRange.font.color = '#0000FF'; // Blue for insertions
              
              await context.sync();
              
              insertedSuccessfully = true;
              applied++;
              console.log(`   ? Inserted successfully (formatted blue)`);
            } else {
              console.log(`   ?? Before-context not found, trying after-context...`);
            }
          }
          
          if (!insertedSuccessfully && op.context.afterContext.length >= 10) {
            const searchText = op.context.afterContext.substring(0, Math.min(50, op.context.afterContext.length));
            
            console.log(`   ? Searching for position before: "${searchText.substring(0, 30)}..."`);
            
            // CRITICAL FIX: Search in document body, not selection
            const body = context.document.body;
            const searchResults = body.search(searchText, {
              matchCase: false,
              matchWholeWord: false
            });
            searchResults.load('items');
            await context.sync();
            
            if (searchResults.items && searchResults.items.length > 0) {
              const afterRange = searchResults.items[0];
              
              // Insert before the after context
              const insertedRange = afterRange.insertText(op.text, Word.InsertLocation.before);
              insertedRange.font.color = '#0000FF'; // Blue for insertions
              
              await context.sync();
              
              insertedSuccessfully = true;
              applied++;
              console.log(`   ? Inserted successfully (formatted blue)`);
            } else {
              console.log(`   ?? After-context not found`);
            }
          }
          
          // LAST RESORT: Try shorter context (5-10 chars) which is more stable after deletions
          if (!insertedSuccessfully && op.context.beforeContext.length >= 5) {
            // Try progressively shorter contexts to find a unique match
            const contextSizes = [10, 8, 5];
            
            for (const size of contextSizes) {
              if (insertedSuccessfully) break;
              
              if (op.context.beforeContext.length >= size) {
                const shortContext = op.context.beforeContext.substring(Math.max(0, op.context.beforeContext.length - size));
                console.log(`   ? Last resort: searching for short context (${size} chars): "${shortContext}"`);
            
            const body = context.document.body;
            const searchResults = body.search(shortContext, {
              matchCase: false,
              matchWholeWord: false
            });
            searchResults.load('items');
            await context.sync();
            
            if (searchResults.items && searchResults.items.length > 0) {
                  // If multiple matches, try to verify we have the right one using after context
                  let selectedRange = searchResults.items[searchResults.items.length - 1]; // Default: use last match
                  
                  if (searchResults.items.length > 1 && op.context.afterContext.length >= 3) {
                    // Try to find the match that has the correct after context
                    const shortAfterCtx = op.context.afterContext.substring(0, Math.min(10, op.context.afterContext.length));
                    
                    for (let i = searchResults.items.length - 1; i >= 0; i--) {
                      try {
                        const match = searchResults.items[i];
                        const afterMatch = match.getRange(Word.RangeLocation.after);
                        afterMatch.load('text');
                        await context.sync();
                        
                        if (afterMatch.text && afterMatch.text.toLowerCase().startsWith(shortAfterCtx.toLowerCase())) {
                          selectedRange = match;
                          break;
                        }
                      } catch (e) {
                        // If we can't check this match, continue
                        continue;
                      }
                    }
                  }
              
              // Insert and format as blue
                  const insertedRange = selectedRange.insertText(op.text, Word.InsertLocation.after);
              insertedRange.font.color = '#0000FF'; // Blue for insertions
              
              await context.sync();
              
              insertedSuccessfully = true;
              applied++;
                  console.log(`   ? Inserted successfully (using short context, ${size} chars)`);
                  break;
                }
              }
            }
          }
          
          // FINAL RESORT: Try searching for key patterns that might still exist
          if (!insertedSuccessfully) {
            console.log(`   ? Final resort: searching for key patterns`);
            
            // Try progressively shorter context sizes
            const contextSizes = [8, 5, 3];
            let found = false;
            
            for (const size of contextSizes) {
              if (found || insertedSuccessfully) break;
              
              if (op.context.beforeContext.length >= size) {
                const contextPart = op.context.beforeContext.substring(Math.max(0, op.context.beforeContext.length - size));
                const body = context.document.body;
                
                try {
                  const contextSearch = body.search(contextPart, {
                    matchCase: false,
                    matchWholeWord: false
                  });
                  contextSearch.load('items');
                  await context.sync();
                  
                  if (contextSearch.items && contextSearch.items.length > 0) {
                    const beforeRange = contextSearch.items[contextSearch.items.length - 1];
                    
                    // Insert and format as blue
                    const insertedRange = beforeRange.insertText(op.text, Word.InsertLocation.after);
                    insertedRange.font.color = '#0000FF';
                    
                    await context.sync();
                    
                    insertedSuccessfully = true;
                    applied++;
                    found = true;
                    console.log(`   ? Inserted successfully (using final resort, context size: ${size})`);
                    break;
                  }
                } catch (contextError) {
                  console.warn(`   ?? Context search (size ${size}) failed: ${getErrorMessage(contextError)}`);
                }
              }
            }
            
            // If still not inserted, try searching for common punctuation or key words
            if (!insertedSuccessfully) {
              // Look for common patterns that might still exist even after deletions
              // Order matters: more specific patterns first
              const keyPatterns = ['configured to:', 'configured to', 'to: ', 'to:', 'processor is', ':'];
              
              for (const pattern of keyPatterns) {
                if (insertedSuccessfully) break;
                
                try {
                  const body = context.document.body;
                  const patternSearch = body.search(pattern, {
                    matchCase: false,
                    matchWholeWord: false
                  });
                  patternSearch.load('items');
                  await context.sync();
                  
                  if (patternSearch.items && patternSearch.items.length > 0) {
                    // Try to narrow down which match to use by checking if it's near the expected position
                    // Since we process in reverse order, we want the match that's closest to where
                    // the insertion should be based on the original position
                    
                    let selectedMatch = patternSearch.items[patternSearch.items.length - 1];
                    
                    // If multiple matches, try to use context to find the right one
                    if (patternSearch.items.length > 1 && op.context.beforeContext.length >= 3) {
                      // Try to find a match that has the before context nearby
                      const shortBeforeCtx = op.context.beforeContext.substring(Math.max(0, op.context.beforeContext.length - 10));
                      
                      for (let i = patternSearch.items.length - 1; i >= 0; i--) {
                        try {
                          const match = patternSearch.items[i];
                          // Get text before this match to see if it matches our context
                          const beforeMatch = match.getRange(Word.RangeLocation.before);
                          beforeMatch.load('text');
                          await context.sync();
                          
                          if (beforeMatch.text && beforeMatch.text.toLowerCase().includes(shortBeforeCtx.toLowerCase())) {
                            selectedMatch = match;
                            break;
                          }
                        } catch (e) {
                          // If we can't check this match, continue to next
                          continue;
                        }
                      }
                    }
                    
                    // Insert after the pattern
                    const insertedRange = selectedMatch.insertText(op.text, Word.InsertLocation.after);
                    insertedRange.font.color = '#0000FF';
                    
                    await context.sync();
                    
                    insertedSuccessfully = true;
                    applied++;
                    console.log(`   ? Inserted successfully (using key pattern: "${pattern}")`);
                    break;
                  }
                } catch (patternError) {
                  console.log(`   ?? Pattern search ("${pattern}") failed: ${getErrorMessage(patternError)}`);
                }
              }
            }
          }
          
          if (!insertedSuccessfully) {
            skipped++;
            console.warn(`   ?? Could not find insertion point`);
          }
        }
        
        if (onProgress && (applied + skipped) % 5 === 0) {
          onProgress(applied + skipped, totalOps, `Progress: ${applied} applied, ${skipped} skipped`);
        }
        
        // Adaptive delay based on text length
        const delay = op.text && op.text.length <= 3 ? 100 : 60;
        await new Promise(resolve => setTimeout(resolve, delay));
        
      } catch (e) {
        skipped++;
        console.error(`? [${applied + skipped}/${totalOps}] Error: ${e.message}`);
      }
      
      console.log('');
    }
    
    // Tracking stayed ON throughout - no need to turn it back on
    console.log('? Track changes mode: ON (stayed on throughout)');
    
    console.log('========================================');
    console.log(`? COMPLETED: ${applied}/${totalOps} operations applied`);
    if (skipped > 0) {
      console.log(`?? SKIPPED: ${skipped}/${totalOps} operations (see details above)`);
      console.log(`   Skipped operations could not be positioned uniquely`);
      console.log(`   Review tracked changes in Word and manually edit if needed`);
    } else {
      console.log(`?? ALL operations applied successfully!`);
    }
    console.log('========================================');
    
      if (onProgress) {
      onProgress(totalOps, totalOps, `Complete: ${applied} applied, ${skipped} skipped`);
    }
    });
  } catch (error) {
    console.error('Error in replaceSelectionWithNativeTrackedRevisions:', error);
    console.error('Error details:', {
      message: error.message,
      name: error.name,
      stack: error.stack,
      debugInfo: error.debugInfo
    });
    throw error;
  }
}

/**
 * Get selected text from Word document
 */
export async function getSelectedText() {
  try {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load('text');
      await context.sync();
      return selection.text;
    });
  } catch (error) {
    console.error('Error getting selected text:', error);
    throw error;
  }
}

/**
 * Insert text at cursor position
 */
export async function insertAtCursor(text) {
  try {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(text, Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (error) {
    console.error('Error inserting text:', error);
    throw error;
  }
}

/**
 * Replace selection with HTML
 */
export async function replaceSelectionWithHtml(html) {
  try {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertHtml(html, Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (error) {
    console.error('Error replacing selection with HTML:', error);
    throw error;
  }
}

/**
 * Replace selection as tracked change
 */
export async function replaceSelectionAsTracked(newText) {
  try {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(newText, Word.InsertLocation.replace);
      await context.sync();
    });
  } catch (error) {
    console.error('Error replacing selection as tracked:', error);
    throw error;
  }
}

/**
 * Get current Track Changes state
 */
export async function getTrackChangesState() {
  try {
    return await Word.run(async (context) => {
      const doc = context.document;
      doc.load('changeTrackingMode');
      await context.sync();
      
      const mode = doc.changeTrackingMode;
      const isTracking = mode === Word.ChangeTrackingMode.trackAll || mode === 'TrackAll';
      return isTracking;
    });
  } catch (error) {
    console.error('Error getting track changes state:', error);
    return null;
  }
}

/**
 * Set Track Changes state
 */
export async function setTrackChangesState(on) {
  try {
    return await Word.run(async (context) => {
      const doc = context.document;
      doc.changeTrackingMode = on ? Word.ChangeTrackingMode.trackAll : Word.ChangeTrackingMode.off;
      await context.sync();
      return on;
    });
  } catch (error) {
    console.error('Error setting track changes state:', error);
    return null;
  }
}

/**
 * Accept all tracked changes in the document
 */
export async function acceptAllTrackedChanges() {
  return Word.run(async (context) => {
    const doc = context.document;
    const body = doc.body;
    const result = { success: false, accepted: 0, insertionsFinalized: 0 };
    
    try {
      console.log('?? Accepting all tracked changes...');
      
      // Step 1: Accept all tracked changes (deletions with red strikethrough)
      const trackedChanges = body.getTrackedChanges();
      trackedChanges.load('items');
      await context.sync();
      
      if (trackedChanges.items && trackedChanges.items.length > 0) {
        console.log(`   Found ${trackedChanges.items.length} tracked changes`);
        for (const change of trackedChanges.items) {
          try {
            change.accept();
            result.accepted++;
          } catch (e) {
            console.warn(`?? Could not accept change: ${e.message}`);
          }
        }
        await context.sync();
        console.log(`? Accepted ${result.accepted} tracked changes`);
      } else {
        console.log('?? No tracked changes found');
      }
      
      // Step 2: Change any blue-colored text to black (inserted text cleanup)
      console.log('?? Cleaning up: Changing blue text to black...');
      try {
        // Find all paragraphs
        const paragraphs = body.paragraphs;
        paragraphs.load('items');
        await context.sync();
        
        for (const para of paragraphs.items) {
          para.font.load('color');
        }
        await context.sync();
        
        // Change blue paragraphs/ranges to black
        for (const para of paragraphs.items) {
          const color = para.font.color;
          if (color && (color.toUpperCase() === '#0000FF' || color.toLowerCase() === 'blue')) {
            para.font.color = '#000000';
            result.insertionsFinalized++;
          }
        }
        await context.sync();
        
        if (result.insertionsFinalized > 0) {
          console.log(`? Changed ${result.insertionsFinalized} blue text ranges to black`);
        }
      } catch (e) {
        console.warn(`?? Could not clean up blue text: ${e.message}`);
      }
      
      // Step 3: Ensure tracking stays ON
      doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      await context.sync();
      console.log('? Track changes mode: ON');
      
      result.success = true;
    } catch (error) {
      console.error('? Error accepting tracked changes:', error);
      result.error = error.message;
    }
    
    return result;
  });
}

/**
 * Accept tracked changes in a specific range or selection
 */
export async function acceptTrackedChangesInRange(rangeOrSelection = 'selection') {
  return Word.run(async (context) => {
    const doc = context.document;
    const range = rangeOrSelection === 'selection' 
      ? context.document.getSelection()
      : rangeOrSelection;
    
    const result = { success: false, accepted: 0, insertionsFinalized: 0 };
    
    try {
      console.log('?? Accepting tracked changes in range...');
      
      // Step 1: Accept all tracked changes in range
      const trackedChanges = range.getTrackedChanges();
      trackedChanges.load('items');
      await context.sync();
      
      if (trackedChanges.items && trackedChanges.items.length > 0) {
        console.log(`   Found ${trackedChanges.items.length} tracked changes in range`);
        for (const change of trackedChanges.items) {
          try {
            change.accept();
            result.accepted++;
          } catch (e) {
            console.warn(`?? Could not accept change: ${e.message}`);
          }
        }
        await context.sync();
        console.log(`? Accepted ${result.accepted} tracked changes`);
      } else {
        console.log('?? No tracked changes found in range');
      }
      
      // Step 2: Change any blue-colored text in range to black
      console.log('?? Cleaning up: Changing blue text to black in range...');
      try {
        const paragraphs = range.paragraphs;
        paragraphs.load('items');
        await context.sync();
        
        for (const para of paragraphs.items) {
          para.font.load('color');
        }
        await context.sync();
        
        for (const para of paragraphs.items) {
          const color = para.font.color;
          if (color && (color.toUpperCase() === '#0000FF' || color.toLowerCase() === 'blue')) {
            para.font.color = '#000000';
            result.insertionsFinalized++;
          }
        }
        await context.sync();
        
        if (result.insertionsFinalized > 0) {
          console.log(`? Changed ${result.insertionsFinalized} blue text ranges to black`);
        }
      } catch (e) {
        console.warn(`?? Could not clean up blue text: ${e.message}`);
      }
      
      // Step 3: Ensure tracking stays ON
      doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      await context.sync();
      console.log('? Track changes mode: ON');
      
      result.success = true;
    } catch (error) {
      console.error('? Error accepting tracked changes in range:', error);
      result.error = error.message;
    }
    
    return result;
  });
}

/**
 * Accept tracked changes based on filter criteria
 */
export async function acceptTrackedChangesByFilter(filter = {}) {
  return Word.run(async (context) => {
    const doc = context.document;
    const body = doc.body;
    const result = { success: false, accepted: 0, skipped: 0, insertionsFinalized: 0 };
    
    try {
      console.log('?? Accepting tracked changes by filter...');
      
      // Step 1: Accept filtered tracked changes
      const trackedChanges = body.getTrackedChanges();
      trackedChanges.load('items');
      await context.sync();
      
      if (trackedChanges.items && trackedChanges.items.length > 0) {
        for (const change of trackedChanges.items) {
          change.load('author,date');
        }
        await context.sync();
        
        for (const change of trackedChanges.items) {
          let shouldAccept = true;
          
          if (filter.author && change.author !== filter.author) {
            shouldAccept = false;
          }
          
          if (filter.after && change.date < filter.after) {
            shouldAccept = false;
          }
          
          if (filter.before && change.date > filter.before) {
            shouldAccept = false;
          }
          
          if (shouldAccept) {
            try {
              change.accept();
              result.accepted++;
            } catch (e) {
              console.warn(`Could not accept change: ${e.message}`);
            }
          } else {
            result.skipped++;
          }
        }
        
        await context.sync();
        console.log(`? Accepted ${result.accepted}, skipped ${result.skipped}`);
      } else {
        console.log('?? No tracked changes found');
      }
      
      // Step 2: Change any blue-colored text to black
      console.log('?? Cleaning up: Changing blue text to black...');
      try {
        const paragraphs = body.paragraphs;
        paragraphs.load('items');
        await context.sync();
        
        for (const para of paragraphs.items) {
          para.font.load('color');
        }
        await context.sync();
        
        for (const para of paragraphs.items) {
          const color = para.font.color;
          if (color && (color.toUpperCase() === '#0000FF' || color.toLowerCase() === 'blue')) {
            para.font.color = '#000000';
            result.insertionsFinalized++;
          }
        }
        await context.sync();
        
        if (result.insertionsFinalized > 0) {
          console.log(`? Changed ${result.insertionsFinalized} blue text ranges to black`);
        }
      } catch (e) {
        console.warn(`?? Could not clean up blue text: ${e.message}`);
      }
      
      // Step 3: Ensure tracking stays ON
      doc.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
      await context.sync();
      console.log('? Track changes mode: ON');
      
      result.success = true;
    } catch (error) {
      console.error('? Error accepting filtered tracked changes:', error);
      result.error = error.message;
    }
    
    return result;
  });
}
