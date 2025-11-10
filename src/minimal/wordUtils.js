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
 * Normalize text for comparison
 * CRITICAL: Remove \r (carriage return) because Word's search API can't match it
 */
function normalizeText(text) {
  if (!text) return '';
  // Remove carriage returns (\r) then normalize Unicode
  return text.replace(/\r/g, '').normalize('NFC');
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
    
    // CRITICAL FIX: Normalize the original text by removing carriage returns
    // Word's search API doesn't match \r characters reliably
    const normalized = normalizeText(originalText);
    
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
        } else {
          // Try to find nearby
          const searchIdx = normalized.indexOf(normText, Math.max(0, currentPos - 30));
          if (searchIdx >= 0 && searchIdx < currentPos + 150) {
            currentPos = searchIdx + normText.length;
          } else {
            console.warn(`?? [Diff ${i}] Could not align equal text (skipping)`);
          }
        }
      } else if (diff.op === 'delete') {
        // Skip deletions that normalize to empty string (e.g., pure \r deletions)
        if (normText.length === 0) {
          console.log(`? [Diff ${i}] Delete normalized to empty (e.g., \\r) - auto-skipping`);
          continue; // Nothing to delete
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
        // Find unique context for insertion point (use normalized text to avoid \r issues)
        const insertContext = findUniqueContext(normalized, currentPos, currentPos, '');
        
        positionMap.push({
          index: i,
          op: 'insert',
          position: currentPos,
          text: diff.text,
          context: insertContext,
          length: normalizeText(diff.text).length
        });
        
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
            console.log(`${prefix} DELETE at pos ${op.start}: "${op.text.substring(0, 40)}..."`);
            console.log(`           Position in original: ${op.start}-${op.end}`);
            
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
                      
                      // Since we process in REVERSE order (end to start), we want matches
                      // that appear LATER in the document for EARLIER deletions in our loop
                      // Use the LAST match as it's most likely to be the one we haven't deleted yet
                      const selectedMatch = directSearch.items[directSearch.items.length - 1];
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
                skipped++;
                console.warn(`   ?? Pattern not found in document`);
              }
            } else {
              skipped++;
              console.warn(`   ?? No unique pattern found`);
            }
          
        } else if (op.op === 'insert') {
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
              console.warn(`   ?? Before-context not found, trying after-context...`);
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
              console.warn(`   ?? After-context not found`);
            }
          }
          
          // LAST RESORT: Try shorter context (5-10 chars) which is more stable after deletions
          if (!insertedSuccessfully && op.context.beforeContext.length >= 5) {
            const shortContext = op.context.beforeContext.substring(Math.max(0, op.context.beforeContext.length - 10));
            console.log(`   ? Last resort: searching for short context: "${shortContext}"`);
            
            const body = context.document.body;
            const searchResults = body.search(shortContext, {
              matchCase: false,
              matchWholeWord: false
            });
            searchResults.load('items');
            await context.sync();
            
            if (searchResults.items && searchResults.items.length > 0) {
              const beforeRange = searchResults.items[searchResults.items.length - 1]; // Use last match (closer to end)
              
              // Insert and format as blue
              const insertedRange = beforeRange.insertText(op.text, Word.InsertLocation.after);
              insertedRange.font.color = '#0000FF'; // Blue for insertions
              
              await context.sync();
              
              insertedSuccessfully = true;
              applied++;
              console.log(`   ? Inserted successfully (using short context)`);
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
