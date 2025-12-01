# Test Examples for Word Track Changes

This file contains test cases to verify the v2.1 insertion algorithm improvements.

## Test Case 1: End-of-Sentence Insertion

**Purpose**: Verify insertions at the end of sentences work correctly (period after newline deletion)

**Old Text**:
```
The quick brown fox jumps over the lazy dog
```

**New Text**:
```
The quick brown fox gracefully leaps over the lazy, sleeping dog.
```

**Expected Diffs**:
1. Delete: "jum"
2. Insert: "gracefully lea"
3. Delete: newline
4. Insert: "."
5. Insert: ", sleeping"

**Expected Result**:
```
The quick brown fox gracefully leaps over the lazy, sleeping dog.
```

**Key Fix**: Period insertion at end after newline deletion (Strategy 3 or 4)

---

## Test Case 2: Complex Sentence Restructuring

**Purpose**: Verify "fall" inserts in correct position between "of the" and "semester"

**Old Text**:
```
Students must complete all assignments by the end of the semester.
```

**New Text**:
```
All enrolled students are required to complete and submit all graded assignments, including the final research project and comprehensive examinations, no later than December 15th, which marks the official conclusion of the fall semester.
```

**Key Operations**:
- Delete: "S" → Insert: "All enrolled s"
- Delete: "must complete all" → Insert: "are required to complete and submit all graded"
- Delete: " by the end" → Insert: ", including the final research project..."
- Insert: "fall " before "semester"

**Expected Result**:
```
All enrolled students are required to complete and submit all graded assignments, including the final research project and comprehensive examinations, no later than December 15th, which marks the official conclusion of the fall semester.
```

**Key Fix**: "fall" insertion uses combined context to find unique position before "semester"

---

## Test Case 3: Partial Word Assembly

**Purpose**: Verify word fragments join correctly without spaces

**Old Text**:
```
The new policy aims to improve workplace productivity and employee satisfaction.
```

**New Text**:
```
The recently implemented flexible work policy, developed through extensive employee surveys and stakeholder consultations, aims to enhance workplace productivity, improve employee satisfaction and retention rates, and promote a healthier work-life balance for all staff members.
```

**Critical Operation**:
- Delete: "improv"
- Insert: "enhanc"
- Equal: "e workplace"

**Expected Result**:
```
...aims to enhance workplace productivity...
```

**NOT**:
```
...aims to enhanc e workplace...  ❌
```

**Key Fix**: Insert BEFORE after-pattern ("e") not AFTER before-pattern (" aims to ")

---

## Test Case 4: Multiple Insertions in Sequence

**Purpose**: Verify left-to-right insertion order maintains accuracy

**Old Text**:
```
The meeting will be held next Tuesday.
```

**New Text**:
```
The quarterly board meeting will be held next Wednesday at 2:30 PM in the executive conference room.
```

**Expected Processing Order**:
1. Insert: "quarterly board " after "The "
2. Insert: "Wednesday at 2:30 PM in the executive conference room" after "next "
3. Delete: "Tuesday"

**Key Fix**: Left-to-right insertion order prevents context invalidation

---

## Test Case 5: Small Context with Duplicates

**Purpose**: Verify handling when context is small but needs to be unique

**Old Text**:
```
The cat sat on the mat. The cat was happy.
```

**New Text**:
```
The big cat sat on the soft mat. The cat was very happy.
```

**Key Operations**:
- Insert: "big " after "The " (first instance)
- Insert: "soft " after "the " (second instance)  
- Insert: "very " after "was " (unique context)

**Expected Result**:
```
The big cat sat on the soft mat. The cat was very happy.
```

**Key Fix**: Combined context or unique match validation prevents wrong insertions

---

## Test Case 6: Long Insertion Followed by Small Insertion

**Purpose**: Verify small insertions don't match within previously inserted long text

**Old Text**:
```
The system works correctly.
```

**New Text**:
```
The advanced machine learning system, developed through years of research and testing, works correctly and efficiently.
```

**Key Operations**:
1. Insert: "advanced machine learning " (large)
2. Insert: ", developed through years of research and testing," (very large)
3. Insert: " and efficiently" at end

**Expected Result**:
```
The advanced machine learning system, developed through years of research and testing, works correctly and efficiently.
```

**Key Fix**: Larger context windows and combined context prevent false matches in newly inserted text

---

## Test Case 7: Edge Case - Single Character Context

**Purpose**: Verify handling when before or after context is just 1-2 characters

**Old Text**:
```
A dog ran.
```

**New Text**:
```
A big dog ran fast.
```

**Key Operations**:
- Insert: "big " after "A " (before context: "A ", 2 chars)
- Insert: " fast" after "ran" (after context: ".", 1 char)

**Expected Result**:
```
A big dog ran fast.
```

**Key Fix**: Strategy 3 handles small contexts with minimum 10-char search patterns

---

## Test Case 8: Newline Insertion and Deletion

**Purpose**: Verify paragraph operations work with tracked changes

**Old Text**:
```
First paragraph. Second paragraph.
```

**New Text**:
```
First paragraph.

Second paragraph with more text.
```

**Key Operations**:
- Delete: " " (space)
- Insert: newline(s)
- Insert: " with more text"

**Expected Result**:
```
First paragraph.

Second paragraph with more text.
```

**Key Fix**: Newline operations use paragraph-based strategy

---

## Test Case 9: Punctuation Changes

**Purpose**: Verify punctuation insertions and deletions

**Old Text**:
```
Hello world
```

**New Text**:
```
Hello, world!
```

**Key Operations**:
- Insert: "," after "Hello"
- Insert: "!" after "world"

**Expected Result**:
```
Hello, world!
```

**Key Fix**: Small insertions use combined context for accuracy

---

## Test Case 10: Complex Multi-Edit Paragraph

**Purpose**: Stress test with many operations in one paragraph

**Old Text**:
```
The company announced record profits this quarter.
```

**New Text**:
```
Despite facing supply chain disruptions and inflationary pressures, the multinational corporation announced unprecedented record-breaking profits this fiscal quarter, exceeding analyst expectations by 23%.
```

**Key Operations**:
- Insert: "Despite facing supply chain disruptions and inflationary pressures, "
- Delete: "company" → Insert: "multinational corporation"
- Insert: "unprecedented record-breaking " before "profits"
- Insert: " fiscal" before "quarter"
- Insert: ", exceeding analyst expectations by 23%" before "."

**Expected Result**:
```
Despite facing supply chain disruptions and inflationary pressures, the multinational corporation announced unprecedented record-breaking profits this fiscal quarter, exceeding analyst expectations by 23%.
```

**Key Fix**: Combination of all strategies - large context windows, combined context, left-to-right order

---

## How to Test

1. Open Word (web or desktop)
2. Load the add-in
3. Insert the "Old Text" into Word
4. Select all the old text
5. Paste the "New Text" into the add-in
6. Click "Preview Changes"
7. Verify the result matches "Expected Result" exactly
8. Check console logs to see which strategies were used
9. Accept tracked changes to finalize

## Success Criteria

✅ All insertions appear in correct positions
✅ No broken words or unwanted spaces
✅ Deletions show as strikethrough in red
✅ Insertions show as underlined in blue
✅ Console shows "Successful: N/N, Failed: 0/N"
✅ Combined context strategy used for most insertions
✅ All operations complete without errors

## Debugging

If a test fails:
1. Check console logs for the failing operation
2. Look for "⚠️ All insertion strategies failed" messages
3. Check the before/after context shown in logs
4. Verify the context should be unique in the document
5. Consider if the context was invalidated by a previous insertion
6. Check if the combined context total size is too small (should be ≥5 chars)

## Performance Benchmarks

Expected performance (approximate):
- Simple edits (3-5 operations): < 1 second
- Medium complexity (10-15 operations): 1-2 seconds
- High complexity (20+ operations): 2-4 seconds

Each insertion syncs immediately for accuracy, so performance scales linearly with operation count.

