# 30 Comprehensive Test Cases for Word.js Track Changes

## Test Case 1: Simple Word Replacement
**Old Text:**
```
The system processes data efficiently.
```

**New Text:**
```
The system handles data efficiently.
```

**Expected:** "processes" deleted, "handles" inserted

---

## Test Case 2: Multiple Word Deletions
**Old Text:**
```
The processor is configured to generate a first visual code for presentation on the signer electronic device.
```

**New Text:**
```
The processor is configured to generate a visual code on the signer electronic device.
```

**Expected:** "first " and "for presentation " deleted

---

## Test Case 3: Phrase Replacement
**Old Text:**
```
The system communicates with the server through a secure connection.
```

**New Text:**
```
The system communicates with the server via a secure connection.
```

**Expected:** "through" deleted, "via" inserted

---

## Test Case 4: Newline Deletion (Colon + Newline)
**Old Text:**
```
The method comprises:
receiving input data.
```

**New Text:**
```
The method comprises receiving input data.
```

**Expected:** ": " deleted, newline deleted, " " inserted

---

## Test Case 5: Newline Insertion
**Old Text:**
```
The method comprises receiving input data.
```

**New Text:**
```
The method comprises:
receiving input data.
```

**Expected:** " " deleted, newline inserted, ": " inserted

---

## Test Case 6: Trailing Newline Deletion
**Old Text:**
```
The system processes data efficiently.

```

**New Text:**
```
The system processes data efficiently.
```

**Expected:** Trailing newline deleted

---

## Test Case 7: Multiple Newlines
**Old Text:**
```
First paragraph.

Second paragraph.

Third paragraph.
```

**New Text:**
```
First paragraph.
Second paragraph.
Third paragraph.
```

**Expected:** Two newlines deleted (one between each paragraph)

---

## Test Case 8: Repeated Words - Delete Middle
**Old Text:**
```
The system processes data and processes information.
```

**New Text:**
```
The system processes data and handles information.
```

**Expected:** Second "processes" deleted, "handles" inserted

---

## Test Case 9: Repeated Words - Delete First
**Old Text:**
```
the the the document contains important information.
```

**New Text:**
```
the the document contains important information.
```

**Expected:** First "the " deleted

---

## Test Case 10: Insertion Only
**Old Text:**
```
The system processes data.
```

**New Text:**
```
The system processes data efficiently.
```

**Expected:** " efficiently" inserted

---

## Test Case 11: Deletion Only
**Old Text:**
```
The system processes data efficiently.
```

**New Text:**
```
The system processes data.
```

**Expected:** " efficiently" deleted

---

## Test Case 12: Single Character Change
**Old Text:**
```
The system processes data.
```

**New Text:**
```
The system process data.
```

**Expected:** "es" deleted (from "processes" to "process")

---

## Test Case 13: Punctuation Changes
**Old Text:**
```
The system, which processes data, is efficient.
```

**New Text:**
```
The system that processes data is efficient.
```

**Expected:** ", which" deleted, "that" inserted, ", " deleted

---

## Test Case 14: Complex Patent-Style Text
**Old Text:**
```
2. The system of claim 1, wherein the processor is further configured to: generate a first visual code for presentation on the signer electronic device, wherein the signer identification confirmation is based on the witness electronic device receiving the first visual code through an interaction between the signer electronic device and the witness electronic device.
```

**New Text:**
```
2. The system of claim 1, wherein the processor is configured to generate a visual code on the signer electronic device, and the signer identification confirmation is based on the witness electronic device receiving the visual code from the signer electronic device.
```

**Expected:** Multiple deletions and insertions

---

## Test Case 15: Whitespace Normalization
**Old Text:**
```
The system  processes   data    efficiently.
```

**New Text:**
```
The system processes data efficiently.
```

**Expected:** Extra spaces deleted

---

## Test Case 16: Leading and Trailing Spaces
**Old Text:**
```
  The system processes data efficiently.  
```

**New Text:**
```
The system processes data efficiently.
```

**Expected:** Leading and trailing spaces deleted

---

## Test Case 17: Tab to Space Conversion
**Old Text:**
```
The system	processes	data.
```

**New Text:**
```
The system processes data.
```

**Expected:** Tabs deleted, spaces inserted

---

## Test Case 18: Multiple Changes in One Sentence
**Old Text:**
```
The processor is further configured to generate a first visual code for presentation on the signer electronic device.
```

**New Text:**
```
The processor is configured to generate a visual code on the signer electronic device.
```

**Expected:** "further ", "first ", "for presentation " deleted

---

## Test Case 19: Complex Multiple Edits
**Old Text:**
```
The method includes the steps of: receiving user input, processing the input data, and generating an output result based on the processed input data.
```

**New Text:**
```
The method includes receiving user input, processing the input, and generating output based on the processed input.
```

**Expected:** Multiple deletions and insertions

---

## Test Case 20: Very Short Text
**Old Text:**
```
A method.
```

**New Text:**
```
A process.
```

**Expected:** "method" deleted, "process" inserted

---

## Test Case 21: Insert at Beginning
**Old Text:**
```
world
```

**New Text:**
```
Hello world
```

**Expected:** "Hello " inserted

---

## Test Case 22: Insert at End
**Old Text:**
```
Hello
```

**New Text:**
```
Hello world
```

**Expected:** " world" inserted

---

## Test Case 23: Delete Entire Text
**Old Text:**
```
Delete me completely.
```

**New Text:**
```

```

**Expected:** All text deleted

---

## Test Case 24: Insert into Empty
**Old Text:**
```

```

**New Text:**
```
New content here.
```

**Expected:** "New content here." inserted

---

## Test Case 25: Newline with Text Changes
**Old Text:**
```
The method comprises:
receiving input data.
```

**New Text:**
```
The method includes:
processing input data.
```

**Expected:** "comprises" deleted, "includes" inserted, "receiving" deleted, "processing" inserted

---

## Test Case 26: Multiple Paragraphs with Newlines
**Old Text:**
```
First section.

Second section.

Third section.
```

**New Text:**
```
First section.
Second section.
Third section.
```

**Expected:** Newlines between sections deleted

---

## Test Case 27: Newline Before Text
**Old Text:**
```
First line.
Second line.
```

**New Text:**
```

First line.
Second line.
```

**Expected:** Newline inserted at beginning

---

## Test Case 28: Colon and Newline Together
**Old Text:**
```
The steps are:
one, two, three.
```

**New Text:**
```
The steps are: one, two, three.
```

**Expected:** Newline deleted, space inserted

---

## Test Case 29: Multiple Consecutive Newlines
**Old Text:**
```
Paragraph one.


Paragraph two.
```

**New Text:**
```
Paragraph one.
Paragraph two.
```

**Expected:** Two consecutive newlines deleted

---

## Test Case 30: Complex Real-World Example
**Old Text:**
```
1. A method for processing data, the method comprising:
   receiving input data from a user interface;
   processing the input data using a first algorithm;
   generating output data based on the processed input data; and
   displaying the output data on a display device.

2. The method of claim 1, wherein the processing step further comprises:
   validating the input data; and
   transforming the validated input data.
```

**New Text:**
```
1. A method for processing data, the method comprising: receiving input data from a user interface; processing the input data using an algorithm; generating output based on the processed input; and displaying the output on a display device.

2. The method of claim 1, wherein the processing step comprises: validating the input data; and transforming the validated input data.
```

**Expected:** Multiple newlines deleted, "first " deleted, "data" changed to "output" in multiple places, "further " deleted, formatting changes

---

## Test Categories Summary

1. **Basic Operations (1-3)**: Simple word/phrase replacements
2. **Newline Handling (4-7)**: Various newline deletion and insertion scenarios
3. **Repeated Text (8-9)**: Handling duplicate words
4. **Single Operations (10-11)**: Insert-only and delete-only
5. **Character-Level (12)**: Single character changes
6. **Punctuation (13)**: Punctuation modifications
7. **Complex Text (14, 19, 30)**: Long, complex patent-style text
8. **Whitespace (15-17)**: Space and tab handling
9. **Multiple Edits (18)**: Multiple changes in one sentence
10. **Edge Cases (20-24)**: Very short text, boundaries, empty text
11. **Newline Combinations (25-29)**: Newlines with other changes
12. **Real-World (30)**: Complex realistic scenario

