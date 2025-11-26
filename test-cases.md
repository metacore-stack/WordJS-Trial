# Test Cases for Track Changes

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

## Test Case 2: Multiple Word Changes
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

## Test Case 4: Complex Patent-Style Text
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

## Test Case 5: Short Deletion (Colon + Newline)
**Old Text:**
```
The method comprises:
receiving input data.
```

**New Text:**
```
The method comprises receiving input data.
```

**Expected:** ": " deleted (colon and newline)

---

## Test Case 6: Repeated Words
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

## Test Case 7: Multiple Changes in Sentence
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

## Test Case 8: Insertion Only
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

## Test Case 9: Deletion Only
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

## Test Case 10: Complex Multiple Edits
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

## Test Case 11: Single Character Change
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

## Test Case 12: Punctuation Changes
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

## Test Case 13: Long Text with Multiple Edits
**Old Text:**
```
The system of claim 1, wherein the processor is further configured to: generate a first visual code for presentation on the signer electronic device, wherein the signer identification confirmation is based on the witness electronic device receiving the first visual code through an interaction between the signer electronic device and the witness electronic device, and wherein the system further comprises a database for storing transaction records.
```

**New Text:**
```
The system of claim 1, wherein the processor is configured to generate a visual code on the signer electronic device, and the signer identification confirmation is based on the witness electronic device receiving the visual code from the signer electronic device, and wherein the system comprises a database for storing transaction records.
```

**Expected:** Multiple deletions and insertions throughout

---

## Test Case 14: Edge Case - Very Short Text
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

## Test Case 15: Whitespace Changes
**Old Text:**
```
The system  processes   data.
```

**New Text:**
```
The system processes data.
```

**Expected:** Extra spaces deleted

