# ThreadSafe Test Workbook Specification

This document describes the structure for an Excel workbook to comprehensively test the thread-safe functions across the three add-ins: ThreadSafeC (C functions), ThreadSafe2 (C# functions), and ThreadSafe1 (test wrappers).

## Workbook Structure

**File Name:** `ThreadSafeTest.xlsx`
**Location:** `C:\Work\ExcelDna\TechnicalSupport\WellsFargo\Test\`

### Sheet 1: "C Functions Direct"
Tests the C functions directly from the ThreadSafeC add-in.

#### Layout:
```
Row 1: Headers
A1: Test Description
B1: Function
C1: Input Value
D1: Result
E1: Thread Info
F1: Performance (ms)

Row 2-20: Various test scenarios
```

#### Test Scenarios:

**Row 2: Basic ThreadSafeCFunction Test**
- A2: `Basic ThreadSafeCFunction`
- B2: `ThreadSafeCFunction`  
- C2: `10`
- D2: `=ThreadSafeCFunction(C2)`
- E2: `=ThreadInfoFunction()`

**Row 3: Basic ThreadSafeCalc Test**
- A3: `Basic ThreadSafeCalc`
- B3: `ThreadSafeCalc`
- C3: `25`
- D3: `=ThreadSafeCalc(C3)`

**Row 4: ThreadSafeXLOPER Test**
- A4: `ThreadSafeXLOPER Test`
- B4: `ThreadSafeXLOPER`
- C4: `15`
- D4: `=ThreadSafeXLOPER(C4)`

**Row 5: AllocatedMemoryFunction Small Array**
- A5: `Small Array (3 elements)`
- B5: `AllocatedMemoryFunction`
- C5: `3`
- D5: `=AllocatedMemoryFunction(C5)`

**Row 6: AllocatedMemoryFunction Large Array**
- A6: `Large Array (20 elements)`
- B6: `AllocatedMemoryFunction`
- C6: `20`
- D6: `=AllocatedMemoryFunction(C6)`

**Row 7: Thread Info**
- A7: `Thread Information`
- B7: `ThreadInfoFunction`
- C7: `N/A`
- D7: `=ThreadInfoFunction()`

**Row 9-18: Stress Test Series (10 identical calls)**
- A9-A18: `Stress Test 1` through `Stress Test 10`
- B9-B18: `ThreadSafeCalc`
- C9-C18: `50, 51, 52, ... 59`
- D9-D18: `=ThreadSafeCalc(C9)` through `=ThreadSafeCalc(C18)`

**Row 20: Performance Test**
- A20: `Performance Comparison`
- B20: `Manual Timing`
- C20: `100`
- D20: `=NOW()` (before calculation)
- E20: `=ThreadSafeCalc(C20)`
- F20: `=NOW()` (after calculation)

---

### Sheet 2: "CS Functions Direct"
Tests the C# functions directly from the ThreadSafe2 add-in.

#### Layout:
```
Row 1: Headers
A1: Test Description
B1: Function
C1: Input Value
D1: Result
E1: Thread Info
F1: Notes

Row 2-20: Various test scenarios
```

#### Test Scenarios:

**Row 2: Basic csThreadSafeCFunction Test**
- A2: `Basic csThreadSafeCFunction`
- B2: `csThreadSafeCFunction`  
- C2: `10`
- D2: `=csThreadSafeCFunction(C2)`
- E2: `=csThreadInfoFunction()`
- F2: `C# version of ThreadSafeCFunction`

**Row 3: Basic csThreadSafeCalc Test**
- A3: `Basic csThreadSafeCalc`
- B3: `csThreadSafeCalc`
- C3: `25`
- D3: `=csThreadSafeCalc(C3)`
- F3: `C# version with managed thread ID`

**Row 4: csThreadSafeXLOPER Test**
- A4: `csThreadSafeXLOPER Test`
- B4: `csThreadSafeXLOPER`
- C4: `15`
- D4: `=csThreadSafeXLOPER(C4)`
- F4: `Returns double directly`

**Row 5: csAllocatedMemoryFunction Small Array**
- A5: `C# Small Array (3 elements)`
- B5: `csAllocatedMemoryFunction`
- C5: `3`
- D5: `=csAllocatedMemoryFunction(C5)`
- F5: `C# managed array`

**Row 6: csAllocatedMemoryFunction Large Array**
- A6: `C# Large Array (15 elements)`
- B6: `csAllocatedMemoryFunction`
- C6: `15`
- D6: `=csAllocatedMemoryFunction(C6)`

**Row 7: C# Thread Info**
- A7: `C# Thread Information`
- B7: `csThreadInfoFunction`
- C7: `N/A`
- D7: `=csThreadInfoFunction()`
- F7: `Managed thread info`

**Row 8: Advanced C# Test**
- A8: `Advanced C# Calculation`
- B8: `csAdvancedTest`
- C8: `10`
- D8: `=csAdvancedTest(C8, 5)`
- F8: `Advanced calculation with iterations`

**Row 10-19: C# Stress Test Series (10 identical calls)**
- A10-A19: `C# Stress Test 1` through `C# Stress Test 10`
- B10-B19: `csThreadSafeCalc`
- C10-C19: `60, 61, 62, ... 69`
- D10-D19: `=csThreadSafeCalc(C10)` through `=csThreadSafeCalc(C19)`
- F10-F19: `Managed thread performance`

---

### Sheet 3: "Test Wrappers"
Tests the wrapper functions from ThreadSafe1 that can call either C or C# versions.

#### Layout:
```
Row 1: Headers
A1: Test Description
B1: Function
C1: Input Value
D1: Use C# Flag
E1: Result
F1: Comparison
G1: Notes

Row 2-30: Comprehensive test scenarios
```

#### Test Scenarios:

**Section 1: Basic Function Tests (Rows 2-11)**

**Row 2: Test C Version of ThreadSafeCFunction**
- A2: `Test ThreadSafeCFunction (C)`
- B2: `TestThreadSafeCFunction`
- C2: `20`
- D2: `FALSE`
- E2: `=TestThreadSafeCFunction(C2, D2)`
- F2: `=ThreadSafeCFunction(C2)`
- G2: `Direct comparison with C function`

**Row 3: Test C# Version of ThreadSafeCFunction**
- A3: `Test ThreadSafeCFunction (C#)`
- B3: `TestThreadSafeCFunction`
- C3: `20`
- D3: `TRUE`
- E3: `=TestThreadSafeCFunction(C3, D3)`
- F3: `=csThreadSafeCFunction(C3)`
- G3: `Direct comparison with C# function`

**Row 4: Test C Version of ThreadSafeCalc**
- A4: `Test ThreadSafeCalc (C)`
- B4: `TestThreadSafeCalc`
- C4: `30`
- D4: `FALSE`
- E4: `=TestThreadSafeCalc(C4, D4)`
- F4: `=ThreadSafeCalc(C4)`
- G4: `C version performance`

**Row 5: Test C# Version of ThreadSafeCalc**
- A5: `Test ThreadSafeCalc (C#)`
- B5: `TestThreadSafeCalc`
- C5: `30`
- D5: `TRUE`
- E5: `=TestThreadSafeCalc(C5, D5)`
- F5: `=csThreadSafeCalc(C5)`
- G5: `C# version performance`

**Row 6: Test C Version of ThreadSafeXLOPER**
- A6: `Test ThreadSafeXLOPER (C)`
- B6: `TestThreadSafeXLOPER`
- C6: `40`
- D6: `FALSE`
- E6: `=TestThreadSafeXLOPER(C6, D6)`
- F6: `=ThreadSafeXLOPER(C6)`
- G6: `C XLOPER12 version`

**Row 7: Test C# Version of ThreadSafeXLOPER**
- A7: `Test ThreadSafeXLOPER (C#)`
- B7: `TestThreadSafeXLOPER`
- C7: `40`
- D7: `TRUE`
- E7: `=TestThreadSafeXLOPER(C7, D7)`
- F7: `=csThreadSafeXLOPER(C7)`
- G7: `C# double version`

**Row 8: Test C Memory Allocation**
- A8: `Test AllocatedMemory (C)`
- B8: `TestAllocatedMemoryFunction`
- C8: `5`
- D8: `FALSE`
- E8: `=TestAllocatedMemoryFunction(C8, D8)`
- F8: `=AllocatedMemoryFunction(C8)`
- G8: `C memory allocation with xlFree`

**Row 9: Test C# Memory Allocation**
- A9: `Test AllocatedMemory (C#)`
- B9: `TestAllocatedMemoryFunction`
- C9: `5`
- D9: `TRUE`
- E9: `=TestAllocatedMemoryFunction(C9, D9)`
- F9: `=csAllocatedMemoryFunction(C9)`
- G9: `C# managed array`

**Row 10: Test C Thread Info**
- A10: `Test ThreadInfo (C)`
- B10: `TestThreadInfoFunction`
- C10: `N/A`
- D10: `FALSE`
- E10: `=TestThreadInfoFunction(D10)`
- F10: `=ThreadInfoFunction()`
- G10: `C native thread info`

**Row 11: Test C# Thread Info**
- A11: `Test ThreadInfo (C#)`
- B11: `TestThreadInfoFunction`
- C11: `N/A`
- D11: `TRUE`
- E11: `=TestThreadInfoFunction(D11)`
- F11: `=csThreadInfoFunction()`
- G11: `C# managed thread info`

**Section 2: Multiple Call Tests (Rows 13-16)**

**Row 13: Multiple Calls (C)**
- A13: `Multiple Calls (C)`
- B13: `TestMultipleThreadSafeCalls`
- C13: `50`
- D13: `FALSE`
- E13: `=TestMultipleThreadSafeCalls(C13, D13)`
- G13: `Tests multiple C functions in sequence`

**Row 14: Multiple Calls (C#)**
- A14: `Multiple Calls (C#)`
- B14: `TestMultipleThreadSafeCalls`
- C14: `50`
- D14: `TRUE`
- E14: `=TestMultipleThreadSafeCalls(C14, D14)`
- G14: `Tests multiple C# functions in sequence`

**Section 3: Performance Tests (Rows 16-25)**

**Row 16: Performance Test (C, 10 iterations)**
- A16: `Performance (C, 10 iter)`
- B16: `TestPerformance`
- C16: `100`
- D16: `FALSE`
- E16: `=TestPerformance(C16, 10, D16)`
- G16: `C version timing`

**Row 17: Performance Test (C#, 10 iterations)**
- A17: `Performance (C#, 10 iter)`
- B17: `TestPerformance`
- C17: `100`
- D17: `TRUE`
- E17: `=TestPerformance(C17, 10, D17)`
- G17: `C# version timing`

**Row 18: Performance Test (C, 50 iterations)**
- A18: `Performance (C, 50 iter)`
- B18: `TestPerformance`
- C18: `100`
- D18: `FALSE`
- E18: `=TestPerformance(C18, 50, D18)`
- G18: `C version extended`

**Row 19: Performance Test (C#, 50 iterations)**
- A19: `Performance (C#, 50 iter)`
- B19: `TestPerformance`
- C19: `100`
- D19: `TRUE`
- E19: `=TestPerformance(C19, 50, D19)`
- G19: `C# version extended`

**Row 20: Direct Performance Comparison**
- A20: `Direct Comparison`
- B20: `ComparePerformance`
- C20: `100`
- D20: `N/A`
- E20: `=ComparePerformance(C20, 25)`
- G20: `Direct C vs C# comparison`

**Section 4: Stress Tests (Rows 22-30)**

**Row 22-30: Alternating C and C# Calls**
- A22: `Stress C 1`, A23: `Stress C# 1`, A24: `Stress C 2`, etc.
- B22-B30: `TestThreadSafeCalc`
- C22-C30: `200, 201, 202, ... 208`
- D22, D24, D26, D28, D30: `FALSE` (C version)
- D23, D25, D27, D29: `TRUE` (C# version)
- E22-E30: `=TestThreadSafeCalc(C22, D22)` through `=TestThreadSafeCalc(C30, D30)`
- G22-G30: `Alternating stress test`

---

## Additional Test Instructions

### Manual Tests to Perform:

1. **Recalculation Test:** Press F9 multiple times to force recalculation and observe thread ID changes
2. **Multi-threading Test:** Enter array formulas across multiple cells simultaneously
3. **Performance Monitoring:** Use Excel's calculation timing features to compare C vs C# performance
4. **Error Handling:** Test with invalid inputs (negative numbers, very large numbers)
5. **Memory Tests:** Monitor memory usage when calling AllocatedMemoryFunction multiple times

### Expected Behaviors:

- **Thread IDs:** Should vary between calls when Excel uses multiple threads
- **C Functions:** May show native thread IDs from Windows
- **C# Functions:** Will show managed thread IDs
- **Performance:** C functions should generally be faster due to native execution
- **Memory:** AllocatedMemoryFunction should properly allocate and free memory
- **Arrays:** Should display properly in Excel cells

### Notes:

- Ensure all three add-ins (ThreadSafeC.xll, ThreadSafe1, ThreadSafe2) are loaded before testing
- Functions marked as thread-safe should work correctly during multi-threaded calculation
- Compare results between direct function calls and wrapper function calls to verify consistency
- Monitor Excel's memory usage during array function tests