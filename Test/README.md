# ThreadSafe Test Suite

This directory contains a comprehensive test suite for testing thread safety across the three Excel add-ins:
- **ThreadSafeC** (C/C++ native XLL)
- **ThreadSafe2** (C# add-in with "cs" prefixed functions)
- **ThreadSafe1** (C# test wrappers that can call either C or C# versions)

## Files in this Directory

### Test Templates:
1. **`ThreadSafeTest_Workbook_Specification.md`** - Complete specification for the Excel workbook
2. **`Sheet1_C_Functions_Direct.csv`** - Template for testing C functions directly
3. **`Sheet2_CS_Functions_Direct.csv`** - Template for testing C# functions directly  
4. **`Sheet3_Test_Wrappers.csv`** - Template for testing wrapper functions with C/C# switching

### How to Create the Test Workbook:

1. **Create a new Excel workbook** named `ThreadSafeTest.xlsx`
2. **Create three worksheets:**
   - Sheet 1: "C Functions Direct"
   - Sheet 2: "CS Functions Direct"  
   - Sheet 3: "Test Wrappers"

3. **Import the CSV templates:**
   - For each sheet, import the corresponding CSV file
   - The formulas are provided as text - you'll need to enter them as actual Excel formulas
   - Replace semicolons (`;`) with commas (`,`) in formulas if using English Excel

## Functions Available for Testing

### C Functions (ThreadSafeC add-in):
- `ThreadSafeCFunction(input)` - Basic function (not thread-safe)
- `ThreadSafeCalc(number)` - Thread-safe calculation with $ flag
- `ThreadSafeXLOPER(input)` - Thread-safe XLOPER12 function
- `AllocatedMemoryFunction(size)` - Returns allocated memory requiring xlFree
- `ThreadInfoFunction()` - Returns native thread information

### C# Functions (ThreadSafe2 add-in):
- `csThreadSafeCFunction(input)` - C# equivalent of ThreadSafeCFunction
- `csThreadSafeCalc(number)` - C# thread-safe calculation
- `csThreadSafeXLOPER(input)` - C# version returning double
- `csAllocatedMemoryFunction(size)` - Returns C# managed array
- `csThreadInfoFunction()` - Returns C# thread information
- `csAdvancedTest(input, iterations)` - Advanced C# test function

### Test Wrapper Functions (ThreadSafe1 add-in):
- `TestThreadSafeCFunction(input, useCSharp)` - Tests either C or C# version
- `TestThreadSafeCalc(number, useCSharp)` - Tests either version
- `TestThreadSafeXLOPER(input, useCSharp)` - Tests either version
- `TestAllocatedMemoryFunction(size, useCSharp)` - Tests memory allocation
- `TestThreadInfoFunction(useCSharp)` - Tests thread info functions
- `TestMultipleThreadSafeCalls(input, useCSharp)` - Tests multiple functions
- `TestPerformance(input, iterations, useCSharp)` - Performance testing
- `ComparePerformance(input, iterations)` - Direct C vs C# comparison

## Testing Procedures

### Prerequisites:
1. Build all three add-in projects (ThreadSafeC, ThreadSafe1, ThreadSafe2)
2. Load all add-ins into Excel:
   - ThreadSafeC.xll (from bin/Debug/x64/ or bin/Release/x64/)
   - ThreadSafe1 (Excel-DNA add-in)
   - ThreadSafe2 (Excel-DNA add-in)

### Basic Tests:
1. **Function Verification:** Enter each formula and verify it returns expected results
2. **Thread Safety:** Press F9 multiple times and observe thread ID changes
3. **Performance:** Compare timing between C and C# versions
4. **Memory:** Test array functions and monitor Excel's memory usage

### Advanced Tests:
1. **Multi-threading:** Create large ranges of formulas and recalculate simultaneously
2. **Stress Testing:** Use the stress test rows to create concurrent calculations
3. **Error Handling:** Test with invalid inputs (negative numbers, zero, very large values)
4. **Comparison Testing:** Use the wrapper functions to directly compare C vs C# results

## Expected Results

### Thread IDs:
- Should vary when Excel uses multiple calculation threads
- C functions show native Windows thread IDs
- C# functions show managed thread IDs

### Performance:
- C functions generally faster due to native execution
- C# functions may show higher managed thread IDs
- Results should be mathematically equivalent between versions

### Memory:
- AllocatedMemoryFunction should properly allocate and free memory
- Arrays should display correctly in Excel
- No memory leaks should occur during repeated testing

## Troubleshooting

### Common Issues:
1. **#NAME? errors:** Verify all add-ins are loaded
2. **#VALUE! errors:** Check input parameters and types
3. **Performance issues:** Reduce iteration counts in performance tests
4. **Memory issues:** Monitor Excel memory usage during array tests

### Formula Notes:
- Use semicolons (`;`) as parameter separators in European Excel
- Use commas (`,`) as parameter separators in US Excel
- Boolean values: TRUE/FALSE (English) or WAHR/FALSCH (German)

## Test Scenarios Summary

- **Basic functionality:** 20+ individual function tests
- **Stress testing:** 10+ concurrent execution tests per sheet
- **Performance comparison:** Direct C vs C# timing comparisons
- **Memory testing:** Array allocation and deallocation tests
- **Thread safety:** Multi-threaded calculation verification
- **Error handling:** Invalid input testing

This comprehensive test suite will validate thread safety, performance, and correctness across all three add-in implementations.