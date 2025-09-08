# Thread-Safe Add-in Tests (Quick Guide)

This page lists the functions and formulas to exercise thread-safe behavior across the C and .NET add-ins, including nested calls and various parameter marshalling styles.

## Add-ins
- C add-in: `ThreadSafeC.xll`
- .NET add-in: `ThreadSafeNet.xll`
- Wrapper/tests: `ThreadSafeTest.xll`

Ensure both add-ins are alongside `ThreadSafeTest.xll` in the Output folder so the wrapper can auto-load them.

## Nested Calls (Thread IDs)
- Same add-in (.NET): `=tsNestedThreadInfo()` → returns `OuterThread:…; InnerThread:…`
- Cross add-in (.NET → .NET): `=tsNestedThreadInfo(TRUE)` → inner comes from `ThreadSafeNet`
- C only: `=cNestedThreadInfo()` → inner is `cInnerThreadInfo`
- C with external inner: `=cNestedThreadInfoEx(1)` → calls `.NET` `csInnerThreadInfo`

## Doubles (no XLOPERs)
- .NET inner: `=csDoubleInner(2, 3)` → `5`
- .NET caller→inner: `=csDoubleCaller(2, 3)` → `5`
- C inner: `=cDoubleInner(2, 3)` → `5`
- C caller→inner: `=cDoubleCaller(2, 3)` → `5`
- C caller (no Temp helpers, per-thread args): `=cDoubleCallerTLS(2, 3)` → `5`

## Doubles inside XLOPERs (C only)
- C inner: `=cXDoubleInner(2, 3)` → `5`
- C caller→inner: `=cXDoubleCaller(2, 3)` → `5`

## Strings inside XLOPERs (C only)
- C inner: `=cXStringInner("abc")` → `Echo:abc`
- C caller→inner: `=cXStringCaller("abc")` → `Echo:abc`

## Existing Core Tests
- C direct:
  - `=ThreadSafeCFunction(A1)` (XLOPER in/out)
  - `=ThreadSafeCalc(A1)` (double)
  - `=ThreadSafeXLOPER(A1)` (XLOPER in/out)
  - `=AllocatedMemoryFunction(5)` (array)
  - `=ThreadInfoFunction()` (string)
- .NET direct:
  - `=csThreadSafeCFunction(A1)` (double)
  - `=csThreadSafeCalc(A1)` (double)
  - `=csThreadSafeXLOPER(A1)` (double)
  - `=csAllocatedMemoryFunction(5)` (array)
  - `=csThreadInfoFunction()` (string)
- Wrapper (ThreadSafeTest):
  - `=TestThreadSafeCFunction(val, useCSharp)`
  - `=TestThreadSafeCalc(val, useCSharp)`
  - `=TestThreadSafeXLOPER(val, useCSharp)`
  - `=TestAllocatedMemoryFunction(size, useCSharp)`
  - `=TestThreadInfoFunction(useCSharp)`
  - `=TestMultipleThreadSafeCalls(val, useCSharp)`

## Notes
- All functions marked as thread-safe (`$` in C registration; `IsThreadSafe=true` in .NET attributes).
- For cross add-in calls, ensure both XLLs are present in the same directory and loaded.
