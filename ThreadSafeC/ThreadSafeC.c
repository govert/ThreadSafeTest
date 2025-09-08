/*
**  ThreadSafeC Example
**  
**  Demonstrates thread-safe Excel function implementation using Excel SDK
*/

#include <windows.h>
#include <math.h>
#include <stdio.h>
#include <xlcall.h>
#include <framewrk.h>

/*
** rgFuncs
**
** This is a table of all functions exported by this module.
** These functions are registered in xlAutoOpen when the XLL loads.
** Format matches the last 7 arguments to REGISTER function.
*/
#define rgFuncsRows 14

static const LPWSTR rgFuncs[rgFuncsRows][7] = {
    {(LPWSTR)L"ThreadSafeCFunction", (LPWSTR)L"QQ$", (LPWSTR)L"ThreadSafeCFunction", (LPWSTR)L"input", (LPWSTR)L"1", (LPWSTR)L"Thread Safe Demo", (LPWSTR)L"Thread-safe version using manual allocation"},
    {(LPWSTR)L"ThreadSafeCalc", (LPWSTR)L"BB$", (LPWSTR)L"ThreadSafeCalc", (LPWSTR)L"number", (LPWSTR)L"1", (LPWSTR)L"Thread Safe Demo", (LPWSTR)L"Thread-safe calculation with $ flag"},
    {(LPWSTR)L"ThreadSafeXLOPER", (LPWSTR)L"QQ$", (LPWSTR)L"ThreadSafeXLOPER", (LPWSTR)L"input", (LPWSTR)L"1", (LPWSTR)L"Thread Safe Demo", (LPWSTR)L"Thread-safe XLOPER12 function"},
    {(LPWSTR)L"AllocatedMemoryFunction", (LPWSTR)L"QQ$", (LPWSTR)L"AllocatedMemoryFunction", (LPWSTR)L"size", (LPWSTR)L"1", (LPWSTR)L"Thread Safe Demo", (LPWSTR)L"Returns allocated memory requiring xlFree"},
    {(LPWSTR)L"ThreadInfoFunction", (LPWSTR)L"Q$", (LPWSTR)L"ThreadInfoFunction", (LPWSTR)L"", (LPWSTR)L"1", (LPWSTR)L"Thread Safe Demo", (LPWSTR)L"Returns thread info - thread safe"},
    {(LPWSTR)L"cInnerThreadInfo", (LPWSTR)L"Q$", (LPWSTR)L"cInnerThreadInfo", (LPWSTR)L"", (LPWSTR)L"1", (LPWSTR)L"Thread Safe Demo", (LPWSTR)L"Inner thread info for nested call"},
    {(LPWSTR)L"cNestedThreadInfo", (LPWSTR)L"Q$", (LPWSTR)L"cNestedThreadInfo", (LPWSTR)L"", (LPWSTR)L"1", (LPWSTR)L"Thread Safe Demo", (LPWSTR)L"Outer+Inner thread info via XlCall"},
    {(LPWSTR)L"cNestedThreadInfoEx", (LPWSTR)L"QB$", (LPWSTR)L"cNestedThreadInfoEx", (LPWSTR)L"external", (LPWSTR)L"1", (LPWSTR)L"Thread Safe Demo", (LPWSTR)L"Outer+Inner thread info, choose external C# call"},
    // Doubles as parameters (no XLOPERs)
    {(LPWSTR)L"cDoubleInner", (LPWSTR)L"BBB$", (LPWSTR)L"cDoubleInner", (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Thread Safe Demo", (LPWSTR)L"Inner double add (no XLOPER)"},
    {(LPWSTR)L"cDoubleCaller", (LPWSTR)L"BBB$", (LPWSTR)L"cDoubleCaller", (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Thread Safe Demo", (LPWSTR)L"Calls cDoubleInner via XlCall (no XLOPER)"},
    // Doubles wrapped inside XLOPER12
    {(LPWSTR)L"cXDoubleInner", (LPWSTR)L"QQQ$", (LPWSTR)L"cXDoubleInner", (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Thread Safe Demo", (LPWSTR)L"Inner double add (XLOPER)"},
    {(LPWSTR)L"cXDoubleCaller", (LPWSTR)L"QQQ$", (LPWSTR)L"cXDoubleCaller", (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Thread Safe Demo", (LPWSTR)L"Calls cXDoubleInner via XlCall (XLOPER)"},
    // Strings inside XLOPER12
    {(LPWSTR)L"cXStringInner", (LPWSTR)L"QQ$", (LPWSTR)L"cXStringInner", (LPWSTR)L"text", (LPWSTR)L"1", (LPWSTR)L"Thread Safe Demo", (LPWSTR)L"Inner string echo (XLOPER)"},
    {(LPWSTR)L"cXStringCaller", (LPWSTR)L"QQ$", (LPWSTR)L"cXStringCaller", (LPWSTR)L"text", (LPWSTR)L"1", (LPWSTR)L"Thread Safe Demo", (LPWSTR)L"Calls cXStringInner via XlCall (XLOPER)"}
};

/*
** ThreadSafeCFunction
**
** Calculates sqrt(input*3) + thread ID to demonstrate thread safety
** Takes XLOPER12 input and returns XLOPER12 result
** Thread-safe implementation using manual allocation
*/
__declspec(dllexport) LPXLOPER12 WINAPI ThreadSafeCFunction(LPXLOPER12 input)
{
    double inputValue = 0.0;
    DWORD threadId = GetCurrentThreadId();
    LPXLOPER12 result;
    
    // Extract double value from XLOPER12
    if (input && input->xltype == xltypeNum)
    {
        inputValue = input->val.num;
    }
    else if (input && input->xltype == xltypeInt)
    {
        inputValue = (double)input->val.w;
    }
    
    // Calculate result
    double value = sqrt(inputValue * 3.0) + (double)threadId;

    // Allocate XLOPER12 result manually (thread-safe) and mark for Excel to free
    result = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (result)
    {
        result->xltype = xltypeNum | xlbitDLLFree;
        result->val.num = value;
    }
    return result;
}

/*
** ThreadSafeCalc
**
** Thread-safe calculation using basic types only
** Registered with $ flag to indicate thread safety
*/
__declspec(dllexport) double WINAPI ThreadSafeCalc(double number)
{
    DWORD threadId = GetCurrentThreadId();
    
    // Simulate some calculation work
    Sleep(10); // Small delay to make threading effects more visible
    
    // Thread-safe calculation using only stack variables
    double result = number * number + sin(number) + (double)threadId;
    
    return result;
}

/*
** ThreadSafeXLOPER
**
** Thread-safe XLOPER12 function using manual memory management
** Registered with $ flag to indicate thread safety
*/
__declspec(dllexport) LPXLOPER12 WINAPI ThreadSafeXLOPER(LPXLOPER12 input)
{
    double inputValue = 0.0;
    DWORD threadId = GetCurrentThreadId();
    LPXLOPER12 result;
    
    // Extract double value from XLOPER12
    if (input && input->xltype == xltypeNum)
    {
        inputValue = input->val.num;
    }
    else if (input && input->xltype == xltypeInt)
    {
        inputValue = (double)input->val.w;
    }
    
    // Allocate memory manually for thread safety (avoid framework functions)
    result = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (result)
    {
        result->xltype = xltypeNum | xlbitDLLFree;  // Mark for Excel to free
        result->val.num = inputValue * 2.0 + (double)threadId;
    }
    
    return result;
}

/*
** AllocatedMemoryFunction
**
** Creates an array of numbers - demonstrates xlbitDLLFree and xlAutoFree12
** Returns allocated memory that Excel must free via xlAutoFree12
*/
__declspec(dllexport) LPXLOPER12 WINAPI AllocatedMemoryFunction(LPXLOPER12 sizeInput)
{
    int size = 5; // Default size
    LPXLOPER12 result;
    LPXLOPER12 arrayData;
    int i;
    DWORD threadId = GetCurrentThreadId();
    
    // Extract size from input
    if (sizeInput && sizeInput->xltype == xltypeNum)
    {
        size = (int)sizeInput->val.num;
        if (size < 1) size = 1;
        if (size > 100) size = 100; // Limit size
    }
    else if (sizeInput && sizeInput->xltype == xltypeInt)
    {
        size = sizeInput->val.w;
        if (size < 1) size = 1;
        if (size > 100) size = 100;
    }
    
    // Allocate the main XLOPER12
    result = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (!result) return NULL;
    
    // Allocate array data
    arrayData = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, size * sizeof(XLOPER12));
    if (!arrayData)
    {
        GlobalFree(result);
        return NULL;
    }
    
    // Fill the array with thread ID + index values
    for (i = 0; i < size; i++)
    {
        arrayData[i].xltype = xltypeNum;
        arrayData[i].val.num = (double)threadId + i;
    }
    
    // Set up the multi array
    result->xltype = xltypeMulti | xlbitDLLFree;
    result->val.array.lparray = arrayData;
    result->val.array.rows = size;
    result->val.array.columns = 1;
    
    return result;
}

/*
** ThreadInfoFunction
**
** Returns information about current thread - thread safe
** Takes no parameters and returns thread info as XLOPER12
*/
__declspec(dllexport) LPXLOPER12 WINAPI ThreadInfoFunction(void)
{
    DWORD threadId = GetCurrentThreadId();
    LPXLOPER12 result;
    wchar_t buffer[256];
    wchar_t* str;
    size_t len;
    
    // Create thread info string
    swprintf_s(buffer, 256, L"Thread: %lu, Time: %lu", threadId, GetTickCount());
    
    // Allocate memory manually for thread safety
    len = wcslen(buffer);
    result = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (result)
    {
        str = (wchar_t*)GlobalAlloc(GMEM_FIXED, (len + 2) * sizeof(wchar_t));
        if (str)
        {
            str[0] = (wchar_t)len;  // Length prefix
            wcscpy_s(&str[1], len + 1, buffer);
            
            result->xltype = xltypeStr | xlbitDLLFree;  // Mark for Excel to free
            result->val.str = str;
        }
        else
        {
            GlobalFree(result);
            result = NULL;
        }
    }
    
    return result;
}

/*
** cInnerThreadInfo
** Returns a string with current (inner) thread id - thread safe
*/
__declspec(dllexport) LPXLOPER12 WINAPI cInnerThreadInfo(void)
{
    DWORD threadId = GetCurrentThreadId();
    LPXLOPER12 result;
    wchar_t buffer[64];
    wchar_t* str;
    size_t len;

    swprintf_s(buffer, 64, L"InnerThread:%lu", threadId);

    len = wcslen(buffer);
    result = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (result)
    {
        str = (wchar_t*)GlobalAlloc(GMEM_FIXED, (len + 2) * sizeof(wchar_t));
        if (str)
        {
            str[0] = (wchar_t)len;
            wcscpy_s(&str[1], len + 1, buffer);
            result->xltype = xltypeStr | xlbitDLLFree;
            result->val.str = str;
        }
        else
        {
            GlobalFree(result);
            result = NULL;
        }
    }
    return result;
}

/*
** cNestedThreadInfo
** Calls cInnerThreadInfo via XlCall and returns a string with both outer and inner thread ids
*/
__declspec(dllexport) LPXLOPER12 WINAPI cNestedThreadInfo(void)
{
    DWORD outerThreadId = GetCurrentThreadId();
    XLOPER12 inner;
    int callRes;
    LPXLOPER12 result;
    wchar_t* outStr;
    wchar_t innerPart[128];
    size_t innerLen = 0;
    wchar_t outerPart[64];
    size_t outerLen;

    // Prepare outer part
    swprintf_s(outerPart, 64, L"OuterThread:%lu; ", outerThreadId);
    outerLen = wcslen(outerPart);

    // Call inner function via Excel
    callRes = Excel12f(xlUDF, &inner, 1, TempStr12(L"cInnerThreadInfo"));
    if (callRes == xlretSuccess && (inner.xltype & xltypeStr) == xltypeStr && inner.val.str)
    {
        // inner.val.str is Excel-style Pascal string [len][chars...]
        innerLen = (size_t)inner.val.str[0];
        if (innerLen > 120) innerLen = 120; // safety
        wcsncpy_s(innerPart, 128, &inner.val.str[1], innerLen);
        innerPart[innerLen] = L'\0';
    }
    else
    {
        wcscpy_s(innerPart, 128, L"InnerError");
        innerLen = wcslen(innerPart);
    }

    // Allocate result string: outerPart + innerPart
    result = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (result)
    {
        size_t totalLen = outerLen + innerLen;
        outStr = (wchar_t*)GlobalAlloc(GMEM_FIXED, (totalLen + 2) * sizeof(wchar_t));
        if (outStr)
        {
            outStr[0] = (wchar_t)totalLen;
            // Copy outer then inner
            wcscpy_s(&outStr[1], totalLen + 1, outerPart);
            wcscat_s(&outStr[1], totalLen + 1, innerPart);
            result->xltype = xltypeStr | xlbitDLLFree;
            result->val.str = outStr;
        }
        else
        {
            GlobalFree(result);
            result = NULL;
        }
    }

    // Free inner result via Excel, if applicable
    if (callRes == xlretSuccess)
    {
        Excel12f(xlFree, 0, 1, (LPXLOPER12)&inner);
    }

    return result;
}

/*
** cNestedThreadInfoEx
** Same as cNestedThreadInfo, but takes a flag (non-zero -> call external C# function csInnerThreadInfo)
*/
__declspec(dllexport) LPXLOPER12 WINAPI cNestedThreadInfoEx(double external)
{
    DWORD outerThreadId = GetCurrentThreadId();
    XLOPER12 inner;
    int callRes;
    LPXLOPER12 result;
    wchar_t* outStr;
    wchar_t innerPart[128];
    size_t innerLen = 0;
    wchar_t outerPart[64];
    size_t outerLen;
    const wchar_t* target = (external != 0.0) ? L"csInnerThreadInfo" : L"cInnerThreadInfo";

    swprintf_s(outerPart, 64, L"OuterThread:%lu; ", outerThreadId);
    outerLen = wcslen(outerPart);

    callRes = Excel12f(xlUDF, &inner, 1, TempStr12((LPWSTR)target));
    if (callRes == xlretSuccess && (inner.xltype & xltypeStr) == xltypeStr && inner.val.str)
    {
        innerLen = (size_t)inner.val.str[0];
        if (innerLen > 120) innerLen = 120;
        wcsncpy_s(innerPart, 128, &inner.val.str[1], innerLen);
        innerPart[innerLen] = L'\0';
    }
    else
    {
        wcscpy_s(innerPart, 128, L"InnerError");
        innerLen = wcslen(innerPart);
    }

    result = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (result)
    {
        size_t totalLen = outerLen + innerLen;
        outStr = (wchar_t*)GlobalAlloc(GMEM_FIXED, (totalLen + 2) * sizeof(wchar_t));
        if (outStr)
        {
            outStr[0] = (wchar_t)totalLen;
            wcscpy_s(&outStr[1], totalLen + 1, outerPart);
            wcscat_s(&outStr[1], totalLen + 1, innerPart);
            result->xltype = xltypeStr | xlbitDLLFree;
            result->val.str = outStr;
        }
        else
        {
            GlobalFree(result);
            result = NULL;
        }
    }

    if (callRes == xlretSuccess)
    {
        Excel12f(xlFree, 0, 1, (LPXLOPER12)&inner);
    }

    return result;
}

// ===== Doubles (no XLOPERs) =====
__declspec(dllexport) double WINAPI cDoubleInner(double x, double y)
{
    return x + y;
}

__declspec(dllexport) double WINAPI cDoubleCaller(double x, double y)
{
    XLOPER12 ret;
    int rc = Excel12f(xlUDF, &ret, 3, TempStr12(L"cDoubleInner"), TempNum12(x), TempNum12(y));
    if (rc == xlretSuccess && (ret.xltype & xltypeNum) == xltypeNum)
        return ret.val.num;
    return 0.0;
}

// ===== Doubles inside XLOPERs =====
__declspec(dllexport) LPXLOPER12 WINAPI cXDoubleInner(LPXLOPER12 x, LPXLOPER12 y)
{
    double xv = 0.0, yv = 0.0;
    if (x)
    {
        if ((x->xltype & xltypeNum) == xltypeNum) xv = x->val.num;
        else if ((x->xltype & xltypeInt) == xltypeInt) xv = (double)x->val.w;
    }
    if (y)
    {
        if ((y->xltype & xltypeNum) == xltypeNum) yv = y->val.num;
        else if ((y->xltype & xltypeInt) == xltypeInt) yv = (double)y->val.w;
    }
    LPXLOPER12 res = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (res)
    {
        res->xltype = xltypeNum | xlbitDLLFree;
        res->val.num = xv + yv;
    }
    return res;
}

__declspec(dllexport) LPXLOPER12 WINAPI cXDoubleCaller(LPXLOPER12 x, LPXLOPER12 y)
{
    XLOPER12 inner;
    int rc = Excel12f(xlUDF, &inner, 3, TempStr12(L"cXDoubleInner"), (LPXLOPER12)x, (LPXLOPER12)y);
    LPXLOPER12 res = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (!res) return NULL;
    if (rc == xlretSuccess && (inner.xltype & xltypeNum) == xltypeNum)
    {
        res->xltype = xltypeNum | xlbitDLLFree;
        res->val.num = inner.val.num;
    }
    else
    {
        res->xltype = xltypeNum | xlbitDLLFree;
        res->val.num = 0.0;
    }
    return res;
}

// ===== Strings inside XLOPERs =====
__declspec(dllexport) LPXLOPER12 WINAPI cXStringInner(LPXLOPER12 s)
{
    const wchar_t* prefix = L"Echo:";
    size_t plen = wcslen(prefix);
    const wchar_t* in = L"";
    size_t ilen = 0;
    if (s && (s->xltype & xltypeStr) == xltypeStr && s->val.str)
    {
        ilen = (size_t)s->val.str[0];
        if (ilen > 240) ilen = 240;
        in = &s->val.str[1];
    }
    LPXLOPER12 res = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (!res) return NULL;
    wchar_t* out = (wchar_t*)GlobalAlloc(GMEM_FIXED, (plen + ilen + 2) * sizeof(wchar_t));
    if (!out)
    {
        GlobalFree(res);
        return NULL;
    }
    out[0] = (wchar_t)(plen + ilen);
    wcscpy_s(&out[1], plen + ilen + 1, prefix);
    wcsncat_s(&out[1], plen + ilen + 1, in, ilen);
    res->xltype = xltypeStr | xlbitDLLFree;
    res->val.str = out;
    return res;
}

__declspec(dllexport) LPXLOPER12 WINAPI cXStringCaller(LPXLOPER12 s)
{
    XLOPER12 inner;
    int rc = Excel12f(xlUDF, &inner, 2, TempStr12(L"cXStringInner"), (LPXLOPER12)s);
    LPXLOPER12 res = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (!res) return NULL;
    if (rc == xlretSuccess && (inner.xltype & xltypeStr) == xltypeStr && inner.val.str)
    {
        size_t ilen = (size_t)inner.val.str[0];
        wchar_t* out = (wchar_t*)GlobalAlloc(GMEM_FIXED, (ilen + 2) * sizeof(wchar_t));
        if (!out)
        {
            GlobalFree(res);
            return NULL;
        }
        out[0] = (wchar_t)ilen;
        wcsncpy_s(&out[1], ilen + 1, &inner.val.str[1], ilen);
        res->xltype = xltypeStr | xlbitDLLFree;
        res->val.str = out;
        Excel12f(xlFree, 0, 1, (LPXLOPER12)&inner);
    }
    else
    {
        wchar_t* out = (wchar_t*)GlobalAlloc(GMEM_FIXED, 2 * sizeof(wchar_t));
        if (!out)
        {
            GlobalFree(res);
            return NULL;
        }
        out[0] = 0;
        out[1] = 0;
        res->xltype = xltypeStr | xlbitDLLFree;
        res->val.str = out;
    }
    return res;
}

/*
** xlAutoOpen
**
** Called by Excel when the XLL is loaded. Registers all functions
** and performs initialization.
*/
__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
    static XLOPER12 xDLL;
    int i;

    // Get the name of this XLL
    Excel12f(xlGetName, &xDLL, 0);

    // Register all functions in the rgFuncs table
    for (i = 0; i < rgFuncsRows; i++) 
    {
        Excel12f(xlfRegister, 0, 1 + 7,
            (LPXLOPER12)&xDLL,
            TempStr12(rgFuncs[i][0]),   // Function name
            TempStr12(rgFuncs[i][1]),   // Type signature
            TempStr12(rgFuncs[i][2]),   // Function text
            TempStr12(rgFuncs[i][3]),   // Argument text
            TempStr12(rgFuncs[i][4]),   // Macro type
            TempStr12(rgFuncs[i][5]),   // Category
            TempStr12(L""),             // Shortcut text
            TempStr12(L""),             // Help topic
            TempStr12(rgFuncs[i][6]),   // Function help
            TempStr12(rgFuncs[i][3])    // Argument help
        );
    }

    // Free temporary memory used by framework
    Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);

    return 1;
}

/*
** xlAutoClose
**
** Called by Excel when the XLL is unloaded. Performs cleanup.
*/
__declspec(dllexport) int WINAPI xlAutoClose(void)
{
    int i;
    
    // Delete function names to clean up Excel's namespace
    for (i = 0; i < rgFuncsRows; i++)
        Excel12f(xlfSetName, 0, 1, TempStr12(rgFuncs[i][2]));
    
    return 1;
}

/*
** xlAutoFree12
**
** Called by Excel to free XLOPER12 memory allocated by our functions.
** This handles memory for functions that use xlbitDLLFree flag.
*/
__declspec(dllexport) void WINAPI xlAutoFree12(LPXLOPER12 pxFree)
{
    if (pxFree == NULL) return;
    
    // Handle different XLOPER12 types that we allocated
    switch (pxFree->xltype & ~xlbitDLLFree)
    {
        case xltypeStr:
            // Free string data allocated by ThreadInfoFunction
            if (pxFree->val.str)
            {
                GlobalFree(pxFree->val.str);
                pxFree->val.str = NULL;
            }
            break;
            
        case xltypeMulti:
            // Free array data allocated by AllocatedMemoryFunction
            if (pxFree->val.array.lparray)
            {
                GlobalFree(pxFree->val.array.lparray);
                pxFree->val.array.lparray = NULL;
            }
            break;
            
        case xltypeNum:
            // For simple numbers (ThreadSafeXLOPER), no additional cleanup needed
            // Just the XLOPER12 structure itself will be freed by Excel
            break;
            
        default:
            // Handle any other types if needed
            break;
    }
    
    // Excel will free the main XLOPER12 structure itself
    return;
}
