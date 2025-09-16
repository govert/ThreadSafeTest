#include <windows.h>
#include <math.h>
#include <stdio.h>
#include <wchar.h>
#include "XLCALL.H"
#include "FRAMEWRK.H"
#include <stdarg.h>

// Debug output helper: writes formatted wide string to debugger output
static void DebugPrintW(const wchar_t* fmt, ...)
{
    wchar_t buffer[512];
    va_list args;
    va_start(args, fmt);
    _vsnwprintf_s(buffer, _countof(buffer), _TRUNCATE, fmt, args);
    va_end(args);
    OutputDebugStringW(buffer);
}

// Functions (thread-safe)
#define rgFuncsRows 12
static const LPWSTR rgFuncs[rgFuncsRows][7] = {
    {(LPWSTR)L"cDoubleInner",  (LPWSTR)L"BBB$", (LPWSTR)L"cDoubleInner",  (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Inner add: returns x+y"},
    {(LPWSTR)L"cDoubleCaller", (LPWSTR)L"BBB$", (LPWSTR)L"cDoubleCaller", (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Caller: calls by name (leaks name XLOPER)"},
    {(LPWSTR)L"cDoubleCallerById", (LPWSTR)L"BBB$", (LPWSTR)L"cDoubleCallerById", (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Caller: calls by register id"},
    {(LPWSTR)L"cDoubleCallerDirect", (LPWSTR)L"BBB$", (LPWSTR)L"cDoubleCallerDirect", (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Caller: calls by name, w/out framework"},
    {(LPWSTR)L"cDoubleCallerDirectById", (LPWSTR)L"BBB$", (LPWSTR)L"cDoubleCallerDirectById", (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Caller: calls by register id, w/out framework"},
    // Excel12Direct test functions: direct MdCallBack12 calls
    {(LPWSTR)L"cDoubleCallerExcel12Direct", (LPWSTR)L"BBB$", (LPWSTR)L"cDoubleCallerExcel12Direct", (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Direct MdCallBack12: calls by name"},
    {(LPWSTR)L"cDoubleCallerExcel12DirectById", (LPWSTR)L"BBB$", (LPWSTR)L"cDoubleCallerExcel12DirectById", (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Direct MdCallBack12: calls by register id"},
    // XLOPER12 string functions: return Q, take two Q args; thread-safe ($)
    {(LPWSTR)L"cStringsInner", (LPWSTR)L"QQQ$", (LPWSTR)L"cStringsInner", (LPWSTR)L"str1,str2", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Inner concat: returns str1+str2"},
    {(LPWSTR)L"cStringsCaller", (LPWSTR)L"QQQ$", (LPWSTR)L"cStringsCaller", (LPWSTR)L"str1,str2", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Caller: calls by name, w/out framework"},
    {(LPWSTR)L"cStringsCallerDirectById", (LPWSTR)L"QQQ$", (LPWSTR)L"cStringsCallerDirectById", (LPWSTR)L"str1,str2", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Caller: calls by register id, w/out framework"},
    // Memory-managed variants
    {(LPWSTR)L"cStringsFreeInner", (LPWSTR)L"QQQ$", (LPWSTR)L"cStringsFreeInner", (LPWSTR)L"str1,str2", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Inner concat: returns str1+str2 (DLLFree)"},
    {(LPWSTR)L"cStringsFreeDirectById", (LPWSTR)L"QQQ$", (LPWSTR)L"cStringsFreeDirectById", (LPWSTR)L"str1,str2", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Caller: by register id (managed)"}
};

// Register id captured for cDoubleInner and cStringsInner
static XLOPER12 g_reg_cDoubleInner = { 0 };
static XLOPER12 g_reg_cStringsInner = { 0 };
static XLOPER12 g_reg_cStringsFreeInner = { 0 };

// Direct MdCallBack12 function pointer and initialization
typedef int (__stdcall *MDCALLBACK12_PROC)(int xlfn, int count, LPXLOPER12 *opers, LPXLOPER12 operRes);
static MDCALLBACK12_PROC g_pMdCallBack12 = NULL;
static HMODULE g_hmoduleExcel = NULL;

// Initialize direct MdCallBack12 access
static int InitMdCallBack12(void)
{
    if (g_pMdCallBack12)
        return 1; // Already initialized

    // Get handle to current process (Excel.exe)
    g_hmoduleExcel = GetModuleHandleA(NULL);
    if (!g_hmoduleExcel)
    {
        DebugPrintW(L"[MultithreadCrash] Failed to get Excel module handle\n");
        return 0;
    }

    // Get the MdCallBack12 function pointer
    g_pMdCallBack12 = (MDCALLBACK12_PROC)GetProcAddress(g_hmoduleExcel, "MdCallBack12");
    if (!g_pMdCallBack12)
    {
        DebugPrintW(L"[MultithreadCrash] Failed to get MdCallBack12 proc address\n");
        return 0;
    }

    DebugPrintW(L"[MultithreadCrash] Successfully initialized MdCallBack12 at %p\n", g_pMdCallBack12);
    return 1;
}

// Excel12Direct: Direct call to MdCallBack12 bypassing framework
static int _cdecl Excel12Direct(int xlfn, LPXLOPER12 operRes, int count, ...)
{
    // Initialize MdCallBack12 if needed
    if (!InitMdCallBack12())
        return xlretFailed;

    // Validate parameters
    if (count < 0 || count > 255)
        return xlretInvCount;

    // Allocate array for XLOPER12 pointers
    LPXLOPER12 opers[256]; // Max 255 args + safety margin
    
    if (count > 0)
    {
        va_list args;
        va_start(args, count);
        
        // Extract variadic arguments into array
        for (int i = 0; i < count; i++)
        {
            opers[i] = va_arg(args, LPXLOPER12);
        }
        
        va_end(args);
    }

    DWORD tid = GetCurrentThreadId();
    DebugPrintW(L"[MultithreadCrash] Thread %lu: Excel12Direct calling MdCallBack12(xlfn=%d, count=%d)\n", 
                tid, xlfn, count);

    // Call MdCallBack12 directly with the signature we determined:
    // int MdCallBack12(int xlfn, int count, LPXLOPER12 *opers, LPXLOPER12 operRes)
    int result = g_pMdCallBack12(xlfn, count, opers, operRes);

    DebugPrintW(L"[MultithreadCrash] Thread %lu: Excel12Direct returned %d\n", tid, result);
    return result;
}

// cDoubleInner: returns x+y
__declspec(dllexport) double WINAPI cDoubleInner(double x, double y)
{
	// Sleep this thread for 100 ms to simulate some work using the Windows API
	Sleep(100);
    return x + y;
}

// cStringsInner: concatenates two strings
__declspec(dllexport) LPXLOPER12 WINAPI cStringsInner(LPXLOPER12 str1, LPXLOPER12 str2)
{
    // Sleep this thread for 50 ms to simulate some work
    // Sleep(50);
    
    // Allocate result XLOPER12 on the heap (intentionally not freed)
    LPXLOPER12 result = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (!result)
        return NULL;
    
    // Check if both inputs are strings
    if ((str1->xltype & xltypeStr) != xltypeStr || (str2->xltype & xltypeStr) != xltypeStr)
    {
        // Return empty string on invalid input
        result->xltype = xltypeStr;
        result->val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, 2 * sizeof(XCHAR));
        if (result->val.str)
        {
            result->val.str[0] = 0; // Length 0
            result->val.str[1] = L'\0';
        }
        return result;
    }
    
    // Get string lengths (first character is length for Excel strings)
    int len1 = str1->val.str[0];
    int len2 = str2->val.str[0];
    int totalLen = len1 + len2;
    
    // Limit total length to avoid excessive allocation
    if (totalLen > 255) totalLen = 255;
    
    // Allocate string buffer (length prefix + string + null terminator)
    result->val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, (totalLen + 2) * sizeof(XCHAR));
    if (!result->val.str)
    {
        GlobalFree(result);
        return NULL;
    }
    
    // Set length prefix
    result->val.str[0] = (XCHAR)totalLen;
    result->xltype = xltypeStr;
    
    // Copy first string
    int copyLen1 = (len1 < totalLen) ? len1 : totalLen;
    if (copyLen1 > 0)
        wcsncpy_s(&result->val.str[1], totalLen + 1, &str1->val.str[1], copyLen1);
    
    // Copy second string
    int copyLen2 = totalLen - copyLen1;
    if (copyLen2 > 0)
        wcsncpy_s(&result->val.str[1 + copyLen1], totalLen + 1 - copyLen1, &str2->val.str[1], copyLen2);
    
    // Null terminate
    result->val.str[totalLen + 1] = L'\0';
    
    return result;
}

// cStringsFreeInner: concatenates two strings and returns a value that Excel will free via xlAutoFree12
__declspec(dllexport) LPXLOPER12 WINAPI cStringsFreeInner(LPXLOPER12 str1, LPXLOPER12 str2)
{
    DWORD tid = GetCurrentThreadId();
    DebugPrintW(L"[MultithreadCrash] Thread %lu: cStringsFreeInner called\n", tid);


    // Determine input validity and lengths
    int isStr1 = (str1 && ((str1->xltype & xltypeStr) == xltypeStr));
    int isStr2 = (str2 && ((str2->xltype & xltypeStr) == xltypeStr));

    int len1 = isStr1 ? str1->val.str[0] : 0;
    int len2 = isStr2 ? str2->val.str[0] : 0;
    int totalLen = len1 + len2;
    if (totalLen > 255) totalLen = 255;

    // Allocate result XLOPER12 and its string; Excel will later call xlAutoFree12 to free
    LPXLOPER12 result = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (!result)
        return NULL;

    result->xltype = xltypeStr | xlbitDLLFree;
    result->val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, (totalLen + 2) * sizeof(XCHAR));
    if (!result->val.str)
    {
        GlobalFree(result);
        return NULL;
    }

    // Build the result string
    result->val.str[0] = (XCHAR)totalLen;

    int copyLen1 = (len1 < totalLen) ? len1 : totalLen;
    if (copyLen1 > 0 && isStr1)
        wcsncpy_s(&result->val.str[1], totalLen + 1, &str1->val.str[1], copyLen1);

    int copyLen2 = totalLen - copyLen1;
    if (copyLen2 > 0 && isStr2)
        wcsncpy_s(&result->val.str[1 + copyLen1], totalLen + 1 - copyLen1, &str2->val.str[1], copyLen2);

    result->val.str[totalLen + 1] = L'\0';
    return result;
}

// cDoubleCaller: calls by NAME using a newly allocated XLOPER string (intentionally leaked). TLS numeric args; no Temp helpers.
__declspec(dllexport) double WINAPI cDoubleCaller(double x, double y)
{
    static __declspec(thread) int tls_init = 0;
    static __declspec(thread) LPXLOPER12 tls_x = NULL;
    static __declspec(thread) LPXLOPER12 tls_y = NULL;

    if (!tls_init)
    {
        DWORD tid = GetCurrentThreadId();
        DebugPrintW(L"[MultithreadCrash] Thread %lu: init TLS numeric args\n", tid);
        tls_x = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
        tls_y = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
        if (tls_x) { tls_x->xltype = xltypeNum; tls_x->val.num = 0.0; }
        if (tls_y) { tls_y->xltype = xltypeNum; tls_y->val.num = 0.0; }
        tls_init = 1;
    }

    if (tls_x) tls_x->val.num = x;
    if (tls_y) tls_y->val.num = y;

    // Allocate a fresh function name XLOPER12 on the heap and DO NOT FREE (intentional leak for test)
    LPXLOPER12 fnArg = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    LPWSTR fnStr = (LPWSTR)GlobalAlloc(GMEM_FIXED, 32 * sizeof(wchar_t));
    if (fnArg && fnStr)
    {
        const wchar_t* name = L"cDoubleInner";
        size_t nlen = wcslen(name); if (nlen > 30) nlen = 30;
        fnStr[0] = (wchar_t)nlen;
        wcsncpy_s(&fnStr[1], 31, name, nlen);
        fnArg->xltype = xltypeStr;
        fnArg->val.str = fnStr;
    }
    // Note: fnArg and fnStr are intentionally not freed to avoid any cross-thread reuse — this leaks memory by design for this test.

    XLOPER12 ret;
    int rc = Excel12f(xlUDF, &ret, 3, fnArg, tls_x, tls_y);
    if (rc == xlretSuccess && (ret.xltype & xltypeNum) == xltypeNum)
        return ret.val.num;
    return 0.0;
}

__declspec(dllexport) double WINAPI cDoubleCallerDirect(double x, double y)
{
    // Allocate an array of XLOPER12s on the heap
    LPXLOPER12 args = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, 3 * sizeof(XLOPER12));
    if (!args)
        return 0.0; // Allocation failed

    // Initialize the first argument: function name
    args[0].xltype = xltypeStr;
    args[0].val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, 14 * sizeof(XCHAR)); // "cDoubleInner" + length prefix
    if (!args[0].val.str)
    {
        GlobalFree(args);
        return 0.0; // Allocation failed
    }
    args[0].val.str[0] = 12; // Length prefix
    wcscpy_s(&args[0].val.str[1], 13, L"cDoubleInner");

    // Initialize the second argument: x
    args[1].xltype = xltypeNum;
    args[1].val.num = x;

    // Initialize the third argument: y
    args[2].xltype = xltypeNum;
    args[2].val.num = y;

    // Prepare the result XLOPER12
    XLOPER12 result;
    ZeroMemory(&result, sizeof(XLOPER12));

    // Call Excel12 directly
    int rc = Excel12(xlUDF, &result, 3, &args[0], &args[1], &args[2]);

    // // Free allocated memory for arguments
    // GlobalFree(args[0].val.str);
    // GlobalFree(args);

    // Check the result and return the value if successful
    if (rc == xlretSuccess && (result.xltype & xltypeNum) == xltypeNum)
    {
        double returnValue = result.val.num;

        return returnValue;
    }

    return 0.0; // Default return value on failure
}

// cDoubleCallerById: calls by REGISTER ID (g_reg_cDoubleInner) and TLS numeric args. No Temp helpers.
__declspec(dllexport) double WINAPI cDoubleCallerById(double x, double y)
{
    static __declspec(thread) int tls_init = 0;
    static __declspec(thread) LPXLOPER12 tls_x = NULL;
    static __declspec(thread) LPXLOPER12 tls_y = NULL;

    if (!tls_init)
    {
        tls_x = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
        tls_y = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
        if (tls_x) { tls_x->xltype = xltypeNum; tls_x->val.num = 0.0; }
        if (tls_y) { tls_y->xltype = xltypeNum; tls_y->val.num = 0.0; }
        tls_init = 1;
    }

    if (tls_x) tls_x->val.num = x;
    if (tls_y) tls_y->val.num = y;

    if ((g_reg_cDoubleInner.xltype & xltypeNum) != xltypeNum)
        return 0.0; // ID not available

    XLOPER12 ret;
    int rc = Excel12f(xlUDF, &ret, 3, (LPXLOPER12)&g_reg_cDoubleInner, tls_x, tls_y);
    if (rc == xlretSuccess && (ret.xltype & xltypeNum) == xltypeNum)
        return ret.val.num;
    return 0.0;
}

// cDoubleCallerDirectById: calls cDoubleInner using its registration ID directly
__declspec(dllexport) double WINAPI cDoubleCallerDirectById(double x, double y)
{
    // Allocate an array of XLOPER12s on the heap
    LPXLOPER12 args = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, 2 * sizeof(XLOPER12));
    if (!args)
        return 0.0; // Allocation failed

    // Initialize the first argument: x
    args[0].xltype = xltypeNum;
    args[0].val.num = x;

    // Initialize the second argument: y
    args[1].xltype = xltypeNum;
    args[1].val.num = y;

    // Prepare the result XLOPER12
    XLOPER12 result;
    ZeroMemory(&result, sizeof(XLOPER12));

    // Call Excel12 directly using the registration ID
    int rc = Excel12(xlUDF, &result, 3, (LPXLOPER12)&g_reg_cDoubleInner, &args[0], &args[1]);

    // Free allocated memory for arguments
    GlobalFree(args);

    // Check the result and return the value if successful
    if (rc == xlretSuccess && (result.xltype & xltypeNum) == xltypeNum)
    {
        double returnValue = result.val.num;

        // Free the result if necessary
        if (result.xltype & xlbitXLFree)
            Excel12(xlFree, 0, 1, &result);

        return returnValue;
    }

    return 0.0; // Default return value on failure
}

// cStringsCaller: calls cStringsInner by NAME, allocating a new name XLOPER12 each call (no frees)
__declspec(dllexport) LPXLOPER12 WINAPI cStringsCaller(LPXLOPER12 str1, LPXLOPER12 str2)
{
    // Allocate and build function name XLOPER12 (intentional leak per call)
    LPXLOPER12 fnArg = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    const wchar_t* fname = L"cStringsInner";
    size_t fnlen = wcslen(fname);
    XCHAR* fnStr = (XCHAR*)GlobalAlloc(GMEM_FIXED, (fnlen + 2) * sizeof(XCHAR));
    if (fnArg && fnStr)
    {
        fnStr[0] = (XCHAR)fnlen;
        wcscpy_s(&fnStr[1], fnlen + 1, fname);
        fnArg->xltype = xltypeStr;
        fnArg->val.str = fnStr;
    }

    // Allocate an array of XLOPER12s on the heap for str1 and str2 (intentional leak)
    LPXLOPER12 args = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, 2 * sizeof(XLOPER12));
    if (!args)
        return NULL;

    // Initialize the first argument: str1 (copy or empty)
    args[0].xltype = xltypeStr;
    if (str1 && ((str1->xltype & xltypeStr) == xltypeStr))
    {
        int len1 = str1->val.str[0];
        args[0].val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, (len1 + 2) * sizeof(XCHAR));
        if (args[0].val.str)
        {
            args[0].val.str[0] = str1->val.str[0];
            wcsncpy_s(&args[0].val.str[1], len1 + 1, &str1->val.str[1], len1);
            args[0].val.str[len1 + 1] = L'\0';
        }
    }
    else
    {
        args[0].val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, 2 * sizeof(XCHAR));
        if (args[0].val.str)
        {
            args[0].val.str[0] = 0;
            args[0].val.str[1] = L'\0';
        }
    }

    // Initialize the second argument: str2 (copy or empty)
    args[1].xltype = xltypeStr;
    if (str2 && ((str2->xltype & xltypeStr) == xltypeStr))
    {
        int len2 = str2->val.str[0];
        args[1].val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, (len2 + 2) * sizeof(XCHAR));
        if (args[1].val.str)
        {
            args[1].val.str[0] = str2->val.str[0];
            wcsncpy_s(&args[1].val.str[1], len2 + 1, &str2->val.str[1], len2);
            args[1].val.str[len2 + 1] = L'\0';
        }
    }
    else
    {
        args[1].val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, 2 * sizeof(XCHAR));
        if (args[1].val.str)
        {
            args[1].val.str[0] = 0;
            args[1].val.str[1] = L'\0';
        }
    }

    // Prepare the result XLOPER12
    XLOPER12 result;
    ZeroMemory(&result, sizeof(XLOPER12));

    // Call Excel12 using function name and two string arguments
    int rc = Excel12(xlUDF, &result, 3, fnArg, &args[0], &args[1]);

    // On success, copy result to a freshly allocated return object (intentional leak)
    if (rc == xlretSuccess && (result.xltype & xltypeStr) == xltypeStr)
    {
        LPXLOPER12 returnValue = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
        if (returnValue)
        {
            int resultLen = result.val.str[0];
            returnValue->xltype = xltypeStr;
            returnValue->val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, (resultLen + 2) * sizeof(XCHAR));
            if (returnValue->val.str)
            {
                returnValue->val.str[0] = result.val.str[0];
                wcsncpy_s(&returnValue->val.str[1], resultLen + 1, &result.val.str[1], resultLen);
                returnValue->val.str[resultLen + 1] = L'\0';
            }
        }
        return returnValue;
    }

    // Failure case: return empty string (intentional leak)
    LPXLOPER12 emptyResult = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (emptyResult)
    {
        emptyResult->xltype = xltypeStr;
        emptyResult->val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, 2 * sizeof(XCHAR));
        if (emptyResult->val.str)
        {
            emptyResult->val.str[0] = 0;
            emptyResult->val.str[1] = L'\0';
        }
    }
    return emptyResult;
}

// cStringsCallerDirectById: calls cStringsInner using its registration ID directly
__declspec(dllexport) LPXLOPER12 WINAPI cStringsCallerDirectById(LPXLOPER12 str1, LPXLOPER12 str2)
{
	// Write Debug info with thread ID
	DWORD tid = GetCurrentThreadId();
	DebugPrintW(L"[MultithreadCrash] Thread %lu: cStringsCallerDirectById called\n", tid);

    // Allocate an array of XLOPER12s on the heap
    LPXLOPER12 args = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, 2 * sizeof(XLOPER12));
    if (!args)
        return NULL; // Allocation failed

    // Initialize the first argument: str1
    args[0].xltype = xltypeStr;
    if ((str1->xltype & xltypeStr) == xltypeStr)
    {
        int len1 = str1->val.str[0];
        args[0].val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, (len1 + 2) * sizeof(XCHAR));
        if (args[0].val.str)
        {
            args[0].val.str[0] = str1->val.str[0]; // Copy length prefix
            wcsncpy_s(&args[0].val.str[1], len1 + 1, &str1->val.str[1], len1);
            args[0].val.str[len1 + 1] = L'\0';
        }
    }
    else
    {
        // Create empty string for invalid input
        args[0].val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, 2 * sizeof(XCHAR));
        if (args[0].val.str)
        {
            args[0].val.str[0] = 0;
            args[0].val.str[1] = L'\0';
        }
    }

    // Initialize the second argument: str2
    args[1].xltype = xltypeStr;
    if ((str2->xltype & xltypeStr) == xltypeStr)
    {
        int len2 = str2->val.str[0];
        args[1].val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, (len2 + 2) * sizeof(XCHAR));
        if (args[1].val.str)
        {
            args[1].val.str[0] = str2->val.str[0]; // Copy length prefix
            wcsncpy_s(&args[1].val.str[1], len2 + 1, &str2->val.str[1], len2);
            args[1].val.str[len2 + 1] = L'\0';
        }
    }
    else
    {
        // Create empty string for invalid input
        args[1].val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, 2 * sizeof(XCHAR));
        if (args[1].val.str)
        {
            args[1].val.str[0] = 0;
            args[1].val.str[1] = L'\0';
        }
    }

    // Prepare the result XLOPER12
    XLOPER12 result;
    ZeroMemory(&result, sizeof(XLOPER12));

    // Call Excel12 directly using the registration ID
    int rc = Excel12(xlUDF, &result, 3, (LPXLOPER12)&g_reg_cStringsInner, &args[0], &args[1]);

    // Note: args and arg strings are intentionally not freed (memory leak by design for test)

    // Check the result and return the value if successful
    if (rc == xlretSuccess && (result.xltype & xltypeStr) == xltypeStr)
    {
        // Allocate a new XLOPER12 for the return value (intentionally not freed)
        LPXLOPER12 returnValue = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
        if (returnValue)
        {
            int resultLen = result.val.str[0];
            returnValue->xltype = xltypeStr;
            returnValue->val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, (resultLen + 2) * sizeof(XCHAR));
            if (returnValue->val.str)
            {
                returnValue->val.str[0] = result.val.str[0]; // Copy length prefix
                wcsncpy_s(&returnValue->val.str[1], resultLen + 1, &result.val.str[1], resultLen);
                returnValue->val.str[resultLen + 1] = L'\0';
            }
        }

        // Free the result if necessary
        if (result.xltype & xlbitXLFree)
            Excel12(xlFree, 0, 1, &result);

        return returnValue;
    }

    // Return empty string on failure
    LPXLOPER12 emptyResult = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (emptyResult)
    {
        emptyResult->xltype = xltypeStr;
        emptyResult->val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, 2 * sizeof(XCHAR));
        if (emptyResult->val.str)
        {
            emptyResult->val.str[0] = 0;
            emptyResult->val.str[1] = L'\0';
        }
    }
    return emptyResult;
}

// cStringsFreeDirectById: calls cStringsFreeInner using its registration ID directly and manages memory
__declspec(dllexport) LPXLOPER12 WINAPI cStringsFreeDirectById(LPXLOPER12 str1, LPXLOPER12 str2)
{
    DWORD tid = GetCurrentThreadId();
    DebugPrintW(L"[MultithreadCrash] Thread %lu: cStringsFreeDirectById called\n", tid);

    // Allocate an array of XLOPER12s on the heap for two string args
    LPXLOPER12 args = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, 2 * sizeof(XLOPER12));
    if (!args)
        return NULL; // Allocation failed

    // Helper lambda-like macros for cleanup
#define FREE_ARG_STR(i) do { if (args[i].xltype == xltypeStr && args[i].val.str) { GlobalFree(args[i].val.str); args[i].val.str = NULL; } } while(0)

    // Initialize arg 0
    args[0].xltype = xltypeStr;
    if (str1 && ((str1->xltype & xltypeStr) == xltypeStr))
    {
        int len1 = str1->val.str[0];
        args[0].val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, (len1 + 2) * sizeof(XCHAR));
        if (args[0].val.str)
        {
            args[0].val.str[0] = str1->val.str[0];
            wcsncpy_s(&args[0].val.str[1], len1 + 1, &str1->val.str[1], len1);
            args[0].val.str[len1 + 1] = L'\0';
        }
    }
    else
    {
        args[0].val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, 2 * sizeof(XCHAR));
        if (args[0].val.str)
        {
            args[0].val.str[0] = 0;
            args[0].val.str[1] = L'\0';
        }
    }

    // Initialize arg 1
    args[1].xltype = xltypeStr;
    if (str2 && ((str2->xltype & xltypeStr) == xltypeStr))
    {
        int len2 = str2->val.str[0];
        args[1].val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, (len2 + 2) * sizeof(XCHAR));
        if (args[1].val.str)
        {
            args[1].val.str[0] = str2->val.str[0];
            wcsncpy_s(&args[1].val.str[1], len2 + 1, &str2->val.str[1], len2);
            args[1].val.str[len2 + 1] = L'\0';
        }
    }
    else
    {
        args[1].val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, 2 * sizeof(XCHAR));
        if (args[1].val.str)
        {
            args[1].val.str[0] = 0;
            args[1].val.str[1] = L'\0';
        }
    }

    // Prepare result holder
    XLOPER12 result;
    ZeroMemory(&result, sizeof(XLOPER12));

    // Call Excel12 using the registration ID of cStringsFreeInner
    int rc = Excel12Direct(xlUDF, &result, 3, (LPXLOPER12)&g_reg_cStringsFreeInner, &args[0], &args[1]);

    // Free argument strings and array now that the call is done
    FREE_ARG_STR(0);
    FREE_ARG_STR(1);
    GlobalFree(args);

    // Success path: copy the result to a new return pointer managed via xlAutoFree12
    if (rc == xlretSuccess && (result.xltype & xltypeStr) == xltypeStr)
    {
        LPXLOPER12 retp = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
        if (retp)
        {
            int rlen = result.val.str ? result.val.str[0] : 0;
            retp->xltype = xltypeStr | xlbitDLLFree;
            retp->val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, (rlen + 2) * sizeof(XCHAR));
            if (retp->val.str)
            {
                retp->val.str[0] = (XCHAR)rlen;
                if (rlen > 0 && result.val.str)
                    wcsncpy_s(&retp->val.str[1], rlen + 1, &result.val.str[1], rlen);
                retp->val.str[rlen + 1] = L'\0';
            }
        }

        // Free any Excel-allocated memory in result
        if (result.xltype & xlbitXLFree)
            Excel12(xlFree, 0, 1, &result);

        return retp;
    }

    // Failure: free Excel-allocated memory if present and return empty string (managed)
    if (result.xltype & xlbitXLFree)
        Excel12(xlFree, 0, 1, &result);

    LPXLOPER12 emptyRet = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    if (emptyRet)
    {
        emptyRet->xltype = xltypeStr | xlbitDLLFree;
        emptyRet->val.str = (XCHAR*)GlobalAlloc(GMEM_FIXED, 2 * sizeof(XCHAR));
        if (emptyRet->val.str)
        {
            emptyRet->val.str[0] = 0;
            emptyRet->val.str[1] = L'\0';
        }
    }
    return emptyRet;
}

// Test functions using Excel12Direct (bypassing framework)

// cDoubleCallerExcel12Direct: calls cDoubleInner by name using Excel12Direct
__declspec(dllexport) double WINAPI cDoubleCallerExcel12Direct(double x, double y)
{
    static __declspec(thread) int tls_init = 0;
    static __declspec(thread) LPXLOPER12 tls_x = NULL;
    static __declspec(thread) LPXLOPER12 tls_y = NULL;

    DWORD tid = GetCurrentThreadId();
    DebugPrintW(L"[MultithreadCrash] Thread %lu: cDoubleCallerExcel12Direct called\n", tid);

    if (!tls_init)
    {
        DebugPrintW(L"[MultithreadCrash] Thread %lu: init TLS numeric args\n", tid);
        tls_x = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
        tls_y = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
        if (tls_x) { tls_x->xltype = xltypeNum; tls_x->val.num = 0.0; }
        if (tls_y) { tls_y->xltype = xltypeNum; tls_y->val.num = 0.0; }
        tls_init = 1;
    }

    if (tls_x) tls_x->val.num = x;
    if (tls_y) tls_y->val.num = y;

    // Allocate a fresh function name XLOPER12 on the heap
    LPXLOPER12 fnArg = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
    LPWSTR fnStr = (LPWSTR)GlobalAlloc(GMEM_FIXED, 32 * sizeof(wchar_t));
    if (fnArg && fnStr)
    {
        const wchar_t* name = L"cDoubleInner";
        size_t nlen = wcslen(name); if (nlen > 30) nlen = 30;
        fnStr[0] = (wchar_t)nlen;
        wcsncpy_s(&fnStr[1], 31, name, nlen);
        fnArg->xltype = xltypeStr;
        fnArg->val.str = fnStr;
    }

    XLOPER12 ret;
    int rc = Excel12Direct(xlUDF, &ret, 3, fnArg, tls_x, tls_y);
    
    DebugPrintW(L"[MultithreadCrash] Thread %lu: Excel12Direct returned %d\n", tid, rc);
    
    if (rc == xlretSuccess && (ret.xltype & xltypeNum) == xltypeNum)
        return ret.val.num;
    return 0.0;
}

// cDoubleCallerExcel12DirectById: calls cDoubleInner by ID using Excel12Direct
__declspec(dllexport) double WINAPI cDoubleCallerExcel12DirectById(double x, double y)
{
    static __declspec(thread) int tls_init = 0;
    static __declspec(thread) LPXLOPER12 tls_x = NULL;
    static __declspec(thread) LPXLOPER12 tls_y = NULL;

    DWORD tid = GetCurrentThreadId();
    DebugPrintW(L"[MultithreadCrash] Thread %lu: cDoubleCallerExcel12DirectById called\n", tid);

    if (!tls_init)
    {
        DebugPrintW(L"[MultithreadCrash] Thread %lu: init TLS numeric args\n", tid);
        tls_x = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
        tls_y = (LPXLOPER12)GlobalAlloc(GMEM_FIXED, sizeof(XLOPER12));
        if (tls_x) { tls_x->xltype = xltypeNum; tls_x->val.num = 0.0; }
        if (tls_y) { tls_y->xltype = xltypeNum; tls_y->val.num = 0.0; }
        tls_init = 1;
    }

    if (tls_x) tls_x->val.num = x;
    if (tls_y) tls_y->val.num = y;

    if ((g_reg_cDoubleInner.xltype & xltypeNum) != xltypeNum)
    {
        DebugPrintW(L"[MultithreadCrash] Thread %lu: Registration ID not available\n", tid);
        return 0.0; // ID not available
    }

    XLOPER12 ret;
    int rc = Excel12Direct(xlUDF, &ret, 3, (LPXLOPER12)&g_reg_cDoubleInner, tls_x, tls_y);
    
    DebugPrintW(L"[MultithreadCrash] Thread %lu: Excel12Direct returned %d\n", tid, rc);
    
    if (rc == xlretSuccess && (ret.xltype & xltypeNum) == xltypeNum)
        return ret.val.num;
    return 0.0;
}

// Registration
__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
    static XLOPER12 xDLL;
    Excel12f(xlGetName, &xDLL, 0);

    // Initialize direct MdCallBack12 access
    if (InitMdCallBack12())
    {
        DebugPrintW(L"[MultithreadCrash] MdCallBack12 initialized successfully in xlAutoOpen\n");
    }
    else
    {
        DebugPrintW(L"[MultithreadCrash] WARNING: Failed to initialize MdCallBack12 in xlAutoOpen\n");
    }

    for (int i = 0; i < rgFuncsRows; i++)
    {
        XLOPER12 regId;
        Excel12f(xlfRegister, &regId, 1 + 7,
            (LPXLOPER12)&xDLL,
            TempStr12(rgFuncs[i][0]),
            TempStr12(rgFuncs[i][1]),
            TempStr12(rgFuncs[i][2]),
            TempStr12(rgFuncs[i][3]),
            TempStr12(rgFuncs[i][4]),
            TempStr12(rgFuncs[i][5]),
            TempStr12(L""),
            TempStr12(L""),
            TempStr12(rgFuncs[i][6]),
            TempStr12(rgFuncs[i][3])
        );

        if (wcscmp(rgFuncs[i][0], L"cDoubleInner") == 0 && (regId.xltype & xltypeNum) == xltypeNum)
        {
            g_reg_cDoubleInner = regId;
            XLOPER12 evalId;
            int evrc = Excel12f(xlfEvaluate, &evalId, 1, TempStr12(L"cDoubleInner"));
            if (evrc == xlretSuccess && (evalId.xltype & xltypeNum) == xltypeNum)
                DebugPrintW(L"[MultithreadCrash] REGISTER vs EVALUATE id: %.0f vs %.0f\n", g_reg_cDoubleInner.val.num, evalId.val.num);
        }
        
        if (wcscmp(rgFuncs[i][0], L"cStringsInner") == 0 && (regId.xltype & xltypeNum) == xltypeNum)
        {
            g_reg_cStringsInner = regId;
            XLOPER12 evalId;
            int evrc = Excel12f(xlfEvaluate, &evalId, 1, TempStr12(L"cStringsInner"));
            if (evrc == xlretSuccess && (evalId.xltype & xltypeNum) == xltypeNum)
                DebugPrintW(L"[MultithreadCrash] cStringsInner REGISTER vs EVALUATE id: %.0f vs %.0f\n", g_reg_cStringsInner.val.num, evalId.val.num);
        }

        if (wcscmp(rgFuncs[i][0], L"cStringsFreeInner") == 0 && (regId.xltype & xltypeNum) == xltypeNum)
        {
            g_reg_cStringsFreeInner = regId;
            XLOPER12 evalId;
            int evrc = Excel12f(xlfEvaluate, &evalId, 1, TempStr12(L"cStringsFreeInner"));
            if (evrc == xlretSuccess && (evalId.xltype & xltypeNum) == xltypeNum)
                DebugPrintW(L"[MultithreadCrash] cStringsFreeInner REGISTER vs EVALUATE id: %.0f vs %.0f\n", g_reg_cStringsFreeInner.val.num, evalId.val.num);
        }
    }

    Excel12f(xlFree, 0, 1, (LPXLOPER12)&xDLL);
    return 1;
}

__declspec(dllexport) int WINAPI xlAutoClose(void)
{
    for (int i = 0; i < rgFuncsRows; i++)
        Excel12f(xlfSetName, 0, 1, TempStr12(rgFuncs[i][2]));
    return 1;
}

// Excel calls this to free results returned with xlbitDLLFree set
__declspec(dllexport) void WINAPI xlAutoFree12(LPXLOPER12 p)
{
    if (!p) return;
    if ((p->xltype & xltypeStr) == xltypeStr)
    {
        if (p->val.str)
            GlobalFree(p->val.str);
    }
    GlobalFree(p);
}
