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
#define rgFuncsRows 5
static const LPWSTR rgFuncs[rgFuncsRows][7] = {
    {(LPWSTR)L"cDoubleInner",  (LPWSTR)L"BBB$", (LPWSTR)L"cDoubleInner",  (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Inner add: returns x+y"},
    {(LPWSTR)L"cDoubleCaller", (LPWSTR)L"BBB$", (LPWSTR)L"cDoubleCaller", (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Caller: calls by name (leaks name XLOPER)"},
    {(LPWSTR)L"cDoubleCallerById", (LPWSTR)L"BBB$", (LPWSTR)L"cDoubleCallerById", (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Caller: calls by register id"},
    {(LPWSTR)L"cDoubleCallerDirect", (LPWSTR)L"BBB$", (LPWSTR)L"cDoubleCallerDirect", (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Caller: calls by name, w/out framework"},
    {(LPWSTR)L"cDoubleCallerDirectById", (LPWSTR)L"BBB$", (LPWSTR)L"cDoubleCallerDirectById", (LPWSTR)L"x,y", (LPWSTR)L"1", (LPWSTR)L"Multithread Crash", (LPWSTR)L"Caller: calls by register id, w/out framework"}
};

// Register id captured for cDoubleInner
static XLOPER12 g_reg_cDoubleInner = { 0 };

// cDoubleInner: returns x+y
__declspec(dllexport) double WINAPI cDoubleInner(double x, double y)
{
	// Sleep this thread for 100 ms to simulate some work using the Windows API
	Sleep(100);
    return x + y;
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

        // Free the result if necessary
        if (result.xltype & xlbitXLFree)
            Excel12(xlFree, 0, 1, &result);

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

// Registration
__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
    static XLOPER12 xDLL;
    Excel12f(xlGetName, &xDLL, 0);

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
