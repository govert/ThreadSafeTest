/*
**  ThreadSafeC Example
**  
**  Demonstrates thread-safe Excel function implementation using Excel SDK
*/

#include <windows.h>
#include <math.h>
#include <xlcall.h>
#include <framewrk.h>

/*
** rgFuncs
**
** This is a table of all functions exported by this module.
** These functions are registered in xlAutoOpen when the XLL loads.
** Format matches the last 7 arguments to REGISTER function.
*/
#define rgFuncsRows 1

static const LPWSTR rgFuncs[rgFuncsRows][7] = {
    {(LPWSTR)L"ThreadSafeCFunction", (LPWSTR)L"QQ", (LPWSTR)L"ThreadSafeCFunction", (LPWSTR)L"input", (LPWSTR)L"1", (LPWSTR)L"Thread Safe Demo", (LPWSTR)L"Calculates sqrt(input*3) + thread ID"}
};

/*
** ThreadSafeCFunction
**
** Calculates sqrt(input*3) + thread ID to demonstrate thread safety
** Takes XLOPER12 input and returns XLOPER12 result
*/
__declspec(dllexport) LPXLOPER12 WINAPI ThreadSafeCFunction(LPXLOPER12 input)
{
    double inputValue = 0.0;
    DWORD threadId = GetCurrentThreadId();
    
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
    double result = sqrt(inputValue * 3.0) + (double)threadId;
    
    // Return result as XLOPER12 using framework function
    return TempNum12(result);
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
** Since we use TempNum12() which manages its own memory through the framework,
** this function doesn't need to do anything explicit.
*/
__declspec(dllexport) void WINAPI xlAutoFree12(LPXLOPER12 pxFree)
{
    // Framework handles memory cleanup automatically for Temp functions
    return;
}