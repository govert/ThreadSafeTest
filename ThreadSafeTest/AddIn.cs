using ExcelDna.Integration;
using System.IO;
using System.Threading;

namespace ThreadSafeTest
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            // Get the directory where this add-in is located
            string addInPath = ExcelDnaUtil.XllPath;
            string addInDirectory = Path.GetDirectoryName(addInPath);
            
            // List of add-ins to load
            string[] addInsToLoad = { "ThreadSafeC.xll", "ThreadSafeNet.xll" };
            
            foreach (string addInFileName in addInsToLoad)
            {
                string fullPath = Path.Combine(addInDirectory, addInFileName);
                
                if (File.Exists(fullPath))
                {
                    try
                    {
                        // Register the add-in using XlCall.Excel
                        var result = XlCall.Excel(XlCall.xlfRegister, fullPath);
                        System.Diagnostics.Debug.WriteLine($"Loaded add-in: {addInFileName}, Result: {result}");
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to load add-in {addInFileName}: {ex.Message}");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"Add-in not found: {fullPath}");
                }
            }
        }

        public void AutoClose()
        {
            // Optional: Unregister add-ins on close
            // Get the directory where this add-in is located
            string addInPath = ExcelDnaUtil.XllPath;
            string addInDirectory = Path.GetDirectoryName(addInPath);
            
            string[] addInsToUnload = { "ThreadSafeC.xll", "ThreadSafeNet.xll" };
            
            foreach (string addInFileName in addInsToUnload)
            {
                string fullPath = Path.Combine(addInDirectory, addInFileName);
                
                if (File.Exists(fullPath))
                {
                    try
                    {
                        // Unregister the add-in using XlCall.Excel
                        var result = XlCall.Excel(XlCall.xlfUnregister, fullPath);
                        System.Diagnostics.Debug.WriteLine($"Unloaded add-in: {addInFileName}, Result: {result}");
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to unload add-in {addInFileName}: {ex.Message}");
                    }
                }
            }
        }
    }

    public static class Functions
    {
        [ExcelFunction(Description = "Inner thread info for nested call test", IsThreadSafe = true)]
        public static string tsInnerThreadInfo()
        {
            var innerThreadId = Thread.CurrentThread.ManagedThreadId;
            return $"InnerThread:{innerThreadId}";
        }

        [ExcelFunction(Description = "Calls an inner function via XlCall. Set callExternal=true to call into ThreadSafeNet.", IsThreadSafe = true)]
        public static object tsNestedThreadInfo(bool callExternal = false)
        {
            var outerThreadId = Thread.CurrentThread.ManagedThreadId;
            var target = callExternal ? "csInnerThreadInfo" : "tsInnerThreadInfo";
            object innerResult;
            try
            {
                innerResult = XlCall.Excel(XlCall.xlUDF, target);
            }
            catch (System.Exception ex)
            {
                innerResult = $"Error:{ex.Message}";
            }

            return $"OuterThread:{outerThreadId}; {innerResult}";
        }

        [ExcelFunction(Description = "Returns a greeting message with thread info", IsThreadSafe = true)]
        public static string ThreadSafeTestFunction(string name)
        {
            var threadId = System.Threading.Thread.CurrentThread.ManagedThreadId;
            return $"Hello {name} from ThreadSafeTest! Thread ID: {threadId}";
        }

        [ExcelFunction(Description = "Test calling ThreadSafeCFunction - use useCSharp=true for C# version", IsThreadSafe = true)]
        public static object TestThreadSafeCFunction(double input, bool useCSharp = false)
        {
            try
            {
                if (useCSharp)
                {
                    return XlCall.Excel(XlCall.xlUDF, "csThreadSafeCFunction", input);
                }
                else
                {
                    return XlCall.Excel(XlCall.xlUDF, "ThreadSafeCFunction", input);
                }
            }
            catch (System.Exception ex)
            {
                return $"Error calling ThreadSafeCFunction: {ex.Message}";
            }
        }

        [ExcelFunction(Description = "Test calling ThreadSafeCalc - use useCSharp=true for C# version", IsThreadSafe = true)]
        public static object TestThreadSafeCalc(double number, bool useCSharp = false)
        {
            try
            {
                if (useCSharp)
                {
                    return XlCall.Excel(XlCall.xlUDF, "csThreadSafeCalc", number);
                }
                else
                {
                    return XlCall.Excel(XlCall.xlUDF, "ThreadSafeCalc", number);
                }
            }
            catch (System.Exception ex)
            {
                return $"Error calling ThreadSafeCalc: {ex.Message}";
            }
        }

        [ExcelFunction(Description = "Test calling ThreadSafeXLOPER - use useCSharp=true for C# version", IsThreadSafe = true)]
        public static object TestThreadSafeXLOPER(double input, bool useCSharp = false)
        {
            try
            {
                if (useCSharp)
                {
                    return XlCall.Excel(XlCall.xlUDF, "csThreadSafeXLOPER", input);
                }
                else
                {
                    return XlCall.Excel(XlCall.xlUDF, "ThreadSafeXLOPER", input);
                }
            }
            catch (System.Exception ex)
            {
                return $"Error calling ThreadSafeXLOPER: {ex.Message}";
            }
        }

        [ExcelFunction(Description = "Test calling AllocatedMemoryFunction - use useCSharp=true for C# version", IsThreadSafe = true)]
        public static object TestAllocatedMemoryFunction(double size, bool useCSharp = false)
        {
            try
            {
                if (useCSharp)
                {
                    return XlCall.Excel(XlCall.xlUDF, "csAllocatedMemoryFunction", size);
                }
                else
                {
                    return XlCall.Excel(XlCall.xlUDF, "AllocatedMemoryFunction", size);
                }
            }
            catch (System.Exception ex)
            {
                return $"Error calling AllocatedMemoryFunction: {ex.Message}";
            }
        }

        [ExcelFunction(Description = "Test calling ThreadInfoFunction - use useCSharp=true for C# version", IsThreadSafe = true)]
        public static object TestThreadInfoFunction(bool useCSharp = false)
        {
            try
            {
                if (useCSharp)
                {
                    return XlCall.Excel(XlCall.xlUDF, "csThreadInfoFunction");
                }
                else
                {
                    return XlCall.Excel(XlCall.xlUDF, "ThreadInfoFunction");
                }
            }
            catch (System.Exception ex)
            {
                return $"Error calling ThreadInfoFunction: {ex.Message}";
            }
        }

        [ExcelFunction(Description = "Test multiple calls to thread-safe functions - use useCSharp=true for C# versions", IsThreadSafe = true)]
        public static object TestMultipleThreadSafeCalls(double input, bool useCSharp = false)
        {
            try
            {
                var managedThreadId = System.Threading.Thread.CurrentThread.ManagedThreadId;
                object result1, result2, threadInfo;
                
                if (useCSharp)
                {
                    result1 = XlCall.Excel(XlCall.xlUDF, "csThreadSafeCalc", input);
                    result2 = XlCall.Excel(XlCall.xlUDF, "csThreadSafeXLOPER", input);
                    threadInfo = XlCall.Excel(XlCall.xlUDF, "csThreadInfoFunction");
                }
                else
                {
                    result1 = XlCall.Excel(XlCall.xlUDF, "ThreadSafeCalc", input);
                    result2 = XlCall.Excel(XlCall.xlUDF, "ThreadSafeXLOPER", input);
                    threadInfo = XlCall.Excel(XlCall.xlUDF, "ThreadInfoFunction");
                }
                
                return $"Managed Thread: {managedThreadId}, Calc: {result1}, XLOPER: {result2}, Info: {threadInfo}";
            }
            catch (System.Exception ex)
            {
                return $"Error in multiple calls: {ex.Message}";
            }
        }

        [ExcelFunction(Description = "Performance test - use useCSharp=true for C# version", IsThreadSafe = true)]
        public static object TestPerformance(double input, int iterations, bool useCSharp = false)
        {
            try
            {
                if (iterations <= 0) iterations = 1;
                if (iterations > 1000) iterations = 1000; // Limit to prevent Excel freezing
                
                var startTime = System.DateTime.Now;
                object lastResult = 0.0;
                string functionName = useCSharp ? "csThreadSafeCalc" : "ThreadSafeCalc";
                
                for (int i = 0; i < iterations; i++)
                {
                    lastResult = XlCall.Excel(XlCall.xlUDF, functionName, input + i);
                }
                
                var endTime = System.DateTime.Now;
                var elapsed = endTime - startTime;
                string version = useCSharp ? "C#" : "C";
                
                return $"{version} - Iterations: {iterations}, Last Result: {lastResult}, Time: {elapsed.TotalMilliseconds:F2}ms";
            }
            catch (System.Exception ex)
            {
                return $"Error in performance test: {ex.Message}";
            }
        }

        [ExcelFunction(Description = "Compare performance between C and C# versions", IsThreadSafe = true)]
        public static object ComparePerformance(double input, int iterations)
        {
            try
            {
                if (iterations <= 0) iterations = 1;
                if (iterations > 500) iterations = 500; // Lower limit for comparison
                
                // Test C version
                var startTimeC = System.DateTime.Now;
                object lastResultC = 0.0;
                for (int i = 0; i < iterations; i++)
                {
                    lastResultC = XlCall.Excel(XlCall.xlUDF, "ThreadSafeCalc", input + i);
                }
                var endTimeC = System.DateTime.Now;
                var elapsedC = endTimeC - startTimeC;
                
                // Test C# version
                var startTimeCS = System.DateTime.Now;
                object lastResultCS = 0.0;
                for (int i = 0; i < iterations; i++)
                {
                    lastResultCS = XlCall.Excel(XlCall.xlUDF, "csThreadSafeCalc", input + i);
                }
                var endTimeCS = System.DateTime.Now;
                var elapsedCS = endTimeCS - startTimeCS;
                
                return $"C: {elapsedC.TotalMilliseconds:F2}ms ({lastResultC}), C#: {elapsedCS.TotalMilliseconds:F2}ms ({lastResultCS})";
            }
            catch (System.Exception ex)
            {
                return $"Error in comparison: {ex.Message}";
            }
        }
    }
}
