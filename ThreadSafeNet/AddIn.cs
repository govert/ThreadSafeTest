using ExcelDna.Integration;
using System;
using System.Threading;

namespace ThreadSafeNet
{
    public class AddIn : IExcelAddIn
    {
        static Dictionary<string, double> _registerIds = new Dictionary<string, double>();

        // Doubles as parameters (no XLOPERs)
        [ExcelFunction(Description = "Inner: add two doubles (no XLOPER)", IsThreadSafe = true)]
        public static double csDoubleInner(double x, double y)
        {
            return x + y;
        }

        [ExcelFunction(Description = "Caller: calls csDoubleInner via XlCall (no XLOPER)", IsThreadSafe = true)]
        public static double csDoubleCaller(double x, double y)
        {
            var res = XlCall.Excel(XlCall.xlUDF, _registerIds[nameof(csDoubleInner)], x, y);
            try { return Convert.ToDouble(res); } catch { return double.NaN; }
        }

        [ExcelFunction(Description = "Inner thread info for cross add-in nested call test", IsThreadSafe = true)]
        public static string csInnerThreadInfo()
        {
            var innerThreadId = Thread.CurrentThread.ManagedThreadId;
            return $"InnerThread:{innerThreadId}";
        }

        [ExcelFunction(Description = "C# version of ThreadSafeCFunction - calculates sqrt(input*3) + thread ID", IsThreadSafe = true)]
        public static double csThreadSafeCFunction(double input)
        {
            var threadId = Thread.CurrentThread.ManagedThreadId;
            Thread.Sleep(10); // Small delay to make threading effects more visible
            return Math.Sqrt(input * 3.0) + threadId;
        }

        [ExcelFunction(Description = "C# version of ThreadSafeCalc - thread-safe calculation", IsThreadSafe = true)]
        public static double csThreadSafeCalc(double number)
        {
            var threadId = Thread.CurrentThread.ManagedThreadId;
            Thread.Sleep(10); // Small delay to make threading effects more visible
            return number * number + Math.Sin(number) + threadId;
        }

        [ExcelFunction(Description = "C# version of ThreadSafeXLOPER - thread-safe calculation returning double", IsThreadSafe = true)]
        public static double csThreadSafeXLOPER(double input)
        {
            var threadId = Thread.CurrentThread.ManagedThreadId;
            return input * 2.0 + threadId;
        }

        [ExcelFunction(Description = "C# version of AllocatedMemoryFunction - returns array of numbers", IsThreadSafe = true)]
        public static object csAllocatedMemoryFunction(double sizeInput)
        {
            var threadId = Thread.CurrentThread.ManagedThreadId;
            int size = (int)sizeInput;
            if (size < 1) size = 1;
            if (size > 100) size = 100; // Limit size

            double[] result = new double[size];
            for (int i = 0; i < size; i++)
            {
                result[i] = threadId + i;
            }

            // Return as vertical array for Excel
            object[,] excelArray = new object[size, 1];
            for (int i = 0; i < size; i++)
            {
                excelArray[i, 0] = result[i];
            }
            
            return excelArray;
        }

        [ExcelFunction(Description = "C# version of ThreadInfoFunction - returns thread information", IsThreadSafe = true)]
        public static string csThreadInfoFunction()
        {
            var threadId = Thread.CurrentThread.ManagedThreadId;
            var tickCount = Environment.TickCount;
            return $"C# Thread: {threadId}, Time: {tickCount}";
        }

        [ExcelFunction(Description = "Advanced C# test - multiple calculations with different data types", IsThreadSafe = true)]
        public static object csAdvancedTest(double input, int iterations)
        {
            var threadId = Thread.CurrentThread.ManagedThreadId;
            if (iterations <= 0) iterations = 1;
            if (iterations > 100) iterations = 100; // Reasonable limit

            double sum = 0;
            for (int i = 0; i < iterations; i++)
            {
                sum += Math.Pow(input + i, 1.5);
            }

            return $"C# Advanced: Thread {threadId}, Sum: {sum:F2}, Iterations: {iterations}";
        }

        public void AutoOpen()
        {
            var funcName = nameof(csDoubleInner);
            var regId = (double)XlCall.Excel(XlCall.xlfEvaluate, funcName);
            _registerIds[funcName] = regId;
        }

        public void AutoClose()
        {
        }
    }
}
