using ExcelDna.Integration;

namespace ThreadSafe2
{
    public static class AddIn
    {
        [ExcelFunction(Description = "Calculates a value with thread safety demonstration")]
        public static double ThreadSafe2Function(double input)
        {
            var threadId = System.Threading.Thread.CurrentThread.ManagedThreadId;
            var result = Math.Sqrt(input * 2) + threadId;
            return result;
        }
    }
}