using ExcelDna.Integration;

namespace ThreadSafe1
{
    public static class AddIn
    {
        [ExcelFunction(Description = "Returns a greeting message with thread info")]
        public static string ThreadSafe1Function(string name)
        {
            var threadId = System.Threading.Thread.CurrentThread.ManagedThreadId;
            return $"Hello {name} from ThreadSafe1! Thread ID: {threadId}";
        }
    }
}