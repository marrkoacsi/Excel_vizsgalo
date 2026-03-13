namespace Excel
{
    public static class DebugLogger
    {
        public static DebugConsole Console;

        public static void Log(string message)
        {
            Console?.WriteLine("[INFO] " + message);
        }

        public static void Error(string message)
        {
            Console?.WriteLine("[ERROR] " + message);
        }
    }
}