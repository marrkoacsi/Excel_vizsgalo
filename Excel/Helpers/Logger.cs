using System;

namespace Excel.Helpers
{
    public static class Logger
    {
        public static Action<string> LogAction;

        public static void Log(string message)
        {
            LogAction?.Invoke($"[{DateTime.Now:HH:mm:ss}] {message}");
        }
    }
}