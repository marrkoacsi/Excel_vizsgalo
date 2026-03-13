using System.Windows;

namespace Excel
{
    public partial class DebugConsole : Window
    {
        public DebugConsole()
        {
            InitializeComponent();
        }

        public void WriteLine(string text)
        {
            ConsoleText.AppendText(text + "\n");
            ConsoleText.ScrollToEnd();
        }
    }
}