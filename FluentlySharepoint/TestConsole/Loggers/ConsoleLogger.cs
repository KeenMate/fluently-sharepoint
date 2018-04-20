using System;
using KeenMate.FluentlySharePoint.Interfaces;

namespace TestConsole.Loggers
{
    /// <summary>
    /// Dummy console logger
    /// </summary>
    public class ConsoleLogger : ILogger
    {
        public void WriteLineInColor(string message, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(message);
            Console.ResetColor();
        }

        public void Trace(string message)
        {
            WriteLineInColor(message, ConsoleColor.Gray);
        }

        public void Debug(string message)
        {
            WriteLineInColor(message, ConsoleColor.DarkGray);
        }

        public void Info(string message)
        {
            WriteLineInColor(message, ConsoleColor.Cyan);
        }

        public void Warn(string message)
        {
            WriteLineInColor(message, ConsoleColor.Yellow);
        }

        public void Warn(Exception ex, string message)
        {
            WriteLineInColor($"{message} - {ex}", ConsoleColor.Yellow);
        }

        public void Error(string message)
        {
            WriteLineInColor($"{message}", ConsoleColor.Red);
        }

        public void Error(Exception ex, string message)
        {
            WriteLineInColor($"{message} - {ex}", ConsoleColor.Red);
        }

        public void Fatal(string message)
        {
            WriteLineInColor($"{message}", ConsoleColor.DarkRed);
        }

        public void Fatal(Exception ex, string message)
        {
            WriteLineInColor($"{message} - {ex}", ConsoleColor.DarkRed);
        }
    }
}