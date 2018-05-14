using System;
using KeenMate.FluentlySharePoint.Interfaces;

namespace KeenMate.FluentlySharePoint.Loggers
{
    public class BlackHoleLogger : ILogger
    {
        public Guid CorrelationId { get; set; }

        public void Trace(string message)
        {
            
        }

        public void Debug(string message)
        {
            
        }

        public void Info(string message)
        {
            
        }

        public void Warn(string message)
        {
            
        }

        public void Warn(Exception ex, string message)
        {
            
        }

        public void Error(string message)
        {
            
        }

        public void Error(Exception ex, string message)
        {
           
        }

        public void Fatal(string message)
        {
            
        }

        public void Fatal(Exception ex, string message)
        {
           
        }
    }
}