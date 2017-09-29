using System;

namespace FluentlySharepoint.Interfaces
{
	public interface ILogger
	{
        Guid CorrelationId { get; set; }

		void Trace(string message);
		void Debug(string message);
		void Info(string message);
		void Warn(string message);
		void Warn(Exception ex, string message);
		void Error(string message);
		void Error(Exception ex, string message);
		void Fatal(string message);
		void Fatal(Exception ex, string message);
	}
}