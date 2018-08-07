using System;

namespace KeenMate.FluentlySharePoint
{
	public class NeedFixException: Exception
	{
		public NeedFixException(string message) : base(message)
		{
		}
	}
}