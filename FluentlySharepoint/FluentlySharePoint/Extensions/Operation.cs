using System;

namespace KeenMate.FluentlySharePoint.Extensions
{
	public static class Operation
	{
		public static CSOMOperation SetCorrelationId(this CSOMOperation operation, Guid correlationId)
		{
			operation.CorrelationId = correlationId;
			return operation;
		}

		public static CSOMOperation NewCorrelationId(this CSOMOperation operation)
		{
			operation.CorrelationId = new Guid();
			return operation;
		}

		public static CSOMOperation ClearCorrelationId(this CSOMOperation operation)
		{
			operation.CorrelationId = Guid.Empty;
			return operation;
		}
	}
}