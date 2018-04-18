namespace KeenMate.FluentlySharePoint
{
	class LevelLock
	{
		public bool Site { get; set; }
		public bool Web { get; set; }
		public bool List { get; set; }
		public bool ContentType { get; set; }

		public void SetLocks(OperationLevels[] levels, bool value)
		{
			foreach (var level in levels)
			{
				switch (level)
				{
					case OperationLevels.Web:
						Web = value;
						break;
					case OperationLevels.Site:
						Site = value;
						break;
					case OperationLevels.List:
						List = value;
						break;
					case OperationLevels.ContentType:
						ContentType = value;
						break;
				}
			}
		}
	}
}