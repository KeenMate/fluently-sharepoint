using System.Collections.Generic;

namespace KeenMate.FluentlySharePoint
{
	public static class DictionaryExtensions
	{
		public static void AddOrUpdate<Tkey, Tvalue>(this Dictionary<Tkey, Tvalue> dictionary, Tkey key, Tvalue value)
		{
			if (dictionary.ContainsKey(key))
			{
				dictionary[key] = value;
			}
			else
			{
				dictionary.Add(key, value);
			}
		}
	}
}