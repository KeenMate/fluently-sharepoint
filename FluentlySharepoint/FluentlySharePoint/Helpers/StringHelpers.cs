using System.Globalization;
using System.Text;

namespace KeenMate.FluentlySharePoint.Helpers
{
	public static class StringExtensions
	{
		public static string RemoveDiacritics(this string s)
		{
			s = s.Normalize(NormalizationForm.FormD);

			var sb = new StringBuilder();

			foreach (var t in s)
			{
				if (CharUnicodeInfo.GetUnicodeCategory(t) != UnicodeCategory.NonSpacingMark) sb.Append(t);
			}

			return sb.ToString();
		}

		public static bool IsNullOrEmpty(this string s)
		{
			return string.IsNullOrEmpty(s);
		}

		public static bool IsNotNullOrEmpty(this string s)
		{
			return !IsNullOrEmpty(s);
		}
	}
}