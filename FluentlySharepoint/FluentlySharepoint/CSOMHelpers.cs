using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace KeenMate.FluentlySharePoint
{
	public static class CSOMHelpers
	{
		public static SecureString ToSecureString(this string text)
		{
			SecureString s = new SecureString();
			foreach (var c in text)
			{
				s.AppendChar(c);
			}

			return s;
		}
	}
}
