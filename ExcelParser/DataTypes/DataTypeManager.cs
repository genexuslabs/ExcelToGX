using Artech.Common.Helpers.Strings;
using Artech.Genexus.Common.Types;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser.DataTypes
{
	public class DataTypeManager
	{
		public static void SetDataType(string text, TransactionAttribute att)
		{
			string normalizedText = text.ToUpper();
			string [] vals = normalizedText.Split('(');
			BasicDataTypeInfo typeInfo = null;
			foreach (BasicDataTypeInfo dtInfo in BasicDataTypeInfo.Types().Values)
			{
				if (dtInfo.Name.ToUpper().StartsWith(vals[0]))
				{
					typeInfo = dtInfo;
					break;
				}
			}
			if (typeInfo != null)
			{
				att.SetDataType(typeInfo);
				if (typeInfo.RegularExpression != null)
				{
					string fullText = normalizedText.Replace(vals[0], typeInfo.Name.ToUpper());
					var match = typeInfo.RegularExpression.Match(fullText);
					if (match.Success)
					{
						if (typeInfo.AllowLength)
						{
							if (match.Groups["length"].Length > 0)
							{
								if (Format.ParseInteger(match.Groups["length"].Value, out int length))
									att.Length = length;
							}
							if (typeInfo.AllowDecimals)
							{
								int decimals = 0;
								// If decimals are not set but the expression is "complete" (e.g. "Numeric(4)"),
								// use 0 instead of the default value).
								if (match.Groups["decimals"].Length > 0)
									Int32.TryParse(match.Groups["decimals"].Value, out decimals);
								else if (text.EndsWith(")"))
									decimals = 0;

								att.Decimals = decimals;
							}

							if (typeInfo.AllowSign && match.Groups["sign"].Length > 0)
								att.Sign = true;
						}
					}
				}
				return;
			}
			throw new Exception($"Could not parse {text} as any known datatype");
		}
	}
}
