using Artech.Common.Helpers.Strings;
using Artech.Genexus.Common.Types;
using System;

namespace ExcelParser.DataTypes
{
	public static class DataTypeManager
	{
		public static void SetDataType(string text, DataTypeElement dte)
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
				dte.SetDataType(typeInfo);
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
									dte.Length = length;
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

								dte.Decimals = decimals;
							}

							if (typeInfo.AllowSign && match.Groups["sign"].Length > 0)
								dte.Sign = true;
						}
					}
				}
				return;
			}
			throw new Exception($"Could not parse {text} as any known datatype");
		}
	}
}
