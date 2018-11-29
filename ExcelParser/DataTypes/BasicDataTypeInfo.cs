using Artech.Common.Helpers.Strings;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Artech.Genexus.Common.Types
{
	public class BasicDataTypeInfo
	{
		#region Data

		private readonly int m_Length;
		private readonly int m_Decimals;
		private readonly bool m_Signed;
		public bool AllowLength { get; set; }
		public bool AllowDecimals { get; private set; }
		public bool AllowSign { get; private set; }

		#endregion

		#region Constructors
		public BasicDataTypeInfo(EDBType type, string name, int length, int decimals, bool signed, bool allowLength)
			: this(type, name, length, decimals, signed, allowLength, false, false)
		{

		}

		public BasicDataTypeInfo(EDBType type, string name, int length, int decimals, bool signed, bool allowLength, bool allowDecimals, bool allowSign)
		{
			Type = type;
			Name = name;
			Description = name;
			AllowLength = allowLength;
			AllowDecimals = allowDecimals;
			AllowSign = allowSign;

			m_Length = length;
			m_Decimals = decimals;
			m_Signed = false;
		}

		private static Dictionary<string, Regex> s_Regexps = new Dictionary<string, Regex>();

		public Regex RegularExpression
		{
			get
			{
				// Calculate the regular expression
				string sPattern = Name.ToString().ToUpper();
				string sExtra = "";

				if (AllowLength)
				{
					sExtra += "(?<length>(" + Format.IntegerRegularExpression + "))";
					if (AllowDecimals)
						sExtra += "(.(?<decimals>[0-9]+))?";
				}

				if (AllowSign)
					sExtra += "(?<sign>-?)";

				if (sExtra.Length > 0)
				{
					sPattern += @"(\(" + sExtra + @"\))?";
					sPattern = "^" + sPattern; // We don't use '$' at end to be forgiving, except if strictMatch. But '^' is necessary to distinguish between varchar and longvarchar.
	
					Regex regexp;
					lock (s_Regexps)
					{
						if (!s_Regexps.TryGetValue(sPattern, out regexp))
						{
							regexp = new Regex(sPattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
							s_Regexps[sPattern] = regexp;
						}
					}
					return regexp;
				}
				return null;

			}
		}


		#endregion

		#region Data Managment

		public string Description { get; private set; }

		#endregion

		#region Static

		private static readonly object m_LockObject = new object();

		public static Dictionary<EDBType, BasicDataTypeInfo> Types()
		{
				Dictionary<EDBType, BasicDataTypeInfo> types = new Dictionary<EDBType, BasicDataTypeInfo>();

				types.Add(EDBType.NUMERIC, new BasicDataTypeInfo(EDBType.NUMERIC, "Numeric", 4, 0, false, true, true, true));
				types.Add(EDBType.CHARACTER, new BasicDataTypeInfo(EDBType.CHARACTER, "Character", 20, 0, false, true));
				types.Add(EDBType.VARCHAR, new BasicDataTypeInfo(EDBType.VARCHAR, "VarChar", 200, 0, false, true));
				types.Add(EDBType.LONGVARCHAR, new BasicDataTypeInfo(EDBType.LONGVARCHAR, "LongVarChar", 200, 0, false, true));
				types.Add(EDBType.DATE, new BasicDataTypeInfo(EDBType.DATE, "Date", 0, 0, false, false));
				types.Add(EDBType.DATETIME, new BasicDataTypeInfo(EDBType.DATETIME, "DateTime", 0, 0, false, false));
				types.Add(EDBType.BINARY, new BasicDataTypeInfo(EDBType.BINARY, "Blob", 0, 0, false, false));
				types.Add(EDBType.BINARYFILE, new BasicDataTypeInfo(EDBType.BINARYFILE, "BlobFile", 0, 0, false, false));
				types.Add(EDBType.BITMAP, new BasicDataTypeInfo(EDBType.BITMAP, "Image", 0, 0, false, false));
				types.Add(EDBType.Boolean, new BasicDataTypeInfo(EDBType.Boolean, "Boolean", 0, 0, false, false));
				types.Add(EDBType.GUID, new BasicDataTypeInfo(EDBType.GUID, "GUID", 0, 0, false, false));
				types.Add(EDBType.VIDEO, new BasicDataTypeInfo(EDBType.VIDEO, "Video", 0, 0, false, false));
				types.Add(EDBType.AUDIO, new BasicDataTypeInfo(EDBType.AUDIO, "Audio", 0, 0, false, false));
				types.Add(EDBType.GEOGRAPHY, new BasicDataTypeInfo(EDBType.GEOGRAPHY, "Geography", 0, 0, false, false));
				types.Add(EDBType.GEOPOINT, new BasicDataTypeInfo(EDBType.GEOPOINT, "GeoPoint", 0, 0, false, false));
				types.Add(EDBType.GEOLINE, new BasicDataTypeInfo(EDBType.GEOLINE, "GeoLine", 0, 0, false, false));
				types.Add(EDBType.GEOPOLYGON, new BasicDataTypeInfo(EDBType.GEOPOLYGON, "GeoPolygon", 0, 0, false, false));

				return types;
		}

		public static bool IsStandardType(EDBType type)
		{
			return (type == EDBType.BINARY
					|| type == EDBType.CHARACTER
					|| type == EDBType.DATE
					|| type == EDBType.DATETIME
					|| type == EDBType.GUID
					|| type == EDBType.INT
					|| type == EDBType.LONGVARCHAR
					|| type == EDBType.NUMERIC
					|| type == EDBType.PACKED
					|| type == EDBType.VARCHAR
					|| type == EDBType.ZONED
					|| type == EDBType.BITMAP
					|| type == EDBType.Boolean
					|| type == EDBType.VIDEO
					|| type == EDBType.AUDIO
					|| type == EDBType.BINARYFILE
					|| type == EDBType.GEOGRAPHY
					|| type == EDBType.GEOPOINT
					|| type == EDBType.GEOLINE
					|| type == EDBType.GEOPOLYGON
					);
		}

		#endregion

		#region IBaseInfo Implementation

		public string Name { get; set; }

		public string Namespace
		{
			get { return null; }
		}

	

		#endregion

		#region ITypedObjectInfo Implementation

		public bool IsBasicType
		{
			get { return IsStandardType(Type); }
		}

	

		public bool IsPublic
		{
			get { return IsStandardType(Type); }
		}

		public string TypeId
		{
			get { return Type.ToString(); }
		}

	

		public object TypedObject
		{
			get { return this; }
		}

		#endregion

		#region ITypedObject Implementation

	

		public EDBType Type { get; set; }

		public int Length
		{
			get { return m_Length; }
		}

		public int Decimals
		{
			get { return m_Decimals; }
		}

		public bool Signed
		{
			get { return m_Signed; }
			
		}

		public bool IsCollection
		{
			get { return false; }
			
		}

		#endregion
	}
}
