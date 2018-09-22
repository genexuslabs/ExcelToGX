using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Security;
using System.Text;
using System.Text.RegularExpressions;

namespace Artech.Common.Helpers.Strings
{
	public static class Format
	{
		private static readonly Regex m_Intercaps = new Regex(@"([^A-Z ])([A-Z])", RegexOptions.Compiled);
		private static readonly Regex m_Intercaps2 = new Regex(@"([^A-Z _])([A-Z])", RegexOptions.Compiled);
		private static readonly string m_InlineDefinitionSeparator = "=";

		#region InterCaps Strings

		/// <summary>
		/// Gets the first constituent part of an intercaps string.
		/// </summary>
		public static string GetNameRoot(string text)
		{
			if (text == null || text.Length == 0)
				return String.Empty;

			Match wordEnd = m_Intercaps.Match(text);
			if (wordEnd.Success)
				return text.Substring(0, wordEnd.Index + 1);
			else
				return text;
		}

		/// <summary>
		/// Gets the concatenation of the constituent parts of an InterCaps string, minus the last one.
		/// If the text contains only one component, the entire string is returned.
		/// </summary>
		public static string GetNameBase(string text)
		{
			if (text == null || text.Length == 0)
				return String.Empty;

			string[] components = GetNameComponents(text);
			if (components.Length > 1)
				return String.Join(String.Empty, components, 0, components.Length - 1);
			else if (components.Length == 1)
				return components[0];

			Debug.Assert(false);
			return String.Empty;
		}

		/// <summary>
		/// Separates an InterCaps string into its constituents.
		/// </summary>
		public static string[] GetNameComponents(string text)
		{
			if (text == null || text.Length == 0)
				return new string[] { };

			int componentStart = 0;
			List<string> parts = new List<string>();
			foreach (Match match in m_Intercaps.Matches(text))
			{
				parts.Add(text.Substring(componentStart, match.Index + 1 - componentStart));
				componentStart = match.Index + 1;
			}

			parts.Add(text.Substring(componentStart, text.Length - componentStart));
			return parts.ToArray();
		}

		/// <summary>
		/// Inserts a space between any two constituents of an intercaps string.
		/// </summary>
		public static string GetDescription(string text)
		{
			if (string.IsNullOrEmpty(text))
				return string.Empty;
			if (text.Trim().Contains(" "))
				return text;
			return m_Intercaps2.Replace(text, @"$1 $2");
		}

		#endregion

		#region Inline Definition Strings

		/// <summary>
		/// The string that is used to separate name from definition by <see cref="ParseInlineDefinition"/>.
		/// </summary>
		public static string InlineDefinitionSeparator
		{
			get { return m_InlineDefinitionSeparator; }
		}

		/// <summary>
		/// Parses the input string as an "inline definition " (of the form A=B) returning both components.
		/// </summary>
		/// <param name="text">Input string.</param>
		/// <param name="name">Left part of the definition if available, or the empty string otherwise.</param>
		/// <param name="definition">Right part of the definition if available, otherwise the empty string.</param>
		public static bool ParseInlineDefinition(string text, out string name, out string definition)
		{
			text = text.Trim();
			int separatorPosition = text.IndexOf(InlineDefinitionSeparator);
			if (separatorPosition != -1)
			{
				name = text.Substring(0, separatorPosition).Trim();
				definition = text.Substring(separatorPosition + 1).Trim();
				return true;
			}
			else
			{
				name = text;
				definition = String.Empty;
				return false;
			}
		}

		/// <summary>
		/// Parses the input string as an "inline order definition " (of the form A WHERE B) returning both components.
		/// </summary>
		/// <param name="text">Input string.</param>
		/// <param name="name">Left part of the definition if available, or the empty string otherwise.</param>
		/// <param name="definition">Right part of the definition if available, otherwise the empty string.</param>
		public static bool ParseInlineOrderDefinition(string text, out string[] names, out string definition)
		{
			text = text.Trim();
			string[] definitionParts = text.ToLower().Split(new string[] { "when" }, StringSplitOptions.None);
			string namesStr;
			if (definitionParts.Length == 2)
			{
				namesStr = text.Substring(0, definitionParts[0].Length);
				definition = text.Substring(text.Length - definitionParts[1].Length).Trim();
			}
			else
			{
				namesStr = text;
				definition = String.Empty;
			}

			names = namesStr.Split(new char[] { ',' });

			for (int i = 0; i < names.Length; i++)
			{
				names[i] = names[i].Trim();
			}

			return definition != String.Empty;
		}

		public static string MakeIdentifier(string s, string fillBlank)
		{
			if (String.IsNullOrEmpty(s))
				throw new ArgumentNullException("s");

			string stFormD = s.Normalize(NormalizationForm.FormD);
			StringBuilder sb = new StringBuilder();
			bool needsCapital = true;

			// Identifiers must begin with letter or underscore.
			// Since digits are not removed by the loop below, prepend an underscore.
			if (stFormD.Length > 0 && Char.IsDigit(stFormD[0]))
				sb.Append("_");

			for (int i = 0; i < stFormD.Length; i++)
			{
				UnicodeCategory uc = CharUnicodeInfo.GetUnicodeCategory(stFormD[i]);
				if (uc != UnicodeCategory.NonSpacingMark)
				{
					if (Char.IsLetter(stFormD[i]) || Char.IsDigit(stFormD[i]))
					{
						sb.Append((needsCapital ? Char.ToUpperInvariant(stFormD[i]) : stFormD[i]));
						needsCapital = false;
					}
					else if (Char.IsWhiteSpace(stFormD[i]))
					{
						if (fillBlank.Length == 0)
							needsCapital = true;
						else
							sb.Append(fillBlank);
					}
				}
			}

			return sb.ToString();
		}

		public static string MakeIdentifier(string s)
		{
			return Format.MakeIdentifier(s, "");
		}


		#endregion

		#region Number String

		#region Categories

		private struct NumberCategory
		{
			public NumberCategory(int multiple, string suffix)
			{
				Multiple = multiple;
				Suffix = suffix;
			}

			public int Multiple;
			public string Suffix;
		}

		private static readonly NumberCategory[] m_NumberCategories = new NumberCategory[] { new NumberCategory(1024 * 1024, "M"), new NumberCategory(1024, "K") };

		#endregion

		public static bool ParseInteger(string text, out int value)
		{
			if (text == null)
				throw new ArgumentNullException("text");

			if (Int32.TryParse(text, out value))
				return true;

			foreach (NumberCategory category in m_NumberCategories)
			{
				if (text.EndsWith(category.Suffix) && Int32.TryParse(text.Substring(0, text.Length - category.Suffix.Length), out value))
				{
					value = value * category.Multiple;
					return true;
				}
			}

			return false;
		}

		public static string FormatInteger(int value)
		{
			if (value != 0)
				foreach (NumberCategory category in m_NumberCategories)
					if (value % category.Multiple == 0)
						return (value / category.Multiple).ToString() + category.Suffix;

			return value.ToString();
		}

		public static string IntegerRegularExpression
		{
			get
			{
				StringBuilder suffixes = new StringBuilder();
				foreach (NumberCategory category in m_NumberCategories)
					suffixes.Append(category.Suffix);

				return String.Format(CultureInfo.InvariantCulture, "[0-9]+[{0}]?", suffixes.ToString());
			}
		}

		#endregion

		#region Misc Formatting

		/// <summary>
		/// Similar to <see cref="String.Join"/> but able to specify different separators.
		/// For example, can be used to generate a string of the form "A, B, C or D" from items A,B,C,D.
		/// </summary>
		public static string Join(string[] items, string commonSeparator, string lastSeparator)
		{
			if (items == null)
				throw new ArgumentNullException("items");

			if (commonSeparator == null)
				throw new ArgumentNullException("commonSeparator");

			if (lastSeparator == null)
				return String.Join(commonSeparator, items);

			if (items.Length == 0)
				return String.Empty;

			if (items.Length == 1)
				return items[0];

			return String.Concat(
				String.Join(commonSeparator, items, 0, items.Length - 1),
				lastSeparator,
				items[items.Length - 1]);
		}

		public static string ToSingleLine(string str)
		{
			if (String.IsNullOrEmpty(str))
				return String.Empty;

			str = str.Replace("\r", "");
			str = str.Replace("\n", " ");
			str = str.Replace("\t", " ");
			while (str.Contains("  "))
				str = str.Replace("  ", " ");

			return str;
		}

		public static string ToSingleLine(string[] str, string separator)
		{
			if (str == null)
				return String.Empty;

			return string.Join(separator, str);
		}

		public static string NormalizeCase(string s)
		{
			if (s != null)
			{
				if (s.Length <= 1)
					return s.ToUpper();
				else
					return s.Substring(0, 1).ToUpper() + s.Substring(1).ToLower();
			}
			return null;
		}

		public static string StringFormat(this string text, params object[] args)
		{
			return string.Format(text, args);
		}

		public static bool IsNullOrEmptyTrim(this string text)
		{
			return string.IsNullOrEmpty(text) || string.IsNullOrEmpty(text.Trim());
		}

		public static SecureString ToSecure(this string current)
		{
			var secure = new SecureString();
			foreach (var c in current.ToCharArray()) secure.AppendChar(c);
			return secure;
		}

		/// <summary>
		/// Taken from http://www.siao2.com/2007/05/14/2629747.aspx
		/// Replaces non-unicode characteres with its unicode counterpart
		/// </summary>
		/// <param name="text">The text to replace</param>
		/// <returns>The replaced string</returns>
		public static string RemoveDiacritics(string text)
		{
			var normalizedString = text.Normalize(NormalizationForm.FormD);
			var stringBuilder = new StringBuilder();

			foreach (var c in normalizedString)
			{
				var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
				if (unicodeCategory != UnicodeCategory.NonSpacingMark)
				{
					stringBuilder.Append(c);
				}
			}

			return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
		}

		#endregion
	}
}
