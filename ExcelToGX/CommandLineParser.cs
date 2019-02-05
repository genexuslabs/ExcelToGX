//#define NUNIT
// CommandLineParser.cs: Contributed by Chris Sells [csells@sellsbrothers.com]
// A command line parser class -- see the sample for usage

#region Copyright © 2002-2004 The Genghis Group

/*
 * This software is provided 'as-is', without any express or implied warranty.
 * In no event will the authors be held liable for any damages arising from the
 * use of this software.
 *
 * Permission is granted to anyone to use this software for any purpose,
 * including commercial applications, subject to the following restrictions:
 *
 * 1. The origin of this software must not be misrepresented; you must not claim
 * that you wrote the original software. If you use this software in a product,
 * an acknowledgment in the product documentation is required, as shown here:
 *
 * Portions Copyright © 2002-2004 The Genghis Group (http://www.genghisgroup.com/).
 *
 * 2. No substantial portion of the source code of this library may be redistributed
 * without the express written permission of the copyright holders, where
 * "substantial" is defined as enough code to be recognizably from this library.
*/

#endregion

#region Features

/*
 * -Parsing command line args from an array of strings or a file
 * -Simple flags and flag/arg pairs
 * -Typed args, both params and flags
 * -Building usages on the fly
 * -Command line args can be read via @argfile
 * -Automatically generates the banner logo from the version attributes
 * -Case insensitive flag comparisions by default, case sensitive as an option
*/

#endregion

#region Limitations

/*
 * -Flags should support +/- at the end to turn on/off
 * -Requires flags to be space separated, e.g.
 *  /efg is a single flag, even if e, f and g are defined and efg isn't
 * -Requires flag and value to be space separated, e.g.
 * /fconfig.sys is treated as a single flag, even if f is defined to take
 *  a value and fconfig.sys isn't defined
 * -No support for separating flag from value via color or equals sign,
 *  as is fairly common
 * -Need better formatting to pad flag/param names that are larger than
 *  16 characters and to support multi-line descriptions.
 * -There is no equivalent (yet) of FileNameValue or
 *  PairValue from the unmanaged code.
*/

#endregion

#region History

/*
 * 11/20/03 (Ethan J. Brown)
 * - Fixed a crash when 0 arguments were supplied on the command line
 * - Fixed a bug when calling parse -- call to Environment.CommandLine has been replaced with Environment.GetCommandLineArgs
 * 
 * 11/11/03
 * -Updated "required arg not found" to use alternate name if one is provided.
 * -Updated core Parse method to be virtual to allow overrides for extensibility.
 * 
 * 08/25/03 (jcmag)
 * - Added a LegalValues parameter to test and restrict values. Use Regular expressions
 * for maximum flexibility and performance.
 * 
 * 08/06/03 - 08/13/03 (jcmag)
 * - Added support for IList fields, user can now specify multiple values for one parameter
 * - Usage now shows default values
 * - User can now use colon in values
 * - Fix a bug with optional parameters
 * 
 * 10/11/02 - 10/13/02
 * -Updated help generation
 *  1) Short usage now shows "[+|-]" for flags
 *  2) Long usage descriptions now formatted better to fit maximum length of each line
 *     Lines are now split on white spaces where possible. 
 *     See FormatMultiLineMaxChars and FormatSingleLineMaxChars
 *  3) Short usage also formatted to fit on line as above
 *  4) Added support for categories (UsageAttribute.Category and ParserUsageAttribute.ShowCategories)
 *  5) Default values show for FlagAttribute with AllowOnOff=true and ValueAttribute for any primitive type
 * 
 * 09/30/02 - 10/05/02
 * -Updated help generation, long lines correctly padded and now supports multi-line descriptions
 * -ParserUsageAttribute.MaxHelpLineLength, user can specify maximum length for help lines
 * -ValueUsageAttribute.Delimiters, allows user to specify which set of delimiters are allowed
 * -FlagUsageAttribute.AllowOnOff, allows usser to specify if flag accepts +/- suffixes
 * -ParserUsageAttribute.EnvironmentDefaults, defaults can now be overriden in environment variable
 * 
 * 9/25/01:
 * -Added support for ':' and '-' name-value separators
 * -Added support for '+' and '+' flag postfix values to turn on/off bool flags
 * 
 * 4/28/01:
 * -Moved the CLP to the Genghis namespace from the Genghis.General namespace.
 *
 * 3/27/01:
 * -Initial port from unmanaged C++ (http://www.sellsbrothers.com/tools/commandlineparser.zip)
*/

#endregion

#region TODO

//TODO ?default delimiter set in attrib / decide on best default (currently default=all)?
//TODO reading from environ variable: need to pass param to Parse(string[]) to indicate that required variables are not required
//TODO what is a PairValue?
//TODO debug code to check if a) variables are declared more than one, b) var names contain invalid chars (e.g. ":\= )
//TODO fix (rewrite) help generation
//  -Single method to return info (name, usage, description, default value, alternate names, enum values) - can use for Short/Long info 
//  -Help gen method formats output for all, no need for co-operative formatting
//  -One of the /? /help /h to generate extended/short help help (e.g. show/hide defaults)
//  -?categories for variables: only group if current var has same cat as previous. So cats could be shown multiple times.
//      this could be improved, e.g. sorting between MatchPosition items
//  -?when showing default values show actual default or the default after parsing the environ var (or both?)

#endregion

using System;
using System.Collections.Specialized;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.ComponentModel;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Artech.Common.Helpers
{
	public class CommandLineParser
	{
		[Flags]
		public enum ValueDelimiters
		{
			Space = 1,
			Colon = 2,
			Equals = 4
		}

		[AttributeUsage(AttributeTargets.Class | AttributeTargets.Struct)]
		public class ParserUsageAttribute : Attribute
		{
			public ParserUsageAttribute(string description)
			{
				Description = description;
			}

			public string Description;
			public string PreferredPrefix = "/";
			public bool AllowArgumentFile = true;
			public string EnvironmentDefaults = "";
			public int MaxHelpLineLength = 79;
			public bool ShowCategories = false;
		}

		[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
		public abstract class UsageAttribute : Attribute
		{
			protected UsageAttribute(string description)
			{
				Description = description;
			}

			public string Description = "";
			public string Name = "";
			public bool IgnoreCase = true;
			public bool Optional = true;
			public string ValueName = "value";
			public bool MatchPosition = false;
			public string Category = "";
			public bool Hidden = false; // Don't show in Usage.
			public string LegalValues; // case-insensitive comparison

			// NOTE: Using two alternates instead of an array because
			// array attributes are not CLS compliant
			public string AlternateName1 = "";
			public string AlternateName2 = "";

			protected bool found;

			protected bool Flag => !MatchPosition;

			public bool Matches(string name, string defaultName)
			{
				Debug.Assert(defaultName.Length > 0);
				return (string.Compare(name, (Name.Length == 0 ? defaultName : Name), IgnoreCase) == 0) ||
					(string.Compare(name, AlternateName1, IgnoreCase) == 0) ||
					(string.Compare(name, AlternateName2, IgnoreCase) == 0);
			}

			// HACK :
			protected internal virtual bool IsFound
			{
				get { return found; }
				set { found = value; }
			}

			protected internal abstract bool ExpectsValue { get; }

			protected internal abstract void ConsumeValue(string s, object container, MemberInfo info);
			protected internal abstract string GetShortUsage(string prefix, string defaultName, object container, MemberInfo info);
			protected internal abstract string GetLongUsage(string prefix, string defaultName, object container, MemberInfo info);

			protected void ConsumeValueHelper(string s, object container, MemberInfo info)
			{
				// TODO: Read value from string into array (efficiently?)
				// TODO: Indexed properties?
				try
				{
					PropertyInfo prop = info as PropertyInfo;
					FieldInfo field = info as FieldInfo;
					Debug.Assert((prop != null) || (field != null));

					if (field != null)
					{
						// Collection field
						IList list = field.GetValue(container) as IList;
						if (list != null)
						{
							list.Add(s);
						}
						// Simple field
						else
						{
							field.SetValue(container, ReadFromString(s, field.FieldType));
						}
					}
					// Property
					else
					{
						Debug.Assert(prop != null);
						prop.SetValue(container, ReadFromString(s, prop.PropertyType), null);
					}
				}
				catch (TargetInvocationException e)
				{
					throw e.InnerException;
				}
			}

			protected object GetValueHelper(object container, MemberInfo info)
			{
				try
				{
					PropertyInfo prop = info as PropertyInfo;
					FieldInfo field = info as FieldInfo;
					Debug.Assert((prop != null) || (field != null));

					if (field != null)
						return field.GetValue(container);
					else
					{
						Debug.Assert(prop != null);
						return prop.GetValue(container, null);
					}
				}
				catch (TargetInvocationException)
				{
					return null;
				}
			}

			// Normally, prefix == "/"
			// defaultName is the name of the field or property obtained via reflection
			protected string GetShortUsageHelper(string prefix, string defaultName, string valueName, object container, MemberInfo info)
			{
				StringBuilder usage = new StringBuilder();
				if (Optional) usage.Append("[");
				usage.Append(prefix);

				// If there's no prefix, assume a param instead of a flag,
				// and wrap the name (great for names with spaces, e.g. "min value" in [/minX <min value>])
				bool wrapName = prefix.Length == 0;
				if (wrapName && !Optional) usage.Append("<");

				Debug.Assert(defaultName.Length > 0);
				usage.Append(Name.Length != 0 ? Name : defaultName);
				if (AlternateName1.Length > 0) usage.Append("|").Append(AlternateName1);
				if (AlternateName2.Length > 0) usage.Append("|").Append(AlternateName2);

				if (valueName.Length != 0) usage.Append(" <").Append(ValueName).Append(">");
				if (wrapName && !Optional) usage.Append(">");
				if (Optional) usage.Append("]");
				return usage.ToString();
			}


			protected string GetLongUsageHelper(string prefix, string defaultName, string valueName, object container, MemberInfo info)
			{
				StringBuilder usage = new StringBuilder();

				if (Description.Length == 0) return "";

				usage.Append(prefix);

				Debug.Assert(defaultName.Length > 0);
				usage.Append(Name.Length != 0 ? Name : defaultName);

				if (valueName.Length != 0) usage.Append(" <").Append(ValueName).Append(">");

				usage.Append('\t').Append(Description);

				return usage.ToString();
			}

			protected object ReadFromString(string s, Type type)
			{
				TypeConverter converter = TypeDescriptor.GetConverter(type);
				return converter.ConvertFromString(s);
			}
		}

		// Presence on the commandline, e.g. /foo
		[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
		public class FlagUsageAttribute : UsageAttribute
		{
			public FlagUsageAttribute(string description)
				: base(description)
			{ }

			public bool AllowOnOff = true;

			protected internal override bool ExpectsValue => false;

			protected internal override void ConsumeValue(string s, object container, MemberInfo info)
			{
				Debug.Assert((s.ToLower() == "true") || (s.ToLower() == "false"), "Flags only consume boolean values");

				ConsumeValueHelper(s, container, info);
				found = true;
			}

			protected internal override string GetShortUsage(string prefix, string defaultName, object container, MemberInfo info)
			{
				string usage = GetShortUsageHelper(prefix, defaultName, "", container, info);
				if (AllowOnOff)
				{
					if (usage.EndsWith("]"))
					{
						usage = usage.Remove(usage.Length - 1, 1);
						usage += "[+|-]]";
					}
					else
					{
						usage += "[+|-]";
					}
				}

				return usage;
			}

			protected internal override string GetLongUsage(string prefix, string defaultName, object container, MemberInfo info)
			{
				string usage = GetLongUsageHelper(prefix, defaultName, "", container, info);
				if (AllowOnOff)
				{
					//Find the 1st tab char and insert "[+|-]"
					int i = usage.IndexOf('\t');
					if (-1 != i) usage = usage.Substring(0, i) + "[+|-]" + usage.Substring(i);

					if (!Optional)
					{
						//Show the default value
						object defaultValue = GetValueHelper(container, info);
						Debug.Assert(defaultValue is bool, "Expecting bool type");
						if (defaultValue is bool) usage += $" ({((bool)defaultValue ? "On" : "Off")} by default)";
					}
				}

				return usage;
			}
		}

		// On the commandline with an associated value,
		// e.g. /foo "fooness" (MatchPosition = false) or "fooness" (MatchPosition = true)
		[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
		public class ValueUsageAttribute : UsageAttribute
		{
			public ValueUsageAttribute(string description)
				: base(description)
			{
			}

			public ValueDelimiters Delimiters = ValueDelimiters.Colon | ValueDelimiters.Equals | ValueDelimiters.Space;

			protected internal override bool ExpectsValue => true;

			protected internal override void ConsumeValue(string s, object container, MemberInfo info)
			{
				ConsumeValueHelper(s, container, info);
				found = true;
			}

			protected internal override string GetShortUsage(string prefix, string defaultName, object container, MemberInfo info)
			{
				// Get the usage string, assuming it will return string whith format of "name <value>" for any
				// arg that takes a value
				string usage = GetShortUsageHelper(prefix, defaultName, Flag ? ValueName : "", container, info);

				// Select a delimiter to show in the usage.
				string delimiter;
				if (0 != (Delimiters & ValueDelimiters.Colon)) delimiter = ":";
				else if (0 != (Delimiters & ValueDelimiters.Equals)) delimiter = "=";
				else delimiter = " ";

				// Replace the space between the arg name and the first '<' char with the appropriate delimiter
				return usage.Replace(" ", delimiter);
			}

			protected internal override string GetLongUsage(string prefix, string defaultName, object container, MemberInfo info)
			{
				// TODO: If we're filling an array or collection, append "..." to ValueName
				// return UsageHelper(sPrefix, true, bFlag ? __T("value") : __T("")) + __T("...");
				string usage = GetLongUsageHelper(prefix, defaultName, Flag ? ValueName : "", container, info);

				// Select a delimiter to show in the usage.
				string delimiter;
				if (0 != (Delimiters & ValueDelimiters.Colon)) delimiter = ":";
				else if (0 != (Delimiters & ValueDelimiters.Equals)) delimiter = "=";
				else delimiter = " ";

				usage = Regex.Replace(usage, " \\<value\\>\t", delimiter + "<value>\t");

				// HACK :
				if (Optional)
				{
					// Get the default value - only display for primitive types and strings
					// Create a new instance to retrieve default values (and not what the user typed)
					object newContainer = Activator.CreateInstance(container.GetType());
					object defaultValue = GetValueHelper(newContainer, info);
					if ((null != defaultValue) && (defaultValue.GetType().IsPrimitive || defaultValue is string))
						usage += " (Default is \"" + defaultValue + "\")";
				}

				return usage;
			}
		}

		// Skip variables marked NoUsage. By default, all fields and properties
		// are assumed to be part of the usage. This makes it easy to get
		// started with this class, i.e. to hand an object full of fields to
		// the parser without any special attributes and have it work.
		// NOTE: Deriving from UsageAttribute makes enumerating
		// our attributes easier, as they all have a common base
		[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
		public class NoUsageAttribute : UsageAttribute
		{
			public NoUsageAttribute()
				: base(null)
			{
			}

			protected internal override bool ExpectsValue => false;

			protected internal override void ConsumeValue(string s, object container, MemberInfo info)
			{
			}

			protected internal override string GetShortUsage(string prefix, string defaultName, object container, MemberInfo info)
			{
				return "";
			}

			protected internal override string GetLongUsage(string prefix, string defaultName, object container, MemberInfo info)
			{
				return "";
			}
		}

		[Serializable]
		public class UsageException : ApplicationException
		{
			public UsageException(string arg, string error)
				: base(error)
			{
				this.arg = arg;
			}

			protected string arg;

			public string Argument => arg;

			public override string Message => arg + ": " + base.Message;
		}

		protected class MemberInfoUsage
		{
			public MemberInfo info;
			public UsageAttribute usage;

			public MemberInfoUsage(MemberInfo info, UsageAttribute usage)
			{
				this.info = info;
				this.usage = usage;
			}
		}

		[NoUsage]
		protected MemberInfoUsage[] members;

		protected MemberInfoUsage[] GetMembers()
		{
			// Return cached members
			if (members != null) return members;

			// Cache members
			ArrayList memberList = new ArrayList();

			// TODO: Parse base class first to get /v and /h shown first
			BindingFlags flags = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;
			foreach (MemberInfo info in GetType().GetMembers(flags))
			{
				// Only doing fields and properties
				if ((info.MemberType != MemberTypes.Field) && (info.MemberType != MemberTypes.Property)) continue;

				// Skip variables w/o usage
				object[] attribs = info.GetCustomAttributes(typeof(UsageAttribute), true);
				UsageAttribute usage = (UsageAttribute)(attribs.Length != 0 ? attribs[0] : null);
				if (usage is NoUsageAttribute) continue;

				// Default settings with no attribute
				if (usage == null)
				{
					PropertyInfo prop = info as PropertyInfo;
					FieldInfo field = info as FieldInfo;
					Debug.Assert((prop != null) || (field != null));

					// If the type is bool, it's probably just a flag
					if (((prop != null) && (prop.PropertyType == typeof(bool))) ||
						((field != null) && (field.FieldType == typeof(bool))))
					{
						usage = new FlagUsageAttribute(info.Name);
					}
					// If the type is not a bool, it probably also have a value
					else
					{
						usage = new ValueUsageAttribute(info.Name);
					}
				}

				memberList.Add(new MemberInfoUsage(info, usage));
			}

			members = (MemberInfoUsage[])memberList.ToArray(typeof(MemberInfoUsage));
			return members;
		}

		public string GetUsage()
		{
			return GetUsage("");
		}

		// Splits a string containing new lines (assumed to be \n or \r\n) and passes each
		// line to FormatSingleLineMaxChars and returns the resulting string array
		public static string[] FormatMultiLineMaxChars(int maxLen, string sLine)
		{
			Debug.Assert(maxLen > 0, "Max must be at least 1");
			if (maxLen <= 0) return null;

			string[] lines = sLine.Replace("\r", "").Split('\n');
			StringCollection formattedLines = new StringCollection();

			foreach (string line in lines)
				formattedLines.AddRange(FormatSingleLineMaxChars(maxLen, line));

			string[] multi = new string[formattedLines.Count];
			formattedLines.CopyTo(multi, 0);
			return multi;
		}

		// Formats a string so that no line will have more than maxLen characters.
		// Input strings may not contain a \r or \n. To format multi-line strings call
		// FormatMultiLineMaxChars
		public static string[] FormatSingleLineMaxChars(int maxLen, string sLine)
		{
			Debug.Assert(maxLen > 0, "Max must be at least 1");
			Debug.Assert(-1 == sLine.IndexOf('\r'));
			Debug.Assert(-1 == sLine.IndexOf('\n'));

			if (maxLen <= 0) return null;
			if (0 == sLine.Length) return new[] { "" };

			int currentStart = 0;
			StringCollection txt = new StringCollection();

			while (currentStart < sLine.Length)
			{
				if (currentStart + maxLen < sLine.Length)
				{
					int currentEnd = Math.Min(currentStart + maxLen, sLine.Length - 1);
					int spaceIdx = sLine.LastIndexOf(" ", currentEnd, currentEnd - currentStart);

					if (-1 == spaceIdx)
					{
						txt.Add(sLine.Substring(currentStart, currentEnd - currentStart));
						currentStart += maxLen;
					}
					else
					{
						txt.Add(sLine.Substring(currentStart, spaceIdx - currentStart));
						currentStart = spaceIdx + 1;
					}
				}
				else
				{
					txt.Add(sLine.Substring(currentStart));
					currentStart += maxLen;
				}
			}

			string[] lines = new string[txt.Count];
			txt.CopyTo(lines, 0);
			return lines;
		}

		public string GetUsage(string err)
		{
			StringBuilder usage = new StringBuilder();

			// Logo
			string logo = GetLogo();
			if (logo.Length != 0)
			{
				usage.Append(logo).Append(Environment.NewLine);
			}

			// Parser prefs, e.g. preferred prefix
			object[] attribs = GetType().GetCustomAttributes(typeof(ParserUsageAttribute), true);
			ParserUsageAttribute parser = (ParserUsageAttribute)(attribs.Length != 0 ? attribs[0] : null);
			string preferredPrefix = parser != null ? parser.PreferredPrefix : "/";
			bool allowArgFile = parser == null || parser.AllowArgumentFile;

			// Error string
			if (err.Length != 0)
			{
				usage.Append(err).Append(Environment.NewLine).Append(Environment.NewLine);
			}

			// Short (name and value name only)
			StringBuilder shortUsage = new StringBuilder();
			shortUsage.Append("Usage: ").Append(GetModuleName()).Append(" ");

			// Long (name and description only)
			StringBuilder longUsage = new StringBuilder();

			// TODO there must be a better way of doing this than looping through the MemberInfo...
			// Find the right-most tab char. 
			int maxTabPos = 0;
			foreach (MemberInfoUsage member in GetMembers())
			{
				if (member.usage != null && member.usage.Hidden)
					continue;

				string prefix = (member.usage.MatchPosition ? "" : preferredPrefix);
				string tmpLongUsage = member.usage.GetLongUsage(prefix, member.info.Name, this, member.info);
				maxTabPos = Math.Max(maxTabPos, tmpLongUsage.IndexOf("\t"));
			}

			// There should be 2 chars after the longest usage, before its description
			maxTabPos += 2;
			string argDscLineFormat = $"{{0,-{maxTabPos}:S}}{{1}}";

			// Max length for the descriptions
			int maxLen = (parser?.MaxHelpLineLength ?? 79) - maxTabPos;

			if (allowArgFile)
			{
				shortUsage.Append("[@argfile]");
				//TODO format line using FormatSingleLineMaxChars
				longUsage.AppendFormat(argDscLineFormat, "@argfile", "Read arguments from a file.").Append(Environment.NewLine);
			}

			//Regex to split line into usage and description
			Regex regexUsage = new Regex(@"(?s)^(?<usage>[^\t]+)(\t)(?<desc>.*)$");

			string lastCategory = "";

			foreach (MemberInfoUsage member in GetMembers())
			{
				if (member.usage != null && member.usage.Hidden)
					continue;

				// NOTE: When matching by position, only the value will be present
				// on the commandline, e.g. "fooness"
				string prefix = member.usage.MatchPosition ? "" : preferredPrefix;

				shortUsage.Append(" ").Append(member.usage.GetShortUsage(prefix, member.info.Name, this, member.info));

				// Get the long usage, which must be of the format usage\tdescription
				// Where 
				//   1) usage may not contain any tabs
				//   2) description can be any length and may contain \n or \r\n chars to indicate a newline
				string tmpLongUsage = member.usage.GetLongUsage(prefix, member.info.Name, this, member.info);

				// Replace all \r chars
				Match m = regexUsage.Match(tmpLongUsage);

				if (null != m)
				{
					// If categories are being displayed and the category has changed then display the category
					// An empty category is not concidered to be a category change
					if ((null != parser) && (parser.ShowCategories) && (0 != member.usage.Category.Length) && (member.usage.Category != lastCategory))
					{
						longUsage.Append(Environment.NewLine).AppendFormat("{0,-" + maxTabPos + ":S}- {1} -", "", member.usage.Category).Append(Environment.NewLine);
						lastCategory = member.usage.Category;
					}

					// Usage
					longUsage.AppendFormat(argDscLineFormat, m.Groups["usage"].Value, string.Empty);

					// Format the string to fit the max line length
					string[] formattedLines = FormatMultiLineMaxChars(maxLen, m.Groups["desc"].Value);

					if (formattedLines.Length > 0) longUsage.Append(formattedLines[0]).Append(Environment.NewLine);
					else longUsage.Append(Environment.NewLine);

					for (int g = 1; g < formattedLines.Length; ++g)
						longUsage.AppendFormat(argDscLineFormat, string.Empty, formattedLines[g]).Append(Environment.NewLine);
				}
			}

			//Format the short usage
			string[] shortLines = FormatMultiLineMaxChars(parser?.MaxHelpLineLength ?? 79, shortUsage.ToString());
			shortUsage.Length = 0;

			if (shortLines.Length > 0) shortUsage.Append(shortLines[0]).Append(Environment.NewLine);
			else shortUsage.Append(Environment.NewLine);

			//subsequent short usage lines to align up under "Usage: " string
			for (int g = 1; g < shortLines.Length; ++g)
				shortUsage.AppendFormat("{0,-7:S}{1}", "", shortLines[g]).Append(Environment.NewLine);

			usage.Append(shortUsage).Append(Environment.NewLine).Append(Environment.NewLine).Append(longUsage).Append(Environment.NewLine);

			if ((null != parser) && ("" != parser.EnvironmentDefaults))
			{
				usage.Append(Environment.NewLine);
				usage.AppendFormat("Switches may be preset in the {0} environment variable.", parser.EnvironmentDefaults).Append(Environment.NewLine);
				usage.Append("Override preset switches by prefixing switches with a - (hyphen)").Append(Environment.NewLine);
				usage.Append(" for example, /W-").Append(Environment.NewLine);
			}
			return usage.ToString();
		}

		private MemberInfoUsage FindFlag(string name)
		{
			foreach (MemberInfoUsage member in GetMembers())
			{
				if (!member.usage.MatchPosition && member.usage.Matches(name, member.info.Name))
				{
					return member;
				}
			}

			return null;
		}

		private MemberInfoUsage GetNextParam()
		{
			foreach (MemberInfoUsage info in GetMembers())
			{
				// HACK :
				if (!info.usage.IsFound && info.usage.MatchPosition) return info;
			}

			return null;
		}

		public void Parse()
		{
			//FIX: Ethan J. Brown - why pass the entire CommandLine when it can be parsed for you?
			Parse(Environment.GetCommandLineArgs(), true);
		}

		public void Parse(String commandLine)
		{
			Parse(commandLine, false);
		}

		/// <summary>
		/// Parses the command line, using a "space" as the argument separator. Use quotes (") to define an argument
		/// with spaces (e.g. a path).
		/// </summary>
		/// <param name="commandLine"></param>
		/// <param name="ignoreFirstArg"></param>
		public void Parse(string commandLine, bool ignoreFirstArg)
		{
			List<string> args = new List<string>();
			string arg = string.Empty;
			bool inQuotes = false;

			int startIdx = 0;
			for (int i = 0; i != commandLine.Length; ++i)
			{
				char c = commandLine[i];
				if (!inQuotes && char.IsWhiteSpace(c))
				{
					if ((i - startIdx) > 0)
						args.Add(commandLine.Substring(startIdx, i - startIdx));
					startIdx = i + 1;
					continue;
				}

				if (c == '"')
				{
					if (inQuotes)
					{
						// Read ahead to check for doublequote pairs
						++i;
						if (i == commandLine.Length)
							break;
						if (commandLine[i] == '"')
							continue;
						--i;
						inQuotes = false;
					}
					else
					{
						inQuotes = true;
					}
				}
			}

			if ((commandLine.Length - startIdx) > 0)
				args.Add(commandLine.Substring(startIdx));
			Parse(args.ToArray(), ignoreFirstArg);
		}

		public void Parse(string[] args)
		{
			Parse(args, false);
		}

		// Return the index of the first inline delimiter - i.e. a ':' or a '='
		// that are passed in as part of the arg name. e.g. "/autoexec:c:\autoexec.bat"
		// Returns -1 if no delimiter is found
		private int FindValueInlineDeliminter(string name)
		{
			int delimiterPos = name.IndexOf(':');

			if (delimiterPos == -1) return name.IndexOf('=');
			else return delimiterPos;
		}

		/// <summary>
		/// Does the actual parsing. Good place to hook up for custom pre/post-parsing.
		/// </summary>
		/// <param name="args">array of arguments to parse</param>
		/// <param name="ignoreFirstArg">whether to ignore first argument or not (it could be the name of the exe)</param>
		protected virtual void Parse(string[] args, bool ignoreFirstArg)
		{
			object[] attribs = GetType().GetCustomAttributes(typeof(ParserUsageAttribute), true);
			ParserUsageAttribute parser = (ParserUsageAttribute)(attribs.Length != 0 ? attribs[0] : null);
			bool allowArgFile = (parser != null ? parser.AllowArgumentFile : true);
			MemberInfoUsage[] members = GetMembers();

			// HACK :
			MemberInfoUsage member = null;

			//FIX: Ethan J. Brown -- used to crash with 0 arguments
			for (int i = ignoreFirstArg ? 1 : 0; i < args.Length; ++i)
			{
				string arg = args[i];
				Debug.WriteLine("Processing arg: " + arg);

				bool isFlag = false;

				// It's a flag
				if ((arg.Length > 1) && ((arg[0] == '/') || (arg[0] == '-')))
				{
					// HACK :
					member = null;
					string flagName;

					// Flags can have a '+' or '-' suffix. If this arg has a prefix remove it,
					// else just remove the '/' or '-'
					bool hasOnOff = arg.EndsWith("-") || arg.EndsWith("+");
					if (hasOnOff) flagName = arg.Substring(1, arg.Length - 2);
					else flagName = arg.Substring(1);

					// Flags can be passed in one of two ways
					//  a) With a param name in 1 arg and a value in the next. e.g. "/flag myValue"
					//  b) Using a delimiter, passed as a single argument. e.g. "/flag:myValue"
					// If using the single arg (a) method extract the arg name
					int delimiterPos = FindValueInlineDeliminter(flagName);

					//default delimiter. A space which is not a delimiter (space delimted values are split into seperate args)
					// will thus have a delimiter value of '\0'
					char delimiter = '\0';
					if (delimiterPos != -1)
					{
						delimiter = flagName[delimiterPos];
						flagName = flagName.Substring(0, delimiterPos);
					}

					// Find the argument by name
					member = FindFlag(flagName);

					if (member != null)
					{
						isFlag = true;


						//OnOff toggle only allowed if: member.usage is FlagUsage and FlagUsage.AllowOnOff is true
						if (hasOnOff)
						{
							FlagUsageAttribute flag = member.usage as FlagUsageAttribute;

							if (null == flag) throw new UsageException(arg, "Only flags support on/off toggles");
							else if (!flag.AllowOnOff) throw new UsageException(arg, "Flag does not allow on/off toggle");
						}

						//Check that only value args have a delimiter and that the correct delimiter has been used
						if (member.usage is ValueUsageAttribute)
						{
							ValueUsageAttribute val = (ValueUsageAttribute)member.usage;

							switch (delimiter)
							{
								case ':':
									if (0 == (val.Delimiters & ValueDelimiters.Colon)) throw new UsageException(arg, "This value param does not support a ':' delimiter");
									break;
								case '=':
									if (0 == (val.Delimiters & ValueDelimiters.Equals)) throw new UsageException(arg, "This value param does not support a '=' delimiter");
									break;
								case '\0':
									if (0 == (val.Delimiters & ValueDelimiters.Space)) throw new UsageException(arg, "This value param does not support a ' ' delimiter");
									break;
								default:
									throw new UsageException(arg, "Unknown delimiter");
							}
						}
						else if ('\0' != delimiter) throw new UsageException(arg, "Only value params support delimiters");
					}
				}
				// It's a file name to process parameters from
				else if ((arg.Length > 1) && (arg[0] == '@') && allowArgFile)
				{
					// HACK :
					member = null;
					ParseFromFile(arg.Substring(1));
					continue;
				}
				// It's a parameter
				else
				{
					// HACK :
					if (member != null)
					{
						FieldInfo field = member.info as FieldInfo;
						if (field != null)
						{
							IList list = field.GetValue(this) as IList;
							if (list == null)
							{
								member.usage.IsFound = true;
								// Find the argument by offset
								member = GetNextParam();
							}
						}
					}
					else
						member = GetNextParam();
				}

				if (member == null) throw new UsageException(arg, "Unrecognized argument");

				// Argument with a value, e.g. /foo bar
				if (member.usage.ExpectsValue)
				{
					string value;

					// If the arg has an inline delimiter (e.g. "/flag:myValue") 
					// then read the value from the current arg else get the value from the next arg
					// HACK : only search delimiters when we have a beginning slash
					if (arg[0] == '/')
					{
						int delimiterPos = FindValueInlineDeliminter(arg);
						if (delimiterPos == -1)
						{
							if (isFlag && (++i == args.Length)) throw new UsageException(arg, "Argument expects a parameter");
							value = args[i];
						}
						else
						{
							if ((arg.Length - 1) == delimiterPos) throw new UsageException(arg, "Argument expects a parameter");
							value = arg.Substring(delimiterPos + 1);
						}
					}
					else
					{ // positional parameter
						value = args[i];
					}
					member.usage.ConsumeValue(value, this, member.info);
				}
				// Argument w/o a value, e.g. /foo
				else
				{
					string value;

					if (isFlag && arg.EndsWith("-")) value = "false";
					else value = "true";

					member.usage.ConsumeValue(value, this, member.info);
				}
				// Test legal values
				if (member.usage.LegalValues != null)
				{
					FieldInfo field = member.info as FieldInfo;
					object val = field.GetValue(this);
					Regex myRegex = new Regex(member.usage.LegalValues);
					if (!myRegex.IsMatch(val.ToString()))
						throw new UsageException(arg, "The value doesn't match one of the legal values");
				}
			}

			// HACK :
			// Check for missing required arguments
			foreach (MemberInfoUsage mbr in members)
			{
				if (!mbr.usage.Optional && !mbr.usage.IsFound)
				{
					throw new UsageException(!string.IsNullOrEmpty(mbr.usage.Name) ? mbr.usage.Name : mbr.info.Name, "Required argument not found");
				}
			}
		}

		// Set (and auto-reset) current working directory
		// We do this so that any file names read out of the file
		// can be relative to the file, not where we're running
		// the app from. We're using the system managed working directory
		// as a "context" for the individual values, i.e. FileNameValue,
		// to be able to compute a complete file name. This is potentially
		// dangerous as the cwd is set per process, not per thread, but
		// since command lines are typically processed before threads are
		// fired off, we should be safe. It saves us from having to pass
		// a virtual cwd to all values as they're parsed.
		// TODO: This doesn't really work if the file/directory names
		// are just strings, since but the time they're turned into
		// full path names, the CWD has already been reset...
		protected class CurrentDir : IDisposable
		{
			public CurrentDir(string newDir)
			{
				m_OldDir = Environment.CurrentDirectory;
				Environment.CurrentDirectory = newDir;
			}

			#region Implementation of IDisposable

			public void Dispose()
			{
				Environment.CurrentDirectory = m_OldDir;
			}

			#endregion

			private readonly string m_OldDir;
		}

		private void ParseFromFile(string fileName)
		{
			// Check if file exists
			if (!(new FileInfo(fileName)).Exists)
			{
				throw new UsageException("argfile", fileName + " not found");
			}

			// Point current directory at the input file
			string argDir = Path.GetDirectoryName(Path.GetFullPath(fileName));
			using (new CurrentDir(argDir))
			using (StreamReader reader = new StreamReader(fileName))
			{
				Parse(reader.ReadToEnd());
			}
		}

		protected virtual string GetLogo()
		{
			// TODO: Refactor this ugly code!
			StringBuilder logo = new StringBuilder();
			string nl = Environment.NewLine;

			Assembly assem = GetAssembly();

			// Title
			string title = GetTitle(assem);

			// Version
			string version = assem.GetName().Version.ToString();

			// Copyright
			var attribs = assem.GetCustomAttributes(typeof(AssemblyCopyrightAttribute), true);
			string copyright = (attribs.Length != 0 ? ((AssemblyCopyrightAttribute)attribs[0]).Copyright : "");

			// Description: Try ParserUsage.Description first, use AssemblyDescription otherwise
			string description;
			attribs = GetType().GetCustomAttributes(typeof(ParserUsageAttribute), true);
			ParserUsageAttribute parser = (ParserUsageAttribute)(attribs.Length != 0 ? attribs[0] : null);
			if (!string.IsNullOrEmpty(parser?.Description))
			{
				description = parser.Description;
			}
			else
			{
				attribs = assem.GetCustomAttributes(typeof(AssemblyDescriptionAttribute), true);
				description = (attribs.Length != 0 ? ((AssemblyDescriptionAttribute)attribs[0]).Description : "");
			}

			// Layout
			logo.Append(title).Append(" v").Append(version).Append(nl);

			if (copyright.Length != 0)
			{
				logo.Append(copyright).Append(nl);
			}

			if (!string.IsNullOrEmpty(description))
			{
				logo.Append(description).Append(nl);
			}

			return logo.ToString();
		}

		public enum ImageSubsystem
		{
			Native = 0x0001, // Currently not supported
			GUI = 0x0002, // Windows EXE or DLL
			CUI = 0x0003, // Console EXE
		}

		// This code heavily inspired by Brent Rector [brent@wiseowl.com]. Thanks, Brent!
		public static ImageSubsystem GetImageSubsystem(string filename)
		{
			using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read))
			using (BinaryReader reader = new BinaryReader(fs))
			{
				// Seek to the beginning of the PE signature offset
				reader.BaseStream.Seek(0x3c, SeekOrigin.Begin);

				// Get the offset to the PE signature
				uint e_lfanew = reader.ReadUInt32();

				// Seek to the beginning of the PE Header
				reader.BaseStream.Seek(e_lfanew, SeekOrigin.Begin);

				// Read the PE signature
				uint PE_SIGNATURE = 0x00004550;
				uint PESignature = reader.ReadUInt32();
				if (PESignature != PE_SIGNATURE) throw new BadImageFormatException("Bad PE signature: " + filename);

				// Seek past the file header
				reader.BaseStream.Seek(0x14, SeekOrigin.Current);

				// Seek to the subsystem in the optional header
				reader.BaseStream.Seek(0x44, SeekOrigin.Current);

				switch (reader.ReadUInt16())
				{
					case (ushort)ImageSubsystem.Native:
						return ImageSubsystem.Native;
					case (ushort)ImageSubsystem.GUI:
						return ImageSubsystem.GUI;
					case (ushort)ImageSubsystem.CUI:
						return ImageSubsystem.CUI;
					default:
						throw new BadImageFormatException("Invalid subsystem: " + filename);
				}
			}
		}

		#region P/Invoke implemention of IsConsole

		/*
		struct SHFILEINFO
		{
			UIntPtr hIcon;          // out: icon
			int     iIcon;          // out: icon index
			uint    dwAttributes;   // out: SFGAO_ flags
			[ MarshalAs( UnmanagedType.ByValArray, SizeConst=260 )]
			char[]  szDisplayName;  // out: display name (or path)
			[ MarshalAs( UnmanagedType.ByValArray, SizeConst=80 )]
			char[]  szTypeName;     // out: type name
		}

		[DllImport("shell32.dll")]
		static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, out SHFILEINFO psfi, uint cbFileInfo, uint uFlags);

		static ushort LOWORD(uint u)
		{
			return ((ushort)(u & 0xffff));
		}

		static ushort HIWORD(uint u)
		{
			return ((ushort)(u >> 16));
		}

		public static bool IsConsole()
		{
			SHFILEINFO  sfi = new SHFILEINFO();
			uint        SHGFI_EXETYPE = 0x000002000;
			uint        IMAGE_NT_SIGNATURE = 0x00004550;    // NOTE: Not on the Mac, though...
			uint        exeType = (uint)SHGetFileInfo(Assembly.GetEntryAssembly().Location, 0, out sfi, (uint)Marshal.SizeOf(typeof(SHFILEINFO)), SHGFI_EXETYPE);
			return (LOWORD(exeType) == IMAGE_NT_SIGNATURE) && (HIWORD(exeType) == 0);
		}
		*/

		#endregion

		public virtual string GetModuleName()
		{
			return Path.GetFileName(GetAssembly().Location);
		}

		public virtual string GetTitle()
		{
			return GetTitle(GetAssembly());
		}

		private string GetTitle(Assembly assem)
		{
			// Title: try AssemblyTitle first, use module name if AssemblyTitle is missing
			object[] attribs = assem.GetCustomAttributes(typeof(AssemblyTitleAttribute), true);
			string title = (attribs.Length != 0 ? ((AssemblyTitleAttribute)attribs[0]).Title : string.Empty);
			if (string.IsNullOrEmpty(title))
				title = GetModuleName();
			return title;
		}

		private static Assembly GetAssembly()
		{
			return Assembly.GetEntryAssembly() ?? Assembly.GetExecutingAssembly();
		}

		private void Show(string s, ConsoleColor color)
		{
			// Message in proporcional output, look align.

			// Always send usage to stdout so it's easy to capture the output
			Console.ForegroundColor = color;
			Console.WriteLine(s);
			Console.ResetColor();
		}

		private bool Continue(string err)
		{
			if (version)
			{
				Show(GetLogo(), ConsoleColor.Cyan);
				return false;
			}
			if (help)
			{
				Show(GetUsage(""), ConsoleColor.Magenta);
				return false;
			}
			if ((err.Length > 0) || help)
			{
				Show(GetUsage(help ? "" : err), ConsoleColor.Red);
				return false;
			}

			return true;
		}

		public bool ParseAndContinue(string[] args)
		{
			string err = "";

			try
			{
				object[] attribs = GetType().GetCustomAttributes(typeof(ParserUsageAttribute), true);
				ParserUsageAttribute parser = (ParserUsageAttribute)(attribs.Length != 0 ? attribs[0] : null);

				//TODO relying on exception being generated last (i.e. after values have been read and assigned) should pass param  
				//Load defaults from the environmental variable
				if (parser != null)
					if ("" != parser.EnvironmentDefaults) try
						{
							Parse(Environment.GetEnvironmentVariable(parser.EnvironmentDefaults));
						}
						catch
						{
						}
				//--------------------------------------------------------

				//Parse the command line
				Parse(args);
			}
			catch (Exception e)
			{
				err = e.Message;
			}
			return Continue(err);
		}

		[FlagUsage("Show usage.", AlternateName1 = "?", AlternateName2 = "h", AllowOnOff = false, Category = "HELP"),]
		public bool help;

		[FlagUsage("Show version.", AlternateName1 = "v", AllowOnOff = false, Category = "HELP")]
		public bool version;
	}
}
