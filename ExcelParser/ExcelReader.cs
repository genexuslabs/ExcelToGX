using Antlr4.StringTemplate;
using Artech.Common.Helpers.Guids;
using ExcelParser.DataTypes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;

namespace ExcelParser
{
	public abstract class ExcelReader
	{
		public abstract void ReadExcel(string[] files, string outputFile, bool continueOnErrors);
	}

	public abstract class ExcelReader<TConfig, IType, TLeafElement, TLevelElement> : ExcelReader
        where TConfig : ExcelReader<TConfig, IType, TLeafElement, TLevelElement>.BaseConfiguration, new()
		where IType : IKBElement
		where TLeafElement : DataTypeElement, IType, new()
		where TLevelElement : LevelElement<IType>, IType, new()
	{
		private string CurrentFileName;
		private bool ContinueOnErrors;

		Dictionary<string, TLevelElement> Levels;
		Dictionary<string, TLeafElement> Leafs;
		Dictionary<string, DataTypeElement> Domains;


		public abstract class BaseConfiguration
		{
			public int ObjectNameRow;
			public int ObjectNameColumn;
			public int ObjectDescColumn;
			public int ObjectDescRow;
			public string DefinitionSheetName;

			public int LevelCheckColumn = 3;
			public int LevelIdColumn = 8;
			public int LevelParentIdColumn = 9;
			public string LevelIdentifierKeyword = "LVL";

			public int DataStartRow = 7;
			public int DataStartColumn = 2;
			public int DataNameColumn = 7;
			public int DataDescriptionColumn = 6;
			public int DataTypeColumn = 8;
			public int DataLengthColumn = 9;
			public int DomainColumn = 9;
		}

		public TConfig Configuration { get; } = new TConfig();
		protected abstract string TemplateFile { get; }
		protected abstract string TemplateRender { get; }

		public override void ReadExcel(string[] files, string outputFile, bool continueOnErrors)
		{
			ContinueOnErrors = continueOnErrors;

			Levels = new Dictionary<string, TLevelElement>();
			Leafs = new Dictionary<string, TLeafElement>();
			Domains = new Dictionary<string, DataTypeElement>(StringComparer.OrdinalIgnoreCase);

			foreach (string fileName in files)
			{
				CurrentFileName = fileName;
				try
				{
					var excel = new ExcelPackage(new FileInfo(fileName));
					ExcelWorksheet sheet = excel.Workbook.Worksheets[Configuration.DefinitionSheetName];
					if (sheet == null)
						throw new Exception($"The xlsx file '{fileName}' was opened but it does not have a definition sheet for the object. Please provide the object definition in a Sheet called '{Configuration.DefinitionSheetName}'");

					string objName = sheet.Cells[Configuration.ObjectNameRow, Configuration.ObjectNameColumn].Value?.ToString();
					if (objName == null)
						throw new Exception($"Could not find the object name at [{Configuration.ObjectNameRow} , {Configuration.ObjectNameColumn}]. Please take a look at the configuration file");
					ProcessFile(sheet, objName);
				}
				catch (Exception ex) when (HandleException(fileName, ex))
				{
				}
			}
			CurrentFileName = null;

			Console.WriteLine($"Generating Export xml to {outputFile}");
			TemplateGroup grp = new TemplateGroupFile(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), TemplateFile));
			Template component_tpl = grp.GetInstanceOf(TemplateRender);

			component_tpl.Add("levels", Levels.Values);
			component_tpl.Add("leafs", Leafs.Values);
			if (Domains.Values.Count > 0)
				component_tpl.Add("domains", Domains.Values);

			using (var stream = new FileStream(outputFile, FileMode.Create, FileAccess.ReadWrite))
			using (var writer = new StreamWriter(stream))
			{
				ITemplateWriter wr = new AutoIndentWriter(writer);
				wr.LineWidth = AutoIndentWriter.NoWrap;
				component_tpl.Write(wr);
			}

			Console.WriteLine($"Success");
		}

		private void ProcessFile(ExcelWorksheet sheet, string objName)
		{
			Dictionary<int, TLevelElement> levels = new Dictionary<int, TLevelElement>();
			string objDescription = sheet.Cells[Configuration.ObjectDescRow, Configuration.ObjectDescColumn].Value?.ToString();
			TLevelElement level = new TLevelElement
			{
				Name = objName,
				Guid = GuidHelper.Create(GuidHelper.IsoOidNamespace, objName, false).ToString(),
				Description = objDescription
			};
			levels[0] = level;
			Console.WriteLine($"Processing object with name '{objName}' and description '{objDescription}'");
			Levels[objName] = level;

			int row = Configuration.DataStartRow;
			int atts = 0;
			while (sheet.Cells[row, Configuration.DataStartColumn].Value != null)
			{
				level = ReadLevel(row, sheet, level, levels, out bool isLevel);
				if (!isLevel)
				{
					try
					{
						TLeafElement leaf = ReadLeaf(sheet, row);
						if (leaf.Domain != null && leaf.Type != null && !Domains.ContainsKey(leaf.Domain))
						{
							DataTypeElement domain = new DataTypeElement
							{
								Name = leaf.Domain,
								Guid = GuidHelper.Create(GuidHelper.UrlNamespace, leaf.Domain, false).ToString()
							};
							DataTypeManager.SetDataType(leaf.Type, domain);
							Domains[domain.Name] = domain;
						}
						if (leaf != null)
						{
							if (Leafs.ContainsKey(leaf.Name) && leaf.ToString() != Leafs[leaf.Name].ToString())
								Console.WriteLine($"{leaf.Name} was already defined with a different data type {Leafs[leaf.Name].ToString()}, taking into account the last one {leaf.ToString()}");
							Leafs[leaf.Name] = leaf;
							atts++;
							level.Items.Add(leaf);
							Console.WriteLine($"Processing Attribute {leaf.Name}");
							row++;
						}
						else
							break;
					}
					catch (Exception ex) when (HandleException($"at row:{row}", ex, true))
					{
					}
				}
				else
				{
					ReadLevelProperties(level, sheet, row);
					row++;
				}
			}
			if (atts == 0)
			{
				throw new Exception($"Definition without content, check the {nameof(Configuration.DataStartRow)} and {nameof(Configuration.DataStartColumn)} values on the config file");
			}
		}

		private TLevelElement ReadLevel(int row, ExcelWorksheet sheet, TLevelElement level, Dictionary<int, TLevelElement> levels, out bool isLevel)
		{
			isLevel = false;
			if (sheet.Cells[row, Configuration.LevelCheckColumn].Value?.ToString().ToLower().Trim() == Configuration.LevelIdentifierKeyword.ToLower().Trim()) //is a level
			{
				isLevel = true;
				int levelId = 0;
				int parentId = 0;
				// Read Level Id
				try
				{
					if (sheet.Cells[row, Configuration.LevelIdColumn].Value != null)
					{
						levelId = int.Parse(sheet.Cells[row, Configuration.LevelIdColumn].Value?.ToString());
					}
					parentId = GetParentLevelId(sheet, row);
				}
				catch (Exception ex)
				{
					throw new Exception($"Invalid identifier for level at {row}, {Configuration.LevelIdColumn} " + ex.Message, ex);
				}
				string levelName = sheet.Cells[row, Configuration.DataNameColumn].Value?.ToString();
				if (levelName == null)
					throw new Exception($"Could not find the Level name at [{row} , {Configuration.DataNameColumn}], please take a look at the configuration file ");
				string levelDesc = sheet.Cells[row, Configuration.DataDescriptionColumn].Value?.ToString();
				TLevelElement newLevel = new TLevelElement
				{
					Name = levelName,
					Guid = GuidHelper.Create(GuidHelper.IsoOidNamespace, levelName, false).ToString(),
					Description = levelDesc
				};
				Debug.Assert(levelId >= 1);
				levels[levelId] = newLevel;

				TLevelElement parentTrn = levels[parentId];
				parentTrn.Items.Add(newLevel);
				return newLevel;
			}
			else
			{
				if (sheet.Cells[row, Configuration.LevelParentIdColumn].Value != null) // Explicit level id
				{
					int parentLevel = GetParentLevelId(sheet, row);
					if (parentLevel >= 0 && levels.ContainsKey(parentLevel))
						return levels[parentLevel];
				}
				return level; //continue in the same level, there is no level at this row, so assume the same.
			}
		}

		protected TLeafElement ReadLeaf(ExcelWorksheet sheet, int row)
		{
			if (string.IsNullOrEmpty(sheet.Cells[row, Configuration.DataNameColumn].Value?.ToString().Trim()))
				return null;
			TLeafElement att = new TLeafElement()
			{
				Name = sheet.Cells[row, Configuration.DataNameColumn].Value?.ToString().Trim(),
				Domain = sheet.Cells[row, Configuration.DomainColumn].Value?.ToString().Trim(),
				Description = sheet.Cells[row, Configuration.DataDescriptionColumn].Value?.ToString().Trim(),
				Type = sheet.Cells[row, Configuration.DataTypeColumn].Value?.ToString().Trim().ToLower(),
			};
			ReadLeafProperties(att, sheet, row);
			att.Guid = GuidHelper.Create(GuidHelper.DnsNamespace, att.Name, false).ToString();
			try
			{
				if (att.Domain == null)
				{
					DataTypeManager.SetDataType(att.Type, att);
					try
					{
						SetLengthAndDecimals(sheet, row, att);
					}
					catch (Exception ex) //never fail because a wrong length/decimals definition, just use the defaults values.
					{
						Console.WriteLine("Error parsing Data Type for row " + row);
						HandleException(sheet.Name, ex);
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine("Error parsing Data Type for row " + row);
				HandleException(sheet.Name, ex);
			}
			return att;
		}

		private void SetLengthAndDecimals(ExcelWorksheet sheet, int row, TLeafElement leaf)
		{
			string lenAndDecimals = sheet.Cells[row, Configuration.DataLengthColumn].Value?.ToString().Trim();
			if (!string.IsNullOrEmpty(lenAndDecimals))
			{
				if (lenAndDecimals.Length > 0 && lenAndDecimals.EndsWith("-"))
				{
					leaf.Sign = true;
					lenAndDecimals = lenAndDecimals.Replace("-", "");
				}
				string[] splitedData;
				if (lenAndDecimals.Contains("."))
					splitedData = lenAndDecimals.Trim().Split('.');
				else
					splitedData = lenAndDecimals.Trim().Split(',');
				if (splitedData.Length >= 1 && int.TryParse(splitedData[0], out int length))
				{
					leaf.Length = length;
				}
				if (splitedData.Length == 2 && int.TryParse(splitedData[1], out int decimals))
				{
					leaf.Decimals = decimals;
				}
			}
		}

		protected abstract void ReadLevelProperties(TLevelElement level, ExcelWorksheet sheet, int row);

		protected abstract void ReadLeafProperties(TLeafElement leaf, ExcelWorksheet sheet, int row);

		private int GetParentLevelId(ExcelWorksheet sheet, int row)
		{
			if (sheet.Cells[row, Configuration.LevelParentIdColumn].Value != null)
			{
				try
				{
					if (int.TryParse(sheet.Cells[row, Configuration.LevelParentIdColumn].Value?.ToString(), out int parentId))
						return parentId;
				}
				catch
				{
				}
			}
			return 0;
		}

		protected bool HandleException(string message, Exception ex, bool showFileName = false)
		{
			if (!ContinueOnErrors)
			{
				Console.WriteLine($"Error: processing {(showFileName ? CurrentFileName + " " + message : message)}, stop processing because continueOnErrors = false");
				return false;
			}
			else
			{
				Console.WriteLine($"Error: processing {(showFileName ? CurrentFileName + " " + message : message)}, continue because continueOnErrors = true");
				Console.WriteLine(ex.Message);
				return true;
			}
		}
    }
}
