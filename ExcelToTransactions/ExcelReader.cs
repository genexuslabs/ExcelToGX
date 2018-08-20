using Antlr4.StringTemplate;
using Artech.Common.Helpers.Guids;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToTransactions
{
	public class ExcelReader
	{
		public class Configuration
		{
			public static int LevelCheckColumn = 3;
			public static int LevelIdColumn = 4;
			public static int LevelParentIdColumn = 5;
			public static string LevelIdentifierKeyword = "LVL";
			public static int TransactionNameRow = 3;
			public static int TransactionNameCol = 7;
			public static int TransactionDescColumn = 11;
			public static int TransactionDescRow = 3;
			public static int AttributeStartRow = 7;
			public static int AttributesStartColumn = 2;
			public static int AttributeNameColumn = 7;
			public static int AttributeDescriptionColumn = 6;
			public static int AttributeNullableColumn = 4;
			public static int AttributeKeyColumn = 3;
			public static int AttributeDataTypeColumn = 8;
			public static int AttributeDataLengthColumn = 9;
			public static string TransactionDefinitionSheetName = "TransactionDefinitionSheet";
			public static int AttributeAutonumberColumn = 12;
		}
		public static void ReadExcel(string fileName, string outputFile)
		{
			Dictionary<int, TransactionLevel> levels = new Dictionary<int, TransactionLevel>();
			var excel = new ExcelPackage(new System.IO.FileInfo(fileName));
			ExcelWorksheet sheet = excel.Workbook.Worksheets[Configuration.TransactionDefinitionSheetName];
			if (sheet == null)
				throw new Exception($"Please provide the transaction definition in a Sheet called {Configuration.TransactionDefinitionSheetName}");

			string trnName = sheet.Cells[Configuration.TransactionNameRow, Configuration.TransactionNameCol].Value?.ToString();
			if (trnName == null)
				throw new Exception($"Could not find the Transaction name at [{Configuration.TransactionNameRow} , {Configuration.TransactionNameCol}], please take a look at the configuration file ");
			string trnDescription = sheet.Cells[Configuration.TransactionDescRow, Configuration.TransactionDescColumn].Value?.ToString();
			List<TransactionAttribute> attributes = new List<TransactionAttribute>();
			TransactionLevel level = new TransactionLevel
			{
				Name = trnName,
				Guid = GuidHelper.Create(GuidHelper.IsoOidNamespace, trnName, false).ToString(),
				Description = trnDescription
			};
			levels[0] = level;
			Console.WriteLine($"Processing Transaction {trnName} with Description {trnDescription}");

			int row = Configuration.AttributeStartRow;
			while (sheet.Cells[row, Configuration.AttributesStartColumn].Value != null)
			{
				level = ReadLevel(row, sheet, level, levels, out bool  isLevel);
				if (!isLevel)
				{
					TransactionAttribute att = ReadAttribute(sheet, row);
					if (att != null)
					{
						attributes.Add(att);
						level.Attributes.Add(att);
						Console.WriteLine($"Processing Attribute {att.Name}");
						row++;
					}
					else
						break;
				}
				else
					row++;
			}
			if (level.Attributes.Count == 0)
			{
				throw new Exception("Transaction without attributes, check the AttributeStartRow and AttributeStartColumn values on the config file");
			}
			Console.WriteLine($"Generating Export xml to {outputFile}");
			TemplateGroup grp = new TemplateGroupString(File.ReadAllText(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "ExportTemplate.stg")));
			Template component_tpl = grp.GetInstanceOf("g_transaction_render");
			component_tpl.Add("component", levels[0]);
			component_tpl.Add("attributes", attributes);
			string xmlExport = component_tpl.Render();
			File.WriteAllText(outputFile, xmlExport);
			Console.WriteLine($"Success");
		}

		private static TransactionLevel ReadLevel(int row, ExcelWorksheet sheet, TransactionLevel level, Dictionary<int, TransactionLevel> levels, out bool  isLevel)
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
					if (sheet.Cells[row, Configuration.LevelParentIdColumn].Value != null)
					{
						try
						{
							parentId = int.Parse(sheet.Cells[row, Configuration.LevelParentIdColumn].Value?.ToString());
						}
						catch
						{
							parentId = 0;
						}
					}
				}
				catch (Exception ex)
				{
					throw new Exception($"Invalid identifier for level at {row}, {Configuration.LevelIdColumn} " + ex.Message, ex);
				}
				string levelName = sheet.Cells[row, Configuration.TransactionNameCol].Value?.ToString();
				if (levelName == null)
					throw new Exception($"Could not find the Level name at [{row} , {Configuration.TransactionNameCol}], please take a look at the configuration file ");
				string levelDesc = sheet.Cells[Configuration.TransactionDescRow, Configuration.TransactionDescColumn].Value?.ToString();
				TransactionLevel newLevel = new TransactionLevel
				{
					Name = levelName,
					Guid = GuidHelper.Create(GuidHelper.IsoOidNamespace, levelName, false).ToString(),
					Description = levelDesc
				};
				Debug.Assert(levelId >= 1);
				levels[levelId] = newLevel;

				TransactionLevel parentTrn = levels[parentId];
				parentTrn.Levels.Add(newLevel);
				return newLevel;
			}
			else
			{
				return level; //continue in the same level, there is no level at this row.
			}
		}

		private static TransactionAttribute ReadAttribute(ExcelWorksheet sheet, int row)
		{
			if (String.IsNullOrEmpty(sheet.Cells[row, Configuration.AttributeNameColumn].Value?.ToString().Trim()))
				return null;
			TransactionAttribute att = new TransactionAttribute
			{
				Name = sheet.Cells[row, Configuration.AttributeNameColumn].Value?.ToString().Trim(),
				Description = sheet.Cells[row, Configuration.AttributeDescriptionColumn].Value?.ToString().Trim(),
				AllowNull = sheet.Cells[row, Configuration.AttributeNullableColumn].Value != null,
				IsKey = sheet.Cells[row, Configuration.AttributeKeyColumn].Value?.ToString().ToLower() == "pk",
				Type = sheet.Cells[row, Configuration.AttributeDataTypeColumn].Value?.ToString().Trim().ToLower(),
				Autonumber = sheet.Cells[row, Configuration.AttributeAutonumberColumn].Value?.ToString().ToLower() == "true"
				
			};
			att.Guid = GuidHelper.Create(GuidHelper.DnsNamespace, att.Name, false).ToString();
			string lenAndDecimals = sheet.Cells[row, Configuration.AttributeDataLengthColumn].Value?.ToString().Trim();
			if (lenAndDecimals != null)
			{
				string[] splitedData;
				if (lenAndDecimals.Contains("."))
					splitedData = lenAndDecimals.Trim().Split('.');
				else
					splitedData = lenAndDecimals.Trim().Split(',');

				att.Decimals = 0;
				if (splitedData.Length >= 1)
					att.Length = int.Parse(splitedData[0]);
				if (splitedData.Length == 2)
					att.Decimals = int.Parse(splitedData[1]);
			}
			else
			{
				att.Decimals = null;
				att.Length = null;
			}
			return att;
		}
	}
}
