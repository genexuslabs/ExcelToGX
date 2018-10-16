using Antlr4.StringTemplate;
using Artech.Common.Helpers.Guids;
using ExcelToTransactions.DataTypes;
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
			public static int LevelIdColumn = 8;
			public static int LevelParentIdColumn = 9;
			public static string LevelIdentifierKeyword = "LVL";
			public static int TransactionNameRow = 3;
			public static int TransactionNameCol = 7;
			public static int TransactionDescColumn = 11;
			public static int TransactionDescRow = 3;
			public static int AttributeStartRow = 7;
			public static int AttributesStartColumn = 2;
			public static int AttributeNameColumn = 7;
			public static int AttributeDomainColumn = 9;
			public static int AttributeDescriptionColumn = 6;
			public static int AttributeNullableColumn = 4;
			public static int AttributeKeyColumn = 3;
			public static int AttributeDataTypeColumn = 8;
			public static int AttributeDataLengthColumn = 9;
			public static string TransactionDefinitionSheetName = "TransactionDefinitionSheet";
			public static int AttributeAutonumberColumn = 12;
			public static string PKValue = "PK";
			public static string NullableValue = "?";
		}
		public static void ReadExcel(string[] files, string outputFile, bool continueOnErrors)
		{
			Dictionary<string, TransactionLevel> transactions = new Dictionary<string, TransactionLevel>();
			Dictionary<string, TransactionAttribute> attributes = new Dictionary<string, TransactionAttribute>();
			Dictionary<string, TransactionAttribute> domains = new Dictionary<string, TransactionAttribute>();

			foreach (string fileName in files)
			{
				try
				{
					ProcessFile(fileName, transactions, attributes, domains, continueOnErrors);
				}
				catch (Exception ex)
				{
					HandleException(fileName, ex, continueOnErrors);
				}
			}
			Console.WriteLine($"Generating Export xml to {outputFile}");
			TemplateGroup grp = new TemplateGroupString(File.ReadAllText(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "ExportTemplate.stg")));
			Template component_tpl = grp.GetInstanceOf("g_transaction_render");
			component_tpl.Add("components", transactions.Values);
			component_tpl.Add("attributes", attributes.Values);
			if (domains.Values != null)
				component_tpl.Add("domains", domains.Values);
			string xmlExport = component_tpl.Render();
			File.WriteAllText(outputFile, xmlExport);
			Console.WriteLine($"Success");
		}

		private static void HandleException(string fileName, Exception ex, bool continueOnErrors)
		{
			if (!continueOnErrors)
			{
				Console.WriteLine($"Error: processing {fileName}, stop processing because continueOnErrors = false");
				throw ex;
			}
			else
			{
				Console.WriteLine($"Error: processing {fileName}, continue because continueOnErrors = true");
				Console.WriteLine(ex.Message);
			}
		}

		private static void ProcessFile(string fileName, Dictionary<string, TransactionLevel> transactions, Dictionary<string, TransactionAttribute> attributes, Dictionary<string, TransactionAttribute> domains, bool continueOnError)
		{
			Dictionary<int, TransactionLevel> levels = new Dictionary<int, TransactionLevel>();
			var excel = new ExcelPackage(new System.IO.FileInfo(fileName));
			ExcelWorksheet sheet = excel.Workbook.Worksheets[Configuration.TransactionDefinitionSheetName];
			if (sheet == null)
				throw new Exception($"The xlsx file {fileName} was open but it have not a definition sheet for the Transaction. Please provide the transaction definition in a Sheet called {Configuration.TransactionDefinitionSheetName}");

			string trnName = sheet.Cells[Configuration.TransactionNameRow, Configuration.TransactionNameCol].Value?.ToString();
			if (trnName == null)
				throw new Exception($"Could not find the Transaction name at [{Configuration.TransactionNameRow} , {Configuration.TransactionNameCol}], please take a look at the configuration file ");
			string trnDescription = sheet.Cells[Configuration.TransactionDescRow, Configuration.TransactionDescColumn].Value?.ToString();
			TransactionLevel level = new TransactionLevel
			{
				Name = trnName,
				Guid = GuidHelper.Create(GuidHelper.IsoOidNamespace, trnName, false).ToString(),
				Description = trnDescription
			};
			levels[0] = level;
			Console.WriteLine($"Processing Transaction {trnName} with Description {trnDescription}");
			transactions[trnName] = level;

			int row = Configuration.AttributeStartRow;
			int atts = 0;
			while (sheet.Cells[row, Configuration.AttributesStartColumn].Value != null)
			{
				level = ReadLevel(row, sheet, level, levels, out bool isLevel);
				if (!isLevel)
				{
					try
					{
						TransactionAttribute att = ReadAttribute(sheet, row, continueOnError);
						if (att.Domain != null && att.Type != null && !domains.ContainsKey(att.Domain))
						{
							TransactionAttribute domain = new TransactionAttribute
							{
								Name = att.Domain,
								Guid = GuidHelper.Create(GuidHelper.UrlNamespace, att.Domain, false).ToString()
							};
							DataTypeManager.SetDataType(att.Type, domain);
							domains[domain.Name.ToLower()] = domain;
						}
						if (att != null)
						{
							if (attributes.ContainsKey(att.Name) && att.ToString() != attributes[att.Name].ToString())
								Console.WriteLine($"{att.Name} was already defined with a different data type {attributes[att.Name].ToString()}, taking into account the last one {att.ToString()}");
							attributes[att.Name] = att;
							atts++;
							level.Items.Add(att);
							Console.WriteLine($"Processing Attribute {att.Name}");
							row++;
						}
						else
							break;
					}
					catch (Exception ex)
					{
						HandleException(fileName + $" at row:{row}", ex, continueOnError);
					}
				}
				else
					row++;
			}
			if (atts == 0)
			{
				throw new Exception("Transaction without attributes, check the AttributeStartRow and AttributeStartColumn values on the config file");
			}

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
					parentId = GetParentLevelId(sheet, row);
				}
				catch (Exception ex)
				{
					throw new Exception($"Invalid identifier for level at {row}, {Configuration.LevelIdColumn} " + ex.Message, ex);
				}
				string levelName = sheet.Cells[row, Configuration.AttributeNameColumn].Value?.ToString();
				if (levelName == null)
					throw new Exception($"Could not find the Level name at [{row} , {Configuration.AttributeNameColumn}], please take a look at the configuration file ");
				string levelDesc = sheet.Cells[row, Configuration.AttributeDescriptionColumn].Value?.ToString();
				TransactionLevel newLevel = new TransactionLevel
				{
					Name = levelName,
					Guid = GuidHelper.Create(GuidHelper.IsoOidNamespace, levelName, false).ToString(),
					Description = levelDesc
				};
				Debug.Assert(levelId >= 1);
				levels[levelId] = newLevel;

				TransactionLevel parentTrn = levels[parentId];
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


		

		private static int GetParentLevelId(ExcelWorksheet sheet, int row)
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

		private static TransactionAttribute ReadAttribute(ExcelWorksheet sheet, int row, bool contineOnErrors)
		{
			if (String.IsNullOrEmpty(sheet.Cells[row, Configuration.AttributeNameColumn].Value?.ToString().Trim()))
				return null;
			TransactionAttribute att = new TransactionAttribute
			{
				Name = sheet.Cells[row, Configuration.AttributeNameColumn].Value?.ToString().Trim(),
				Domain = sheet.Cells[row, Configuration.AttributeDomainColumn].Value?.ToString().Trim(),
				Description = sheet.Cells[row, Configuration.AttributeDescriptionColumn].Value?.ToString().Trim(),
				AllowNull = sheet.Cells[row, Configuration.AttributeNullableColumn].Value?.ToString().ToLower() == Configuration.NullableValue.ToLower(),
				IsKey = sheet.Cells[row, Configuration.AttributeKeyColumn].Value?.ToString().ToLower() == Configuration.PKValue.ToLower(),
				Type = sheet.Cells[row, Configuration.AttributeDataTypeColumn].Value?.ToString().Trim().ToLower(),
				Autonumber = sheet.Cells[row, Configuration.AttributeAutonumberColumn].Value?.ToString().ToLower() == "true"
			};
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
                        HandleException(sheet.Name, ex, contineOnErrors);
                    }
                }
			}
			catch (Exception ex)
			{
                Console.WriteLine("Error parsing Data Type for row " + row);
                HandleException(sheet.Name, ex, contineOnErrors);
			}
			return att;
		}

        private static void SetLengthAndDecimals(ExcelWorksheet sheet, int row, TransactionAttribute att)
        {
            string lenAndDecimals = sheet.Cells[row, Configuration.AttributeDataLengthColumn].Value?.ToString().Trim();
            if (!String.IsNullOrEmpty(lenAndDecimals))
            {
                string[] splitedData;
                if (lenAndDecimals.Contains("."))
                    splitedData = lenAndDecimals.Trim().Split('.');
                else
                    splitedData = lenAndDecimals.Trim().Split(',');
                int length, decimals;
                if (splitedData.Length >= 1 && int.TryParse(splitedData[0], out length))
                {
                    att.Length = length;
                }
                if (splitedData.Length == 2 && int.TryParse(splitedData[1], out decimals))
                {
                    att.Decimals = decimals;
                }
            }
        }
    }
}
