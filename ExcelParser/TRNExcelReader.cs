using Antlr4.StringTemplate;
using Artech.Common.Helpers.Guids;
using ExcelParser.DataTypes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelParser
{
	public class TRNExcelReader : ExcelReader<TRNExcelReader.TRNConfiguration>
	{
		Dictionary<string, TransactionLevel> Transactions;
		Dictionary<string, TransactionAttribute> Attributes;
		Dictionary<string, TransactionAttribute> Domains;

		public class TRNConfiguration : BaseConfiguration
		{
			public TRNConfiguration()
			{
				DefinitionSheetName = "TransactionDefinitionSheet";
			}
		}

		protected override void Initialize()
		{
			Transactions = new Dictionary<string, TransactionLevel>();
			Attributes = new Dictionary<string, TransactionAttribute>();
			Domains = new Dictionary<string, TransactionAttribute>();
		}

		protected override string TemplateFile => "ExportTRNTemplate.stg";
		protected override string TemplateRender => "g_transaction_render";

		protected override void SetTemplateProperties(Template component_tpl)
		{
			base.SetTemplateProperties(component_tpl);
			component_tpl.Add("components", Transactions.Values);
			component_tpl.Add("attributes", Attributes.Values);
			if (Domains.Values.Count > 0)
				component_tpl.Add("domains", Domains.Values);
		}

		protected override void ProcessFile(ExcelWorksheet sheet, string objName)
		{
			Dictionary<int, TransactionLevel> levels = new Dictionary<int, TransactionLevel>();
			string trnDescription = sheet.Cells[Configuration.ObjectDescRow, Configuration.ObjectDescColumn].Value?.ToString();
			TransactionLevel level = new TransactionLevel
			{
				Name = objName,
				Guid = GuidHelper.Create(GuidHelper.IsoOidNamespace, objName, false).ToString(),
				Description = trnDescription
			};
			levels[0] = level;
			Console.WriteLine($"Processing Transaction {objName} with Description {trnDescription}");
			Transactions[objName] = level;

			int row = Configuration.AttributeStartRow;
			int atts = 0;
			while (sheet.Cells[row, Configuration.AttributesStartColumn].Value != null)
			{
				level = ReadLevel(row, sheet, level, levels, out bool isLevel);
				if (!isLevel)
				{
					try
					{
						TransactionAttribute att = ReadAttribute(sheet, row);
						if (att.Domain != null && att.Type != null && !Domains.ContainsKey(att.Domain))
						{
							TransactionAttribute domain = new TransactionAttribute
							{
								Name = att.Domain,
								Guid = GuidHelper.Create(GuidHelper.UrlNamespace, att.Domain, false).ToString()
							};
							DataTypeManager.SetDataType(att.Type, domain);
							Domains[domain.Name.ToLower()] = domain;
						}
						if (att != null)
						{
							if (Attributes.ContainsKey(att.Name) && att.ToString() != Attributes[att.Name].ToString())
								Console.WriteLine($"{att.Name} was already defined with a different data type {Attributes[att.Name].ToString()}, taking into account the last one {att.ToString()}");
							Attributes[att.Name] = att;
							atts++;
							level.Items.Add(att);
							Console.WriteLine($"Processing Attribute {att.Name}");
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
					row++;
			}
			if (atts == 0)
			{
				throw new Exception("Transaction without attributes, check the AttributeStartRow and AttributeStartColumn values on the config file");
			}

		}

		private TransactionLevel ReadLevel(int row, ExcelWorksheet sheet, TransactionLevel level, Dictionary<int, TransactionLevel> levels, out bool isLevel)
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

		private TransactionAttribute ReadAttribute(ExcelWorksheet sheet, int row)
		{
			if (string.IsNullOrEmpty(sheet.Cells[row, Configuration.AttributeNameColumn].Value?.ToString().Trim()))
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

		private void SetLengthAndDecimals(ExcelWorksheet sheet, int row, TransactionAttribute att)
		{
			string lenAndDecimals = sheet.Cells[row, Configuration.AttributeDataLengthColumn].Value?.ToString().Trim();
			if (!string.IsNullOrEmpty(lenAndDecimals))
			{
				if (lenAndDecimals.Length > 0 && lenAndDecimals.EndsWith("-"))
				{
					att.Sign = true;
					lenAndDecimals = lenAndDecimals.Replace("-", "");
				}
				string[] splitedData;
				if (lenAndDecimals.Contains("."))
					splitedData = lenAndDecimals.Trim().Split('.');
				else
					splitedData = lenAndDecimals.Trim().Split(',');
				if (splitedData.Length >= 1 && int.TryParse(splitedData[0], out int length))
				{
					att.Length = length;
				}
				if (splitedData.Length == 2 && int.TryParse(splitedData[1], out int decimals))
				{
					att.Decimals = decimals;
				}
			}
		}
	}
}
