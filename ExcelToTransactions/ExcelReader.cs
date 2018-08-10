using Antlr4.StringTemplate;
using Artech.Common.Helpers.Guids;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
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
			public static int TransactionNameRow = 3;
			public static int TransactionNameCol = 7;
			public static int TransactionDescColumn = 11;

			public static int AttributeStartRow = 7;
			public static int AttributesStartColumn = 2;

			public static int AttributeNameColumn = 7;
			public static int AttributeDescriptionColumn = 6;
			public static int AttributeNullableColumn = 4;
			public static int AttributeKeyColumn = 3;
			public static int AttributeDataTypeColumn = 8;
			public static int AttributeDataLengthColumn = 9;

			public static string TransactionDefinitionSheetName = "TransactionDefinitionSheet";
		}
		public static void ReadExcel(string fileName, string outputFile)
		{
			var excel = new ExcelPackage(new System.IO.FileInfo(fileName));
			ExcelWorksheet sheet = excel.Workbook.Worksheets[Configuration.TransactionDefinitionSheetName];
			if (sheet == null)
				throw new Exception("Please provide the transaction definition in a Sheet called TransactionDefinitionSheet");

			string trnName = sheet.Cells[Configuration.TransactionNameRow, Configuration.TransactionNameCol].Value?.ToString();
			if (trnName == null)
				throw new Exception($"Could not find the Transaction name at [{Configuration.TransactionNameRow} , {Configuration.TransactionNameCol}], please take a look at the configuration file ");
			string trnDescription = sheet.Cells[Configuration.TransactionNameRow, Configuration.TransactionDescColumn].Value?.ToString();
			TransactionLevel level = new TransactionLevel
			{
				Name = trnName,
				Guid = GuidHelper.Create(GuidHelper.IsoOidNamespace, trnName, false).ToString(),
				Description = trnDescription
			};
			Console.WriteLine($"Processing Transaction {trnName} with Description {trnDescription}");

			int row = Configuration.AttributeStartRow;
			while (sheet.Cells[row, Configuration.AttributesStartColumn].Value != null)
			{
				TransactionAttribute att = ReadAttribute(sheet, row);
				if (att != null)
				{
					level.Attributes.Add(att);
					Console.WriteLine($"Processing Attribute {att.Name}");
					row++;
				}
				else
					break;
			}
			if (level.Attributes.Count == 0)
			{
				throw new Exception("Transaction without attributes, check the AttributeStartRow and AttributeStartColumn values on the config file");
			}
			Console.WriteLine($"Generating Export xml to {outputFile}");
			TemplateGroup grp = new TemplateGroupString(File.ReadAllText(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "ExportTemplate.stg")));
			Template component_tpl = grp.GetInstanceOf("g_transaction_render");
			component_tpl.Add("component", level);
			string xmlExport = component_tpl.Render();
			File.WriteAllText(outputFile, xmlExport);
			Console.WriteLine($"Success");
		}

		private static TransactionAttribute ReadAttribute(ExcelWorksheet sheet, int row)
		{
			if (String.IsNullOrEmpty(sheet.Cells[row, 6].Value?.ToString().Trim()))
				return null;
			TransactionAttribute att = new TransactionAttribute
			{
				Name = sheet.Cells[row, Configuration.AttributeNameColumn].Value?.ToString().Trim(),
				Description = sheet.Cells[row, Configuration.AttributeDescriptionColumn].Value?.ToString().Trim(),
				AllowNull = sheet.Cells[row, Configuration.AttributeNullableColumn].Value != null,
				IsKey = sheet.Cells[row, Configuration.AttributeKeyColumn].Value?.ToString().ToLower() == "pk",
				Type = sheet.Cells[row, Configuration.AttributeDataTypeColumn].Value?.ToString().Trim().ToLower(),
				
			};
			att.Guid = GuidHelper.Create(GuidHelper.DnsNamespace, att.Name, false).ToString();
			string lenAndDecimals = sheet.Cells[row, Configuration.AttributeDataLengthColumn].Value?.ToString().Trim();
			if (lenAndDecimals != null)
			{
				string[] splitedData = lenAndDecimals.Trim().Split(',');
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
