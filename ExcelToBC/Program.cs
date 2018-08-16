using Artech.Common.Helpers;
using ExcelToBC.Properties;
using ExcelToTransactions;
using System;
using System.IO;
using System.Reflection;
using static ExcelToTransactions.ExcelReader;

namespace GeneXus.Utilities
{
	class Program
	{
		static void Main(string[] args)
		{
			ConvertCommandLine cmd = new ConvertCommandLine();
			try
			{
				cmd.Parse(args);
			}
			catch
			{
				Console.WriteLine(cmd.GetUsage());
				return;
			}
			// Read Configuration from App.config
			ReadConfiguration();

			try
			{
				string inputXlsx = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), cmd.ExcelFile);
				string outputXml = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), cmd.OutputFile);
				ExcelReader.ReadExcel(inputXlsx, outputXml);
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);
			}
		}

		private static void ReadConfiguration()
		{
			Configuration.TransactionNameRow = Settings.Default.TransactionNameRow;
			Configuration.TransactionNameCol = Settings.Default.TransactionNameCol;
			Configuration.TransactionDescColumn = Settings.Default.TransactionDescCol;
			Configuration.TransactionDescRow = Settings.Default.TransactionDescRow;

			Configuration.AttributeStartRow = Settings.Default.AttributeStartRow;
			Configuration.AttributesStartColumn = Settings.Default.AttributeStartColumn;
			Configuration.AttributeNameColumn = Settings.Default.AttributeNameColumn;
			Configuration.AttributeDescriptionColumn = Settings.Default.AttributeDescriptionColumn;
			Configuration.AttributeDataTypeColumn = Settings.Default.AttributeDataTypeColumn;
			Configuration.AttributeDataLengthColumn = Settings.Default.AttributeDataLengthColumn;
			Configuration.AttributeKeyColumn = Settings.Default.AttributeKeyColumn;
			Configuration.AttributeNullableColumn = Settings.Default.AttributeNullableColumn;
			Configuration.TransactionDefinitionSheetName = Settings.Default.TransactionDefinitionSheetName;
		}
	}

	public class ConvertCommandLine : CommandLineParser
	{
		[ValueUsage("Uri of Excel File, could be relative to this exe or absoulte", Optional = false, AlternateName1 = "x")]
		public string ExcelFile;

		[ValueUsage("The relative or full path to the output file, the output is in xml format",  Optional = true, AlternateName1 = "o")]
		public string OutputFile = "Transaction.xml";
	}
}
