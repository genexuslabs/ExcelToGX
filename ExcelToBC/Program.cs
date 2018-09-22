using Artech.Common.Helpers;
using ExcelToBC.Properties;
using ExcelToTransactions;
using System;
using System.Collections.Generic;
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
				string inputDirectory = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), cmd.ExcelFile);
				string outputXml = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), cmd.OutputFile);
				if (!String.IsNullOrEmpty(inputXlsx))
				{
					if (!File.Exists(inputXlsx))
						Console.WriteLine($"Input File {inputXlsx} does not exists, please specify an existing xlsx file");
					else
						ExcelReader.ReadExcel(new string[] { inputXlsx }, outputXml, cmd.ContinueOnErrors);
				}
				else
				{
					if (!String.IsNullOrEmpty(inputDirectory) && Directory.Exists(inputDirectory))
					{
						string[] files = Directory.GetFiles(inputDirectory, "*.xlsx");
						List<string> fullPaths = new List<string>();
						foreach (string file in files)
							fullPaths.Add(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), file));
					
						if (files.Length > 0)
							ExcelReader.ReadExcel(fullPaths.ToArray(), outputXml, cmd.ContinueOnErrors);
						else
							Console.WriteLine("Could not find any xlsx on the given directory " + inputDirectory);
					}
					else
					{
						Console.WriteLine("Please specify a valid directory to explore for xlsx files");
					}
				}
				
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
			
			Configuration.AttributeKeyColumn = Settings.Default.AttributeKeyColumn;
			Configuration.AttributeNullableColumn = Settings.Default.AttributeNullableColumn;
			Configuration.TransactionDefinitionSheetName = Settings.Default.TransactionDefinitionSheetName;
			Configuration.LevelCheckColumn = Settings.Default.LevelCheckColumn;
			Configuration.LevelIdColumn = Settings.Default.LevelIdColumn;
			Configuration.LevelIdentifierKeyword = Settings.Default.LevelIdentifierKeyword;
			Configuration.LevelParentIdColumn = Settings.Default.LevelParentIdColumn;
			// Values for PK and Nullable
			Configuration.PKValue = Settings.Default.PKValue;
			Configuration.NullableValue = Settings.Default.NullableValue;
			Configuration.AttributeDomainColumn = Settings.Default.DomainColumn;
		}
	}

	public class ConvertCommandLine : CommandLineParser
	{
		[ValueUsage("Uri of Excel File, could be relative to this exe or absoulte", Optional = true, AlternateName1 = "x")]
		public string ExcelFile;

		[ValueUsage("Directory to process all xlsx files inside, could be relative to this exe or absoulte", Optional = true, AlternateName1 = "d")]
		public string Directory;

		[ValueUsage("The relative or full path to the output file, the output is in xml format",  Optional = true, AlternateName1 = "o")]
		public string OutputFile = "Transaction.xml";

		[ValueUsage("Specify if the the tool must continue converting even errors are detected ", Optional = true, AlternateName1 = "c")]
		public bool ContinueOnErrors = false;
	}
}
