using Antlr4.StringTemplate;
using OfficeOpenXml;
using System;
using System.IO;
using System.Reflection;

namespace ExcelParser
{
	public abstract class ExcelReader<TConfig>
        where TConfig : ExcelReader<TConfig>.BaseConfiguration, new()
	{
		private string CurrentFileName;
		private bool ContinueOnErrors;

		public abstract class BaseConfiguration
		{
			public int ObjectNameRow = 3;
			public int ObjectNameColumn = 7;
			public int ObjectDescColumn = 11;
			public int ObjectDescRow = 3;
			public string DefinitionSheetName;

			public int LevelCheckColumn = 3;
			public int LevelIdColumn = 8;
			public int LevelParentIdColumn = 9;
			public string LevelIdentifierKeyword = "LVL";
			public int AttributeStartRow = 7;
			public int AttributesStartColumn = 2;
			public int AttributeNameColumn = 7;
			public int AttributeDomainColumn = 9;
			public int AttributeDescriptionColumn = 6;
			public int AttributeNullableColumn = 4;
			public int AttributeKeyColumn = 3;
			public int AttributeDataTypeColumn = 8;
			public int AttributeDataLengthColumn = 9;
			public int AttributeAutonumberColumn = 12;
			public string PKValue = "PK";
			public string NullableValue = "?";
		}

        public TConfig Configuration { get; } = new TConfig();
		protected abstract string TemplateFile { get; }
		protected abstract string TemplateRender { get; }

		public void ReadExcel(string[] files, string outputFile, bool continueOnErrors)
		{
			ContinueOnErrors = continueOnErrors;
			Initialize();

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
			SetTemplateProperties(component_tpl);

			using (var stream = new FileStream(outputFile, FileMode.Create, FileAccess.ReadWrite))
			using (var writer = new StreamWriter(stream))
			{
				ITemplateWriter wr = new AutoIndentWriter(writer);
				wr.LineWidth = AutoIndentWriter.NoWrap;
				component_tpl.Write(wr);
			}

			Console.WriteLine($"Success");
		}

		protected abstract void ProcessFile(ExcelWorksheet sheet, string objName);

		protected virtual void SetTemplateProperties(Template component_tpl)
		{
		}

		protected virtual void Initialize() { }

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
