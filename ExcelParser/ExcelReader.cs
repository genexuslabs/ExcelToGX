using Antlr4.StringTemplate;
using Artech.Common.Helpers.Guids;
using OfficeOpenXml;
using System;
using System.IO;
using System.Reflection;

namespace ExcelParser
{
	public abstract class ExcelReader
	{
		public abstract void ReadExcel(string[] files, string outputFile, bool continueOnErrors);

		public abstract class BaseConfiguration
		{
			public int ObjectNameRow;
			public int ObjectNameColumn;
			public int ObjectDescColumn;
			public int ObjectDescRow;
			public string DefinitionSheetName;

			public int DataStartRow = 7;
			public int DataStartColumn = 2;
			public int DataNameColumn = 7;
			public int DataDescriptionColumn = 6;

			public bool Guid_CompatibilityMode;
		}
	}

	public abstract class ExcelReader<TConfig, TObject> : ExcelReader
		where TConfig : ExcelReader.BaseConfiguration, new()
		where TObject : IKBElement
	{
		private string CurrentFileName;
		private bool ContinueOnErrors;
		public TConfig Configuration { get; } = new TConfig();
		protected abstract string TemplateFile { get; }
		protected abstract string TemplateRender { get; }
		protected abstract Guid ObjectTypeGuid { get; }
		protected abstract TObject CreateObject(string name, Guid guid, string description);

		public sealed override void ReadExcel(string[] files, string outputFile, bool continueOnErrors)
		{
			ContinueOnErrors = continueOnErrors;

			BeforeFileList();

			foreach (string fileName in files)
			{
				CurrentFileName = fileName;
				try
				{
					var excel = new ExcelPackage(new FileInfo(fileName));
					ExcelWorksheet sheet = excel.Workbook.Worksheets[Configuration.DefinitionSheetName];
					if (sheet is null)
						throw new Exception($"The xlsx file '{fileName}' was opened but it does not have a definition sheet for the object. Please provide the object definition in a Sheet called '{Configuration.DefinitionSheetName}'");

					string objName = sheet.Cells[Configuration.ObjectNameRow, Configuration.ObjectNameColumn].Value?.ToString();
					if (objName is null)
						throw new Exception($"Could not find the object name at [{Configuration.ObjectNameRow} , {Configuration.ObjectNameColumn}]. Please take a look at the configuration file");

					string objDescription = sheet.Cells[Configuration.ObjectDescRow, Configuration.ObjectDescColumn].Value?.ToString();
					Console.WriteLine($"Processing object with name '{objName}' and description '{objDescription}'");
					TObject obj = CreateObject
					(
						name: objName,
						description: objDescription,
						guid: GuidHelper.Create(Configuration.Guid_CompatibilityMode ? GuidHelper.LegacyGuids.IsoOidNamespace : ObjectTypeGuid, objName, false)
					);

					ProcessFile(sheet, obj);
				}
				catch (Exception ex) when (HandleException(fileName, ex))
				{
				}
			}
			CurrentFileName = null;

			Console.WriteLine($"Generating Export xml to {outputFile}");
			TemplateGroup grp = new TemplateGroupFile(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), TemplateFile));
			Template component_tpl = grp.GetInstanceOf(TemplateRender);

			AddTemplateArguments(component_tpl);

			using (var stream = new FileStream(outputFile, FileMode.Create, FileAccess.ReadWrite))
			using (var writer = new StreamWriter(stream))
			{
				ITemplateWriter wr = new AutoIndentWriter(writer);
				wr.LineWidth = AutoIndentWriter.NoWrap;
				component_tpl.Write(wr);
			}

			Console.WriteLine($"Success");
		}

		protected virtual void BeforeFileList()
		{
		}

		protected abstract void ProcessFile(ExcelWorksheet sheet, TObject obj);

		protected abstract void AddTemplateArguments(Template component_tpl);

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
