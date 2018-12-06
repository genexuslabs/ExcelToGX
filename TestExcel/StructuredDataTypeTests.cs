using ExcelParser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace TestExcel
{
	[TestClass]
	public class StructuredDataTypeTests
	{
		private void Config(SDTExcelReader reader)
		{
			reader.Configuration.ObjectNameRow = 3;
			reader.Configuration.ObjectNameColumn = 4;
			reader.Configuration.ObjectDescColumn = 5;
			reader.Configuration.ObjectDescRow = 3;
			reader.Configuration.LevelCheckColumn = 3;
			reader.Configuration.LevelIdentifierKeyword = "LV";
			reader.Configuration.LevelParentIdColumn = 4;
			reader.Configuration.LevelIdColumn = 2;
			reader.Configuration.DomainColumn = 7;
			reader.Configuration.DataStartRow = 6;
			reader.Configuration.DataNameColumn = 5;
			reader.Configuration.DataDescriptionColumn = 6;
			reader.Configuration.DataTypeColumn = 8;
			reader.Configuration.DataLengthColumn = 9;
			reader.Configuration.ItemIsCollectionColumn = 10;
			reader.Configuration.CollectionIdentifierKeyword = "Y";
			reader.Configuration.CollectionItemNameColumn = 11;
		}


		[TestMethod]
		public void TestReadOneFile()
		{
			SDTExcelReader reader = new SDTExcelReader();
			Config(reader);
			reader.Configuration.DefinitionSheetName = "SDT Definition";
			reader.ReadExcel(new string[] { Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "SDT_Test.xlsx") }
				, Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "sdts.xml"), false);
		}
	}
}
