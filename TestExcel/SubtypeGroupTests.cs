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
	public class SubtypeGroupTests
	{
		private void Config(GRPExcelReader reader)
		{
			reader.Configuration.ObjectNameRow = 2;
			reader.Configuration.ObjectNameColumn = 1;
			reader.Configuration.ObjectDescColumn = 2;
			reader.Configuration.ObjectDescRow = 2;
			reader.Configuration.DataStartRow = 5;
			reader.Configuration.DataNameColumn = 1;
			reader.Configuration.DataDescriptionColumn = 2;
            reader.Configuration.SupertypeColumn = 3;
		}


		[TestMethod]
		public void TestReadOneFile()
		{
			GRPExcelReader reader = new GRPExcelReader();
			Config(reader);
			reader.Configuration.DefinitionSheetName = "GroupDefinitionSheet";
			reader.ReadExcel(new string[] { Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "GRP_Test.xlsx") }
				, Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "groups.xml"), false);
		}
	}
}
