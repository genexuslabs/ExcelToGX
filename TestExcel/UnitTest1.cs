using System;
using System.IO;
using System.Reflection;
using ExcelToTransactions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static ExcelToTransactions.ExcelReader;

namespace TestExcel
{
	[TestClass]
	public class UnitTest1
	{
		[TestMethod]
		public void TestMethod1()
		{
			Configuration.TransactionNameRow = 2;
			Configuration.TransactionNameCol = 1;
			Configuration.TransactionDescColumn = 2;
			Configuration.TransactionDescRow = 2;

			Configuration.AttributeStartRow = 5;
			Configuration.AttributesStartColumn = 1;
			Configuration.AttributeNameColumn = 1;
			Configuration.AttributeDescriptionColumn = 2;
			Configuration.AttributeDataTypeColumn = 4;
			Configuration.AttributeDataLengthColumn = 5;
			Configuration.AttributeKeyColumn = 3;
			Configuration.AttributeNullableColumn = 6;
			Configuration.TransactionDefinitionSheetName = "TransactionDefinitionSheet";
			ExcelReader.ReadExcel(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Test.xlsx")
				, Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "trns.xml"));
		}
	}
}
