using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using ExcelToTransactions;
using ExcelToTransactions.DataTypes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using static ExcelToTransactions.ExcelReader;

namespace TestExcel
{
	[TestClass]
	public class UnitTest1
	{

		private void Config()
		{
			Configuration.TransactionNameRow = 2;
			Configuration.TransactionNameCol = 1;
			Configuration.TransactionDescColumn = 2;
			Configuration.TransactionDescRow = 2;
			Configuration.LevelCheckColumn = 3;
			Configuration.LevelIdentifierKeyword = "LVL";
			Configuration.LevelParentIdColumn = 9;
			Configuration.LevelIdColumn = 8;
			Configuration.AttributeDomainColumn = 10;
			Configuration.AttributeStartRow = 5;
			Configuration.AttributesStartColumn = 1;
			Configuration.AttributeNameColumn = 1;
			Configuration.AttributeDescriptionColumn = 2;
			Configuration.AttributeDataTypeColumn = 4;
			Configuration.AttributeDataLengthColumn = 5;
			Configuration.AttributeKeyColumn = 3;
			Configuration.AttributeNullableColumn = 6;
		}
		[TestMethod]
		public void TestReadOneFile()
		{
			Config();
			Configuration.TransactionDefinitionSheetName = "TransactionDefinitionSheet";
			ExcelReader.ReadExcel(new string[] { Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Test.xlsx") }
				, Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "trns.xml"), false);
		}

		[TestMethod]
		public void TestReadFiles()
		{
			Config();

			Configuration.TransactionDefinitionSheetName = "TransactionDefinitionSheet";
			ExcelReader.ReadExcel(new string[] { Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Test.xlsx"), Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Test3.xlsx") }
				, Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "trns.xml"), false);
		}

		[TestMethod]
		public void SetTypes()
		{
			TransactionAttribute att = new TransactionAttribute();

			DataTypeManager.SetDataType("Numeric(8.2)", att);
			Assert.IsTrue(att.Type == "Numeric" && att.Length == 8 && att.Decimals == 2 && att.Sign == false);

			DataTypeManager.SetDataType("Num(8.2)", att);
			Assert.IsTrue(att.Type == "Numeric" && att.Length == 8 && att.Decimals == 2 && att.Sign == false);

			DataTypeManager.SetDataType("Num(8)", att);
			Assert.IsTrue(att.Type == "Numeric" && att.Length == 8 && att.Decimals == 0 && att.Sign == false);

			DataTypeManager.SetDataType("Num(8-)", att);
			Assert.IsTrue(att.Type == "Numeric" && att.Length == 8 && att.Decimals == 0 && att.Sign == true);

			DataTypeManager.SetDataType("Num(12.4-)", att);
			Assert.IsTrue(att.Type == "Numeric" && att.Length == 12 && att.Decimals == 4 && att.Sign == true);

			DataTypeManager.SetDataType("DateTime(12.4-)", att);
			Assert.IsTrue(att.Type == "DateTime" && att.Length == null && att.Decimals == null && att.Sign == null);

			DataTypeManager.SetDataType("DateTime", att);
			Assert.IsTrue(att.Type == "DateTime" && att.Length == null && att.Decimals == null && att.Sign == null);

			DataTypeManager.SetDataType("Video", att);
			Assert.IsTrue(att.Type == "Video" && att.Length == null && att.Decimals == null && att.Sign == null);

			DataTypeManager.SetDataType("VIDEO", att);
			Assert.IsTrue(att.Type == "Video" && att.Length == null && att.Decimals == null && att.Sign == null);

			DataTypeManager.SetDataType("Char(20.2)", att);
			Assert.IsTrue(att.Type == "Character" && att.Length == 20 && att.Decimals == null && att.Sign == null);

			DataTypeManager.SetDataType("VarChar(200)", att);
			Assert.IsTrue(att.Type == "VarChar" && att.Length == 200 && att.Decimals == null && att.Sign == null);

			DataTypeManager.SetDataType("VarChar()", att);
			Assert.IsTrue(att.Type == "VarChar" && att.Length == 200 && att.Decimals == null && att.Sign == null);


			//DataTypeManager.SetDataType("Num(8.2)", att);
			//DataTypeManager.SetDataType("Num(8.2)", att);
		}

	}
}
