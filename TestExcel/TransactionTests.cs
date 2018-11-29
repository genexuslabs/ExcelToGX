using ExcelParser;
using ExcelParser.DataTypes;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Reflection;

namespace TestExcel
{
	[TestClass]
	public class TransactionTests
	{
		private void Config(TRNExcelReader reader)
		{
			reader.Configuration.ObjectNameRow = 2;
			reader.Configuration.ObjectNameColumn = 1;
			reader.Configuration.ObjectDescColumn = 2;
			reader.Configuration.ObjectDescRow = 2;
			reader.Configuration.LevelCheckColumn = 3;
			reader.Configuration.LevelIdentifierKeyword = "LVL";
			reader.Configuration.LevelParentIdColumn = 9;
			reader.Configuration.LevelIdColumn = 8;
			reader.Configuration.AttributeDomainColumn = 10;
			reader.Configuration.AttributeStartRow = 5;
			reader.Configuration.AttributesStartColumn = 1;
			reader.Configuration.AttributeNameColumn = 1;
			reader.Configuration.AttributeDescriptionColumn = 2;
			reader.Configuration.AttributeDataTypeColumn = 4;
			reader.Configuration.AttributeDataLengthColumn = 5;
			reader.Configuration.AttributeKeyColumn = 3;
			reader.Configuration.AttributeNullableColumn = 6;
		}

		private void ConfigJapan(TRNExcelReader reader)
		{
			reader.Configuration.DefinitionSheetName = "トランザクション定義書";
			reader.Configuration.ObjectNameRow = 3;
			reader.Configuration.ObjectNameColumn = 12;
			reader.Configuration.ObjectDescRow = 3;
			reader.Configuration.ObjectDescColumn = 13;
			reader.Configuration.AttributeStartRow = 6;
			reader.Configuration.AttributesStartColumn = 2;
			reader.Configuration.AttributeKeyColumn = 7;
			reader.Configuration.PKValue = "PK";
			reader.Configuration.AttributeNameColumn = 5;
			reader.Configuration.AttributeDescriptionColumn = 6;
			reader.Configuration.AttributeDomainColumn = 8;
			reader.Configuration.AttributeDataTypeColumn = 9;
			reader.Configuration.AttributeDataLengthColumn = 10;
			reader.Configuration.AttributeNullableColumn = 11;
			reader.Configuration.NullableValue = "Y";
			reader.Configuration.LevelCheckColumn = 3;
			reader.Configuration.LevelIdentifierKeyword = "レベル";
			reader.Configuration.LevelIdColumn = 2;
			reader.Configuration.LevelParentIdColumn = 4;
		}

		[TestMethod]
		public void TestReadOneFile()
		{
			TRNExcelReader reader = new TRNExcelReader();
			Config(reader);
			reader.Configuration.DefinitionSheetName = "TransactionDefinitionSheet";
			reader.ReadExcel(new string[] { Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Test.xlsx") }
				, Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "trns.xml"), false);
		}

		//[TestMethod]
		//public void TestJapanFile()
		//{
		//	TRNExcelReader reader = new TRNExcelReader();
		//	ConfigJapan(reader);
		//	reader.ReadExcel(new string[] { Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "TransactionDefinition.xlsx") }
		//		, Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "trnsJapan.xml"), false);
		//}

		[TestMethod]
		public void TestReadFiles()
		{
			TRNExcelReader reader = new TRNExcelReader();
			Config(reader);
			reader.Configuration.DefinitionSheetName = "TransactionDefinitionSheet";
			reader.ReadExcel(new string[] { Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Test.xlsx"), Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Test3.xlsx") }
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
