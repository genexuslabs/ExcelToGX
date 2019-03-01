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
            reader.Configuration.BaseTypeColumn = 10;
            reader.Configuration.DataStartRow = 5;
            reader.Configuration.DataStartColumn = 1;
            reader.Configuration.DataNameColumn = 1;
            reader.Configuration.DataDescriptionColumn = 2;
            reader.Configuration.DataTypeColumn = 4;
            reader.Configuration.DataLengthColumn = 5;
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
            reader.Configuration.DataStartRow = 6;
            reader.Configuration.DataStartColumn = 2;
            reader.Configuration.AttributeKeyColumn = 7;
            reader.Configuration.PKValue = "PK";
            reader.Configuration.DataNameColumn = 5;
            reader.Configuration.DataDescriptionColumn = 6;
            reader.Configuration.BaseTypeColumn = 8;
            reader.Configuration.DataTypeColumn = 9;
            reader.Configuration.DataLengthColumn = 10;
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
            DataTypeElement dte = new DataTypeElement();

            DataTypeManager.SetDataType("Numeric(8.2)", dte);
            Assert.IsTrue(dte.Type == "Numeric" && dte.Length == 8 && dte.Decimals == 2 && dte.Sign == false);

            DataTypeManager.SetDataType("Num(8.2)", dte);
            Assert.IsTrue(dte.Type == "Numeric" && dte.Length == 8 && dte.Decimals == 2 && dte.Sign == false);

            DataTypeManager.SetDataType("Num(8)", dte);
            Assert.IsTrue(dte.Type == "Numeric" && dte.Length == 8 && dte.Decimals == 0 && dte.Sign == false);

            DataTypeManager.SetDataType("Num(8-)", dte);
            Assert.IsTrue(dte.Type == "Numeric" && dte.Length == 8 && dte.Decimals == 0 && dte.Sign == true);

            DataTypeManager.SetDataType("Num(12.4-)", dte);
            Assert.IsTrue(dte.Type == "Numeric" && dte.Length == 12 && dte.Decimals == 4 && dte.Sign == true);

            DataTypeManager.SetDataType("DateTime(12.4-)", dte);
            Assert.IsTrue(dte.Type == "DateTime" && dte.Length == null && dte.Decimals == null && dte.Sign == null);

            DataTypeManager.SetDataType("DateTime", dte);
            Assert.IsTrue(dte.Type == "DateTime" && dte.Length == null && dte.Decimals == null && dte.Sign == null);

            DataTypeManager.SetDataType("Video", dte);
            Assert.IsTrue(dte.Type == "Video" && dte.Length == null && dte.Decimals == null && dte.Sign == null);

            DataTypeManager.SetDataType("VIDEO", dte);
            Assert.IsTrue(dte.Type == "Video" && dte.Length == null && dte.Decimals == null && dte.Sign == null);

            DataTypeManager.SetDataType("Char(20.2)", dte);
            Assert.IsTrue(dte.Type == "Character" && dte.Length == 20 && dte.Decimals == null && dte.Sign == null);

            DataTypeManager.SetDataType("VarChar(200)", dte);
            Assert.IsTrue(dte.Type == "VarChar" && dte.Length == 200 && dte.Decimals == null && dte.Sign == null);

            DataTypeManager.SetDataType("VarChar()", dte);
            Assert.IsTrue(dte.Type == "VarChar" && dte.Length == 200 && dte.Decimals == null && dte.Sign == null);


            //DataTypeManager.SetDataType("Num(8.2)", att);
            //DataTypeManager.SetDataType("Num(8.2)", att);
        }
    }
}
