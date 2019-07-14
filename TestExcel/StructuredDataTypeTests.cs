using ExcelParser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Reflection;

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
            reader.Configuration.BaseTypeColumn = 7;
            reader.Configuration.DataStartRow = 6;
            reader.Configuration.DataNameColumn = 5;
            reader.Configuration.DataDescriptionColumn = 6;
            reader.Configuration.DataTypeColumn = 8;
            reader.Configuration.DataLengthColumn = 9;
            reader.Configuration.ItemIsCollectionColumn = 10;
            reader.Configuration.CollectionIdentifierKeyword = "Y";
            reader.Configuration.CollectionItemNameColumn = 11;
        }

        private void ConfigJP(SDTExcelReader reader)
        {
            reader.Configuration.ObjectNameRow = 3;
            reader.Configuration.ObjectNameColumn = 4;
            reader.Configuration.ObjectDescColumn = 5;
            reader.Configuration.ObjectDescRow = 3;
            reader.Configuration.LevelCheckColumn = 3;
            reader.Configuration.LevelIdentifierKeyword = "レベル";
            reader.Configuration.LevelParentIdColumn = 4;
            reader.Configuration.LevelIdColumn = 2;
            reader.Configuration.BaseTypeColumn = 8;
            reader.Configuration.DataStartColumn = 2;
            reader.Configuration.DataStartRow = 6;
            reader.Configuration.DataNameColumn = 5;
            reader.Configuration.DataDescriptionColumn = 6;
            reader.Configuration.DataTypeColumn = 9;
            reader.Configuration.DataLengthColumn = 10;
            reader.Configuration.ItemIsCollectionColumn = 11;
            reader.Configuration.CollectionIdentifierKeyword = "Y";
            reader.Configuration.CollectionItemNameColumn = 12;
        }

        private string GetFilePath(string fileName) =>
            Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), fileName);


        [TestMethod]
        public void TestReadOneFile()
        {
            SDTExcelReader reader = new SDTExcelReader();
            Config(reader);
            reader.Configuration.DefinitionSheetName = "SDT Definition";
            reader.ReadExcel(new string[] { GetFilePath("SDT_Test.xlsx") }
                , GetFilePath("sdts.xml"), false);
        }

        [TestMethod]
        public void TestReadSDTCollection()
        {
            SDTExcelReader reader = new SDTExcelReader();
            Config(reader);
            reader.Configuration.DefinitionSheetName = "SDT Definition";
            reader.ReadExcel(new string[] { GetFilePath("SDT_Test_LVL.xlsx") }
                , GetFilePath("sdts_col.xml"), false);
        }

        [TestMethod]
        public void TestReadJPFile()
        {
            SDTExcelReader reader = new SDTExcelReader();
            ConfigJP(reader);
            reader.Configuration.DefinitionSheetName = "定義シート";
            reader.ReadExcel(new string[] { GetFilePath("SDT_Test_JP.xlsx") }
                , GetFilePath("sdts_jp.xml"), false);
        }

        [TestMethod]
        public void Test4()
        {
            SDTExcelReader reader = new SDTExcelReader();
            reader.Configuration.DataStartRow = 5;
            reader.Configuration.DataStartColumn = 1;
            reader.Configuration.DataNameColumn = 1;
            reader.Configuration.DataDescriptionColumn = 2;
            reader.Configuration.DataTypeColumn = 4;
            reader.Configuration.LevelCheckColumn = 3;
            reader.Configuration.LevelIdColumn = 8;
            reader.Configuration.LevelParentIdColumn = 9;
            reader.Configuration.LevelIdentifierKeyword = "LVL";
            reader.Configuration.BaseTypeColumn = 10;
            reader.Configuration.DataLengthColumn = 5;
            reader.Configuration.ObjectNameRow = 2;
            reader.Configuration.ObjectNameColumn = 1;
            reader.Configuration.ObjectDescRow = 2;
            reader.Configuration.ObjectDescColumn = 2;
            reader.Configuration.DefinitionSheetName = "TransactionDefinitionSheet";
            reader.Configuration.ItemIsCollectionColumn = 10;
            reader.Configuration.CollectionIdentifierKeyword = "Y";
            reader.Configuration.CollectionItemNameColumn = 11;
            reader.ReadExcel(new string[] { GetFilePath("Test4_JP_SDT1.xlsx"), GetFilePath("Test4_JP_SDT2.xlsx") }
                , GetFilePath("sdts_jp_4.xml"), false);
        }
    }
}
