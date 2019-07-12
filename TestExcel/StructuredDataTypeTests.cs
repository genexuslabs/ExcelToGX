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

        [TestMethod]
        public void TestReadOneFile()
        {
            SDTExcelReader reader = new SDTExcelReader();
            Config(reader);
            reader.Configuration.DefinitionSheetName = "SDT Definition";
            reader.ReadExcel(new string[] { Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "SDT_Test.xlsx") }
                , Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "sdts.xml"), false);
        }

        [TestMethod]
        public void TestReadSDTCollection()
        {
            SDTExcelReader reader = new SDTExcelReader();
            Config(reader);
            reader.Configuration.DefinitionSheetName = "SDT Definition";
            reader.ReadExcel(new string[] { Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "SDT_Test_LVL.xlsx") }
                , Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "sdts_col.xml"), false);
        }

        [TestMethod]
        public void TestReadJPFile()
        {
            SDTExcelReader reader = new SDTExcelReader();
            ConfigJP(reader);
            reader.Configuration.DefinitionSheetName = "定義シート";
            reader.ReadExcel(new string[] { Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "SDT_Test_JP.xlsx") }
                , Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "sdts_jp.xml"), false);
        }
    }
}
