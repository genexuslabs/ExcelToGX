using OfficeOpenXml;

namespace ExcelParser
{
	public class SDTExcelReader : ExcelReader<SDTExcelReader.SDTConfiguration, ISDTElement, SDTItem, SDTLevel>
	{
		public class SDTConfiguration : BaseConfiguration
		{
			public SDTConfiguration()
			{
				ObjectNameRow = 3;
				ObjectNameColumn = 4;
				ObjectDescRow = 3;
				ObjectDescColumn = 5;
				DefinitionSheetName = "SDT Definition";

				LevelCheckColumn = 3;
				LevelIdColumn = 2;
				LevelParentIdColumn = 4;
				LevelIdentifierKeyword = "LV";
			}

			public int ItemIsCollectionColumn = 10;
			public string CollectionIdentifierKeyword = "Y";
			public int CollectionItemNameColumn = 11;
		}

		protected override string TemplateFile => "ExportSDTTemplate.stg";

		protected override string TemplateRender => "g_structureddatatype_render";

		protected override void ReadLevelProperties(SDTLevel level, ExcelWorksheet sheet, int row)
		{
			ReadIsCollection(level, sheet, row);
		}

		protected override void ReadLeafProperties(SDTItem item, ExcelWorksheet sheet, int row)
		{
			ReadIsCollection(item, sheet, row);
		}

		private void ReadIsCollection(ISDTElement element, ExcelWorksheet sheet, int row)
		{
			if (sheet.Cells[row, Configuration.ItemIsCollectionColumn].Value?.ToString().ToLower().Trim() == Configuration.CollectionIdentifierKeyword.ToLower().Trim()) //is collection
			{
				element.IsCollection = true;
				string collectionItemName = sheet.Cells[row, Configuration.CollectionItemNameColumn].Value?.ToString();
				if (!string.IsNullOrEmpty(collectionItemName))
					element.CollectionItemName = collectionItemName;
			}
		}
	}
}
