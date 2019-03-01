using OfficeOpenXml;
using System;

namespace ExcelParser
{
	public class SDTExcelReader : DataExcelReader<SDTExcelReader.SDTConfiguration, ISDTElement, SDTItem, SDTLevel>
	{
		public class SDTConfiguration : BaseDataConfiguration
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

			public string DomainPrefixKeyword = "dom";
			public string AttributePrefixKeyword = "att";
			public string SDTPrefixKeyword = "sdt";
			public string DefaultBaseTypePrefixKeyword = "dom";
		}

		protected override string TemplateFile => "ExportSDTTemplate.stg";

		protected override string TemplateRender => "g_structureddatatype_render";

		protected override void ReadLevelProperties(SDTLevel level, ExcelWorksheet sheet, int row)
		{
			ReadIsCollection(level, sheet, row);
		}

		protected override void ReadLeafProperties(SDTItem item, ExcelWorksheet sheet, int row)
		{
			string objectName = item.BaseType;
			if (!string.IsNullOrEmpty(objectName))
			{
				string typeName = null;
				int pos = item.BaseType.IndexOf(':');
				if (pos > -1)
				{
					typeName = objectName.Substring(0, pos);
					objectName = objectName.Substring(pos + 1);
					item.BaseType = objectName;
				}

				if (string.IsNullOrEmpty(typeName))
					typeName = Configuration.DefaultBaseTypePrefixKeyword;

				string lowerTypeName = typeName.ToLower();
				item.BaseTypeObject = objectName;
				if (lowerTypeName == Configuration.SDTPrefixKeyword.ToLower())
				{
					item.BaseTypeProperty = "ATTCUSTOMTYPE";
					item.BaseTypePrefix = "sdt";
					item.BaseType = null;
				}
				else
				{
					item.BaseTypeProperty = "idBasedOn";
					if (lowerTypeName == Configuration.DomainPrefixKeyword.ToLower())
						item.BaseTypePrefix = "Domain";
					else if (lowerTypeName == Configuration.AttributePrefixKeyword.ToLower())
					{
						item.BaseTypePrefix = "Attribute";
						item.BaseType = null;
					}
					else
						throw new InvalidOperationException($"The prefix '{typeName}' used in [{row} , {Configuration.DataTypeColumn}] is not defined as an expected type name prefix.");
				}
			}

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
