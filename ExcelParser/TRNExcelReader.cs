using Antlr4.StringTemplate;
using Artech.Common.Helpers.Guids;
using ExcelParser.DataTypes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelParser
{
	public class TRNExcelReader : ExcelReader<TRNExcelReader.TRNConfiguration, ITransactionElement, TransactionAttribute, TransactionLevel>
	{
		public class TRNConfiguration : BaseConfiguration
		{
			public TRNConfiguration()
			{
				ObjectNameRow = 3;
				ObjectNameColumn = 7;
				ObjectDescColumn = 11;
				ObjectDescRow = 3;
				DefinitionSheetName = "TransactionDefinitionSheet";

				LevelCheckColumn = 3;
				LevelIdColumn = 8;
				LevelParentIdColumn = 9;
				LevelIdentifierKeyword = "LVL";
			}

			public int AttributeNullableColumn = 4;
			public int AttributeKeyColumn = 3;
			public int AttributeAutonumberColumn = 12;
			public string PKValue = "PK";
			public string NullableValue = "?";
		}

		protected override string TemplateFile => "ExportTRNTemplate.stg";
		protected override string TemplateRender => "g_transaction_render";

		protected override void ReadLevelProperties(TransactionLevel level, ExcelWorksheet sheet, int row)
		{
		}

		protected override void ReadLeafProperties(TransactionAttribute att, ExcelWorksheet sheet, int row)
		{
			att.AllowNull = sheet.Cells[row, Configuration.AttributeNullableColumn].Value?.ToString().ToLower() == Configuration.NullableValue.ToLower();
			att.IsKey = sheet.Cells[row, Configuration.AttributeKeyColumn].Value?.ToString().ToLower() == Configuration.PKValue.ToLower();
			att.Autonumber = sheet.Cells[row, Configuration.AttributeAutonumberColumn].Value?.ToString().ToLower() == "true";
		}
	}
}
