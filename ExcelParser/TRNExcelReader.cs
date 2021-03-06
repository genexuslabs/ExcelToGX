﻿using System;
using Artech.Common.Helpers.Guids;
using OfficeOpenXml;

namespace ExcelParser
{
	public class TRNExcelReader : DataExcelReader<TRNExcelReader.TRNConfiguration, ITransactionElement, TransactionAttribute, TransactionLevel>
	{
		public class TRNConfiguration : BaseDataConfiguration
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
			public int AttributeTitleColumn = 11;
			public int AttributeColumnTitleColumn = 12;
			public int AttributeContextualTitleColumn = 13;

			public int FormulaCheckColumn = 3;
			public string FormulaIdentifierKeyword = "FRM";
			public int FormulaColumn = 7;
		}

		protected override string TemplateFile => "ExportTRNTemplate.stg";
		protected override string TemplateRender => "g_transaction_render";

		protected override Guid LevelTypeGuid => GuidHelper.ObjClass.TransactionLevel;

		protected override Guid LeafTypeGuid => GuidHelper.ObjClass.Attribute;
		protected override bool UseParentKeyInLeafGuid => false;

		protected override Guid ObjectTypeGuid => GuidHelper.ObjClass.Transaction;

		protected override TransactionLevel CreateLevelElement(string name, Guid guid, string description, string parentKeyPath) =>
			new TransactionLevel(parentKeyPath)
			{
				Name = name,
				Guid = guid.ToString(),
				Description = description
			};

		protected override void ReadLevelProperties(TransactionLevel level, ExcelWorksheet sheet, int row)
		{
		}

		protected override void ReadLeafProperties(TransactionAttribute att, ExcelWorksheet sheet, int row)
		{
			att.IsKey = sheet.Cells[row, Configuration.AttributeKeyColumn].Value?.ToString().ToLower() == Configuration.PKValue.ToLower();
			if (!att.IsKey)
			{
				att.IsFormula = sheet.Cells[row, Configuration.FormulaCheckColumn].Value?.ToString().ToLower() == Configuration.FormulaIdentifierKeyword.ToLower();
				if (att.IsFormula)
					att.Formula = sheet.Cells[row, Configuration.FormulaColumn].Value?.ToString();
				else
					att.AllowNull = sheet.Cells[row, Configuration.AttributeNullableColumn].Value?.ToString().ToLower() == Configuration.NullableValue.ToLower();
			}
			if (!att.IsFormula)
				att.Autonumber = sheet.Cells[row, Configuration.AttributeAutonumberColumn].Value?.ToString().ToLower() == "true";
			att.Title = sheet.Cells[row, Configuration.AttributeTitleColumn].Value?.ToString();
			att.ColumnTitle = sheet.Cells[row, Configuration.AttributeColumnTitleColumn].Value?.ToString();
			att.ContextualTitle = sheet.Cells[row, Configuration.AttributeContextualTitleColumn].Value?.ToString();
		}
	}
}
