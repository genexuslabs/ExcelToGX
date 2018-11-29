using System;

namespace ExcelParser
{
	public abstract class ExcelReader<TConfig>
        where TConfig : ExcelReader<TConfig>.BaseConfiguration, new()
	{
		public abstract class BaseConfiguration
		{
			public int ObjectNameRow = 3;
			public int ObjectNameColumn = 7;
			public int ObjectDescColumn = 11;
			public int ObjectDescRow = 3;
			public string DefinitionSheetName;

			public int LevelCheckColumn = 3;
			public int LevelIdColumn = 8;
			public int LevelParentIdColumn = 9;
			public string LevelIdentifierKeyword = "LVL";
			public int AttributeStartRow = 7;
			public int AttributesStartColumn = 2;
			public int AttributeNameColumn = 7;
			public int AttributeDomainColumn = 9;
			public int AttributeDescriptionColumn = 6;
			public int AttributeNullableColumn = 4;
			public int AttributeKeyColumn = 3;
			public int AttributeDataTypeColumn = 8;
			public int AttributeDataLengthColumn = 9;
			public int AttributeAutonumberColumn = 12;
			public string PKValue = "PK";
			public string NullableValue = "?";
		}

        public TConfig Configuration { get; } = new TConfig();

		protected static void HandleException(string fileName, Exception ex, bool continueOnErrors)
		{
			if (!continueOnErrors)
			{
				Console.WriteLine($"Error: processing {fileName}, stop processing because continueOnErrors = false");
				throw ex;
			}
			else
			{
				Console.WriteLine($"Error: processing {fileName}, continue because continueOnErrors = true");
				Console.WriteLine(ex.Message);
			}
		}
    }
}
