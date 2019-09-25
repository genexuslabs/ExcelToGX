using Artech.Genexus.Common.Types;
using System.Collections.Generic;
using System.Security;

namespace ExcelParser
{
	public interface IKBElement
	{
		string Name { get; set; }
		string Description { get; set; }
		string Guid { get; set; }
	}

	public interface ITransactionElement : IKBElement
	{
		bool IsAttribute { get; }
	}

	public class LevelElement<IType> : IKBElement
		where IType : IKBElement
	{
		protected LevelElement(string parentKeyPath)
		{
			ParentKeyPath = parentKeyPath;
		}

		private string ParentKeyPath { get; }
		public string KeyPath => (ParentKeyPath is string parentKeyPath ? parentKeyPath + "::" : "") + Name;

		public string Name { get; set; }
		public string Description { get; set; }
		public string Guid { get; set; }
		public bool IsAttribute => false;

		public List<IType> Items = new List<IType>();
	}

	public class TransactionLevel : LevelElement<ITransactionElement>, ITransactionElement
	{
		public TransactionLevel(string parentKeyPath)
			: base(parentKeyPath)
		{
		}
	}

	public class DataTypeElement : IKBElement
	{
		public string Guid { get; set; }
		public string Name { get; set; }
		public string Description { get; set; }
		public string Type { get; set; }
		public string BaseType { get; set; }
		public int? Length { get; set; }
		public int? Decimals { get; set; }
		public bool? Sign { get; set; }

		internal void SetDataType(BasicDataTypeInfo dtInfo)
		{
			Type = dtInfo.Name;
			if (dtInfo.AllowLength)
				Length = dtInfo.Length;
			else
				Length = null;
			if (dtInfo.AllowSign)
				Sign = dtInfo.Signed;
			else
				Sign = null;
			if (dtInfo.AllowDecimals)
				Decimals = dtInfo.Decimals;
			else
				Decimals = null;
		}
	}

	public class TransactionAttribute : DataTypeElement, ITransactionElement
	{
		public bool IsKey { get; set; }
		public bool Autonumber { get; set; }
		public bool IsFormula { get; set; }
		public string Formula { get; set; }
		public bool AllowNull { get; set; }
		public bool IsAttribute => true;

		public string EscapeFormula => SecurityElement.Escape(Formula).Replace("&apos;", "'");

		public string Title { get; internal set; }
		public string ColumnTitle { get; internal set; }
		public string ContextualTitle { get; internal set; }

		public override string ToString() => $"{Type} {Description}";
	}

	public interface ISDTElement : IKBElement
	{
		bool IsCollection { get; set; }
		string CollectionItemName { get; set; }
		bool IsItem { get; }
	}

	public class SDTItem : DataTypeElement, ISDTElement
	{
		public bool IsCollection { get; set; }
		public string CollectionItemName { get; set; }
		public bool IsItem => true;

		public string BaseTypeObject { get; set; }
		public string BaseTypeProperty { get; set; }
		public string BaseTypePrefix { get; set; }
	}

	public class SDTLevel : LevelElement<ISDTElement>, ISDTElement
	{
		public SDTLevel(string parentKeyPath)
			: base(parentKeyPath)
		{
		}

		public bool IsCollection { get; set; }
		public string CollectionItemName { get; set; }
		public bool IsItem => false;
	}

	public class GroupAttribute : IKBElement
	{
		public string Name { get; set; }
		public string Description { get; set; }
		public string Guid { get; set; }
		public string Supertype { get; set; }

		public override string ToString() => $"{Name} {Description} {Supertype}";
	}

	public class GroupStructure : LevelElement<GroupAttribute>, IKBElement
	{
		public GroupStructure()
			: base(null)
		{
		}
	}
}
