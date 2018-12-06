using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Artech.Genexus.Common.Types;

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
		public string Name { get; set; }
		public string Description { get; set; }
		public string Guid { get; set; }
		public bool IsAttribute => false;

		public List<IType> Items = new List<IType>();
	}

	public class TransactionLevel : LevelElement<ITransactionElement>, ITransactionElement
	{
	}

	public class DataTypeElement : IKBElement
	{
		public string Guid { get; set; }
		public string Name { get; set; }
		public string Description { get; set; }
		public string Type { get; set; }
		public string Domain { get; set; }
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
		public bool Autonumber { get; set; } = false;
		public bool AllowNull { get; set; }
		public bool IsAttribute => true;

		public override string ToString()
		{
			return $"{Type} {Description}";
		}
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
	}

	public class SDTLevel : LevelElement<ISDTElement>, ISDTElement
	{
		public bool IsCollection { get; set; }
		public string CollectionItemName { get; set; }
		public bool IsItem => false;
	}
}
