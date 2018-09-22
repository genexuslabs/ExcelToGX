using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Artech.Genexus.Common.Types;

namespace ExcelToTransactions
{
	public interface ITransactionElement
	{
		string Name { get; set; }
		string Description { get; set; }
		string Guid { get; set; }


	}
	public class TransactionLevel : ITransactionElement
	{
		public string Name { get; set; }
		public string Description { get; set; }
		public string Guid { get; set; }
		public bool IsAttribute { get; } = false;

		public List<ITransactionElement> Items = new List<ITransactionElement>();
	}

	public class TransactionAttribute : ITransactionElement
	{
		public string Guid { get; set; }
		public bool IsKey { get; set; }
		public bool Autonumber { get; set; } = false;
		public bool AllowNull { get; set; }
		public string Name { get; set; }
		public string Description { get; set; }
		public string Type { get; set; }
		public string Domain { get; set; }
		public int? Length { get; set; }
		public int? Decimals { get; set; }
		public bool? Sign { get; set; }
		public bool IsAttribute { get; } = true;

		public override string ToString()
		{
			return $"{Type} {Description}";
		}

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
}
