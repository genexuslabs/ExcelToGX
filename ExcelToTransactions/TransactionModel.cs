using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToTransactions
{
	public class TransactionLevel
	{
		public string Name { get; set; }
		public string Description { get; internal set; }
		public string Guid { get; set; }
		public List<TransactionAttribute> Attributes { get; set; } = new List<TransactionAttribute>();
	}

	public class TransactionAttribute
	{
		public string Guid { get; set; }
		public bool IsKey { get; set; }
		public bool AllowNull { get; set; }
		public string Name { get; set; }
		public string Description { get; set; }
		public string Type { get; set; }
		public int? Length { get; set; }
		public int? Decimals { get; set; }
	}
}
