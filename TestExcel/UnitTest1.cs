using System;
using System.IO;
using System.Reflection;
using ExcelToTransactions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestExcel
{
	[TestClass]
	public class UnitTest1
	{
		[TestMethod]
		public void TestMethod1()
		{
			ExcelReader.ReadExcel(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "TransactionDefinitionSheet_MovementDetail.xlsx")
				, Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "trns.xml"));
		}
	}
}
