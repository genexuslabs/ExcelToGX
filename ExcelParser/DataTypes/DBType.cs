//	DBType.cs
//		Initially generated (16/04/2004) applying basicobject.dkt on DBType.xml
//
//	$Log: DBType.cs,v $
//	Revision 1.2  2004/06/25 18:06:27  jlr
//		Big architectural changes:
//			Simple data-only objects are used for inter-layer comunications.
//			Mappers and Providers (which deal with dataobjects) are used between layers.
//			Managers and Loaders (which deal with objects) are used inside each layer.
//			No Key interfaces. The keys are directly in the architecture.
//	
//	Revision 1.1  2004/06/25 16:45:32  jlr
//		Initial revision.
//	
//

using System;
using System.Collections;
using Artech.Common;

namespace Artech.Genexus.Common
{
	/* TemplateName = basicobject.dkt */
	/* Folder = Architecture\Basic\Common\Objects */
	/* OtroTest = special */
	[Serializable]
	public class DBType
	{

		protected EDBType m_Type;
		protected int m_Length;
		protected int m_Decimal;
		protected bool m_Sign;

		public DBType()
		{
		}

		public DBType(DBType parmDBType)
		{
			m_Type = parmDBType.Type;
			m_Length = parmDBType.Length;
			m_Decimal = parmDBType.Decimal;
			m_Sign = parmDBType.Sign;
		}

		private bool Compare(DBType parmDBType)
		{
			return (m_Type == parmDBType.m_Type
				&& m_Length == parmDBType.m_Length
				&& m_Decimal == parmDBType.m_Decimal
				&& m_Sign == parmDBType.m_Sign
				&& true);
		}

		#region DBType Simple Members

		protected EDBType Type
		{
			get { return m_Type; }
			set
			{
				m_Type = value;

			}
		}

		protected int Length
		{
			get { return m_Length; }
			set
			{
				m_Length = value;

			}
		}

		protected int Decimal
		{
			get { return m_Decimal; }
			set
			{
				m_Decimal = value;

			}
		}

		protected bool Sign
		{
			get { return m_Sign; }
			set
			{
				m_Sign = value;

			}
		}

		#endregion


		public override int GetHashCode()
		{
			return Convert.ToInt32(m_Type)
				^ Convert.ToInt32(m_Length)
				^ Convert.ToInt32(m_Decimal)
				^ Convert.ToInt32(m_Sign)
				^ 0;
		}

		public override bool Equals(object obj)
		{
			if (base.Equals(obj)) return true;
			if (obj == null) return false;
			if (GetType() != obj.GetType()) return false;

			DBType data = (DBType)obj;

			return Compare(data);
		}

		public static bool operator ==(DBType left, DBType right)
		{
			if ((object)left == null && (object)right == null) return true;
			if ((object)left == null) return false;
			return left.Equals(right);
		}

		public static bool operator !=(DBType left, DBType right)
		{
			return !(left == right);
		}
	}
}
