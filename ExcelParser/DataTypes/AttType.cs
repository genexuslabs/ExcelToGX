using System;
using System.Linq;

namespace Artech.Genexus.Common
{
	public enum EDBType
	{
		NUMERIC = 4,
		CHARACTER,
		DATE,             // Same as DATE or DATE_TP
		BITMAP,           // Bitmap
		LONGVARCHAR,      // Long Varchar
		PACKED,
		ZONED,
		INT,              // Attri type -> Int
		DATETIME,         // Date&Time
		VARCHAR,          // (Short) Varchar
		BINARY,           // bit stream or BLOb (Binary Large Object)
		Boolean,          // Boolean
		COLLECTION = 44,  // for deklarit
		GUID = 45,        // GUI Data Type
		VIDEO = 92,       // Multimedia (Video)
		AUDIO = 93,       // Multimedia (Audio)
		BINARYFILE = 112, // Binary File
		GEOGRAPHY = 102,  // Geospatial
		GEOPOINT = 103,   // Geospatial
		GEOLINE = 104,    // Geospatial
		GEOPOLYGON = 105, // Geospatial

		// Se deja espacio para los tipos GX (gxtypes.gxt)
		GX_DATASELECTOR = 244, // Dataselector
		SDOBJECT = 245,       // Dashboard, WorkWith or SDPanel
		GX_ARR_REF = 247,  // Array
		GX_VAR_REF,        // Variable
		GX_DOM_REF,        // Domain
		DT_Program,        // Program
		GX_BUSCOMP_LEVEL,  // Business component
		GX_BUSCOMP,        // Business component (root level)
		GX_ATT_REF,        // Attribute
		GX_SDT,            // Structure data type
		GX_USRDEFTYP,      // Tipo de usuario definido por el usuario
		GX_EXTERNAL_OBJECT,
		NONE               // Usado para void en metodos de External Object
	};

	public class EDBTypeConstants
	{
		public const int MAX_NUMERIC_LENGTH = 18;

		private static EDBType[] CHARACTER_TYPES = { EDBType.CHARACTER, EDBType.VARCHAR, EDBType.LONGVARCHAR };
		public static bool IsString(EDBType dbtype)
		{
			return CHARACTER_TYPES.Contains(dbtype);
		}
	}
}
