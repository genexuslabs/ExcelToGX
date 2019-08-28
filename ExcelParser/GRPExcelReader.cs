using Antlr4.StringTemplate;
using Artech.Common.Helpers.Guids;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace ExcelParser
{
	public class GRPExcelReader : ExcelReader<GRPExcelReader.GRPConfiguration, GroupStructure>
	{
		Dictionary<string, GroupAttribute> Attributes;
		Dictionary<string, GroupStructure> Groups;

		public class GRPConfiguration : BaseConfiguration
		{
			public int SupertypeColumn;
		}

		protected override string TemplateFile => "ExportGRPTemplate.stg";
		protected override string TemplateRender => "g_subtypegroup_render";

		protected override Guid ObjectTypeGuid => GuidHelper.ObjClass.SubtypeGroup;

		protected override void BeforeFileList()
		{
			base.BeforeFileList();
			Attributes = new Dictionary<string, GroupAttribute>();
			Groups = new Dictionary<string, GroupStructure>();
		}

		protected override void AddTemplateArguments(Template component_tpl)
		{
			component_tpl.Add("groups", Groups.Values);
			component_tpl.Add("attributes", Attributes.Values);
		}

		protected override void ProcessFile(ExcelWorksheet sheet, GroupStructure group)
		{
			Groups[group.Name] = group;

			int row = Configuration.DataStartRow;
			int atts = 0;
			string subtypeName = sheet.Cells[row, Configuration.DataNameColumn].Value?.ToString().Trim();
			while (!string.IsNullOrEmpty(subtypeName))
			{
				Console.WriteLine($"Processing subtype {subtypeName}");
				try
				{
					GroupAttribute subtype = new GroupAttribute()
					{
						Name = subtypeName,
						Description = sheet.Cells[row, Configuration.DataDescriptionColumn].Value?.ToString().Trim(),
						Guid = GuidHelper.Create(Configuration.Guid_CompatibilityMode ? GuidHelper.LegacyGuids.DnsNamespace : GuidHelper.ObjClass.Attribute, subtypeName, false).ToString(),
						Supertype = sheet.Cells[row, Configuration.SupertypeColumn].Value?.ToString().Trim(),
					};
					if (Attributes.TryGetValue(subtypeName, out var other) && subtype.ToString() != other.ToString())
						Console.WriteLine($"{subtype.Name} was already defined with a different supertype {other.ToString()}, taking into account the last one {subtype.ToString()}");
					Attributes[subtypeName] = subtype;
					atts++;
					group.Items.Add(subtype);
					row++;
				}
				catch (Exception ex) when (HandleException($"at row:{row}", ex, true))
				{
				}
				subtypeName = sheet.Cells[row, Configuration.DataNameColumn].Value?.ToString().Trim();
			}
			if (atts == 0)
			{
				throw new Exception($"Definition without content, check the {nameof(Configuration.DataStartRow)} and {nameof(Configuration.DataStartColumn)} values on the config file");
			}
		}
	}
}
