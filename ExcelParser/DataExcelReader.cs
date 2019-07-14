using Antlr4.StringTemplate;
using Artech.Common.Helpers.Guids;
using ExcelParser.DataTypes;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace ExcelParser
{
    public abstract class DataExcelReader<TConfig, IType, TLeafElement, TLevelElement> : ExcelReader<TConfig, TLevelElement>
        where TConfig : DataExcelReader<TConfig, IType, TLeafElement, TLevelElement>.BaseDataConfiguration, new()
        where IType : IKBElement
        where TLeafElement : DataTypeElement, IType, new()
        where TLevelElement : LevelElement<IType>, IType, new()
    {
        Dictionary<string, TLevelElement> Levels;
        Dictionary<string, TLeafElement> Leafs;
        Dictionary<string, DataTypeElement> Domains;

        public abstract class BaseDataConfiguration : BaseConfiguration
        {
            public int LevelCheckColumn = 3;
            public int LevelIdColumn = 8;
            public int LevelParentIdColumn = 9;
            public string LevelIdentifierKeyword = "LVL";

            public int DataTypeColumn = 8;
            public int DataLengthColumn = 9;
            public int BaseTypeColumn = 9;
        }

        protected override void BeforeFileList()
        {
            base.BeforeFileList();
            Levels = new Dictionary<string, TLevelElement>();
            Leafs = new Dictionary<string, TLeafElement>();
            Domains = new Dictionary<string, DataTypeElement>(StringComparer.OrdinalIgnoreCase);
        }

        protected override void AddTemplateArguments(Template component_tpl)
        {
            component_tpl.Add("levels", Levels.Values);
            component_tpl.Add("leafs", Leafs.Values);
            if (Domains.Values.Count > 0)
                component_tpl.Add("domains", Domains.Values);
        }

        protected override void ProcessFile(ExcelWorksheet sheet, TLevelElement obj)
        {
            Dictionary<int, TLevelElement> levels = new Dictionary<int, TLevelElement>();
            TLevelElement level = obj;
            levels[0] = level;
            Levels[level.Name] = level;

            int row = Configuration.DataStartRow;
            int atts = 0;
            while (sheet.Cells[row, Configuration.DataStartColumn].Value != null)
            {
                level = ReadLevel(row, sheet, level, levels, out bool isLevel);
                if (!isLevel)
                {
                    try
                    {
                        TLeafElement leaf = ReadLeaf(sheet, row);
                        if (leaf.BaseType != null && leaf.Type != null && !Domains.ContainsKey(leaf.BaseType))
                        {
                            DataTypeElement domain = new DataTypeElement
                            {
                                Name = leaf.BaseType,
                                Guid = GuidHelper.Create(GuidHelper.UrlNamespace, leaf.BaseType, false).ToString()
                            };
                            DataTypeManager.SetDataType(leaf.Type, domain);
                            SetLengthAndDecimals(sheet, row, domain);
                            Domains[domain.Name] = domain;
                        }
                        if (leaf != null)
                        {
                            if (Leafs.ContainsKey(leaf.Name) && leaf.ToString() != Leafs[leaf.Name].ToString())
                                Console.WriteLine($"{leaf.Name} was already defined with a different data type {Leafs[leaf.Name].ToString()}, taking into account the last one {leaf.ToString()}");
                            Leafs[leaf.Name] = leaf;
                            atts++;
                            level.Items.Add(leaf);
                            Console.WriteLine($"Processing Attribute {leaf.Name}");
                            row++;
                        }
                        else
                            break;
                    }
                    catch (Exception ex) when (HandleException($"at row:{row}", ex, true))
                    {
                    }
                }
                else
                {
                    ReadLevelProperties(level, sheet, row);
                    row++;
                }
            }
            if (atts == 0)
            {
                throw new Exception($"Definition without content, check the {nameof(Configuration.DataStartRow)} and {nameof(Configuration.DataStartColumn)} values on the config file");
            }
        }

        private TLevelElement ReadLevel(int row, ExcelWorksheet sheet, TLevelElement level, Dictionary<int, TLevelElement> levels, out bool isLevel)
        {
            int id;
            string idValue = sheet.Cells[row, Configuration.LevelIdColumn].Value?.ToString();
            if (idValue is null)
                id = -1;
            else
            {
                try
                {
                    id = int.Parse(sheet.Cells[row, Configuration.LevelIdColumn].Value?.ToString());
                }
                catch (Exception ex)
                {
                    throw new Exception($"Invalid identifier for level at {row}, {Configuration.LevelIdColumn} " + ex.Message, ex);
                }
            }
            isLevel = false;
            if (sheet.Cells[row, Configuration.LevelCheckColumn].Value?.ToString().ToLower().Trim() == Configuration.LevelIdentifierKeyword.ToLower().Trim()) //is a level
            {
                isLevel = true;

                if (id == 0)
                    return levels[0];

                int parentId = 0;
                // Read Level Id
                parentId = GetParentLevelId(sheet, row);
                string levelName = sheet.Cells[row, Configuration.DataNameColumn].Value?.ToString();
                if (levelName is null)
                    throw new Exception($"Could not find the Level name at [{row} , {Configuration.DataNameColumn}], please take a look at the configuration file ");
                string levelDesc = sheet.Cells[row, Configuration.DataDescriptionColumn].Value?.ToString();

                TLevelElement newLevel = new TLevelElement
                {
                    Name = levelName,
                    Guid = GuidHelper.Create(GuidHelper.IsoOidNamespace, levelName, false).ToString(),
                    Description = levelDesc
                };
                Debug.Assert(id >= 1);
                levels[id] = newLevel;

                TLevelElement parentTrn = levels[parentId];
                parentTrn.Items.Add(newLevel);
                return newLevel;
            }
            else
            {
                if (id == 0)
                    throw new Exception($"Invalid identifier for item at {row}, {Configuration.LevelIdColumn}. Only the root level element can have id 0 ");

                if (sheet.Cells[row, Configuration.LevelParentIdColumn].Value != null) // Explicit level id
                {
                    int parentLevel = GetParentLevelId(sheet, row);
                    if (parentLevel >= 0 && levels.ContainsKey(parentLevel))
                        return levels[parentLevel];
                }
                return level; //continue in the same level, there is no level at this row, so assume the same.
            }
        }

        protected TLeafElement ReadLeaf(ExcelWorksheet sheet, int row)
        {
            string leafName = sheet.Cells[row, Configuration.DataNameColumn].Value?.ToString().Trim();
            if (string.IsNullOrEmpty(leafName))
                return null;
            TLeafElement leaf = new TLeafElement()
            {
                Name = leafName,
                Description = sheet.Cells[row, Configuration.DataDescriptionColumn].Value?.ToString().Trim(),
                Guid = GuidHelper.Create(GuidHelper.DnsNamespace, leafName, false).ToString(),
                BaseType = sheet.Cells[row, Configuration.BaseTypeColumn].Value?.ToString().Trim(),
                Type = sheet.Cells[row, Configuration.DataTypeColumn].Value?.ToString().Trim().ToLower(),
            };
            ReadLeafProperties(leaf, sheet, row);
            try
            {
                if (leaf.BaseType is null && !string.IsNullOrEmpty(leaf.Type))
                {
                    DataTypeManager.SetDataType(leaf.Type, leaf);
                    SetLengthAndDecimals(sheet, row, leaf);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error parsing Data Type for row " + row);
                HandleException(sheet.Name, ex);
            }
            return leaf;
        }

        private void SetLengthAndDecimals(ExcelWorksheet sheet, int row, DataTypeElement dte)
        {
            try
            {
                string lenAndDecimals = sheet.Cells[row, Configuration.DataLengthColumn].Value?.ToString().Trim();
                if (!string.IsNullOrEmpty(lenAndDecimals))
                {
                    if (lenAndDecimals.Length > 0 && lenAndDecimals.EndsWith("-"))
                    {
                        dte.Sign = true;
                        lenAndDecimals = lenAndDecimals.Replace("-", "");
                    }
                    string[] splitedData;
                    if (lenAndDecimals.Contains("."))
                        splitedData = lenAndDecimals.Trim().Split('.');
                    else
                        splitedData = lenAndDecimals.Trim().Split(',');
                    if (splitedData.Length >= 1 && int.TryParse(splitedData[0], out int length))
                    {
                        dte.Length = length;
                    }
                    if (splitedData.Length == 2 && int.TryParse(splitedData[1], out int decimals))
                    {
                        dte.Decimals = decimals;
                    }
                }
            }
            catch (Exception ex) //never fail because a wrong length/decimals definition, just use the defaults values.
            {
                Console.WriteLine("Error parsing Data Type for row " + row);
                HandleException(sheet.Name, ex);
            }
        }

        protected abstract void ReadLevelProperties(TLevelElement level, ExcelWorksheet sheet, int row);

        protected abstract void ReadLeafProperties(TLeafElement leaf, ExcelWorksheet sheet, int row);

        private int GetParentLevelId(ExcelWorksheet sheet, int row)
        {
            if (sheet.Cells[row, Configuration.LevelParentIdColumn].Value != null)
            {
                try
                {
                    if (int.TryParse(sheet.Cells[row, Configuration.LevelParentIdColumn].Value?.ToString(), out int parentId))
                        return parentId;
                }
                catch
                {
                }
            }
            return 0;
        }
    }
}
