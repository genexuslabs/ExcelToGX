using Artech.Common.Helpers;
using ExcelParser;
using ExcelToGX.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace GeneXus.Utilities
{
    class Program
    {
        static void Main(string[] args)
        {
            ConvertCommandLine cmd = new ConvertCommandLine();
            try
            {
                cmd.Parse(args);
                if (cmd.help)
                {
                    Console.WriteLine(cmd.GetUsage());
                    return;
                }
            }
            catch
            {
                Console.WriteLine(cmd.GetUsage());
                return;
            }

            ExcelReader reader = CreateReader(cmd);
            if (reader is null)
                return;

            try
            {
                if (cmd.ExcelFile == null && cmd.Directory == null)
                {
                    Console.WriteLine("Specify the Directory to scan for several xlsx files or the ExcelFile option for an specific file");
                    return;
                }
                string outputXml = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), cmd.OutputFile);
                if (cmd.ExcelFile != null)
                {
                    string inputXlsx = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), cmd.ExcelFile);
                    if (!string.IsNullOrEmpty(inputXlsx))
                    {
                        if (!File.Exists(inputXlsx))
                            Console.WriteLine($"Input File {inputXlsx} does not exists, please specify an existing xlsx file");
                        else
                            reader.ReadExcel(new string[] { inputXlsx }, outputXml, cmd.ContinueOnErrors);
                    }
                }
                else if (cmd.Directory != null)
                {
                    string inputDirectory = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), cmd.Directory);
                    if (!string.IsNullOrEmpty(inputDirectory) && Directory.Exists(inputDirectory))
                    {
                        string[] files = Directory.GetFiles(inputDirectory, "*.xlsx");
                        List<string> fullPaths = new List<string>();
                        foreach (string file in files)
                            fullPaths.Add(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), file));

                        if (files.Length > 0)
                            reader.ReadExcel(fullPaths.ToArray(), outputXml, cmd.ContinueOnErrors);
                        else
                            Console.WriteLine("Could not find any xlsx on the given directory " + inputDirectory);
                    }
                }
                else
                {
                    Console.WriteLine("Please specify a valid directory to explore for xlsx files");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static ExcelReader CreateReader(ConvertCommandLine cmd)
        {
            switch (cmd.Type)
            {
                case "TRN":
                    var trnReader = new TRNExcelReader();
                    if (ReadConfiguration(trnReader))
                        return trnReader;
                    return null;
                case "SDT":
                    var sdtReader = new SDTExcelReader();
                    if (ReadConfiguration(sdtReader))
                        return sdtReader;
                    return null;
                case "GRP":
                    var grpReader = new GRPExcelReader();
                    if (ReadConfiguration(grpReader))
                        return grpReader;
                    return null;
                default:
                    Console.WriteLine("Invalid value for Type argument");
                    return null;
            }
        }

        public static bool CheckConfigFileIsPresent()
        {
            return File.Exists(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
        }

        private static bool ReadBaseConfiguration<TConfig, TObject>(ExcelReader<TConfig, TObject> reader)
            where TConfig : ExcelReader.BaseConfiguration, new()
            where TObject : IKBElement, new()
        {
            if (!CheckConfigFileIsPresent())
            {
                Console.WriteLine("Missing Config file, Config file is used to configurate all settings for this tool, see documentation on https://github.com/genexuslabs/ExcelToGX");
                return false;
            }

            reader.Configuration.ObjectNameRow = Settings.Default.ObjectNameRow;
            reader.Configuration.ObjectNameColumn = Settings.Default.ObjectNameColumn;
            reader.Configuration.ObjectDescRow = Settings.Default.ObjectDescRow;
            reader.Configuration.ObjectDescColumn = Settings.Default.ObjectDescColumn;
            reader.Configuration.DefinitionSheetName = Settings.Default.DefinitionSheetName;

            reader.Configuration.DataStartRow = Settings.Default.DataStartRow;
            reader.Configuration.DataStartColumn = Settings.Default.DataStartColumn;
            reader.Configuration.DataNameColumn = Settings.Default.DataNameColumn;
            reader.Configuration.DataDescriptionColumn = Settings.Default.DataDescriptionColumn;

            return true;
        }

        private static bool ReadBaseDataConfiguration<TConfig, IType, TLeafElement, TLevelElement>(DataExcelReader<TConfig, IType, TLeafElement, TLevelElement> reader)
            where TConfig : DataExcelReader<TConfig, IType, TLeafElement, TLevelElement>.BaseDataConfiguration, new()
            where IType : IKBElement
            where TLeafElement : DataTypeElement, IType, new()
            where TLevelElement : LevelElement<IType>, IType, new()
        {
            if (!ReadBaseConfiguration(reader))
                return false;

            reader.Configuration.LevelCheckColumn = Settings.Default.LevelCheckColumn;
            reader.Configuration.LevelIdColumn = Settings.Default.LevelIdColumn;
            reader.Configuration.LevelIdentifierKeyword = Settings.Default.LevelIdentifierKeyword;
            reader.Configuration.LevelParentIdColumn = Settings.Default.LevelParentIdColumn;

            reader.Configuration.DataTypeColumn = Settings.Default.DataTypeColumn;
            reader.Configuration.DataLengthColumn = Settings.Default.DataLengthColumn;

            reader.Configuration.BaseTypeColumn = Settings.Default.BaseTypeColumn;
            return true;
        }

        private static bool ReadConfiguration(TRNExcelReader reader)
        {
            if (!ReadBaseDataConfiguration(reader))
                return false;

            reader.Configuration.AttributeKeyColumn = Settings.Default.AttributeKeyColumn;
            reader.Configuration.AttributeNullableColumn = Settings.Default.AttributeNullableColumn;
            reader.Configuration.AttributeAutonumberColumn = Settings.Default.AttributeAutonumberColumn;
            reader.Configuration.AttributeTitleColumn = Settings.Default.AttributeTitleColumn;
            reader.Configuration.AttributeColumnTitleColumn = Settings.Default.AttributeColumnTitleColumn;
            reader.Configuration.AttributeContextualTitleColumn = Settings.Default.AttributeContextualTitleColumn;
            // Values for PK and Nullable
            reader.Configuration.PKValue = Settings.Default.PKValue;
            reader.Configuration.NullableValue = Settings.Default.NullableValue;
            // Formula
            reader.Configuration.FormulaCheckColumn = Settings.Default.FormulaCheckColumn;
            reader.Configuration.FormulaIdentifierKeyword = Settings.Default.FormulaIdentifierKeyword;
            reader.Configuration.FormulaColumn = Settings.Default.FormulaColumn;
            return true;
        }

        private static bool ReadConfiguration(SDTExcelReader reader)
        {
            if (!ReadBaseDataConfiguration(reader))
                return false;

            reader.Configuration.ItemIsCollectionColumn = Settings.Default.ItemIsCollectionColumn;
            reader.Configuration.CollectionIdentifierKeyword = Settings.Default.CollectionIdentifierKeyword;
            reader.Configuration.CollectionItemNameColumn = Settings.Default.CollectionItemNameColumn;

            reader.Configuration.DomainPrefixKeyword = Settings.Default.DomainPrefixKeyword;
            reader.Configuration.AttributePrefixKeyword = Settings.Default.AttributePrefixKeyword;
            reader.Configuration.SDTPrefixKeyword = Settings.Default.SDTPrefixKeyword;
            reader.Configuration.DefaultBaseTypePrefixKeyword = Settings.Default.DefaultBaseTypePrefixKeyword;

            return true;
        }

        private static bool ReadConfiguration(GRPExcelReader reader)
        {
            if (!ReadBaseConfiguration(reader))
                return false;

            reader.Configuration.SupertypeColumn = Settings.Default.SupertypeColumn;
            return true;
        }
    }

    public class ConvertCommandLine : CommandLineParser
    {
        [ValueUsage("Uri of Excel File, could be relative to this exe or absoulte", Optional = true, AlternateName1 = "x")]
        public string ExcelFile;

        [ValueUsage("Directory to process all xlsx files inside, could be relative to this exe or absoulte", Optional = true, AlternateName1 = "d")]
        public string Directory;

        [ValueUsage("The relative or full path to the output file, the output is in xml format", Optional = true, AlternateName1 = "o")]
        public string OutputFile = "Transaction.xml";

        [ValueUsage("Specify if the tool must continue converting after errors are detected", Optional = true, AlternateName1 = "c")]
        public bool ContinueOnErrors = false;

        [ValueUsage(@"Specify the type of worksheet to parse.
Values: TRN | SDT | GRP
Default: TRN", Optional = true, AlternateName1 = "t")]
        public string Type = "TRN";
    }
}
