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
			TRNExcelReader reader = new TRNExcelReader();
            if (!ReadConfiguration(reader))
                return;

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

        public static bool CheckConfigFileIsPresent()
        {
            return File.Exists(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile);
        }

        private static bool ReadConfiguration(TRNExcelReader reader)
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
            reader.Configuration.AttributeStartRow = Settings.Default.AttributeStartRow;
            reader.Configuration.AttributesStartColumn = Settings.Default.AttributeStartColumn;
            reader.Configuration.AttributeNameColumn = Settings.Default.AttributeNameColumn;
            reader.Configuration.AttributeDescriptionColumn = Settings.Default.AttributeDescriptionColumn;
            reader.Configuration.AttributeDataTypeColumn = Settings.Default.AttributeDataTypeColumn;
			reader.Configuration.AttributeDataLengthColumn = Settings.Default.AttributeDataLengthColumn;

            reader.Configuration.AttributeKeyColumn = Settings.Default.AttributeKeyColumn;
            reader.Configuration.AttributeNullableColumn = Settings.Default.AttributeNullableColumn;
            reader.Configuration.DefinitionSheetName = Settings.Default.DefinitionSheetName;
            reader.Configuration.LevelCheckColumn = Settings.Default.LevelCheckColumn;
            reader.Configuration.LevelIdColumn = Settings.Default.LevelIdColumn;
            reader.Configuration.LevelIdentifierKeyword = Settings.Default.LevelIdentifierKeyword;
			reader.Configuration.LevelParentIdColumn = Settings.Default.LevelParentIdColumn;
            // Values for PK and Nullable
            reader.Configuration.PKValue = Settings.Default.PKValue;
            reader.Configuration.NullableValue = Settings.Default.NullableValue;
			reader.Configuration.AttributeDomainColumn = Settings.Default.DomainColumn;
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

        [ValueUsage("Specify if the the tool must continue converting even errors are detected ", Optional = true, AlternateName1 = "c")]
        public bool ContinueOnErrors = false;
    }
}
