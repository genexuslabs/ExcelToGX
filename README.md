# ExcelToGX

Command Line utility to allow declaring a GeneXus Transaction or Structured Data Type in an Excel file and convert it to a GeneXus export file. The type of object to be scanned from the Excel file must be specified as an argument to the tool.

The compiled binaries of the tool are located on the Bin directory of this repository. 

## Sample Executions

Convert a xlsx file to a GeneXus export
```
ExcelToGX.exe /x:TRN_Test.xlsx /o:TRN_Export.xml
ExcelToGX.exe /x:SDT_Test.xlsx /o:SDT_Export.xml /t:SDT
```
Note that the default type of definition file to analyze is for Transactions.
Scan the given directory looking for .xlsx files and create a file with all the transaction found.
```
ExcelToGX.exe /d:MyDefinitionsDirectory /o:MyExport.xml
```
When creating a export file from several xlsx files, all xlsx must be for the same type of definition, either Transactions or Structured Data Types, but not both at the same time. In the case a conflict is identified attribute definitions, for instance when the same attribute is in different files with different data types, the first definition for the attribute is used and a warning is raised to inform the conflict.


## Configuration

A key aspect to make it work is the configuration where you specify the locations of certain key cells in the Excel file.

You need to configure the ExcelToGX.exe.config file with the values for:

- The name of the Sheet where the structure is declared.
- Row and Column for the name and description of the object being declared.
- Row and Column of the start where the structure is specified.
- Columns for certain type of definitions. For instance, in which column expect to find type declarations, or what column to check if a new level started, and many more.


### Data 
- Specify the column for Data Name, Data Description, Data Type, Domain

#### Type
The Data Type can be specified in the same way you write in the GeneXus Transaction editor. Just by using Data Type Column.

For example: 

Num(8.2), Numeric(8.2) , DateTime, Numeric(4-), Character(20), Char(20), VarChar(300), Numeric(7.2-), etc

Or is possible to use a separated column for Length and Decimals. In this case you use the Data Type Column just for the Type name and the DataLengthColumn in order to configure the Data Length, Decimals and Sign.

#### Domain
When you specify a value for the Domain column the attribute or item is defined based on this Domain. In general the Type column should be empty, it depends if you are just referencing the Domain or if you want to define the Data Type for the Domain.
When the Domain column has a value the Type column is considered the Data Type for the given Domain. 

- In order to specify when an Attribute is a PK there is an AttributeKeyColumn setting that specify wich column to check and a PKValue to specify what value to search for in this column that say that is Key. The default value is "PK"
- In order to specify when an Attribute allows null there is an AttributeNullableColumn and a NullableValue with default value "?"

### Levels
- In order to identify a Level we must specify the following settings:
  - LevelIdentifierKeyword and LevelCheckColumn, basically they state in what column we need to check for the keyword specified.
  For example the identifier keyword could be "LVL" and the column could be the first one.
   - LevelIdColumn , each level must have an identifier to be referenced by other levels in the LevelParentIdColumn.
   - The value in the LevelIdColumn must be a number (int), the value in LevelParentIdColumn must be a number or empty.  When empty it means that its parent level is the root level.
   0 means the root level too.
   Take into account that the Parent Id can be specified for Attributes too, so you can define several levels and after define for example an attribute of the root level.
   
   - The Level name and description are taken from the columns of Data name and description.

### Sample Configuration File

For a Excel like the following:

![Image of Sample](https://github.com/genexuslabs/ExcelToBC/blob/master/sample.png)

The imported Transaction in GeneXus will be

![Image of Result](https://github.com/genexuslabs/ExcelToBC/blob/master/importedTrn.png)


The Configuration File should be

```xml
<?xml version="1.0" encoding="utf-8" ?>
<configuration>
  <configSections>
    <sectionGroup name="applicationSettings" type="System.Configuration.ApplicationSettingsGroup, System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" >
      <section name="ExcelToGX.Properties.Settings" type="System.Configuration.ClientSettingsSection, System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" requirePermission="false" />
    </sectionGroup>
  </configSections>
  <startup> 
    <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.7.1" />
  </startup>
  <applicationSettings>
    <ExcelToGX.Properties.Settings>
      <setting name="DataStartRow" serializeAs="String">
        <value>5</value>
      </setting>
      <setting name="DataStartColumn" serializeAs="String">
        <value>1</value>
      </setting>
      <setting name="DataNameColumn" serializeAs="String">
        <value>1</value>
      </setting>
      <setting name="DataDescriptionColumn" serializeAs="String">
        <value>2</value>
      </setting>
      <setting name="AttributeNullableColumn" serializeAs="String">
        <value>6</value>
      </setting>
      <setting name="AttributeKeyColumn" serializeAs="String">
        <value>3</value>
      </setting>
      <setting name="DataTypeColumn" serializeAs="String">
        <value>4</value>
      </setting>
      <setting name="LevelCheckColumn" serializeAs="String">
        <value>3</value>
      </setting>
      <setting name="LevelIdColumn" serializeAs="String">
        <value>8</value>
      </setting>
      <setting name="LevelParentIdColumn" serializeAs="String">
        <value>9</value>
      </setting>
      <setting name="LevelIdentifierKeyword" serializeAs="String">
        <value>LVL</value>
      </setting>
      <setting name="PKValue" serializeAs="String">
        <value>PK</value>
      </setting>
      <setting name="NullableValue" serializeAs="String">
        <value>?</value>
      </setting>
      <setting name="DomainColumn" serializeAs="String">
        <value>10</value>
      </setting>
      <setting name="DataLengthColumn" serializeAs="String">
        <value>5</value>
      </setting>
      <setting name="ObjectNameRow" serializeAs="String">
        <value>2</value>
      </setting>
      <setting name="ObjectNameColumn" serializeAs="String">
        <value>1</value>
      </setting>
      <setting name="ObjectDescRow" serializeAs="String">
        <value>2</value>
      </setting>
      <setting name="ObjectDescColumn" serializeAs="String">
        <value>2</value>
      </setting>
      <setting name="DefinitionSheetName" serializeAs="String">
        <value>TransactionDefinitionSheet</value>
      </setting>
    </ExcelToGX.Properties.Settings>
  </applicationSettings>
</configuration>
```
## Command Line Tool Specification

The ExcelToGX is a command line tool with the following specification

```
ExcelToGX v1.0.0.0
Copyright GeneXus c  2018
Allow to convert definitions in Excel to a Genexus Export file

Usage: ExcelToGX.exe [@argfile] [/ExcelFile|x:<value>] [/Directory|d:<value>]
       [/OutputFile|o:<value>] [/ContinueOnErrors|c:<value>] [/Type|t:<value>] [/help|?|h] [/version|v]


@argfile                   Read arguments from a file.
/ExcelFile:<value>         Uri of Excel File, could be relative to this exe or
                           absoulte
/Directory:<value>         Directory to process all xlsx files inside, could be
                           relative to this exe or absoulte
/OutputFile:<value>        The relative or full path to the output file, the
                           output is in xml format (Default is
                           "Transaction.xml")
/ContinueOnErrors:<value>  Specify if the tool must continue converting
                           after errors are detected (Default is "False")
/Type:<value>              Specify the type of worksheet to parse.
                           Values: TRN | SDT
                           (Default is "TRN")
/help                      Show usage.
/version                   Show version.
```


