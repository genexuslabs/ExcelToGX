# ExcelToBC

Command Line utility to allow declaring a GeneXus Transaction in an Excel file and converting it to a GeneXus export file.

You can use just the binary located on the Bin directory of this repository. 

## Sample Execution

```
ExcelToBC.exe /x:Test.xlsx /o:MyExport.xml
```


## Configuration

A key aspect to make it work is the configuration where you specify the locations of certain key cells in the Excel file.

You need to configure the ExcelToBC.exe.config file with the values for:

- The name of the Sheet where the Transaction is declared.
- Row and Column for the TransactionName, Row and Column for the Transaction Description
- Row and Column of the start where the the collection of attributes are specified. 


### Attribute 
- Specify the column for Attribute Name, Attribute Description, Attribute Data Type, Attribute Domain

#### Type
The Data Type can be specified in the same way you write in the GeneXus Transaction editor.

For example: 

Num(8.2), Numeric(8.2) , DateTime, Numeric(4-), Character(20), Char(20), VarChar(300), Numeric(7.2-), etc

- In order to specify when an Attribute is a PK there is an AttributeKeyColumn setting that specify wich column to check and a PKValue to specify what value to search for in this column that say that is Key. The default value is "PK"
- In order to specify when an Attribute allows null there is an AttributeNullableColumn and a NullableValue with default value "?"

### Levels
- In order to identify a Level we must specify the following settings:
  - LevelIdentifierKeyword and LevelCheckColumn, basically you said in what column we need to check for the keyword specified.
  For example the identifier keyword could be "LVL" and the column could be the first one.
   - LevelIdColumn , each level must have an identifier to be referenced by other levels in the LevelParentIdColumn.
   - The value in the LevelIdColumn must be a number (int), the value in LevelParentIdColumn must be a number or empty.  When empty it means that its parent level is the root level.
   0 means the root level too.
   Take into account that the Parent Id can be specified for Attributes too, so you can define several levels and after define for example an attribute of the root level.
   
   - The Level name and description are taken from the columns of Attribute name and description.

### Sample Configuration File

For a Excel like the following:

![Image of Sample](https://github.com/genexuslabs/ExcelToBC/blob/master/sample.png)

The Configuration File should be>

```
<?xml version="1.0" encoding="utf-8" ?>
<configuration>
    <configSections>
        <sectionGroup name="applicationSettings" type="System.Configuration.ApplicationSettingsGroup, System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" >
            <section name="ExcelToBC.Properties.Settings" type="System.Configuration.ClientSettingsSection, System, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089" requirePermission="false" />
        </sectionGroup>
    </configSections>
    <startup> 
        <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.7.1" />
    </startup>
  <applicationSettings>
        <ExcelToBC.Properties.Settings>
            <setting name="TransactionNameRow" serializeAs="String">
                <value>3</value>
            </setting>
            <setting name="TransactionNameCol" serializeAs="String">
                <value>7</value>
            </setting>
            <setting name="TransactionDescCol" serializeAs="String">
                <value>11</value>
            </setting>
            <setting name="AttributeStartRow" serializeAs="String">
                <value>7</value>
            </setting>
            <setting name="AttributeStartColumn" serializeAs="String">
                <value>2</value>
            </setting>
            <setting name="AttributeNameColumn" serializeAs="String">
                <value>7</value>
            </setting>
            <setting name="AttributeDescriptionColumn" serializeAs="String">
                <value>6</value>
            </setting>
            <setting name="AttributeNullableColumn" serializeAs="String">
                <value>4</value>
            </setting>
            <setting name="AttributeKeyColumn" serializeAs="String">
                <value>3</value>
            </setting>
            <setting name="AttributeDataTypeColumn" serializeAs="String">
                <value>8</value>
            </setting>
            <setting name="TransactionDefinitionSheetName" serializeAs="String">
                <value>てすと</value>
            </setting>
            <setting name="TransactionDescRow" serializeAs="String">
                <value>3</value>
            </setting>
            <setting name="LevelCheckColumn" serializeAs="String">
                <value>3</value>
            </setting>
            <setting name="LevelIdColumn" serializeAs="String">
                <value>2</value>
            </setting>
            <setting name="LevelParentIdColumn" serializeAs="String">
                <value>7</value>
            </setting>
            <setting name="LevelIdentifierKeyword" serializeAs="String">
                <value> レベル1</value>
            </setting>
            <setting name="PKValue" serializeAs="String">
                <value>PK</value>
            </setting>
            <setting name="NullableValue" serializeAs="String">
                <value>〇</value>
            </setting>
            <setting name="DomainColumn" serializeAs="String">
                <value>10</value>
            </setting>
        </ExcelToBC.Properties.Settings>
    </applicationSettings>
   </configuration>
```
## Command Line Tool Specification

The ExcelToBC is a command line tool with the following specification


ExcelToBC v1.0.0.0
Copyright GeneXus c  2018
Allow to convert a transaction specified in Excel to a Genexus Export file

Usage: ExcelToBC.exe [@argfile] [/ExcelFile|x:<value>] [/Directory|d:<value>]
       [/OutputFile|o:<value>] [/ContinueOnErrors|c:<value>] [/help|?|h] [/version|v]


@argfile                   Read arguments from a file.
/ExcelFile:<value>         Uri of Excel File, could be relative to this exe or
                           absoulte
/Directory:<value>         Directory to process all xlsx files inside, could be
                           relative to this exe or absoulte
/OutputFile:<value>        The relative or full path to the output file, the
                           output is in xml format (Default is
                           "Transaction.xml")
/ContinueOnErrors:<value>  Specify if the the tool must continue converting
                           even errors are detected  (Default is "False")
/help                      Show usage.
/version                   Show version.
```


