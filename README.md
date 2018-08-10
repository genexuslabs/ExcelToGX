# ExcelToBC
Utility to create a Genexus BC from an excel file



This utility allow to convert an excel file where a Transaction is described in an XML file.

In order to works correctly you need to configure the ExcelToBC.config file with the values for:

- Row and Column for the TransactionName, Column for the Transaction Description (row is assumed the same)

- Row and Column of the start where the the collection of attributes are specified. 

- Specify the column for Attribute Name, Attribute Description, Attribute Data Type, Attribute Length (length is specified as <length>[,<decimals>]

Sample Configuration File

```
<?xml version='1.0' encoding='utf-8'?>
<SettingsFile xmlns="http://schemas.microsoft.com/VisualStudio/2004/01/settings" CurrentProfile="(Default)" GeneratedClassNamespace="ExcelToBC.Properties" GeneratedClassName="Settings">
  <Profiles />
  <Settings>
    <Setting Name="TransactionNameRow" Type="System.Int32" Scope="Application">
      <Value Profile="(Default)">3</Value>
    </Setting>
    <Setting Name="TransactionNameCol" Type="System.Int32" Scope="Application">
      <Value Profile="(Default)">7</Value>
    </Setting>
    <Setting Name="TransactionDescCol" Type="System.Int32" Scope="Application">
      <Value Profile="(Default)">11</Value>
    </Setting>
    <Setting Name="AttributeStartRow" Type="System.Int32" Scope="Application">
      <Value Profile="(Default)">7</Value>
    </Setting>
    <Setting Name="AttributeStartColumn" Type="System.Int32" Scope="Application">
      <Value Profile="(Default)">2</Value>
    </Setting>
    <Setting Name="AttributeNameColumn" Type="System.Int32" Scope="Application">
      <Value Profile="(Default)">7</Value>
    </Setting>
    <Setting Name="AttributeDescriptionColumn" Type="System.Int32" Scope="Application">
      <Value Profile="(Default)">6</Value>
    </Setting>
    <Setting Name="AttributeNullableColumn" Type="System.Int32" Scope="Application">
      <Value Profile="(Default)">4</Value>
    </Setting>
    <Setting Name="AttributeKeyColumn" Type="System.Int32" Scope="Application">
      <Value Profile="(Default)">3</Value>
    </Setting>
    <Setting Name="AttributeDataTypeColumn" Type="System.Int32" Scope="Application">
      <Value Profile="(Default)">8</Value>
    </Setting>
    <Setting Name="AttributeDataLengthColumn" Type="System.Int32" Scope="Application">
      <Value Profile="(Default)">9</Value>
    </Setting>
    <Setting Name="TransactionDefinitionSheetName" Type="System.String" Scope="Application">
      <Value Profile="(Default)">TransactionDefinitionSheet</Value>
    </Setting>
  </Settings>
</SettingsFile>
```
