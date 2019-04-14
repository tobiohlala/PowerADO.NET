# PowerADO.NET

[![Codacy Badge](https://api.codacy.com/project/badge/Grade/fa05ad4e54b345c88807ff4454c845b0)](https://www.codacy.com/app/off-world/PowerADO.NET?utm_source=github.com&amp;utm_medium=referral&amp;utm_content=off-world/PowerADO.NET&amp;utm_campaign=Badge_Grade)

Query relational databases from PowerShell using SQL

## Description

With ADO.NET the .NET Framework provides a collection of classes for accessing relational databases. Existing data providers
such as for example OleDB or Odbc allow you to work with a variety of relational databases - from local Microsoft Access
databases and Excel spreadsheets to Microsoft SQL Server, MySQL databases or Oracle databases. Because PowerShell provides
access to .NET classes, ADO.NET can also be addressed through PowerShell.

In order to establish a connection to a database and execute SQL queries, two pieces of information are required which can be
communicated to the Cmdlet via parameters. These are the **data provider** and the **connection string** of the database. 
Currently, the following default data providers are supported by the `System.Data` namespace:
-   SqlClient
-   OleDb
-   Odbc
-   OracleClient
-   EntityClient
-   SqlCeClient

A collection of common connection strings for different database technologies can be found [here](www.connectionstrings.com).

SQL queries can be passed as strings to the Cmdlet using the PowerShell Pipeline and will be processed sequentially. In turn,
a string can contain multiple queries separated by a semicolon. If it is a SELECT query, the requested data is returned from 
the database in the form of PowerShell objects, allowing for easy further processing. All other requests are executed without
any return value. However, you can use the `-Verbose` parameter to see how many rows were affected by the query.

## Installation

Install from [PowerShell Gallery](https://www.powershellgallery.com/packages/PowerADO.NET)

```Powershell
Install-Module -Name PowerADO.NET
```

## Usage

```Powershell
Import-Module PowerADO.NET
```

## Examples

```Powershell
Get-Help Invoke-SqlQuery -Examples
```
