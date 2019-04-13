# Copyright 2017 Tobias Heilig
#
# Redistribution and use in source and binary forms, with or without modification, are
# permitted provided that the following conditions are met:
#
# 1. Redistributions of source code must retain the above copyright notice, this list
# of conditions and the following disclaimer.
# 2. Redistributions in binary form must reproduce the above copyright notice, this list
# of conditions and the following disclaimer in the documentation and/or other materials
# provided with the distribution.
# 3. Neither the name of the copyright holder nor the names of its contributors may be used
# to endorse or promote products derived from this software without specific prior written
# permission.
#
# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS
# OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY
# AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR
# CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
# DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
# DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER
# IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT
# OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

<# 
    .SYNOPSIS 
    Query a database. 
 
    .DESCRIPTION 
    Execute SQL queries using ADO.NET. 
 
    .PARAMETER Query 
    The SQL query to execute which could be any string containing simple valid SQL. Can consist of
    multiple SQL commands separated by semicolons which will be processed sequentially by the Cmdlet.
 
    .PARAMETER Provider 
    The ADO.NET dataprovider used for accessing the data source to query. Can be any of the following:
    Sql - Provides data access for Microsoft SQL Server
    OleDb - For data sources exposed by using OLE DB
    Odbc - For data sources exposed by using ODBC
    Oracle - For Oracle data sources
    Entity - Provides data access for Entity Data Model (EDM) applications
    SqlCe - Provides data access for Microsoft SQL Server Compact 4.0
 
    .PARAMETER ConnectionString 
    The database connection string.
	
    .OUTPUTS
    On SELECT commands, matching rows will be returned from the database as PowerShell objects.
    Other commands like INSERT or DELETE are executed without any return value. The amount
    of rows in the database affected by those commands will be shown when the -Verbose switch
    is specified, though.
	
    .COMPONENT
    ADO.NET
 
    .EXAMPLE 
    'SELECT * FROM Customers' | Invoke-SqlQuery -Provider OleDb -ConnectionString 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=.\db1.accdb' 
 
    .EXAMPLE 
    $query = "INSERT INTO Customers VALUES ('Doe','John');SELECT * FROM Customers WHERE Surname='John'" 
    $query | Invoke-SqlQuery -Provider Sql -ConnectionString 'Server=db1.contoso.com;Database=CustomersDb;User Id=Admin;Password=pwd' 
 
    .EXAMPLE 
    $query1 = "INSERT INTO Customers VALUES ('Doe','John')" 
    $query2 = "SELECT * FROM Customers WHERE Surname='John'" 
    $query1, query2 | query Sql 'Server=db1.contoso.com;Database=CustomersDb;User Id=Admin;Password=pwd' 
#>
function Invoke-SqlQuery
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [String]
        $Query,

        [Parameter(Mandatory=$true, Position=0)]
        [ValidateSet('Sql', 'OleDb', 'Odbc', 'Oracle', 'Entity', 'SqlCe')]
        [String]
        $Provider,

        [Parameter(Mandatory=$true, Position=1)]
        [String]
        $ConnectionString
    )

    begin
    {
        $connection, $command = switch ($Provider)
        {
            'Sql'
            {
                New-Object System.Data.SqlClient.SqlConnection $ConnectionString
                New-Object System.Data.SqlClient.SqlCommand
            }
            'OleDb'
            {
                New-Object System.Data.OleDb.OleDbConnection $ConnectionString
                New-Object System.Data.OleDb.OleDbCommand
            }
            'Odbc'
            {
                New-Object System.Data.Odbc.OdbcConnection $ConnectionString
                New-Object System.Data.Odbc.OdbcCommand
            }
            'Oracle'
            {
                New-Object System.Data.OracleClient.OracleConnection $ConnectionString
                New-Object System.Data.OracleClient.OracleCommand
            }
            'Entity'
            {
                New-Object System.Data.EntityClient.EntityConnection $ConnectionString
                New-Object System.Data.EntityClient.EntityCommand
            }
            'SqlCe'
            {
                New-Object System.Data.SqlServerCe.SqlCeConnection $ConnectionString
                New-Object System.Data.SqlServerCe.SqlCeCommand
            }
        }
        $connection.Open()
        $command.Connection = $connection
    }

    process
    {
        $_ -split ';' | ForEach-Object {
            if (-not [String]::IsNullOrWhiteSpace($_))
            {
                $command.CommandText = $_.Trim()

                if ($command.CommandText -like 'SELECT*')
                {
                    $reader = $command.ExecuteReader()
                    while ($reader.Read())
                    {
                        $row = [ordered]@{}
                        0..($reader.FieldCount-1) | ForEach-Object {
                            $row.Add($reader.GetName($_) , $reader[$_])
                        }
                        New-Object PSObject -Property $row
                    }
                    $reader.Close()
                }
                else
                {
                    $affectedRows = $command.ExecuteNonQuery()
                    Write-Verbose "$affectedRows row(s) affected by $(($command.CommandText -split '\s+')[0]) command"
                }
            }
        }
    }

    end
    {
        $connection.Close()
    }
}

New-Alias -Name query -Value Invoke-SqlQuery
