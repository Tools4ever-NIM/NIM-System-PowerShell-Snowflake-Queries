#
# Snowflake Queries.ps1 - IDM System PowerShell script for Snowflake database
#

$Log_MaskableKeys = @(
    'password'
)

# Check if ODBC Driver installed
$driverPrefix = "Snowflake"
$drivers64 = Get-ItemProperty -Path "HKLM:\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers" -ErrorAction SilentlyContinue
$matchingDrivers = $drivers64.PSObject.Properties | Where-Object { $_.Name -like "$driverPrefix*" }

$Global:ModuleStatus = "<b><div class=`"alert alert-danger`" role=`"alert`">$($driverPrefix) ODBC Driver not installed.</div></b>"

if ($matchingDrivers.Count -gt 0) {
    "Found matching ODBC driver(s):"
    $matchingDrivers | ForEach-Object { " - $($_.Name)" }
	$Global:ModuleStatus = "<b><div class=`"alert alert-success`" role=`"alert`">$($matchingDrivers[0].Name) is installed.</div></b>"
} else {
    "No ODBC drivers found starting with '$($driverPrefix)'."
}

#
# System functions
#

function Idm-SystemInfo {
    param (
        # Operations
        [switch] $Connection,
        [switch] $TestConnection,
        [switch] $Configuration,
        # Parameters
        [string] $ConnectionParams
    )

    Log verbose "-Connection=$Connection -TestConnection=$TestConnection -Configuration=$Configuration -ConnectionParams='$ConnectionParams'"
    
    if ($Connection) {
        @(      
			@{
				name = 'ModuleStatus'
				type = 'text'
				label = 'ODBC Status'
				text = $Global:ModuleStatus
			}
			@{
                name = 'connection_header'
                type = 'text'
                text = 'Connection'
				tooltip = 'Connection information for the database'
            }
			@{
                name = 'host_name'
                type = 'textbox'
                label = 'Server'
                description = 'IP or Hostname of Server'
                value = ''
            }
            @{
                name = 'database'
                type = 'textbox'
                label = 'Database'
                description = 'Name of database'
                value = ''
            }
            @{
                name = 'user'
                type = 'textbox'
                label = 'Username'
                label_indent = $true
                description = 'User account name to access server'
            }
            @{
                name = 'password'
                type = 'textbox'
                password = $true
                label = 'Password'
                label_indent = $true
                description = 'User account password to access server'
            }
            @{
                name = 'query_timeout'
                type = 'textbox'
                label = 'Query Timeout'
                description = 'Time it takes for the query to timeout'
                value = '1800'
            }
			@{
                name = 'connection_timeout'
                type = 'textbox'
                label = 'Connection Timeout'
                description = 'Time it takes for the ODBC Connection to timeout'
                value = '3600'
            }
			@{
                name = 'session_header'
                type = 'text'
                text = 'Session Options'
				tooltip = 'Options for system session'
            }
			@{
                name = 'nr_of_sessions'
                type = 'textbox'
                label = 'Max. number of simultaneous sessions'
                tooltip = ''
                value = 1
            }
            @{
                name = 'sessions_idle_timeout'
                type = 'textbox'
                label = 'Session cleanup idle time (minutes)'
                tooltip = ''
                value = 1
            }
			@{
                name = 'table_header'
                type = 'text'
                text = 'Tables'
				tooltip = 'Query to Table mapping'
            }
			@{
                name = 'table_1_header'
                type = 'text'
                text = 'Table 1'
				tooltip = 'Table 1 Config'
            }
            @{
                name = 'table_1_name'
                type = 'textbox'
                label = 'Table 1 Name'
                description = ''
            }
            @{
                name = 'table_1_query'
                type = 'textbox'
                label = 'Table 1 Query'
                description = ''
				disabled = '!table_1_name'
                hidden = '!table_1_name'
            }
			
			@{
				name = 'table_2_header'
				type = 'text'
				text = 'Table 2'
				tooltip = 'Table 2 Config'
				disabled = '!table_1_name'
				hidden = '!table_1_name'
			}
			@{
				name = 'table_2_name'
				type = 'textbox'
				label = 'Table 2 Name'
				description = ''
				disabled = '!table_1_name'
				hidden = '!table_1_name'
			}
			@{
				name = 'table_2_query'
				type = 'textbox'
				label = 'Table 2 Query'
				description = ''
				disabled = '!table_1_name'
				hidden = '!table_1_name'
			}

			@{
				name = 'table_3_header'
				type = 'text'
				text = 'Table 3'
				tooltip = 'Table 3 Config'
				disabled = '!table_2_name'
				hidden = '!table_2_name'
			}
			@{
				name = 'table_3_name'
				type = 'textbox'
				label = 'Table 3 Name'
				description = ''
				disabled = '!table_2_name'
				hidden = '!table_2_name'
			}
			@{
				name = 'table_3_query'
				type = 'textbox'
				label = 'Table 3 Query'
				description = ''
				disabled = '!table_2_name'
				hidden = '!table_2_name'
			}

			@{
				name = 'table_4_header'
				type = 'text'
				text = 'Table 4'
				tooltip = 'Table 4 Config'
				disabled = '!table_3_name'
				hidden = '!table_3_name'
			}
			@{
				name = 'table_4_name'
				type = 'textbox'
				label = 'Table 4 Name'
				description = ''
				disabled = '!table_3_name'
				hidden = '!table_3_name'
			}
			@{
				name = 'table_4_query'
				type = 'textbox'
				label = 'Table 4 Query'
				description = ''
				disabled = '!table_3_name'
				hidden = '!table_3_name'
			}

			@{
				name = 'table_5_header'
				type = 'text'
				text = 'Table 5'
				tooltip = 'Table 5 Config'
				disabled = '!table_4_name'
				hidden = '!table_4_name'
			}
			@{
				name = 'table_5_name'
				type = 'textbox'
				label = 'Table 5 Name'
				description = ''
				disabled = '!table_4_name'
				hidden = '!table_4_name'
			}
			@{
				name = 'table_5_query'
				type = 'textbox'
				label = 'Table 5 Query'
				description = ''
				disabled = '!table_4_name'
				hidden = '!table_4_name'
			}

			@{
				name = 'table_6_header'
				type = 'text'
				text = 'Table 6'
				tooltip = 'Table 6 Config'
				disabled = '!table_5_name'
				hidden = '!table_5_name'
			}
			@{
				name = 'table_6_name'
				type = 'textbox'
				label = 'Table 6 Name'
				description = ''
				disabled = '!table_5_name'
				hidden = '!table_5_name'
			}
			@{
				name = 'table_6_query'
				type = 'textbox'
				label = 'Table 6 Query'
				description = ''
				disabled = '!table_5_name'
				hidden = '!table_5_name'
			}

			@{
				name = 'table_7_header'
				type = 'text'
				text = 'Table 7'
				tooltip = 'Table 7 Config'
				disabled = '!table_6_name'
				hidden = '!table_6_name'
			}
			@{
				name = 'table_7_name'
				type = 'textbox'
				label = 'Table 7 Name'
				description = ''
				disabled = '!table_6_name'
				hidden = '!table_6_name'
			}
			@{
				name = 'table_7_query'
				type = 'textbox'
				label = 'Table 7 Query'
				description = ''
				disabled = '!table_6_name'
				hidden = '!table_6_name'
			}

			@{
				name = 'table_8_header'
				type = 'text'
				text = 'Table 8'
				tooltip = 'Table 8 Config'
				disabled = '!table_7_name'
				hidden = '!table_7_name'
			}
			@{
				name = 'table_8_name'
				type = 'textbox'
				label = 'Table 8 Name'
				description = ''
				disabled = '!table_7_name'
				hidden = '!table_7_name'
			}
			@{
				name = 'table_8_query'
				type = 'textbox'
				label = 'Table 8 Query'
				description = ''
				disabled = '!table_7_name'
				hidden = '!table_7_name'
			}

			@{
				name = 'table_9_header'
				type = 'text'
				text = 'Table 9'
				tooltip = 'Table 9 Config'
				disabled = '!table_8_name'
				hidden = '!table_8_name'
			}
			@{
				name = 'table_9_name'
				type = 'textbox'
				label = 'Table 9 Name'
				description = ''
				disabled = '!table_8_name'
				hidden = '!table_8_name'
			}
			@{
				name = 'table_9_query'
				type = 'textbox'
				label = 'Table 9 Query'
				description = ''
				disabled = '!table_8_name'
				hidden = '!table_8_name'
			}

			@{
				name = 'table_10_header'
				type = 'text'
				text = 'Table 10'
				tooltip = 'Table 10 Config'
				disabled = '!table_9_name'
				hidden = '!table_9_name'
			}
			@{
				name = 'table_10_name'
				type = 'textbox'
				label = 'Table 10 Name'
				description = ''
				disabled = '!table_9_name'
				hidden = '!table_9_name'
			}
			@{
				name = 'table_10_query'
				type = 'textbox'
				label = 'Table 10 Query'
				description = ''
				disabled = '!table_9_name'
				hidden = '!table_9_name'
			}

			@{
				name = 'table_11_header'
				type = 'text'
				text = 'Table 11'
				tooltip = 'Table 11 Config'
				disabled = '!table_10_name'
				hidden = '!table_10_name'
			}
			@{
				name = 'table_11_name'
				type = 'textbox'
				label = 'Table 11 Name'
				description = ''
				disabled = '!table_10_name'
				hidden = '!table_10_name'
			}
			@{
				name = 'table_11_query'
				type = 'textbox'
				label = 'Table 11 Query'
				description = ''
				disabled = '!table_10_name'
				hidden = '!table_10_name'
			}

			@{
				name = 'table_12_header'
				type = 'text'
				text = 'Table 12'
				tooltip = 'Table 12 Config'
				disabled = '!table_11_name'
				hidden = '!table_11_name'
			}
			@{
				name = 'table_12_name'
				type = 'textbox'
				label = 'Table 12 Name'
				description = ''
				disabled = '!table_11_name'
				hidden = '!table_11_name'
			}
			@{
				name = 'table_12_query'
				type = 'textbox'
				label = 'Table 12 Query'
				description = ''
				disabled = '!table_11_name'
				hidden = '!table_11_name'
			}

			@{
				name = 'table_13_header'
				type = 'text'
				text = 'Table 13'
				tooltip = 'Table 13 Config'
				disabled = '!table_12_name'
				hidden = '!table_12_name'
			}
			@{
				name = 'table_13_name'
				type = 'textbox'
				label = 'Table 13 Name'
				description = ''
				disabled = '!table_12_name'
				hidden = '!table_12_name'
			}
			@{
				name = 'table_13_query'
				type = 'textbox'
				label = 'Table 13 Query'
				description = ''
				disabled = '!table_12_name'
				hidden = '!table_12_name'
			}

			@{
				name = 'table_14_header'
				type = 'text'
				text = 'Table 14'
				tooltip = 'Table 14 Config'
				disabled = '!table_13_name'
				hidden = '!table_13_name'
			}
			@{
				name = 'table_14_name'
				type = 'textbox'
				label = 'Table 14 Name'
				description = ''
				disabled = '!table_13_name'
				hidden = '!table_13_name'
			}
			@{
				name = 'table_14_query'
				type = 'textbox'
				label = 'Table 14 Query'
				description = ''
				disabled = '!table_13_name'
				hidden = '!table_13_name'
			}

			@{
				name = 'table_15_header'
				type = 'text'
				text = 'Table 15'
				tooltip = 'Table 15 Config'
				disabled = '!table_14_name'
				hidden = '!table_14_name'
			}
			@{
				name = 'table_15_name'
				type = 'textbox'
				label = 'Table 15 Name'
				description = ''
				disabled = '!table_14_name'
				hidden = '!table_14_name'
			}
			@{
				name = 'table_15_query'
				type = 'textbox'
				label = 'Table 15 Query'
				description = ''
				disabled = '!table_14_name'
				hidden = '!table_14_name'
			}

			@{
				name = 'table_16_header'
				type = 'text'
				text = 'Table 16'
				tooltip = 'Table 16 Config'
				disabled = '!table_15_name'
				hidden = '!table_15_name'
			}
			@{
				name = 'table_16_name'
				type = 'textbox'
				label = 'Table 16 Name'
				description = ''
				disabled = '!table_15_name'
				hidden = '!table_15_name'
			}
			@{
				name = 'table_16_query'
				type = 'textbox'
				label = 'Table 16 Query'
				description = ''
				disabled = '!table_15_name'
				hidden = '!table_15_name'
			}

			@{
				name = 'table_17_header'
				type = 'text'
				text = 'Table 17'
				tooltip = 'Table 17 Config'
				disabled = '!table_16_name'
				hidden = '!table_16_name'
			}
			@{
				name = 'table_17_name'
				type = 'textbox'
				label = 'Table 17 Name'
				description = ''
				disabled = '!table_16_name'
				hidden = '!table_16_name'
				
			}
			@{
				name = 'table_17_query'
				type = 'textbox'
				label = 'Table 17 Query'
				description = ''
				disabled = '!table_16_name'
				hidden = '!table_16_name'
			}

			@{
				name = 'table_18_header'
				type = 'text'
				text = 'Table 18'
				tooltip = 'Table 18 Config'
				disabled = '!table_17_name'
				hidden = '!table_17_name'
			}
			@{
				name = 'table_18_name'
				type = 'textbox'
				label = 'Table 18 Name'
				description = ''
				disabled = '!table_17_name'
				hidden = '!table_17_name'
			}
			@{
				name = 'table_18_query'
				type = 'textbox'
				label = 'Table 18 Query'
				description = ''
				disabled = '!table_17_name'
				hidden = '!table_17_name'
			}

			@{
				name = 'table_19_header'
				type = 'text'
				text = 'Table 19'
				tooltip = 'Table 19 Config'
				disabled = '!table_18_name'
				hidden = '!table_18_name'
			}
			@{
				name = 'table_19_name'
				type = 'textbox'
				label = 'Table 19 Name'
				description = ''
				disabled = '!table_18_name'
				hidden = '!table_18_name'
			}
			@{
				name = 'table_19_query'
				type = 'textbox'
				label = 'Table 19 Query'
				description = ''
				disabled = '!table_18_name'
				hidden = '!table_18_name'
			}

			@{
				name = 'table_20_header'
				type = 'text'
				text = 'Table 20'
				tooltip = 'Table 20 Config'
				disabled = '!table_19_name'
				hidden = '!table_19_name'
			}
			@{
				name = 'table_20_name'
				type = 'textbox'
				label = 'Table 20 Name'
				description = ''
				disabled = '!table_19_name'
				hidden = '!table_19_name'
			}
			@{
				name = 'table_20_query'
				type = 'textbox'
				label = 'Table 20 Query'
				description = ''
				disabled = '!table_19_name'
				hidden = '!table_19_name'
			}

        )
    }

    if ($TestConnection) {
        Open-SnowFlakeConnection $ConnectionParams
    }

    if ($Configuration) {
        @()
    }

    Log verbose "Done"
}


function Idm-OnUnload {
    Close-SnowFlakeConnection
}


#
# CRUD functions
#

$ColumnsInfoCache = @{}


function Idm-Dispatcher {
    param (
        # Optional Class/Operation
        [string] $Class,
        [string] $Operation,
        # Mode
        [switch] $GetMeta,
        # Parameters
        [string] $SystemParams,
        [string] $FunctionParams
    )

    Log verbose "-Class='$Class' -Operation='$Operation' -GetMeta=$GetMeta -SystemParams='$SystemParams' -FunctionParams='$FunctionParams'"
    $connection_params = ConvertFrom-Json2 $SystemParams

    if ($Class -eq '') {

        if ($GetMeta) {
            #
            # Get all tables and views in database
            #

            Open-SnowFlakeConnection $SystemParams

            #
            # Output list of supported operations per table/view (named Class)
            #
            for ($i = 0; $i -lt 21; $i++)
            {
                if($connection_params."table_$($i)_name".length -gt 0)
                {
                    @(
                        [ordered]@{
                            Class = $connection_params."table_$($i)_name"
                            Operation = 'Read'
                            'Source type' = 'Query'
                            'Primary key' = ''
                            'Supported operations' = 'R'
                            'Query' = $connection_params."table_$($i)_query"
                        }
                    )
                    }
            }

        }
        else {
            # Purposely no-operation.
        }

    }
    else {

        if ($GetMeta) {
            #
            # Get meta data
            #

            Open-SnowFlakeConnection $SystemParams

            @()

        }
        else {
            #
            # Execute function
            #

            Open-SnowFlakeConnection $SystemParams

            for ($i = 0; $i -lt 21; $i++)
            {
                if($connection_params."table_$($i)_name" -eq $class)
                {
                    $class_query = $connection_params."table_$($i)_query"
                    break
                }
            }

		    $column_query = $class_query + " FETCH FIRST 5 ROWS ONLY"
        
            $columns = Fill-SqlInfoCache -Query $column_query -Timeout $SystemParams.query_timeout
        
            $Global:ColumnsInfoCache[$Class] = @{
                primary_keys = @($columns | Where-Object { $_.is_primary_key } | ForEach-Object { $_.name })
                identity_col = @($columns | Where-Object { $_.is_identity    } | ForEach-Object { $_.name })[0]
            }

            $primary_keys = $Global:ColumnsInfoCache[$Class].primary_keys
            $identity_col = $Global:ColumnsInfoCache[$Class].identity_col

            $function_params = ConvertFrom-Json2 $FunctionParams

            # Replace $null by [System.DBNull]::Value
            $keys_with_null_value = @()
            foreach ($key in $function_params.Keys) { if ($function_params[$key] -eq $null) { $keys_with_null_value += $key } }
            foreach ($key in $keys_with_null_value) { $function_params[$key] = [System.DBNull]::Value }
            
            $sql_command = New-SnowFlakeCommand $class_query
            $sql_command.CommandTimeout = $SystemParams.query_timeout
            Invoke-SnowFlakeCommand $sql_command
            Dispose-SnowFlakeCommand $sql_command

        }

    }

    Log verbose "Done"
}


#
# Helper functions
#

function Fill-SqlInfoCache {
    param (
        [switch] $Force,
        [string] $Query,
        [string] $Class,
        [string] $Timeout
    )

    # Refresh cache
	Log verbose "Executing Query: $($Query)"
    $sql_command = New-SnowFlakeCommand $Query
    $sql_command.CommandTimeout = $Timeout
    $result = (Invoke-SnowFlakeCommand $sql_command) | Get-Member -MemberType Properties | Select-Object Name
    
    Dispose-SnowFlakeCommand $sql_command

    $objects = New-Object System.Collections.ArrayList
    $object = @{}
    # Process in one pass
    foreach ($row in $result) {
            $object = @{
                full_name = $Class
                type      = 'Query'
                columns   = New-Object System.Collections.ArrayList
            }

        $object.columns.Add(@{
            name           = $row.Name
            is_primary_key = $false
            is_identity    = $false
            is_computed    = $false
            is_nullable    = $true
        }) | Out-Null
    }

    if ($object.full_name -ne $null) {
        $objects.Add($object) | Out-Null
    }
    @($objects)
}

function New-SnowFlakeCommand {
    param (
        [string] $CommandText
    )

    $sql_command = New-Object System.Data.Odbc.OdbcCommand($CommandText, $Global:SnowFlakeConnection)
    return $sql_command
}


function Dispose-SnowFlakeCommand {
    param (
        [System.Data.Odbc.OdbcCommand] $SqlCommand
    )

    $SqlCommand.Dispose()
}

function Invoke-SnowFlakeCommand {
    param (
        [System.Data.Odbc.OdbcCommand] $SqlCommand
    )

    function Invoke-SnowFlakeCommand-ExecuteReader {
        param (
            [System.Data.Odbc.OdbcCommand] $SqlCommand
        )
        $data_reader = $SqlCommand.ExecuteReader()
        $column_names = @($data_reader.GetSchemaTable().ColumnName)

        if ($column_names) {

            $hash_table = [ordered]@{}

            foreach ($column_name in $column_names) {
                $hash_table[$column_name] = ""
            }

#           $obj = [PSCustomObject]$hash_table
            $obj = New-Object -TypeName PSObject -Property $hash_table

            # Read data
	    if($data_reader.HasRows) {
	            while ($data_reader.Read()) {
	                foreach ($column_name in $column_names) {
	                    $obj.$column_name = if ($data_reader[$column_name] -is [System.DBNull]) { $null } else { $data_reader[$column_name] }
	                }
	
	                # Output data
	                $obj
	            }
	    } else { [PSCustomObject]$hash_table }

        }

        $data_reader.Close()

    }

    try {
        Invoke-SnowFlakeCommand-ExecuteReader $SqlCommand
    }
    catch {
        Log error "Query Failure: $_"
        throw $_
    }

    Log debug "Done"

}

function Open-SnowFlakeConnection {
    param (
        [string] $ConnectionParams
    )

    $connection_params = ConvertFrom-Json2 $ConnectionParams
    $connection_string = ("Driver={{SnowflakeDSIIDriver}};Server={0};Database={1};UID={2};PWD={3}" -f $connection_params.host_name, $connection_params.database, $connection_params.user, $connection_params.password)
    
    Log verbose $connection_string
    
    if ($Global:SnowFlakeConnection -and $connection_string -ne $Global:SnowFlakeConnectionString) {
        Log verbose "SnowFlakeConnection connection parameters changed"
        Close-SnowFlakeConnection
    }

    if ($Global:SnowFlakeConnection -and $Global:SnowFlakeConnection.State -ne 'Open') {
        Log warn "SnowFlakeConnection State is '$($Global:SnowFlakeConnection.State)'"
        Close-SnowFlakeConnection
    }

    Log verbose "Opening SnowFlakeConnection '$connection_string'"

    try {
        $connection = (new-object System.Data.Odbc.OdbcConnection);
        $connection.connectionstring = $connection_string
		$connection.ConnectionTimeout = $connection_params.connection_timeout
        $connection.open();

        $Global:SnowFlakeConnection       = $connection
        $Global:SnowFlakeConnectionString = $connection_string

        $Global:ColumnsInfoCache = @{}
    }
    catch {
        Log error "Connection Failure: $($_)"
        throw $_
    }

    Log verbose "Done"
    
}


function Close-SnowFlakeConnection {
    if ($Global:SnowFlakeConnection) {
        Log verbose "Closing SnowFlakeConnection"

        try {
            $Global:SnowFlakeConnection.Close()
            $Global:SnowFlakeConnection = $null
        }
        catch {
            # Purposely ignoring errors
        }

        Log verbose "Done"
    }
}
