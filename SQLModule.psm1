function Connect-SQLServer{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        [string]$InstanceName,
        [Parameter(Mandatory)]
        [string]$DatabaseName,
        [Parameter(Mandatory)]
        [bool]$IntergratedSecurity,
        [Parameter()]
        [string]$Username,
        [Parameter()]
        [string]$Password
    )

    try{
        $SQLConnection=New-Object System.Data.SqlClient.SqlConnection
        if($IntergratedSecurity){
            $SQLConnection.ConnectionString="Server=$InstanceName;Database=$DatabaseName;Integrated Security=true"
        }else{
            $SQLConnection.ConnectionString="Server=$InstanceName;Database=$DatabaseName;User=$Username;Password=$Password"
        }

        $SQLConnection.Open()

        return $SQLConnection
    }catch{
        Write-Error $_.exception.message
    }

}

function Close-SQLServerConnection{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        $Connection
    )

    try{
        $Connection.close()
    }catch{
        Write-Error $_.exception.message
    }
}


function Invoke-SQLSelect{
    [CmdletBinding()]
    Param(
       [Parameter(Mandatory)]
       $Connection,
       [Parameter(Mandatory)]
       [string]$SelectStatement
    )

    try{
        $SQLCommand=$Connection.CreateCommand()

        $SQLCommand.CommandText="$SelectStatement"

        $SQLDataAdapter=New-Object System.Data.SqlClient.SqlDataAdapter $SQLCommand
        $DataSet=New-Object System.Data.DataSet

        [void]$SQLDataAdapter.Fill($DataSet)

        $Data=$DataSet.Tables[0]

        return $Data

    }catch{
        Write-Error $_.exception.message
    }

}

function Invoke-SQLInsert{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        $Connection,
        [Parameter(Mandatory)]
        [string]$InsertStatement
    )

    try{
        $SQLCommand=$Connection.CreateCommand()

        $SQLCommand.CommandText=$InsertStatement

        $Result=$SQLCommand.ExecuteNonQuery()

        if($Result -eq 1){
            Write-Verbose "Insert succesful"
        }else{
            Write-Verbose "Insert failed"
        }
    }catch{
        Write-Error $_.exception.message
    }
}

function Invoke-SQLUpdate{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        $Connection,
        [Parameter(Mandatory)]
        [string]$UpdateStatement
    )

    try{
        $SQLCommand=$Connection.CreateCommand()

        $SQLCommand.CommandText=$UpdateStatement

        $Result=$SQLCommand.ExecuteNonQuery()

        if($Result -ge 1){
            Write-Verbose "Update succesful"
        }else{
            Write-Verbose "Update failed"
        }
    }catch{
        Write-Error $_.exception.message
    }
}

function Invoke-SQLDelete{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory)]
        $Connection,
        [Parameter(Mandatory)]
        [string]$DeleteStatement
    )

    try{
        $SQLCommand=$Connection.CreateCommand()

        $SQLCommand.CommandText=$DeleteStatement

        $Result=$SQLCommand.ExecuteNonQuery()

        if($Result -ge 1){
            Write-Verbose "Delete succesful"
        }else{
            Write-Verbose "Delete failed"
        }
    }catch{
        Write-Error $_.exception.message
    }   
}
