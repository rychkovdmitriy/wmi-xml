﻿[string] $COMP_NAME             = $env:COMPUTERNAME
[string] $DB_PROVIDER_MSSQL      = "System.Data.SqlClient"
[string] $DB_SERVER_NAME_MSSQL   = "you-mssql-server-name"
[string] $DB_NAME_MSSQL          = "Computers"


class clsWmiXmlDataTable
{
    [string] $strComputerName

    clsWmiXmlDataTable()
    {
        $this.strComputerName = "."
    }

    clsWmiXmlDataTable([string] $sComputerName)
    {
        $this.strComputerName = $sComputerName
    }

    [System.Data.DataTable] CreateTable([string] $sTabName,[System.Object[]] $arrCol)
    {
        [System.Data.DataTable] $tab = [System.Data.DataTable]::New($sTabName)
 
        foreach ($col in $arrCol ) 
        {
           $tab.Columns.Add($col,([string]))
        }
        return $tab
    }


	
    [System.Data.DataTable] GetWmiTab([string] $wmiClassName,[System.Object[]] $wmiAttr)
    {

        $tabWmi = $this.CreateTable($wmiClassName,$wmiAttr)
        $wmiObj = Get-CimInstance -Class $wmiClassName  -ComputerName $this.strComputerName | Select-Object $wmiAttr
        if($wmiObj -eq $null)
        {
            return $null
        }
       
        return  $this.InsertWmiInTable($tabWmi,$wmiObj)
    }

    [System.Data.DataTable] GetWmiTab([string] $wmiClassName,[System.Object[]] $wmiAttr,[string]  $wmiFilter )
    {

        $tabWmi = $this.CreateTable($wmiClassName,$wmiAttr)
        $wmiObj = Get-CimInstance -Class $wmiClassName  -ComputerName $this.strComputerName -Filter $wmiFilter| Select-Object $wmiAttr
        if($wmiObj -eq $null)
        {
            return $null
        }
       
        return  $this.InsertWmiInTable($tabWmi,$wmiObj)
    }

    [System.Data.DataTable] InsertWmiInTable([System.Data.DataTable] $tabWmi, $wmiObj)
    {
        $wmiObj |%{
            $row = $tabWmi.NewRow()
            foreach($col in $tabWmi.Columns)
            {
                $colName = $col.ToString()
                Try
                {
                    $value =  $_.($colName)
                }
                catch
                {
                    $value = ""
                }
                $row[$colName] = $this.GetXmlValue($value)

            }
            $tabWmi.Rows.Add($row)
        }
        return $tabWmi
    }

    [string] GetXmlValue($value)
    {
        [string] $xmlValue = ""
        if( $value  -eq $null) 
        {
            return $xmlValue
        }
        if( $value  -is [array]) 
        {
            $xmlValue = ($value -join ",") 
        } 
        else 
        {
            $xmlValue  = $value 
        }
        return ($xmlValue.ToString().Trim() -replace '[\x00-\x1F\x7F<>&"]',"*")
    }

    [System.Data.DataTable] GetIpTab()
    {
        $attribIp = @("PSComputerName","DNSHostName","Description","Caption","DHCPEnabled","IPAddress","IPSubnet","MACAddress","DNSServerSearchOrder","DNSDomainSuffixSearchOrder","DefaultIPGateway")
        $filter = "IPEnabled=true"
        return $this.GetWmiTab("Win32_NetworkAdapterConfiguration",$attribIp,$filter)
    }

    [System.Data.DataTable] GetCpuTab()
    {
        $attribCpu = @("PSComputerName","Manufacturer","Name","DeviceID","NumberOfCores","NumberOfLogicalProcessors","CurrentClockSpeed","L2CacheSize","L3CacheSize")
        return $this.GetWmiTab("Win32_Processor",$attribCpu)
    }

    [System.Data.DataTable] GetBiosTab()
    {
        $attribBios = @("PSComputerName","Manufacturer","Name","SerialNumber","Version","Description","SMBIOSBIOSVersion","SMBIOSMajorVersion")
        return $this.GetWmiTab("Win32_BIOS",$attribBios)
    }

    [System.Data.DataTable] GetVideoTab()
    {
        $attribVideo = @("PSComputerName","AdapterCompatibility","AdapterDACType","AdapterRAM","Description","DriverDate","DriverVersion","Name","VideoModeDescription","VideoProcessor")
        return $this.GetWmiTab("Win32_VideoController",$attribVideo)
    }

    [System.Data.DataTable] GetHddTab()
    {
        $attribHdd = @("PSComputerName","Description","FirmwareRevision","Model","Manufacturer","Partitions","SerialNumber","Size","Status","SCSIBus","SCSILogicalUnit","SCSIPort","SCSITargetId","InterfaceType")
        return $this.GetWmiTab("Win32_DiskDrive",$attribHdd)
    }

    [System.Data.DataTable] GetRamTab()
    {
        $attribRam = @("PSComputerName","BankLabel","Capacity","DeviceLocator","Model","Manufacturer","FormFactor","PartNumber","SerialNumber","Speed")
        return $this.GetWmiTab("Win32_PhysicalMemory",$attribRam)
    }

    [System.Data.DataTable] GetUsersTab()
    {
        $tabWmi = $this.CreateTable("Win32_LoggedOnUser",@("PSComputerName","StartTime","DomainName","UserName"))
        $tabWmi.PrimaryKey = @($tabWmi.Columns["UserName"],$tabWmi.Columns["StartTime"])
        $LoggedOnUser =  Get-CimInstance -ComputerName $this.strComputerName -ClassName Win32_LoggedOnUser
         Get-CimInstance  -ComputerName $this.strComputerName -ClassName Win32_LogonSession | ? { $_.LogonType -eq 2 -or  $_.LogonType -eq 10} | %{
                  $id = $_.LogonId
                  $usr = $LoggedOnUser | ? { $_.Dependent.LogonId -eq $id}
                            if($usr -ne $null) 
                            {                        
                                try
                                {
                
                                    $tabWmi.Rows.Add($this.strComputerName,$_.StartTime.ToString(),$usr.Antecedent.Domain, $usr.Antecedent.Name )
                                }
                                catch 
                                {
                                    Write-Host "Row not need add"
                                }
                            }

                  } 
        return $tabWmi
    }

    [string] GetXmlString( [System.Data.DataTable] $tab)
    {
         [System.IO.StringWriter] $writer = [System.IO.StringWriter]::new()
         [System.Data.DataSet] $ds = [System.Data.DataSet]::New("Objects")
         $ds.Tables.Add($tab)
         $ds.WriteXml($writer)
         return $writer.ToString()
    }


}



class clsDb
{
    [System.Data.Common.DbProviderFactory]          $factory
    [System.Data.Common.DbConnection]               $conn
    [System.Data.Common.DbConnectionStringBuilder]  $csb
    [string] $strProviderName
    [string] $strServer
    [string] $strDatabase
    [string] $strUserID
    [string] $strPass


    clsDb()
    {
        $this.factory = $null
        $this.conn    = $null
        $this.csb     = $null
        $this.strProviderName = ""
        $this.strServer       = ""
        $this.strDatabase     = ""
    }

    clsDb([string] $ProviderName,[string] $Server,[string] $Database)
    {
        $this.factory = $null
        $this.conn    = $null
        $this.csb     = $null
        $this.strProviderName = $ProviderName
        $this.strServer       = $Server
        $this.strDatabase     = $Database
        $this.conn =  $this.CreateConnection($this.strServer,$this.strDatabase)
    }

    clsDb([string]  $UserID, [string] $Pass,[string] $ProviderName,[string] $Server,[string] $Database)
    {
        $this.factory = $null
        $this.conn    = $null
        $this.csb     = $null
        $this.strProviderName = $ProviderName
        $this.strServer       = $Server
        $this.strDatabase     = $Database
        $this.strUserID = $UserID
        $this.strPass = $Pass
        $this.conn =  $this.CreateConnection( $this.strUserID,$this.strPass,$this.strServer,$this.strDatabase)
    }

    [System.Data.Common.DbProviderFactory] GetFactory()
    {

        if($this.factory -eq $null)
        {
            try 
            {
                $this.factory = [System.Data.Common.DbProviderFactories]::GetFactory($this.strProviderName)
                Write-Host "clsDb:GFactory = " $this.factory.GetType()
                return $this.factory
            } 
            catch [System.ArgumentException] 
            {
                if($this.strProviderName -eq "MySql.Data.MySqlClient")
                {
                    $this.factory = new-object "MySql.Data.MySqlClient.MySqlClientFactory"
                    Write-Host "clsDb:GFactory = " $this.factory.GetType()
                    return $this.factory
                }
                else
                {
                    Write-Host "clsDb:GFactory: Null"
                    return $null
                }
            }
        }
        else
        {
            return $this.factory
        }
        
    }

    [System.Data.Common.DbConnection] GetConnection()
    {
        return $this.conn
    }

    [System.Data.Common.DbConnection] CreateConnection([string] $UserID, [string] $Password,[string] $Server,[string]  $Database)
    {
            try 
            {
                $this.csb = $this.GetFactory().CreateConnectionStringBuilder()
                $this.csb.Add("User ID", $UserID)
                $this.csb.Add("Password", $Password)
                $this.csb.Add("Server", $Server)
                $this.csb.Add("Database", $Database)
                if($this.strProviderName -eq "MySql.Data.MySqlClient")
                {
                     $this.csb.Add("charset", "utf8")
                }
                Write-Host "clsDb:CreateConnection: " $this.csb.ConnectionString

                $this.conn =  $this.GetFactory().CreateConnection()
                $this.conn.ConnectionString = $this.csb.ConnectionString
                return $this.conn
            }
            catch [System.ArgumentException] 
            {
                return $null
            }
    }

    [System.Data.Common.DbConnection] CreateConnection([string] $Server,[string]  $Database)
    {
            try 
            {
                $this.csb = $this.GetFactory().CreateConnectionStringBuilder()
                $this.csb.Add("Server", $Server)
                $this.csb.Add("Database", $Database)
                $this.csb.Add("Integrated Security", $true)
                
                if($this.strProviderName -eq "MySql.Data.MySqlClient")
                {
                     $this.csb.Add("charset", "utf8")
                }
                Write-Host "clsDb:CreateConnection: " $this.csb.ConnectionString

                $this.conn =  $this.GetFactory().CreateConnection()
                $this.conn.ConnectionString = $this.csb.ConnectionString
                return $this.conn
            }
            catch [System.ArgumentException] 
            {
                return $null
            }
    }

    [System.Data.Common.DbCommand] CreateCommandSp([string] $CommandText,[string]  $ParameterName)
    {
            try 
            {
                [System.Data.Common.DbCommand] $cmd = $this.GetFactory().CreateCommand()
                $cmd.Connection = $this.GetConnection()
                $cmd.CommandType = [System.Data.CommandType]::StoredProcedure;
                $cmd.CommandText = $CommandText
                $cmd.Parameters.Add($this.CreateParameter($ParameterName,[System.Data.DbType]::String))
                return  $cmd
            }
            catch [System.ArgumentException] 
            {
                return $null
            }
    }

    [System.Data.Common.DbCommand] CreateCommandSp([string] $CommandText,[string]  $ParameterName, [string] $val)
    {
            try 
            {
                [System.Data.Common.DbCommand] $cmd = $this.GetFactory().CreateCommand()
                $cmd.Connection = $this.GetConnection()
                $cmd.CommandType = [System.Data.CommandType]::StoredProcedure;
                $cmd.CommandText = $CommandText
                $cmd.Parameters.Add($this.CreateParameter($ParameterName,[System.Data.DbType]::String,$val))
                return  $cmd
            }
            catch [System.ArgumentException] 
            {
                return $null
            }
    }

    [System.Data.Common.DbParameter] CreateParameter([string] $ParameterName, [System.Data.DbType] $DbType)
    {
            try 
            {
                [System.Data.Common.DbParameter] $param =  $this.GetFactory().CreateParameter()
                $param.ParameterName = $ParameterName
                $param.DbType = $DbType
                return $param
            }
            catch [System.ArgumentException] 
            {
                return $null
            }
               
    }

    [System.Data.Common.DbParameter] CreateParameter([string] $ParameterName, [string] $SourceColumn, [System.Data.DbType] $DbType)
    {
            try 
            {
                [System.Data.Common.DbParameter] $param =  $this.GetFactory().CreateParameter()
                $param.ParameterName = $ParameterName
                $param.SourceColumn = $SourceColumn
                $param.DbType = $DbType
                return $param
            }
            catch [System.ArgumentException] 
            {
                return $null
            }
               
    }

    [System.Data.Common.DbParameter] CreateParameter([string] $ParameterName, [System.Data.DbType] $DbType, [string] $val)
    {
            try 
            {
                [System.Data.Common.DbParameter] $param =  $this.GetFactory().CreateParameter()
                $param.ParameterName = $ParameterName
                $param.DbType = $DbType
                $param.Value = $val
                return $param
            }
            catch [System.ArgumentException] 
            {
                return $null
            }
               
    }

    [System.Data.Common.DbDataAdapter] InsertDataTable([string] $sp,[System.Data.DataTable] $tab)
    {
         try 
            {
                [System.Data.Common.DbDataAdapter] $adapter =  $this.GetFactory().CreateDataAdapter()


                $adapter.InsertCommand = $this.conn.CreateCommand()
                $adapter.InsertCommand.CommandType = [System.Data.CommandType]::StoredProcedure;
                $adapter.InsertCommand.CommandText = $sp
                foreach($col in $tab.Columns )
                {
                    $adapter.InsertCommand.Parameters.Add($this.CreateParameter("@p" + $col.ToString(),$col.ToString(),[System.Data.DbType]::String))
                }
                $adapter.Update($tab)
                return $adapter
            }
            catch [System.ArgumentException] 
            {
                return $null
            }
    }

    [System.Data.Common.DbDataAdapter] InsertDataTable([System.Data.DataTable] $tab)
    {
        return  $this.InsertDataTable("spInsert$($tab.TableName)Log",$tab)
    }


    [bool] ExecuteCmd($cmd)
    {
        if ($cmd -ne $null) 
        {
            Write-Host $cmd.CommandText
            ForEach ($param in $cmd.Parameters) 
            {
                 Write-Host "ParameterName = " $param.ParameterName
                 Write-Host "ParameterValue = " $param.Value
            }
            
            if($cmd.Connection.State -eq [System.Data.ConnectionState]::Closed)
            {
                $cmd.Connection.Open()
                Write-Host "EXEC" $cmd.ExecuteNonQuery()
                return $true
            }
            else
            {
                $cmd.ExecuteNonQuery()
                return $true
            }
        }
        else
        {
            Write-Host "cmd is null!"
            return $false
        }
  
	}
    
    [string] CreateSql([System.Data.DataTable] $tab)
    {
        [System.Text.StringBuilder] $sb = [System.Text.StringBuilder]::new()
        $tabName = $tab.TableName.ToString()
        $tabColumns =  $tab.Columns
        $sb.AppendLine($this.CreateTableSql($tabName,$tabColumns))
        $sb.AppendLine($this.CreateTableInsertSpSql($tabName,$tabColumns))
        return  $sb.ToString()
    }
	
    [string] CreateTableSql([string] $tabName,$tabColumns)
    {
        [System.Text.StringBuilder] $sb = [System.Text.StringBuilder]::new()
        $sb.AppendLine("CREATE TABLE [dbo].[$tabName](")
        $sb.AppendLine("   [id] INT IDENTITY(1,1) PRIMARY KEY,")
        for ($j=0; $j -ilt  $tabColumns.Count; $j++)
        {
            $sb.Append("   [" + $tabColumns[$j].ToString() + "]  ")
            $sb.AppendLine(" NVARCHAR(250) NULL,")
        }
        $sb.AppendLine("   [DateAdd] datetime NULL DEFAULT (getdate()),")
        $sb.AppendLine("   [DateUpdate] datetime NULL);")
        $sb.AppendLine(" GO")
        return  $sb.ToString()
    }


    [string] CreateTableInsertSpSql([string] $tabName,$tabColumns)
    {
        [System.Text.StringBuilder] $sb = [System.Text.StringBuilder]::new()
        $sb.AppendLine("CREATE PROCEDURE [dbo].[spInsert$($tabName)Log]")
        for ($j=0; $j -ilt  $tabColumns.Count - 1; $j++)
        {
            $sb.Append("   @p" +  $tabColumns[$j].ToString())
            $sb.AppendLine(" NVARCHAR(250),")
        }
        $sb.Append("   @p" + $tabColumns[$tabColumns.Count - 1].ToString())
        $sb.AppendLine(" NVARCHAR(250)")
        $sb.AppendLine("AS")
        $sb.AppendLine("BEGIN")

        $sb.AppendLine("  SET NOCOUNT ON;")
	    $sb.AppendLine("  DECLARE @idOld AS INT;")
        $sb.Append("  SELECT TOP 1 @idOld = [id] from [$tabName]  where ")

        for($j=0;$j -ilt $tabColumns.Count - 1; $j++)
        {
            $sb.Append("   [" + $tabColumns[$j].ToString() + "] = ")
            $sb.AppendLine(" @p" + $tabColumns[$j].ToString() + " and ")
        }
        $sb.Append("   [" + $tabColumns[$tabColumns.Count - 1].ToString() + "] = ")
        $sb.AppendLine(" @p" + $tabColumns[$tabColumns.Count - 1].ToString() + ";")

        $sb.AppendLine("if  @idOld IS NOT NULL")
        $sb.AppendLine("   UPDATE [$tabName] SET  DateUpdate = GetDate()  WHERE ID = @idOld;")
        $sb.AppendLine("else")
        $sb.Append("   INSERT INTO [$tabName] (")
        for($j=0;$j -ilt $tabColumns.Count - 1; $j++)
        {
            $sb.AppendLine("    [" + $tabColumns[$j].ToString() + "], ")
        }
        $sb.AppendLine(" [" + $tabColumns[$tabColumns.Count - 1].ToString() + "]) VALUES ")
        $sb.Append("(")
        for($j=0;$j -ilt $tabColumns.Count - 1; $j++)
        {
            $sb.AppendLine("     @p" + $tabColumns[$j].ToString() + ", ")
        }
        $sb.AppendLine(" @p" + $tabColumns[$tabColumns.Count - 1].ToString() + ")")
        $sb.AppendLine("END;")
        $sb.AppendLine("GO") 
        $sb.AppendLine("GRANT EXECUTE ON OBJECT::[$($this.strDatabase)].[dbo].[spInsert$($tabName)Log] ")
        $sb.AppendLine("TO [$($env:USERDOMAIN)\Domain Computers]; ")
        $sb.AppendLine("GO") 
        return  $sb.ToString()     
    }

}



[clsWmiXmlDataTable] $wmi = [clsWmiXmlDataTable]::new($COMP_NAME)


<#$wmi.GetXmlString($wmi.GetBiosTab())
$wmi.GetXmlString($wmi.GetCpuTab())
$wmi.GetXmlString($wmi.GetVideoTab())
$wmi.GetXmlString($wmi.GetHddTab())
$wmi.GetXmlString($wmi.GetRamTab())
$wmi.GetXmlString($wmi.GetIpTab())
$wmi.GetXmlString($wmi.GetUsersTab())#>

<#$wmi.GetBiosTab()
$wmi.GetCpuTab()
$wmi.GetVideoTab()
$wmi.GetHddTab()
$wmi.GetRamTab()
$wmi.GetIpTab()
$wmi.GetUsersTab()#>

<#[clsDb] $mssql = [clsDb]::new($DB_PROVIDER_MSSQL,$DB_SERVER_NAME_MSSQL,$DB_NAME_MSSQL )
[System.Text.StringBuilder] $sb = [System.Text.StringBuilder]::new()
$sb.AppendLine($mssql.CreateSql($wmi.GetBiosTab())) | Out-Null
$sb.AppendLine($mssql.CreateSql($wmi.GetCpuTab()))  | Out-Null
$sb.AppendLine($mssql.CreateSql($wmi.GetVideoTab()))  | Out-Null
$sb.AppendLine($mssql.CreateSql($wmi.GetHddTab()))  | Out-Null
$sb.AppendLine($mssql.CreateSql($wmi.GetRamTab()))  | Out-Null
$sb.AppendLine($mssql.CreateSql($wmi.GetIpTab()))  | Out-Null
$sb.AppendLine($mssql.CreateSql($wmi.GetUsersTab()))  | Out-Null
$sb.ToString() | Out-File -FilePath "C:\distr\sqlDataTable.sql"#>


[clsDb] $mssql = [clsDb]::new($DB_PROVIDER_MSSQL,$DB_SERVER_NAME_MSSQL,$global:DB_NAME_MSSQL )
$mssql.InsertDataTable($wmi.GetBiosTab())
$mssql.InsertDataTable($wmi.GetCpuTab())
$mssql.InsertDataTable($wmi.GetHddTab())
$mssql.InsertDataTable($wmi.GetRamTab())
$mssql.InsertDataTable($wmi.GetIpTab())
$mssql.InsertDataTable($wmi.GetUsersTab())
$mssql.InsertDataTable($wmi.GetVideoTab())

