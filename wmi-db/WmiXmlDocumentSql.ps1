[string] $COMP_NAME             = $env:COMPUTERNAME
[string] $DB_PROVIDER_MSSQL      = "System.Data.SqlClient"
[string] $DB_SERVER_NAME_MSSQL   = "you-mssql-server-name"
[string] $DB_NAME_MSSQL          = "Computers"

class clsWmiXmlDocument
{
    [string] $strComputerName

    clsWmiXmlDocument()
    {
        $this.strComputerName = "."
    }

    clsWmiXmlDocument([string] $sComputerName)
    {
        $this.strComputerName = $sComputerName
    }


    [xml] GetWmiXml([string] $wmiClassName,[System.Object[]] $wmiAttr)
    {

        $xml=[xml] '<Objects/>'
        $wmiObj = Get-CimInstance -Class $wmiClassName  -ComputerName $this.strComputerName | Select-Object $wmiAttr
        if($wmiObj -eq $null)
        {
            return $xml
        }
       
        $wmiObj |%{
            $n=$xml.DocumentElement.AppendChild($xml.CreateElement($wmiClassName))
            foreach($atr in $wmiAttr)
            {
                $c=$n.AppendChild($xml.CreateElement($atr))
                $c.InnerText = $this.GetXmlValue($_.($atr))
            }
        }
        return $xml

    }

    [xml] GetWmiXml([string] $wmiClassName,[System.Object[]] $wmiAttr,[string]  $wmiFilter )
    {

        $xml=[xml]'<Objects/>'
        $wmiObj = Get-CimInstance -Class $wmiClassName  -ComputerName $this.strComputerName -Filter $wmiFilter| Select-Object $wmiAttr
        if($wmiObj -eq $null)
        {
            return $xml
        }
       
        
        $wmiObj |%{
            $n=$xml.DocumentElement.AppendChild($xml.CreateElement($wmiClassName))
            foreach($atr in $wmiAttr)
            {
                $c=$n.AppendChild($xml.CreateElement($atr))
                $c.InnerText = $this.GetXmlValue($_.($atr))
               
            }
        }
        return $xml
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

    [xml] GetIpXml()
    {
        $attribIp = @("PSComputerName","DNSHostName","Description","Caption","DHCPEnabled","IPAddress","IPSubnet","MACAddress","DNSServerSearchOrder","DNSDomainSuffixSearchOrder","DefaultIPGateway")
        $filter = "IPEnabled=true"
        return $this.GetWmiXml("Win32_NetworkAdapterConfiguration",$attribIp,$filter)
    }

    [xml] GetCpuXml()
    {
        $attribCpu = @("PSComputerName","Manufacturer","Name","DeviceID","NumberOfCores","NumberOfLogicalProcessors","CurrentClockSpeed","L2CacheSize","L3CacheSize")
        return $this.GetWmiXml("Win32_Processor",$attribCpu)
    }

    [xml] GetBiosXml()
    {
        $attribBios = @("PSComputerName","Manufacturer","Name","SerialNumber","Version","Description","SMBIOSBIOSVersion","SMBIOSMajorVersion")
        return $this.GetWmiXml("Win32_BIOS",$attribBios)
    }

    [xml] GetVideoXml()
    {
        $attribVideo = @("PSComputerName","AdapterCompatibility","AdapterDACType","AdapterRAM","Description","DriverDate","DriverVersion","Name","VideoModeDescription","VideoProcessor")
        return $this.GetWmiXml("Win32_VideoController",$attribVideo)
    }

    [xml] GetHddXml()
    {
        $attribHdd = @("PSComputerName","Description","FirmwareRevision","Model","Manufacturer","Partitions","SerialNumber","Size","Status","SCSIBus","SCSILogicalUnit","SCSIPort","SCSITargetId","InterfaceType")
        return $this.GetWmiXml("Win32_DiskDrive",$attribHdd)
    }

    [xml] GetRamXml()
    {
        $attribRam = @("PSComputerName","BankLabel","Capacity","DeviceLocator","Model","Manufacturer","FormFactor","PartNumber","SerialNumber","Speed")
        return $this.GetWmiXml("Win32_PhysicalMemory",$attribRam)
    }

    [xml] GetUsersXml()
    {
        $xml=[xml]'<Objects/>'
        $LoggedOnUser =  Get-CimInstance -ComputerName $this.strComputerName -ClassName Win32_LoggedOnUser
         Get-CimInstance  -ComputerName $this.strComputerName -ClassName Win32_LogonSession | ? { $_.LogonType -eq 2 -or  $_.LogonType -eq 10} | %{
                  $n=$xml.DocumentElement.AppendChild($xml.CreateElement("Win32_LoggedOnUser"))
                  $id = $_.LogonId
                  $usr = $LoggedOnUser | ? { $_.Dependent.LogonId -eq $id}
                            if($usr -ne $null) 
                            {                        
                                try
                                {
                                    $PSComputerName=$n.AppendChild($xml.CreateElement("PSComputerName"))
                                    $PSComputerName.InnerText= $this.strComputerName

                                    $StartTime=$n.AppendChild($xml.CreateElement("StartTime"))
                                    $StartTime.InnerText = $_.StartTime.ToString()

                                    $DomainName=$n.AppendChild($xml.CreateElement("DomainName"))
                                    $DomainName.InnerText = $usr.Antecedent.Domain

                                    $UserName=$n.AppendChild($xml.CreateElement("UserName"))
                                    $UserName.InnerText = $usr.Antecedent.Name
                
                                }
                                catch 
                                {
                                    Write-Host "Row not need add"
                                }
                            }

                  } 
        return $xml
    }

    [string] GetXmlString( [xml] $xml)
    {
         return $xml.OuterXml
    }

    [System.Data.DataTable] GetDataTable([xml] $xml)
    {
        [System.Data.DataSet] $tab = [System.Data.DataSet]::new()
        $stream = [System.IO.MemoryStream]::New()
        $writer = [System.IO.StreamWriter]::New($stream)
        $writer.Write($xml.OuterXml)
        $writer.Flush()
        $stream.Position = 0
        $tab.ReadXml($stream)
        return $tab.Tables[0]
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

    [bool] InsertXmlTable([string] $sp,[xml] $tab)
    {
        $this.ExecuteCmd($this.CreateCommandSp( $sp,  "@xmlText",$tab.OuterXml))
        return $true
    }

    [bool] InsertXmlTable([xml] $tab)
    {
            $root = $tab.FirstChild

        if ($root.HasChildNodes)
        {
            $sp = "[spInsert$($root.ChildNodes[0].ToString())XmlLog]"
            $this.ExecuteCmd($this.CreateCommandSp( $sp,  "@xmlText",$tab.OuterXml))
            return $true
        }
         return $false
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

    [string] CreateSql([xml] $tab)
    {
        [System.Text.StringBuilder] $sb = [System.Text.StringBuilder]::new()
        $root = $tab.FirstChild

        if ($root.HasChildNodes)
        {
              for ($i=0; $i -ilt  $root.ChildNodes.Count; $i++)
              {
                $tabName = $root.ChildNodes[$i].ToString()
                $tabColumns =  $root.ChildNodes[$i].ChildNodes
                $sb.AppendLine($this.CreateTableSql($tabName,$tabColumns))
                $sb.AppendLine($this.CreateTableInsertSpSql($tabName,$tabColumns))
                $sb.AppendLine($this.CreateTableInsertSpXmlSql($tabName,$tabColumns))
              }
        }
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

    [string] CreateTableInsertSpXmlSql([string] $tabName,$tabColumns)
    {
        [System.Text.StringBuilder] $sb = [System.Text.StringBuilder]::new()
        $indexLastCol = $tabColumns.Count - 1
        $sb.AppendLine("CREATE PROCEDURE [dbo].[spInsert$($tabName)XmlLog]")
        $sb.AppendLine(" @xmlText nvarchar(max)")
        $sb.AppendLine("AS")
        $sb.AppendLine("BEGIN")

        $sb.AppendLine("  SET NOCOUNT ON;")
        $sb.AppendLine("  DECLARE @dateUpdate DateTime;")
        $sb.AppendLine("  SET @dateUpdate = GetDate();")
	    $sb.AppendLine("  DECLARE @xml xml")
 	    $sb.AppendLine("  SET @xml = TRY_CONVERT(xml,@xmlText);")

        $sb.Append("  DECLARE ")
  	 

        for ($j=0; $j -ilt  $indexLastCol; $j++)
        {
            $sb.Append("    @" +  $tabColumns[$j].ToString())
            $sb.AppendLine(" NVARCHAR(250),")
        }
        $sb.Append("    @" +  $tabColumns[$indexLastCol].ToString())
        $sb.AppendLine(" NVARCHAR(250);")

        $sb.AppendLine("DECLARE cur CURSOR FOR")
        $sb.AppendLine("   SELECT")
        for($j=0;$j -ilt $tabColumns.Count - 1; $j++)
        {
            $sb.Append("      tab.col.value")
            $sb.AppendLine("('string((" + $tabColumns[$j].ToString() + "/text())[1])',	'NVARCHAR(250)') AS  c" + $tabColumns[$j].ToString() + ",")
        }
        $sb.Append("      tab.col.value")
        $sb.AppendLine("('string((" + $tabColumns[$indexLastCol].ToString() + "/text())[1])',	'NVARCHAR(250)') AS  c" + $tabColumns[$indexLastCol].ToString())
        $sb.AppendLine("   FROM")
        $sb.AppendLine(" @xml.nodes('/Objects/" + $tabName + "')AS tab(col) ")
		$sb.AppendLine(" OPEN cur")
		
        $sb.AppendLine(" FETCH NEXT FROM cur")
        $sb.Append("INTO ")
        for ($j=0; $j -ilt  $indexLastCol; $j++)
        {
            $sb.Append(" @" +  $tabColumns[$j].ToString() + ",")
        }
        $sb.AppendLine(" @" +  $tabColumns[$indexLastCol].ToString())
        
        $sb.AppendLine(" WHILE @@FETCH_STATUS = 0 ")
        $sb.AppendLine(" BEGIN ")
        $sb.Append(" EXEC  [dbo].[spInsert$($tabName)Log]  ")
        
        for ($j=0; $j -ilt  $indexLastCol; $j++)
        {
            $sb.AppendLine("   @p" +  $tabColumns[$j].ToString() + " = @" +  $tabColumns[$j].ToString() + ",")
        }
        $sb.AppendLine("   @p" +  $tabColumns[$indexLastCol].ToString() + " = @" +  $tabColumns[$indexLastCol].ToString() + ";")
        	
		$sb.AppendLine("   FETCH NEXT FROM cur")
        $sb.Append("   INTO ")
        for ($j=0; $j -ilt $indexLastCol; $j++)
        {
            $sb.Append(" @" +  $tabColumns[$j].ToString() + ",")
        }
        $sb.AppendLine(" @" +  $tabColumns[$indexLastCol].ToString())
        $sb.AppendLine(" END ")
        $sb.AppendLine(" CLOSE cur;")
		$sb.AppendLine(" DEALLOCATE cur;")
        $sb.AppendLine("END;")
        $sb.AppendLine("GO") 
        $sb.AppendLine("GRANT EXECUTE ON OBJECT::[$($this.strDatabase)].[dbo].[spInsert$($tabName)XmlLog] ")
        $sb.AppendLine("TO [$($env:USERDOMAIN)\Domain Computers]; ")
        $sb.AppendLine("GO") 
        return  $sb.ToString()     
    }
}

[clsWmiXmlDocument] $wmi = [clsWmiXmlDocument]::new($COMP_NAME)

<#$wmi.GetXmlString($wmi.GetCpuXml())
$wmi.GetXmlString($wmi.GetBiosXml())
$wmi.GetXmlString($wmi.GetVideoXml())
$wmi.GetXmlString($wmi.GetHddXml())
$wmi.GetXmlString($wmi.GetRamXml())
$wmi.GetXmlString($wmi.GetIpXml())
$wmi.GetXmlString($wmi.GetUsersXml())#>

<#$wmi.GetDataTable($wmi.GetCpuXml())
$wmi.GetDataTable($wmi.GetBiosXml())
$wmi.GetDataTable($wmi.GetVideoXml())
$wmi.GetDataTable($wmi.GetHddXml())
$wmi.GetDataTable($wmi.GetRamXml())
$wmi.GetDataTable($wmi.GetIpXml())
$wmi.GetDataTable($wmi.GetUsersXml())#>


<#[clsDb] $mssql = [clsDb]::new($DB_PROVIDER_MSSQL,$DB_SERVER_NAME_MSSQL,$DB_NAME_MSSQL )
[System.Text.StringBuilder] $sb = [System.Text.StringBuilder]::new()
$sb.AppendLine($mssql.CreateSql($wmi.GetBiosXml()))   | Out-Null
$sb.AppendLine($mssql.CreateSql($wmi.GetCpuXml()))    | Out-Null
$sb.AppendLine($mssql.CreateSql($wmi.GetVideoXml()))  | Out-Null
$sb.AppendLine($mssql.CreateSql($wmi.GetHddXml()))    | Out-Null
$sb.AppendLine($mssql.CreateSql($wmi.GetRamXml()))    | Out-Null
$sb.AppendLine($mssql.CreateSql($wmi.GetIpXml()))     | Out-Null
$sb.AppendLine($mssql.CreateSql($wmi.GetUsersXml()))  | Out-Null
$sb.ToString() | Out-File -FilePath "C:\distr\sqlXmlDocument.sql"
#>

[clsDb] $mssql = [clsDb]::new($DB_PROVIDER_MSSQL,$DB_SERVER_NAME_MSSQL,$DB_NAME_MSSQL )
$mssql.InsertXmlTable($wmi.GetRamXml())
$mssql.InsertXmlTable($wmi.GetBiosXml())
$mssql.InsertXmlTable($wmi.GetCpuXml())
$mssql.InsertXmlTable($wmi.GetVideoXml())
$mssql.InsertXmlTable($wmi.GetHddXml())
$mssql.InsertXmlTable($wmi.GetIpXml())
$mssql.InsertXmlTable($wmi.GetUsersXml())
