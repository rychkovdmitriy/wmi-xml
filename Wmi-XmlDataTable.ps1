[string] $COMP_NAME             = $env:COMPUTERNAME

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



[clsWmiXmlDataTable] $wmi = [clsWmiXmlDataTable]::new($COMP_NAME)



$wmi.GetXmlString($wmi.GetCpuTab())
$wmi.GetXmlString($wmi.GetBiosTab())
$wmi.GetXmlString($wmi.GetVideoTab())
$wmi.GetXmlString($wmi.GetHddTab())
$wmi.GetXmlString($wmi.GetRamTab())
$wmi.GetXmlString($wmi.GetIpTab())
$wmi.GetXmlString($wmi.GetUsersTab())


$wmi.GetCpuTab()
$wmi.GetBiosTab()
$wmi.GetVideoTab()
$wmi.GetHddTab()
$wmi.GetRamTab()
$wmi.GetIpTab()
$wmi.GetUsersTab()



