[string] $COMP_NAME             = $env:COMPUTERNAME

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



[clsWmiXmlDocument] $wmi = [clsWmiXmlDocument]::new($COMP_NAME)
$wmi.GetXmlString($wmi.GetCpuXml())
$wmi.GetXmlString($wmi.GetBiosXml())
$wmi.GetXmlString($wmi.GetVideoXml())
$wmi.GetXmlString($wmi.GetHddXml())
$wmi.GetXmlString($wmi.GetRamXml())
$wmi.GetXmlString($wmi.GetIpXml())
$wmi.GetXmlString($wmi.GetUsersXml())

$wmi.GetDataTable($wmi.GetCpuXml())
$wmi.GetDataTable($wmi.GetBiosXml())
$wmi.GetDataTable($wmi.GetVideoXml())
$wmi.GetDataTable($wmi.GetHddXml())
$wmi.GetDataTable($wmi.GetRamXml())
$wmi.GetDataTable($wmi.GetIpXml())
$wmi.GetDataTable($wmi.GetUsersXml())


