[string] $COMP_NAME             = $env:COMPUTERNAME

class clsWmiPs
{
    [string] $strComputerName

    clsWmiPs()
    {
        $this.strComputerName =  $env:COMPUTERNAME
    }

    clsWmiPs([string] $sComputerName)
    {
        $this.strComputerName = $sComputerName
    }

    [xml] GetIpXml()
    {
        return Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=$true -ComputerName $this.strComputerName | Select-Object PSComputerName,DNSHostName,Description,Caption,DHCPEnabled,IPAddress,IPSubnet,MACAddress,DNSServerSearchOrder,DNSDomainSuffixSearchOrder,DefaultIPGateway | ConvertTo-XML -NoTypeInformation
    }

    [xml] GetCpuXml()
    {
        return Get-WmiObject -Class Win32_Processor -ComputerName $this.strComputerName | Select-Object  PSComputerName,Manufacturer,Name,DeviceID,NumberOfCores,NumberOfLogicalProcessors,CurrentClockSpeed,L2CacheSize,L3CacheSize | ConvertTo-XML -NoTypeInformation
    }

    [xml] GetBiosXml()
    {
        return Get-WmiObject -Class Win32_BIOS -ComputerName $this.strComputerName  | Select-Object PSComputerName,Name,SerialNumber,Version,Description,SMBIOSBIOSVersion,SMBIOSMajorVersion | ConvertTo-XML -NoTypeInformation
    }

    [xml] GetVideoXml()
    {
        return Get-WmiObject -Class Win32_VideoController -ComputerName $this.strComputerName  | Select-Object PSComputerName,AdapterCompatibility,AdapterDACType,AdapterRAM,Description,DriverDate,DriverVersion,Name,VideoModeDescription,VideoProcessor |  ConvertTo-XML -NoTypeInformation
    }

    [xml] GetHddXml()
    {
        return Get-WmiObject -Class Win32_DiskDrive -ComputerName $this.strComputerName  | Select-Object PSComputerName,Description,FirmwareRevision,Model,Manufacturer,Partitions,SerialNumber,Size,Status,SCSIBus,SCSILogicalUnit,SCSIPort,SCSITargetId,InterfaceType | ConvertTo-XML -NoTypeInformation
    }

    [xml] GetRamXml()
    {
        return Get-WmiObject -Class Win32_PhysicalMemory -ComputerName $this.strComputerName  | Select-Object PSComputerName,BankLabel,Capacity,DeviceLocator,Model,Manufacturer,FormFactor,PartNumber,SerialNumber,Speed | ConvertTo-XML -NoTypeInformation
    }

    [xml] GetUsersXml()
    {
        $LoggedOnUser =  Get-CimInstance -ComputerName $this.strComputerName -ClassName Win32_LoggedOnUser
        return ( Get-CimInstance  -ComputerName $this.strComputerName -ClassName Win32_LogonSession | ? { $_.LogonType -eq 2 -or  $_.LogonType -eq 10} | %{
                  $id = $_.LogonId
                  $usr = $LoggedOnUser | ? { $_.Dependent.LogonId -eq $id}
                            if($usr -ne $null)
                            {
                                New-Object -TypeName psobject -Property @{
                                    PSComputerName = $this.strComputerName
                                    StartTime =  $_.StartTime.ToString()
                                    DomainName = $usr.Antecedent.Domain
                                    UserName = $usr.Antecedent.Name
                                }
                            }

                  }|   ConvertTo-XML -NoTypeInformation)
    }

    [string] GetXmlString( [xml] $xml)
    {
         return $xml.OuterXml
    }

}


[clsWmiPs] $wmi = [clsWmiPs]::new($COMP_NAME)

$wmi.GetXmlString( $wmi.GetCpuXml())
$wmi.GetXmlString( $wmi.GetBiosXml())
$wmi.GetXmlString( $wmi.GetVideoXml())
$wmi.GetXmlString( $wmi.GetHddXml())
$wmi.GetXmlString( $wmi.GetRamXml())
$wmi.GetXmlString( $wmi.GetIpXml())
$wmi.GetXmlString( $wmi.GetUsersXml())
