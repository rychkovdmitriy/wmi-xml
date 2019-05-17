# wmi-xml

Данные скрипты разными способоами формируют XML файлы:
1) Файл Wmi-XmlPs.ps1
Использует WMI  запрос  можно сказать в одну строчку получаем то что нужно:

Get-WmiObject -Class Win32_BIOS  | 
Select-Object PSComputerName,Name,SerialNumber,Version,Description,SMBIOSBIOSVersion,SMBIOSMajorVersion
| ConvertTo-XML -NoTypeInformation -as String

В результате получится файл вида:

<?xml version="1.0" encoding="utf-8"?>
<Objects>
  <Object>
    <Property Name="PSComputerName">WIN-E3RPT5J1UC0</Property>
    <Property Name="Name">Default System BIOS</Property>
    <Property Name="SerialNumber">SN0988MPQ</Property>
    <Property Name="Version">030717 - 20170307</Property>
    <Property Name="Description">Default System BIOS</Property>
    <Property Name="SMBIOSBIOSVersion">080016 </Property>
    <Property Name="SMBIOSMajorVersion">2</Property>
  </Object>

Такой вариант получается самым простым сопобом с помощью  ConvertTo-XML

2) Файл Wmi-XmlDocument.ps1

Получаем WMI данные и заполняем  XmlDocument
Для удобства есть функция которая получает класс WMI и массив атрибутов:
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

Благодаря данной функции  получаем XML файл, достаточно кратко:

        $attribCpu = @("PSComputerName","Manufacturer","Name","DeviceID",
        "NumberOfCores","NumberOfLogicalProcessors","CurrentClockSpeed","L2CacheSize","L3CacheSize")
        
        GetWmiXml("Win32_Processor",$attribCpu)
        
        
Упрощенный пример:

class clsWmiXmlDocument
{
        [xml] GetWmiXml([string] $wmiClassName,[System.Object[]] $wmiAttr)
        {

            $xml=[xml] '<Objects/>'
            $wmiObj = Get-CimInstance -Class $wmiClassName -ComputerName  $env:COMPUTERNAME  | Select-Object $wmiAttr
            if($wmiObj -eq $null)
            {
                return $xml
            }
       
            $wmiObj |%{
                $n=$xml.DocumentElement.AppendChild($xml.CreateElement($wmiClassName))
                foreach($atr in $wmiAttr)
                {
                    $c=$n.AppendChild($xml.CreateElement($atr))
                    $c.InnerText = $_.($atr)
                }
            }
            return $xml

        }


        [xml] GetBiosXml()
        {
            $attribBios = @("PSComputerName","Manufacturer","Name","SerialNumber","Version","Description","SMBIOSBIOSVersion","SMBIOSMajorVersion")
            return $this.GetWmiXml("Win32_BIOS",$attribBios)
        }
    }

    [clsWmiXmlDocument] $wmi = [clsWmiXmlDocument]::new()
    $wmi.GetBiosXml().OuterXml


В результате получится файл вида:

<Objects>
  <Win32_BIOS>
    <PSComputerName>WIN-E3RPT5J1UC0</PSComputerName>
    <Manufacturer>American Megatrends Inc.</Manufacturer>
    <Name>Default System BIOS</Name>
    <SerialNumber>SN0988MPQ</SerialNumber>
    <Version>030717 - 20170307</Version>
    <Description>Default System BIOS</Description>
    <SMBIOSBIOSVersion>080016 </SMBIOSBIOSVersion>
    <SMBIOSMajorVersion>2</SMBIOSMajorVersion>
  </Win32_BIOS>
</Objects>

