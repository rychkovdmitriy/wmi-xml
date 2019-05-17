# wmi-xml

Данные скрипты разными способоами формируют XML файлы:
1) Файл **Wmi-XmlPs.ps1**
Использует WMI  запрос  можно сказать в одну строчку получаем то что нужно:
```powershell
Get-WmiObject -Class Win32_BIOS  | 
Select-Object PSComputerName,Name,SerialNumber,Version,Description,SMBIOSBIOSVersion,SMBIOSMajorVersion
| ConvertTo-XML -NoTypeInformation -as String
```
В результате получится файл вида:
```xml
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
```
Такой вариант получается самым простым сопобом с помощью  ConvertTo-XML

2) Файл **Wmi-XmlDocument.ps1**

Получаем WMI данные и заполняем  XmlDocument
Для удобства есть функция которая получает класс WMI и массив атрибутов:
```powershell
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
```
Благодаря данной функции  получаем XML файл, достаточно кратко:
```powershell
        $attribCpu = @("PSComputerName","Manufacturer","Name","DeviceID",
        "NumberOfCores","NumberOfLogicalProcessors","CurrentClockSpeed","L2CacheSize","L3CacheSize")
        
        GetWmiXml("Win32_Processor",$attribCpu)
```
        
Упрощенный пример:
```powershell
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

```
В результате получится файл вида:
```xml
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

```
2) Файл **Wmi-XmlDocument.ps1**

Аналагично предидущему, получаем с помощью запроса WMI данные , и заполняем уже таблицу DataTable
Это может быть удобно, если данные хотим передать в базу данных.

Простой пример:
```powershell
class clsWmiXmlDataTable
{
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
        $wmiObj = Get-CimInstance -Class $wmiClassName  -ComputerName $env:COMPUTERNAME | Select-Object $wmiAttr
        if($wmiObj -eq $null)
        {
            return $null
        }
       
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
                $row[$colName] = $value

            }
            $tabWmi.Rows.Add($row)
        }
        return $tabWmi

    }
    
    [System.Data.DataTable] GetBiosTab()
    {
        $attribBios = @("PSComputerName","Manufacturer","Name","SerialNumber","Version","Description","SMBIOSBIOSVersion","SMBIOSMajorVersion")
        return $this.GetWmiTab("Win32_BIOS",$attribBios)
    }

 }

 
[clsWmiXmlDataTable] $wmi = [clsWmiXmlDataTable]::new()
$wmi.GetBiosTab()
```

Результат будет таким:
```
PSComputerName     : WIN-E3RPT5J1UC0
Manufacturer       : American Megatrends Inc.
Name               : Default System BIOS
SerialNumber       : SN0988MPQ
Version            : 030717 - 20170307
Description        : Default System BIOS
SMBIOSBIOSVersion  : 080016 
SMBIOSMajorVersion : 2
```

Из таблицы так же можно получить XML документ, с помощью метода:
```powershell
    [string] GetXmlString( [System.Data.DataTable] $tab)
    {
         [System.IO.StringWriter] $writer = [System.IO.StringWriter]::new()
         [System.Data.DataSet] $ds = [System.Data.DataSet]::New("Objects")
         $ds.Tables.Add($tab)
         $ds.WriteXml($writer)
         return $writer.ToString()
    }
```

В результате опять XML:
```xml
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
```
