[string] $COMP_NAME             = $env:COMPUTERNAME

function ConvertTo-DataTable
{
    <# 
    .Synopsis 
        Creates a DataTable from an object 
    .Description 
        Creates a DataTable from an object, containing all properties 
    .Example 
        Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=$true | Select-Object DNSServerSearchOrder,DHCPEnabled,IPAddress |  ConvertTo-DataTable
    .Link 
        Select-DataTable 
    #> 
    [CmdletBinding()]
    [OutputType([System.Data.DataTable])]
    param(
    # The input objects
    [PARAMETER(Position=0, Mandatory=$True, ValueFromPipeline = $true)]  [PSObject[]] $InputObject,
	[PARAMETER(Position=1, Mandatory=$False, HelpMessage = "Table Name")][String]    $TableName='Object'
    ) 
 
    begin 
    { 
        $outputDataTable = [System.Data.DataTable]::new($TableName)
        $types = @(
                'System.Boolean',
                'System.Byte',
                'System.Char',
                'System.Datetime',
                'System.Decimal',
                'System.Double',
                'System.Guid',
                'System.Int16',
                'System.Int32',
                'System.Int64',
                'System.Single',
                'System.String',
                'System.UInt16',
                'System.UInt32',
                'System.UInt64')
    } 

    process {         
               
        foreach ($In in $InputObject) 
        { 
            $DataRow = $outputDataTable.NewRow()   
             
            foreach($property in $In.PsObject.properties) 
            {   

                $propName     =  $property.Name
                $propValue    =  $property.Value
                $propType     =  'System.Object'
                $isSimpleType =  $types -contains $property.TypeNameOfValue
                if ($isSimpleType) 
                {
                    $propType = $property.TypeNameOfValue
                } 

                if (-not $outputDataTable.Columns.Contains($propName)) 
                {   
                   $outputDataTable.Columns.Add($propName,[System.Type]::GetType($propType)) | Out-Null
                }                   
                
                Try
                {
                    $DataRow.Item($propName) = if ($isSimpleType -and $propValue) 
                    {
                        $propValue
                    }
                    elseif ($property.TypeNameOfValue -eq "System.String[]" -and $propValue) 
                    { 
                        $propValue -join ","
                    }
                    elseif ($property.TypeNameOfValue.ToString().Contains("[]") -and $propValue) 
                    { 
                        $propValue | ConvertTo-XML -As String -NoTypeInformation -Depth 1
                    }
                    elseif ($propValue) 
                    {
                        [PSObject]$propValue
                    } 
                    else 
                    {
                        [DBNull]::Value
                    }
                }
                Catch
                {
                    Write-Error "Could not add property $propName with value $propValue and type $propType"
                    continue
                }

            }
            Try
            {   
                $outputDataTable.Rows.Add($DataRow)   
            }
            Catch
            {
                Write-Error "Failed to add row '$($DataRow | Out-String)':`n$_"
            }
        } 
    }  
    end 
    { 
        ,$outputDataTable
    } 
 
}



  Get-CimInstance -Class Win32_NetworkAdapterConfiguration -Filter IPEnabled=$true -ComputerName $COMP_NAME   | 
    Select-Object PSComputerName,DNSHostName,Description,Caption,DHCPEnabled,IPAddress,IPSubnet,MACAddress,DNSServerSearchOrder,DNSDomainSuffixSearchOrder,DefaultIPGateway |   ConvertTo-DataTable -TableName "Win32_NetworkAdapterConfiguration"
  
  Get-CimInstance -Class Win32_Processor -ComputerName $COMP_NAME  |
    Select-Object  PSComputerName,Manufacturer,Name,DeviceID,NumberOfCores,NumberOfLogicalProcessors,CurrentClockSpeed,L2CacheSize,L3CacheSize | ConvertTo-DataTable -TableName  "Win32_Processor"

  Get-CimInstance -Class Win32_BIOS -ComputerName $COMP_NAME  | 
    Select-Object PSComputerName,Name,SerialNumber,Version,Description,SMBIOSBIOSVersion,SMBIOSMajorVersion |  ConvertTo-DataTable   -TableName  "Win32_BIOS"  
  
  Get-CimInstance -Class Win32_VideoController -ComputerName $COMP_NAME  | 
    Select-Object PSComputerName,AdapterCompatibility,AdapterDACType,AdapterRAM,Description,DriverDate,DriverVersion,Name,VideoModeDescription,VideoProcessor |  ConvertTo-DataTable  -TableName "Win32_VideoController"
  
  Get-CimInstance -Class Win32_DiskDrive -ComputerName $COMP_NAME  | 
    Select-Object PSComputerName,Description,FirmwareRevision,Model,Manufacturer,Partitions,SerialNumber,Size,Status,SCSIBus,SCSILogicalUnit,SCSIPort,SCSITargetId,InterfaceType | ConvertTo-DataTable  -TableName "Win32_DiskDrive"
  
  Get-CimInstance -Class Win32_PhysicalMemory -ComputerName $COMP_NAME  | 
    Select-Object PSComputerName,BankLabel,Capacity,DeviceLocator,Model,Manufacturer,FormFactor,PartNumber,SerialNumber,Speed | ConvertTo-DataTable  -TableName  "Win32_PhysicalMemory"


   $LoggedOnUser =  Get-CimInstance -ComputerName $COMP_NAME -ClassName Win32_LoggedOnUser 
   Get-CimInstance  -ComputerName $COMP_NAME -ClassName Win32_LogonSession | ? { $_.LogonType -eq 2 -or  $_.LogonType -eq 10} | %{
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

                  }|    ConvertTo-DataTable  -TableName  "Win32_LoggedOnUser"
