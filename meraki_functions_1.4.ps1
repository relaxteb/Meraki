###############################################################################################################################################
#                               MERAKI UPDATE
###############################################################################################################################################
#   This scripts contains Meraki Functions
#
#
##############################################################################################################################################



######################################################################################################################################################
#region webrequests
# Faultcode array 
# Source: https://documentation.meraki.com/zGeneral_Administration/Other_Topics/The_Cisco_Meraki_Dashboard_API#Status_and_Error_Codes
$faultcodes = @(
    ("400","Bad Request","You did something wrong, e.g. a malformed request or missing parameter."),
    ("403","Forbidden","You don't have permission to do that."),
    ("404","Not found","No such URL, or you don't have access to the API or organization at all.")
)


#endregion
######################################################################################################################################################

######################################################################################################################################################
#region General
function Get-MerakiOrganizations {
    <#
 .SYNOPSIS
 Get Meraki Organizations from the Meraki API
 
 .DESCRIPTION
 Get Meraki Organizations from the Meraki API
 The script retuns an array of Organizations
 
 .EXAMPLE
 get-merakiorganizations -endpointurl https://www.example.com/api/v0 -apikey 12345
 
 .NOTES
 General notes
 #>

    param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $endpointurl,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $apikey
    )

    $api = @{
        "method"   = "GET"
        "endpoint" = $endpointurl
        "url"      = '/organizations'
    }
 
    $header = @{
        "X-Cisco-Meraki-API-Key" = $apikey
        "Content-Type"           = 'application/json'
    }
 
    $uri = $api.endpoint + $api.url
 

    try {
        $request = Invoke-RestMethod -Method $api.method -Uri $uri -Headers $header 
        
    }
    catch {
        
        $_.Exception  
        Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__  -foreground red
        Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription -foreground red
        Pause
    }
    
 
    return $request
}

function Get-MerakiNetworks {
    <#
 .SYNOPSIS
 Get Meraki Networks for a specific organization
 
 .DESCRIPTION
 Get Meraki Networks from the Meraki API
 The script retuns an array of Networks
 
 .EXAMPLE
 get-merakinetworks -endpointurl https://www.example.com/api/v0 -apikey 12345 -organizationid Vecozo
 
 .NOTES
 General notes
 #>

    param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $endpointurl,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $apikey,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $organizationid
 
    )

    if ($organizationid) {

        $api = @{
            "method"   = "GET"
            "endpoint" = $endpointurl
            "url"      = '/organizations/' + $organizationid + '/networks'
        }

        $header = @{
            "X-Cisco-Meraki-API-Key" = $apikey
            "Content-Type"           = 'application/json'
        }

        $uri = $api.endpoint + $api.url
 

        try {
            $request = Invoke-RestMethod -Method $api.method -Uri $uri -Headers $header 
            
        }
        catch {
            
            $_.Exception  
            Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__  -foreground red
            Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription -foreground red
            Pause
        }
        
        return $request
    }
 
    else {
        Write-Host "Organization not entered." -ForegroundColor Red
        pause
    }
}

#endregion
######################################################################################################################################################

######################################################################################################################################################
#region Devices
function Get-MerakiDevice {
    <#
    .SYNOPSIS
    Get a specific device
    
    .DESCRIPTION
    Get the attributes of specific device by using it's name
    
    .PARAMETER endpointurl
    Parameter description
    
    .PARAMETER apikey
    Parameter description
    
    .PARAMETER networkid
    Parameter description
    
    .PARAMETER switchname
    Parameter description
    
    .EXAMPLE
    Get-MerakiDevice -endpointurl https://test.com -apikey 12345 -networkid 12345 -switchname testswitch
    
    .NOTES
    General notes
    #>
    

    param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $endpointurl,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $apikey,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$networkid,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$switchname
    )
    foreach ($device in $devices) {
        if ($device.name -eq $switchname) { 
            write-host "Serial found: " $device.serial -ForegroundColor green
            $serial = $device.serial
        }
    }
    
    
    if ($serial) {
        $api = @{
            "method"   = "GET"
            "endpoint" = $endpointurl
            "url"      = "/networks/" + $networkid + "/devices/" + $serial
        }

        $header = @{
            "X-Cisco-Meraki-API-Key" = $apikey
            "Content-Type"           = 'application/json'
        }
       
        $uri = $api.endpoint + $api.url

        try {
            $request = Invoke-RestMethod -Method $api.method -Uri $uri -Headers $header
        }
        catch {
            
            $_.Exception  
            Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__  -foreground red
            Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription -foreground red
            Pause
        }
        
        return $request
    
    }
    else {
        Write-Host "Serial not found for switch $switch_name." -ForegroundColor Red
    }
}

function Get-MerakiDevices {
    <#
    .SYNOPSIS
    Short description
    
    .DESCRIPTION
    Long description
    
    .PARAMETER endpointurl
    Parameter description
    
    .PARAMETER apikey
    Parameter description
    
    .PARAMETER networkid
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    General notes
    #>
    
    param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $endpointurl,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $apikey,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $networkid
    )

    write-host "Getting devices for network: " $networkid
    if ($networkid) {

        $api = @{
            "method"   = "GET"
            "endpoint" = $endpointurl
            "url"      = '/networks/' + $networkid + '/devices'
        }

        $header = @{
            "X-Cisco-Meraki-API-Key" = $apikey
            "Content-Type"           = 'application/json'
        }

        $uri = $api.endpoint + $api.url 

        try {
            $request = Invoke-RestMethod -Method $api.method -Uri $uri -Headers $header            
        }
        catch {
            
            $_.Exception  
            Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__  -foreground red
            Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription -foreground red
            Pause
        }
        return $request
    }
    else {
        Write-Host "Network ID not entered." -ForegroundColor Red
    }



}

function read-MerakiDevicesFromExcel {
    <#
.SYNOPSIS
Read Excel sheet  containing devices

.DESCRIPTION
Long description

.PARAMETER filename
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
    param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $filename,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $devicesheetname
    )
    # READ DEVICE info from Excel

    write-host "Reading hosts from:" $filename
    #https://docs.microsoft.com/en-us/office/vba/api/excel.range.specialcells
    $lastCell = 11 
    #$startRow, $model, $serial, $nOF, $ext = 1, 2, 3, 6, 7
    $startRow, $excelcolumnname, $excelcolumnserial, $excelcolumnvecid = 2, 2, 3, 6

    $excel = New-Object -ComObject excel.application
    $wb = $excel.workbooks.open($filename)
    #$sh = $wb.Sheets.Item(1)
    $sh = $wb.WorkSheets.Item($devicesheetname)
    $endRow = $sh.UsedRange.SpecialCells($lastCell).Row

    #$exceldevices = 
    for ($i = 1; $i -le $endRow; $i++) {
        $pName = $sh.Cells.Item($startRow, $excelcolumnname).Value2
        $pSerial = $sh.Cells.Item($startRow, $excelcolumnserial).Value2
        $pVECID = $sh.Cells.Item($startRow, $excelcolumnvecid).Value2
        if ($pName) {     
            write-host "Add switch:" $pName -ForegroundColor Green
            New-Object PSObject -Property @{ Name = $pName; Serial = $pSerial; VECID = $pVECID}#; NameOnPhone = $nameOnPhone; Extension = $extension }
        }
        $startRow++
    }

    write-host "Closing workbook..."
    $wb.close($true)
    Start-Sleep -Seconds 1

    write-host "Quit Excel..."
    $excel.quit()
    Start-Sleep -Seconds 1

    write-host "Releasing objects..."
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sh) | out-null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | out-null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | out-null
    Start-Sleep -Seconds 1

    write-host "Removing variables..."
    remove-variable sh
    Remove-Variable wb
    Remove-Variable excel
    Start-Sleep -Seconds 1

}    
function write-MerakiDevices2Excel {
    <#
.SYNOPSIS
Write Meraki devices from array to excel file

.DESCRIPTION
Write Meraki devices array to a specific sheet and file

.PARAMETER filename
Excel filename

.PARAMETER devicesheetname
Excel sheet containing devices

.PARAMETER devicearray
The array to be parsed

.EXAMPLE
write_merakidevices_to_excelfile -filename test.xls -devicesheetname devices -devicearray $devicearray

.NOTES
General notes
#>
  
    param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $filename,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $devicesheetname,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [Array] $devicearray
    )

    write-host "Writing Devices to Excel File:" $filename
    
    $excel = New-Object -ComObject excel.application
    $wb = $excel.workbooks.open($filename)

    $sh = $wb.WorkSheets.Item($devicesheetname)

    # Determine range of document
    #https://docs.microsoft.com/en-us/office/vba/api/excel.range.specialcells
                
    # Determine colums
    $ColumnName = $sh.Cells.Find('Name').Column
    $ColumnSerial = $sh.Cells.Find('Serial').Column
    $ColumnModel = $sh.Cells.Find('Model').Column
    $ColumnLanIP = $sh.Cells.Find('LanIP').Column
    $ColumnnetworkId = $sh.Cells.Find('networkId').Column
    $Columnlat = $sh.Cells.Find('lat').Column
    $Columnlng = $sh.Cells.Find('lng').Column
    $Columnmac = $sh.Cells.Find('mac').Column
    $Columnnotes = $sh.Cells.Find('notes').Column
    $ColumnbeaconIDParams = $sh.Cells.Find('beaconIDParams').Column
    $Columntags = $sh.Cells.Find('tags').Column

    # Reset Start Row
    #$startRow = 2 # First row is title bar

    for ($i = 0; $i -le $devicearray.Length - 1; $i++) {
        write-host "writing " $devicearray[$i].name $devicearray[$i].Serial
        $sh.Cells.Item($i + 2, $ColumnName) = $devicearray[$i].Name
        $sh.Cells.Item($i + 2, $ColumnSerial) = $devicearray[$i].Serial
        $sh.Cells.Item($i + 2, $ColumnModel) = $devicearray[$i].Model
        $sh.Cells.Item($i + 2, $ColumnLanIP) = $devicearray[$i].LanIP
        $sh.Cells.Item($i + 2, $ColumnnetworkId) = $devicearray[$i].networkId
        $sh.Cells.Item($i + 2, $Columnlat) = $devicearray[$i].lat
        $sh.Cells.Item($i + 2, $Columnlng) = $devicearray[$i].lng
        $sh.Cells.Item($i + 2, $Columnmac) = $devicearray[$i].mac
        $sh.Cells.Item($i + 2, $Columnnotes) = $devicearray[$i].notes
        #$sh.Cells.Item($i + 2, $ColumnbeaconIDParams) = $devicearray[$i].beaconIDParams
        $sh.Cells.Item($i + 2, $Columntags) = $devicearray[$i].tags
    } 
    

        
    write-host "Closing workbook..."
    $wb.close($true)
    Start-Sleep -Seconds 1

    write-host "Quit Excel..."
    $excel.quit()
    Start-Sleep -Seconds 1

    write-host "Releasing objects..."
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sh) | out-null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | out-null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | out-null
    Start-Sleep -Seconds 1

    write-host "Removing variables..."
    remove-variable sh
    Remove-Variable wb
    Remove-Variable excel
    Start-Sleep -Seconds 1
}      

function Update-MerakiDevices {
    param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $endpointurl,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $apikey,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $filename,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $devicesheetname,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $networkid
    )

    # READ EXCEL
    write-host "Writing Devices to API" 
 
    $excel = New-Object -ComObject excel.application
    $wb = $excel.workbooks.open($filename)
    $sh = $wb.WorkSheets.Item($devicesheetname)
 
    # Determine range of document
    #https://docs.microsoft.com/en-us/office/vba/api/excel.range.specialcells
    $lastCell = 11
 
    $endRow = $sh.UsedRange.SpecialCells($lastCell).Row

 
    $startRow = 2 # First row is title bar

 
    $endRow--

    # Determine colums
    $ColumnName = $sh.Cells.Find('Name').Column
    $ColumnSerial = $sh.Cells.Find('Serial').Column
    $ColumnnetworkId = $sh.Cells.Find('networkId').Column
    $Columnlat = $sh.Cells.Find('lat').Column
    $Columnlng = $sh.Cells.Find('lng').Column
    $Columnnotes = $sh.Cells.Find('notes').Column
    $Columntags = $sh.Cells.Find('tags').Column											
    #$Columnbeaconidparams = $sh.Cells.Find('beaconidparams').Column
    

    for ($startrow; $startrow -le $endRow; $startrow++) {
        
        
        $Name = $sh.Cells.Item($startRow, $ColumnName).Value2
        $Serial = $sh.Cells.Item($startRow, $ColumnSerial).Value2                
        $networkId = $sh.Cells.Item($startRow, $ColumnnetworkId).Value2             
        $lat = $sh.Cells.Item($startRow, $Columnlat).Value2                   
        $lng = $sh.Cells.Item($startRow, $Columnlng).Value2                   
        $notes = $sh.Cells.Item($startRow, $Columnnotes).Value2                 
        $tags = $sh.Cells.Item($startRow, $Columntags).Value2 
        


        write-host "Updating: " $name $serial

        if ($networkid) {
            if ($serial) {
                $api = @{
                    "method"   = "PUT"
                    "endpoint" = $endpointurl
                    "url"      = '/networks/' + $networkid + '/devices/' + $serial
                    }   
                $header = @{
                    "X-Cisco-Meraki-API-Key" = $apikey
                    "Content-Type"           = 'application/json'
                    }

            
                $body = @{
                    "name"          = $name
                    "tags"          = $tags
                    "lat"           = $lat
                    "lng"           = $lng
                    "notes"         = $notes
                    "moveMapMarker" = "false"
                    } | ConvertTo-Json

                $uri = $api.endpoint + $api.url 

                try {
                 $request = Invoke-RestMethod -Method $api.method -Uri $uri -Headers $header -Body $body
                    }

                catch {
                    
                    $_.Exception  
                    Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__  -foreground red
                    Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription -foreground red
                    Pause

                    }
                # Max 5 Calls per second
                # See https://documentation.meraki.com/zGeneral_Administration/Other_Topics/The_Cisco_Meraki_Dashboard_API
                Start-Sleep -Milliseconds 200
                #return $request
            }
        
            else {
            Write-Host "Serial not entered." -ForegroundColor Red
                }
        }
    
        else {
        Write-Host "Network ID not entered." -ForegroundColor Red   
            }
    }

        write-host "Closing workbook..."
        $wb.close($true)
        Start-Sleep -Seconds 1
    
        write-host "Quit Excel..."
        $excel.quit()
        Start-Sleep -Seconds 1
    
        write-host "Releasing objects..."
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sh) | out-null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | out-null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | out-null
        Start-Sleep -Seconds 1
    
        write-host "Removing variables..."
        remove-variable sh
        Remove-Variable wb
        Remove-Variable excel
        Start-Sleep -Seconds 1

    }

function Get-ManagementInterfaceSetting {
    <#
    .SYNOPSIS
    Get a specific device
    
    .DESCRIPTION
    Get the management interface settings of specific device by using it's serial
    
    .PARAMETER endpointurl
    Parameter description
    
    .PARAMETER apikey
    Parameter description
    
    .PARAMETER networkid
    Parameter description
    
    .PARAMETER switchname
    Parameter description
    
    .EXAMPLE
    Get-ManagementInterfaceSetting -endpointurl https://test.com -apikey 12345 -networkid 12345 -switchserial s43afsa3
    
    .NOTES
    General notes
    #>
    

    param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $endpointurl,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $apikey,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$networkid,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$switchserial
    )
    foreach ($device in $devices) {
        if ($device.serial -eq $switchserial) { 
            write-host "Serial found: " $device.serial -ForegroundColor green
            $serial = $device.serial
        }
    }
    
    
    if ($serial) {
        $api = @{
            "method"   = "GET"
            "endpoint" = $endpointurl
            "url"      = "/networks/" + $networkid + "/devices/" + $serial + "/managementInterfaceSettings"
            
        }

        $header = @{
            "X-Cisco-Meraki-API-Key" = $apikey
            "Content-Type"           = 'application/json'
        }
       
        $uri = $api.endpoint + $api.url

        try {
            $request = Invoke-RestMethod -Method $api.method -Uri $uri -Headers $header
        }
        catch {
            
            $_.Exception  
            Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__  -foreground red
            Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription -foreground red
            Pause
        }
        
        return $request
    
    }
    else {
        Write-Host "Serial not found for switch $switch_name." -ForegroundColor Red
    }
}




#endregion
######################################################################################################################################################

######################################################################################################################################################
#region Switches

function Get-MerakiSwitchPort {
    <#
    .SYNOPSIS
    Get specific switch port info
    
    .DESCRIPTION
    Get the specific info of a port on a switch
    
    .PARAMETER endpointurl
    Parameter description
    
    .PARAMETER apikey
    Parameter description
    
    .PARAMETER port
    Parameter description
    
    .PARAMETER switchname
    Parameter description
    
    .EXAMPLE
    get-merakiswitchport -endpointurl https://test.com -apikey 12345 -port 1 -switchname coreswitch
    
    .NOTES
    General notes
    #>
    

    param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $endpointurl,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $apikey,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$port,
        [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$switchname
    )
    
    foreach ($device in $devices) {
        if ($device.name -eq $switchname) { 
            $serial = $device.serial
        }
    }
    
    if ($port) {
        if ($serial) {

            $api = @{
                "method"   = "GET"
                "endpoint" = $endpointurl
                "url"      = "/devices/" + $serial + "/switchPorts/" + $port
            }

            $header = @{
                "X-Cisco-Meraki-API-Key" = $apikey
                "Content-Type"           = 'application/json'
            }

       
            $uri = $api.endpoint + $api.url

            try {
                $request = Invoke-RestMethod -Method $api.method -Uri $uri -Headers $header
            }
            catch {
                
                $_.Exception  
                Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__  -foreground red
                Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription -foreground red
                Pause
            }
            
            return $request
    
        }
        else {
            Write-Host "Serial not found for switch $switch_name." -ForegroundColor Red
        }
    }
    else {
        Write-Host "Port not found for switch $switch_name." -ForegroundColor Red
    }


}

function Get-MerakiSwitchPorts {
    <#
 .SYNOPSIS
 Get Meraki Switchports for a specific switch
 
 .DESCRIPTION
 Get Meraki Switchports for a specific switch from the Meraki API
 The script retuns an array of ports
 
 .EXAMPLE
 get-merakinetworks -endpointurl https://www.example.com/api/v0 -apikey 12345 -switchserial xxxxxxxx
 
 .NOTES
 General notes
 #>
 
    param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $endpointurl,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $apikey,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $switchserial,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $switchname
    )
 
    if ($switchserial) {
 
        $api = @{
            "method"   = "GET"
            "endpoint" = $endpointurl
            "url"      = "/devices/" + $switchserial + "/switchPorts"
        }
 
        $header = @{
            "X-Cisco-Meraki-API-Key" = $apikey
            "Content-Type"           = 'application/json'
        } 
 
        $uri = $api.endpoint + $api.url
        
        try {
            $request = Invoke-RestMethod -Method $api.method -Uri $uri -Headers $header
        }
        catch {
            
            $_.Exception  
            Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__  -foreground red
            Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription -foreground red
            Pause
        }
        
        # Add Column for Switch name
        $request | ForEach-Object {$_ | add-member -membertype NoteProperty -Name "switch" -value  $switchname}
        
        return $request
    }
 
    else {
        Write-Host "Serial not found for switch $switchserial." -ForegroundColor Red
        pause
    }
 
}
 
function Update-MerakiSwitchPort {
    <#
 .SYNOPSIS
 Update a specific switchport on a specific switch
 
 .DESCRIPTION
 This function uses the switchport variables to configure a switchport using the API
 
 .PARAMETER endpointurl
 PUT endpoint needed for API

 .PARAMETER apikey
 API key needed

 .PARAMETER port
 Port number identification ##
 
 .PARAMETER serial
 Switch serial number
 
 .PARAMETER name
 The name of the switch port
 
 .PARAMETER tags
 The tags of the switch port
 
 .PARAMETER enabled
 The status of the switch port
 
 .PARAMETER type
 The type of the switch port ("access" or "trunk")
 
 .PARAMETER vlan
 The VLAN of the switch port
 
 .PARAMETER voiceVlan
 The voice VLAN of the switch port. Only applicable to access ports.
 
 .PARAMETER allowedVlans
 The VLANs allowed on the switch port. Only applicable to trunk ports.
 
 .PARAMETER poeEnabled
 The PoE status of the switch port
 
 .PARAMETER isolationEnabled
 The isolation status of the switch port
 
 .PARAMETER rstpEnabled
 The rapid spanning tree protocol status
 
 .PARAMETER stpGuard
 The state of the STP guard ("disabled", "Root guard", "BPDU guard", "Loop guard")
 
 .PARAMETER accessPolicyNumber
 The number of the access policy of the switch port. Only applicable to access ports.
 
 .PARAMETER linkNegotiation
 The link speed for the switch port
 
 .EXAMPLE
 Update-MerakiSwitchPort -endpointurl https://www.test.test `
 -apikey 1234 `
 -port 1 `
 -name "My switch port" `
 -tags "tag1 tag2" `
 -enabled true `
 -poeEnabled true `
 -type "access" `
 -vlan 10 `
 -voiceVlan 20 `
 -isolationEnabled false `
 -rstpEnabled true `
 -stpGuard disabled `
 -accessPolicyNumber 1234 `
 -linkNegotiation "Auto negotiate"
 
 .NOTES
 General notes
 #>
 
 
    param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()]  [String] $endpointurl,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()]  [String] $apikey,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()]  [String] $port, 
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()]  [String] $serial,
        [Parameter(Mandatory = $false)]                            [String] $name,	 
        [Parameter(Mandatory = $false)]                            [String] $tags,	 
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [String] $enabled, 	 
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [String] $type, 	 
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [String] $vlan,	 
        [Parameter(Mandatory = $false)]                            [String] $voiceVlan,	 
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [String] $allowedVlans,	 
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [String] $poeEnabled,	 
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [String] $isolationEnabled,	 
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [String] $rstpEnabled,	 
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [String] $stpGuard, 	 
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [String] $accessPolicyNumber,	
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [String] $linkNegotiation	 
 
    )
    if ($port) {
        if ($serial) {
            $api = @{
 
                "method"   = "PUT"
                "endpoint" = $endpointurl
                "url"      = '/devices/' + $serial + '/switchPorts/' + $port
            } 
            $header = @{
                "X-Cisco-Meraki-API-Key" = $apikey
                "Content-Type"           = 'application/json'
            }
 
            $body = @{ 
                "name"               = $name
                "tags"               = $tags
                "enabled"            = $enabled
                "type"               = $type
                "vlan"               = $vlan
                "voiceVlan"          = $voiceVlan
                "allowedVlans"       = $allowedVlans
                "poeEnabled"         = $poeEnabled
                "isolationEnabled"   = $isolationEnabled
                "rstpEnabled"        = $rstpEnabled
                "stpGuard"           = $stpGuard
                "accessPolicyNumber" = $accessPolicyNumber
            } | ConvertTo-Json
 
            $uri = $api.endpoint + $api.url 

            try {
                $request = Invoke-RestMethod -Method $api.method -Uri $uri -Headers $header -Body $body
            }
            catch {
                
                $_.Exception  
                Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__  -foreground red
                Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription -foreground red
                Pause
            }
            
            return $request
        }
        else {
            Write-Host "Serial not entered." -ForegroundColor Red
        }
    }
    else {
        Write-Host "Port not entered." -ForegroundColor Red 
    }
}

function write-MerakiSwitchports2Excel {
    <#
 .SYNOPSIS
 Writing array of switchports to excel document
 
 .DESCRIPTION
 This function needs an array of switchports and will write them to an excel file / sheet specified
 
 .PARAMETER filename
 Parameter description
 
 .PARAMETER firewallrulearray
 Parameter description
 
 .EXAMPLE
 write_firewallrules_to_excelfile -filename test.xls -firewallrulearray $firewallrules -firewallsheetname testsheet
 
 .NOTES
 General notes
 #>
 
    param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $filename,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [Array] $switchportsarray,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $switchportsheetname
    )
 
    write-host "Writing Switchports to Excel File:" $filename
 
    $excel = New-Object -ComObject excel.application
    $wb = $excel.workbooks.open($filename)
 
    $sh = $wb.WorkSheets.Item($switchportsheetname)
 
    # Determine range of document
    #https://docs.microsoft.com/en-us/office/vba/api/excel.range.specialcells
    $lastCell = 11
 
    $endRow = $sh.UsedRange.SpecialCells($lastCell).Row
 
    $startRow = 2 # First row is title bar
 
    # Determine colums
    $Columnswitch = $sh.Cells.Find('switch').Column
    $Columnnumber = $sh.Cells.Find('number').Column
    $Columnname = $sh.Cells.Find('name').Column
    $Columntags = $sh.Cells.Find('tags').Column
    $Columnenabled = $sh.Cells.Find('enabled').Column
    $ColumnpoeEnabled = $sh.Cells.Find('poeEnabled').Column
    $Columntype = $sh.Cells.Find('type').Column
    $Columnvlan = $sh.Cells.Find('vlan').Column
    $ColumnvoiceVlan = $sh.Cells.Find('voiceVlan').Column
    $Columnnativevlan = $sh.Cells.Find('nativevlan').Column
    $ColumnallowedVlans = $sh.Cells.Find('allowedVlans').Column
    $Columntrusted = $sh.Cells.Find('trusted').Column
    $Columnudld = $sh.Cells.Find('udld').Column
    $ColumnaccessPolicyNumber = $sh.Cells.Find('accessPolicyNumber').Column
    $ColumnwhitelistedMacs = $sh.Cells.Find('whitelistedMacs').Column
    $Columnwhitelistsizelimit = $sh.Cells.Find('whitelistsizelimit').Column
    $ColumnrstpEnabled = $sh.Cells.Find('rstpEnabled').Column
    $ColumnstpGuard = $sh.Cells.Find('stpGuard').Column
    $ColumnisolationEnabled = $sh.Cells.Find('isolationEnabled').Column
    $Columnportschedule = $sh.Cells.Find('portschedule').Column
    $ColumnlinkNegotiation = $sh.Cells.Find('linkNegotiation').Column
    
    
 
    # Remove old content
    for ($startrow; $startrow -le $endRow; $startrow++) {
        $range = $sh.Cells.Item($startRow, 1).entirerow
        write-host "Deleting row: " $startRow
        $range.Delete()
    }
 
    # Reset Start Row
    $startRow = 2 # First row is title bar
 
    write-host "Array lenth: " $switchportsarray.Length
    for ($i = 0; $i -le $switchportsarray.Length - 1; $i++) {

        write-host "writing " $switchportsarray[$i].switch $switchportsarray[$i].number
        $sh.Cells.Item($i + 2, $Columnswitch) = $switchportsarray[$i].switch				
        $sh.Cells.Item($i + 2, $Columnnumber) = $switchportsarray[$i].number			 	
        $sh.Cells.Item($i + 2, $Columnname) = $switchportsarray[$i].name		 	
        $sh.Cells.Item($i + 2, $Columntags) = $switchportsarray[$i].tags		 	
        $sh.Cells.Item($i + 2, $Columnenabled) = $switchportsarray[$i].enabled		 	
        $sh.Cells.Item($i + 2, $ColumnpoeEnabled) = $switchportsarray[$i].poeEnabled			
        $sh.Cells.Item($i + 2, $Columntype) = $switchportsarray[$i].type	 	
        $sh.Cells.Item($i + 2, $Columnvlan) = $switchportsarray[$i].vlan		
        $sh.Cells.Item($i + 2, $ColumnvoiceVlan) = $switchportsarray[$i].voiceVlan		
        $sh.Cells.Item($i + 2, $Columnnativevlan) = $switchportsarray[$i].nativevlan		
        $sh.Cells.Item($i + 2, $ColumnallowedVlans) = $switchportsarray[$i].allowedVlans 		
        $sh.Cells.Item($i + 2, $Columntrusted) = $switchportsarray[$i].trusted		 	
        $sh.Cells.Item($i + 2, $Columnudld) = $switchportsarray[$i].udld	 	
        $sh.Cells.Item($i + 2, $ColumnaccessPolicyNumber) = $switchportsarray[$i].accessPolicyNumber 	
        $sh.Cells.Item($i + 2, $ColumnwhitelistedMacs) = $switchportsarray[$i].whitelistedMacs	
        $sh.Cells.Item($i + 2, $Columnwhitelistsizelimit) = $switchportsarray[$i].whitelistsizelimit
        $sh.Cells.Item($i + 2, $ColumnrstpEnabled) = $switchportsarray[$i].rstpEnabled			
        $sh.Cells.Item($i + 2, $ColumnstpGuard) = $switchportsarray[$i].stpGuard		
        $sh.Cells.Item($i + 2, $ColumnisolationEnabled) = $switchportsarray[$i].isolationEnabled
        $sh.Cells.Item($i + 2, $Columnportschedule) = $switchportsarray[$i].portschedule		
        $sh.Cells.Item($i + 2, $ColumnlinkNegotiation) = $switchportsarray[$i].linkNegotiation
		         
		
    } 
 
    write-host "Closing workbook..."
    $wb.close($true)
    Start-Sleep -Seconds 1

    write-host "Quit Excel..."
    $excel.quit()
    Start-Sleep -Seconds 1

    write-host "Releasing objects..."
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sh) | out-null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | out-null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | out-null
    Start-Sleep -Seconds 1

    write-host "Removing variables..."
    remove-variable sh
    Remove-Variable wb
    Remove-Variable excel
    Start-Sleep -Seconds 1
} 

function read-MerakiSwitchPortsFromExcel {

    param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $filename,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $switchportsheetname
        
    )
    # READ Port info from Excel
    write-host "Reading ports from:" $filename

    #https://docs.microsoft.com/en-us/office/vba/api/excel.range.specialcells
    $lastCell = 11 
    
    $startRow = 2 # First row is title bar
    


    $excel = New-Object -ComObject excel.application
    
    $wb = $excel.workbooks.open($filename)
    
    $sh = $wb.WorkSheets.Item($switchportsheetname)

    $ColumnSwitch = $sh.Cells.Find('Switch').Column
    $ColumnNumber = $sh.Cells.Find('Number').Column
    $ColumnName = $sh.Cells.Find('Name').Column
    $ColumnTags = $sh.Cells.Find('Tags').Column
    $ColumnEnabled = $sh.Cells.Find('Enabled').Column
    $ColumnPOEEnabled = $sh.Cells.Find('POEEnabled').Column
    $ColumnType = $sh.Cells.Find('Type').Column
    $Columnvlan = $sh.Cells.Find('vlan').Column
    $Columnvoicevlan = $sh.Cells.Find('voicevlan').Column
    $ColumnNativeVlan = $sh.Cells.Find('NativeVlan').Column
    $ColumnallowedVlans = $sh.Cells.Find('allowedVlans').Column
    $ColumnTrusted = $sh.Cells.Find('Trusted').Column
    $ColumnUDLD = $sh.Cells.Find('UDLD').Column
    $ColumnaccessPolicyNumber = $sh.Cells.Find('accessPolicyNumber').Column
    $ColumnWhitelistedMACs = $sh.Cells.Find('WhitelistedMACs').Column
    $ColumnWhitelistSizeLimit = $sh.Cells.Find('WhitelistSizeLimit').Column
    $ColumnrstpEnabled = $sh.Cells.Find('rstpEnabled').Column
    $ColumnstpGuard = $sh.Cells.Find('stpGuard').Column
    $ColumnisolationEnabled = $sh.Cells.Find('isolationEnabled').Column
    $ColumnPortSchedule = $sh.Cells.Find('PortSchedule').Column
    $ColumnlinkNegotiation = $sh.Cells.Find('linkNegotiation').Column

    $endRow = $sh.UsedRange.SpecialCells($lastCell).Row
        
    $excelports = for ($i = 1; $i -le $endRow; $i++) {
        $pSwitch = $sh.Cells.Item($startRow, $ColumnSwitch).Value2
        $pNumber = $sh.Cells.Item($startRow, $ColumnNumber).Value2
        $pName = $sh.Cells.Item($startRow, $ColumnName).Value2
        $pTags = $sh.Cells.Item($startRow, $ColumnTags).Value2
        $pEnabled = $sh.Cells.Item($startRow, $ColumnEnabled).Value2
        $pPOEEnabled = $sh.Cells.Item($startRow, $ColumnPOEEnabled).Value2
        $pType = $sh.Cells.Item($startRow, $ColumnType).Value2
        $pvlan = $sh.Cells.Item($startRow, $Columnvlan).Value2
        $pvoicevlan = $sh.Cells.Item($startRow, $Columnvoicevlan).Value2
        $pNativeVlan = $sh.Cells.Item($startRow, $ColumnNativeVlan).Value2
        $pallowedVlans = $sh.Cells.Item($startRow, $ColumnallowedVlans).Value2
        $pTrusted = $sh.Cells.Item($startRow, $ColumnTrusted).Value2
        $pUDLD = $sh.Cells.Item($startRow, $ColumnUDLD).Value2
        $paccessPolicyNumber = $sh.Cells.Item($startRow, $ColumnaccessPolicyNumber).Value2
        $pWhitelistedMACs = $sh.Cells.Item($startRow, $ColumnWhitelistedMACs).Value2
        $pWhitelistSizeLimit = $sh.Cells.Item($startRow, $ColumnWhitelistSizeLimit).Value2
        $prstpEnabled = $sh.Cells.Item($startRow, $ColumnrstpEnabled).Value2
        $pstpGuard = $sh.Cells.Item($startRow, $ColumnstpGuard).Value2
        $pisolationEnabled = $sh.Cells.Item($startRow, $ColumnisolationEnabled).Value2
        $pPortSchedule = $sh.Cells.Item($startRow, $ColumnPortSchedule).Value2
        $plinkNegotiation = $sh.Cells.Item($startRow, $ColumnlinkNegotiation).Value2

        if ($pSwitch) {     
            write-host "Reading port:" $pSwitch $pNumber -ForegroundColor Green
            New-Object PSObject -Property @{
                Switch             = $pSwitch;
                Number             = $pNumber;
                Name               = $pName;
                Tags               = $pTags;
                Enabled            = $pEnabled;
                POEEnabled         = $pPOEEnabled;
                Type               = $pType;
                vlan               = $pvlan;
                voicevlan          = $pvoicevlan;
                NativeVlan         = $pNativeVlan;
                allowedVlans       = $pallowedVlans;
                Trusted            = $pTrusted;
                UDLD               = $pUDLD;
                accessPolicyNumber = $paccessPolicyNumber;
                WhitelistedMACs    = $pWhitelistedMACs;
                WhitelistSizeLimit = $pWhitelistSizeLimit;
                rstpEnabled        = $prstpEnabled;
                stpGuard           = $pstpGuard;
                isolationEnabled   = $pisolationEnabled;
                PortSchedule       = $pPortSchedule;
                linkNegotiation    = $plinkNegotiation
            }#; NameOnPhone = $nameOnPhone; Extension = $extension }
        }
        $startRow++
    }
    
    write-host "Closing workbook..."
    $wb.close($true)
    Start-Sleep -Seconds 1

    write-host "Quit Excel..."
    $excel.quit()
    Start-Sleep -Seconds 1

    write-host "Releasing objects..."
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sh) | out-null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | out-null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | out-null
    Start-Sleep -Seconds 1

    write-host "Removing variables..."
    remove-variable sh
    Remove-Variable excel
    Remove-Variable wb
    Start-Sleep -Seconds 1

    return $excelports
}   

#endregion
######################################################################################################################################################

######################################################################################################################################################
#region Firewalls
function Get-MerakiMXFirewallRules {
<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER endpointurl
Parameter description

.PARAMETER apikey
Parameter description

.PARAMETER networkid
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>

    
    param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $endpointurl,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $apikey,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $networkid    
    # [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$port,
    # [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$switchname
    )
    write-host "Getting Firewall Rules from API"
 
    $api = @{
        "method"   = "GET"
        "endpoint" = $endpointurl
        "url"      = "/networks/" + $networkid + "/l3FirewallRules" 
    }
 
    $header = @{
        "X-Cisco-Meraki-API-Key" = $apikey
        "Content-Type"           = 'application/json'
    }
 
    $uri = $api.endpoint + $api.url

    try {
        $request = Invoke-RestMethod -Method $api.method -Uri $uri -Headers $header
    }
    catch {
        
        $_.Exception  
        Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__  -foreground red
        Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription -foreground red
        Pause
    }
    
    return $request
 
}

function Update-MerakiFirewallrules {
    <#
 .SYNOPSIS
 Update the MX L3 Firewall rules

 .DESCRIPTION
 Updates the entire MX L3 Firewall rules by using an array
 
 .PARAMETER endpointurl
 PUT endpoint needed for API

 .PARAMETER apikey
 API key needed
 
 .PARAMETER filename
 Excel filename needed for importing the rules
 
 .PARAMETER firewallsheetname
 Excel sheetname with firewall rules
 
 .PARAMETER networkid
 The network ID
 
 .EXAMPLE
 Update-MerakiFirewallrules -endpointurl http://test.com -apikey xxxx -filename test.xls -firewallsheetname firewall -networkid 12345
 
 .NOTES
 General notes
 #>
 
    param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $endpointurl,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $apikey,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $filename,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $firewallsheetname,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $networkid
    )
 
    write-host "Writing Firewall Rules to API" 
 
    $excel = New-Object -ComObject excel.application
    $wb = $excel.workbooks.open($filename)
    $sh = $wb.WorkSheets.Item($firewallsheetname)
 
    # Determine range of document
    #https://docs.microsoft.com/en-us/office/vba/api/excel.range.specialcells
    $lastCell = 11
 
    # Last Implicit firewall rules is explicitly defined in excel to display the real world configuration
    $endRow = $sh.UsedRange.SpecialCells($lastCell).Row
    write-host "end row = " $endRow
    $endRow--
    write-host "last row containing non default rule = " $endrow

    $startRow = 2 # First row is title bar
    write-host "start row = " $startRow
 
    
    # Determine colums
    $ColumnComment = $sh.Cells.Find('comment').Column
    $ColumnPolicy = $sh.Cells.Find('policy').Column
    $ColumnProtocol = $sh.Cells.Find('protocol').Column
    $ColumnSrcPort = $sh.Cells.Find('srcPort').Column
    $ColumnSrcCidr = $sh.Cells.Find('srcCidr').Column
    $ColumnDestPort = $sh.Cells.Find('destPort').Column 
    $ColumnDestCidr = $sh.Cells.Find('destCidr').Column 
    $ColumnSyslogEnabled = $sh.Cells.Find('syslogEnabled').Column
 
    $firewallrulearray = for ($startrow; $startrow -le $endRow; $startrow++) {
        $pComment = $sh.Cells.Item($startRow, $ColumnComment).Value2
        $pPolicy = $sh.Cells.Item($startRow, $ColumnPolicy).Value2
        $pProtocol = $sh.Cells.Item($startRow, $ColumnProtocol).Value2
        $pSrcPort = $sh.Cells.Item($startRow, $ColumnSrcPort).Value2
        $pSrcCidr = $sh.Cells.Item($startRow, $ColumnSrcCidr).Value2
        $pDestPort = $sh.Cells.Item($startRow, $ColumnDestPort).Value2
        $pDestCidr = $sh.Cells.Item($startRow, $ColumnDestCidr).Value2
        $pSyslogEnabled = $sh.Cells.Item($startRow, $ColumnSyslogEnabled).Value2
 
        write-host "Add firewall rule:" $pPolicy $pProtocol $pSrcPort $pSrcCidr $pDestPort $pDestCidr $pSyslogEnabled $pComment -ForegroundColor Green
        New-Object PSObject -Property @{ 
            comment       = $pComment;
            policy        = $pPolicy;
            protocol      = $pProtocol;
            srcPort       = $pSrcPort;
            srcCidr       = $pSrcCidr;
            destPort      = $pDestPort;
            destCidr      = $pDestCidr;
            syslogEnabled = $pSyslogEnabled
        }
    }
 
    write-host "Closing workbook..."
    $wb.close($true)
    Start-Sleep -Seconds 1

    write-host "Quit Excel..."
    $excel.quit()
    Start-Sleep -Seconds 1

    write-host "Releasing objects..."
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sh) | out-null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | out-null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | out-null
    Start-Sleep -Seconds 1

    write-host "Removing variables..."
    remove-variable sh
    Remove-Variable wb
    Remove-Variable excel

    Start-Sleep -Seconds 1
 
    # PUT/networks/[networkId]/l3FirewallRules
    $api = @{
        "method"   = "PUT"
        "endpoint" = $endpointurl
        "url"      = "/networks/" + $networkid + "/l3FirewallRules" 
    }
 
    $header = @{
        "X-Cisco-Meraki-API-Key" = $apikey
        "Content-Type"           = 'application/json'
    }
 
    #$body = $firewallrulearray | ConvertTo-Json
    $body = ConvertTo-Json @{rules = @( $firewallrulearray)}
 
    $uri = $api.endpoint + $api.url

    try {
        $request = Invoke-RestMethod -Method $api.method -Uri $uri -Headers $header -body $body
    }
    catch {
        
        $_.Exception  
        Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__  -foreground red
        Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription -foreground red
        Pause
    }
    
    return $request
 
} 

function write-MerakiMXFirewallRules2Excel {
    <#
 .SYNOPSIS
 Writing array of firewall rules to excel document
 
 .DESCRIPTION
 This function needs an array of firewall rules and will write them to an excel file / sheet specified
 
 .PARAMETER filename
 Parameter description
 
 .PARAMETER firewallrulearray
 Parameter description
 
 .EXAMPLE
 write_firewallrules_to_excelfile -filename test.xls -firewallrulearray $firewallrules -firewallsheetname testsheet
 
 .NOTES
 General notes
 #>
 
    param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $filename,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [Array] $firewallrulearray,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $firewallsheetname
    )
 
    write-host "Writing Firewall Rules to Excel File:" $filename
 
    $excel = New-Object -ComObject excel.application
    $wb = $excel.workbooks.open($filename)
 
    $sh = $wb.WorkSheets.Item($firewallsheetname)
 
    # Determine range of document
    #https://docs.microsoft.com/en-us/office/vba/api/excel.range.specialcells
    $lastCell = 11
 
    $endRow = $sh.UsedRange.SpecialCells($lastCell).Row
 
    $startRow = 2 # First row is title bar
 
    # Determine colums
    $ColumnComment = $sh.Cells.Find('comment').Column
    $ColumnPolicy = $sh.Cells.Find('policy').Column
    $ColumnProtocol = $sh.Cells.Find('protocol').Column
    $ColumnSrcPort = $sh.Cells.Find('srcPort').Column
    $ColumnSrcCidr = $sh.Cells.Find('srcCidr').Column
    $ColumnDestPort = $sh.Cells.Find('destPort').Column 
    $ColumnDestCidr = $sh.Cells.Find('destCidr').Column 
    $ColumnSyslogEnabled = $sh.Cells.Find('syslogEnabled').Column
 
    # Remove old content
    for ($startrow; $startrow -le $endRow; $startrow++) {
        $range = $sh.Cells.Item($startRow, 1).entirerow
        write-host "Deleting row: " $startRow
        $range.Delete()
    }
 
    # Reset Start Row
    $startRow = 2 # First row is title bar
 
    for ($i = 0; $i -le $firewallrulearray.Length - 1; $i++) {
 
        $sh.Cells.Item($i + 2, $ColumnComment) = $firewallrulearray[$i].comment
        $sh.Cells.Item($i + 2, $ColumnPolicy ) = $firewallrulearray[$i].policy 
        $sh.Cells.Item($i + 2, $ColumnProtocol) = $firewallrulearray[$i].protocol
        $sh.Cells.Item($i + 2, $ColumnSrcPort) = $firewallrulearray[$i].srcPort
        $sh.Cells.Item($i + 2, $ColumnSrcCidr ) = $firewallrulearray[$i].srcCidr 
        $sh.Cells.Item($i + 2, $ColumnDestPort) = $firewallrulearray[$i].destPort 
        $sh.Cells.Item($i + 2, $ColumnDestCidr) = $firewallrulearray[$i].destCidr 
        $sh.Cells.Item($i + 2, $ColumnSyslogEnabled ) = $firewallrulearray[$i].syslogEnabled
    } 
 
    write-host "Closing workbook..."
    $wb.close($true)
    Start-Sleep -Seconds 1

    write-host "Quit Excel..."
    $excel.quit()
    Start-Sleep -Seconds 1

    write-host "Releasing objects..."
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sh) | out-null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | out-null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | out-null
    Start-Sleep -Seconds 1

    write-host "Removing variables..."
    remove-variable sh
    Remove-Variable wb
    Remove-Variable excel

    Start-Sleep -Seconds 1
} 

#endregion
######################################################################################################################################################

######################################################################################################################################################
#region WIFI

function Get-MerakiWifiSSIDs {
    <#
    .SYNOPSIS
    Short description
    
    .DESCRIPTION
    Long description
    
    .PARAMETER endpointurl
    Parameter description
    
    .PARAMETER apikey
    Parameter description
    
    .PARAMETER networkid
    Parameter description
    
    .EXAMPLE
    An example
    
    .NOTES
    General notes
    #>
    
        
        param (
            [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $endpointurl,
            [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $apikey,
            [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $networkid    
        # [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$port,
        # [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][String]$switchname
        )
        write-host "Getting SSIDs from API"
     
        $api = @{
            "method"   = "GET"
            "endpoint" = $endpointurl
            "url"      = "/networks/" + $networkid + "/ssids" 
        }
     
        $header = @{
            "X-Cisco-Meraki-API-Key" = $apikey
            "Content-Type"           = 'application/json'
        }
     
        $uri = $api.endpoint + $api.url
    
        try {
            $request = Invoke-RestMethod -Method $api.method -Uri $uri -Headers $header
        }
        catch {
            
            $_.Exception  
            Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__  -foreground red
            Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription -foreground red
            Pause
        }
        
        return $request
     
    }
    
   

function Update-MerakiWifiSSIDs {
        <#
     .SYNOPSIS
     Update the MX L3 Firewall rules
    
     .DESCRIPTION
     Updates the entire MX L3 Firewall rules by using an array
     
     .PARAMETER endpointurl
     PUT endpoint needed for API
    
     .PARAMETER apikey
     API key needed
     
     .PARAMETER filename
     Excel filename needed for importing the rules
     
     .PARAMETER firewallsheetname
     Excel sheetname with firewall rules
     
     .PARAMETER networkid
     The network ID
     
     .EXAMPLE
     Update-MerakiFirewallrules -endpointurl http://test.com -apikey xxxx -filename test.xls -firewallsheetname firewall -networkid 12345
     
     .NOTES
     General notes
     #>
     
        param (
            [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $endpointurl,
            [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $apikey,
            [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $filename,
            [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $ssidsheetname,
            [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $networkid
            
            
        )
     
        write-host "Reading SSIDs from excel to array" 
     
        $excel = New-Object -ComObject excel.application
        $wb = $excel.workbooks.open($filename)
        $sh = $wb.WorkSheets.Item($ssidsheetname)
     
        # Determine range of document
        #https://docs.microsoft.com/en-us/office/vba/api/excel.range.specialcells
        $lastCell = 11
     
        # Last Implicit firewall rules is explicitly defined in excel to display the real world configuration
        $endRow = $sh.UsedRange.SpecialCells($lastCell).Row
        $endRow--
        $startRow = 2 # First row is title bar
                
        # Determine colums
        $ColumnNumber = $sh.Cells.Find('number').Column
        $ColumnName = $sh.Cells.Find('name').Column
        $ColumnEnabled = $sh.Cells.Find('enabled').Column
        $ColumnSplashPage = $sh.Cells.Find('splashPage').Column
        $ColumnSsidAdminAccessible = $sh.Cells.Find('ssidAdminAccessible').Column
        $ColumnAuthMode = $sh.Cells.Find('authMode').Column 
        $ColumnIpAssignmentMode = $sh.Cells.Find('ipAssignmentMode').Column 
        $ColumnBandSelection = $sh.Cells.Find('bandSelection').Column
        $ColumnPerClientBandwidthLimitUp = $sh.Cells.Find('perClientBandwidthLimitUp').Column
        $ColumnPerClientBandwidthLimitDown = $sh.Cells.Find('perClientBandwidthLimitDown').Column
     
        $ssidarray = for ($startrow; $startrow -le $endRow; $startrow++) {
            $pNumber						= $sh.Cells.Item($startRow, $ColumnNumber).Value2
            $pName         					= $sh.Cells.Item($startRow, $ColumnName).Value2
            $pEnabled        				= $sh.Cells.Item($startRow, $ColumnEnabled).Value2
            $pSplashPage        			= $sh.Cells.Item($startRow, $ColumnSplashPage).Value2
            $pSsidAdminAccessible           = $sh.Cells.Item($startRow, $ColumnSsidAdminAccessible).Value2
            $pAuthMode                      = $sh.Cells.Item($startRow, $ColumnAuthMode).Value2
            $pIpAssignmentMode         		= $sh.Cells.Item($startRow, $ColumnIpAssignmentMode).Value2
            $pBandSelection        			= $sh.Cells.Item($startRow, $ColumnBandSelection).Value2
            $pPerClientBandwidthLimitUp     = $sh.Cells.Item($startRow, $ColumnPerClientBandwidthLimitUp).Value2
            $pPerClientBandwidthLimitDown   = $sh.Cells.Item($startRow, $ColumnPerClientBandwidthLimitDown).Value2
     
            write-host "Add SSID :" $pNumber $pName $pEnabled $pSplashPage $pSsidAdminAccessible $pAuthMod $pIpAssignmentMode $pBandSelection $pPerClientBandwidthLimitUp $pPerClientBandwidthLimitDown              -ForegroundColor Green
            New-Object PSObject -Property @{ 
                Number						= $pNumber;
                Name         				= $pName;      					
                Enabled        				= $pEnabled;
                SplashPage        			= $pSplashPage;      
                SsidAdminAccessible         = $pSsidAdminAccessible;
                AuthMode                    = $pAuthMode;  			
                IpAssignmentMode         	= $pIpAssignmentMode;         		
                BandSelection        		= $pBandSelection;        			
                PerClientBandwidthLimitUp   = $pPerClientBandwidthLimitUp;        
                PerClientBandwidthLimitDown = $pPerClientBandwidthLimitDown
            }
        }
     
        write-host "Closing workbook..."
        $wb.close($true)
        Start-Sleep -Seconds 1
    
        write-host "Quit Excel..."
        $excel.quit()
        Start-Sleep -Seconds 1
    
        write-host "Releasing objects..."
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sh) | out-null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | out-null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | out-null
        Start-Sleep -Seconds 1
    
        write-host "Removing variables..."
        remove-variable sh
        Remove-Variable wb
        Remove-Variable excel
    
        Start-Sleep -Seconds 1
     
        # PUT/networks/[networkId]/ssids/[number]

        foreach ($ssidentry in $ssidarray){
            
            write-host "Updating " $ssidentry.number $ssidentry.name -foreground green

            $api = @{
                "method"   = "PUT"
                "endpoint" = $endpointurl
                "url"      = "/networks/" + $networkid + "/ssids/" + $ssidentry.number
                }
     
            $header = @{
            "X-Cisco-Meraki-API-Key" = $apikey
            "Content-Type"           = 'application/json'
                }
     
        
            $body = ConvertTo-Json $ssidentry
     
            $uri = $api.endpoint + $api.url
    
            try {
            
                $request = Invoke-RestMethod -Method $api.method -Uri $uri -Headers $header -body $body
                }
            catch {
            
                $_.Exception  
                Write-Host "StatusCode:" $_.Exception.Response.StatusCode.value__  -foreground red
                Write-Host "StatusDescription:" $_.Exception.Response.StatusDescription -foreground red
                Pause
                }       
        
            $request
           
         }
     
    } 


function write-MerakiWifiSSIDs2Excel {
        <#
     .SYNOPSIS
     Writing array of firewall rules to excel document
     
     .DESCRIPTION
     This function needs an array of firewall rules and will write them to an excel file / sheet specified
     
     .PARAMETER filename
     Parameter description
     
     .PARAMETER firewallrulearray
     Parameter description
     
     .EXAMPLE
     write_firewallrules_to_excelfile -filename test.xls -firewallrulearray $firewallrules -firewallsheetname testsheet
     
     .NOTES
     General notes
     #>
     
        param (
            [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $filename,
            [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [Array] $ssidarray,
            [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String] $ssidsheetname
        )
     
        write-host "Writing SSIDs to Excel File:" $filename
     
        $excel = New-Object -ComObject excel.application
        $wb = $excel.workbooks.open($filename)
     
        $sh = $wb.WorkSheets.Item($ssidsheetname)
     
        # Determine range of document
        #https://docs.microsoft.com/en-us/office/vba/api/excel.range.specialcells
        $lastCell = 11
     
        $endRow = $sh.UsedRange.SpecialCells($lastCell).Row
     
        $startRow = 2 # First row is title bar
     
        # Determine colums
        $ColumnNumber = $sh.Cells.Find('number').Column
        $ColumnName = $sh.Cells.Find('name').Column
        $ColumnEnabled = $sh.Cells.Find('enabled').Column
        $ColumnSplashPage = $sh.Cells.Find('splashPage').Column
        $ColumnssidAdminAccessible = $sh.Cells.Find('ssidAdminAccessible').Column
        $ColumnAuthMode = $sh.Cells.Find('authMode').Column 
        $ColumnIpAssignmentMode = $sh.Cells.Find('ipAssignmentMode').Column 
        $ColumnBandSelection = $sh.Cells.Find('bandSelection').Column
        $ColumnPerClientBandwidthLimitUp = $sh.Cells.Find('perClientBandwidthLimitUp').Column
        $ColumnPerClientBandwidthLimitDown = $sh.Cells.Find('perClientBandwidthLimitDown').Column
     
        # Remove old content
        for ($startrow; $startrow -le $endRow; $startrow++) {
            $range = $sh.Cells.Item($startRow, 1).entirerow
            write-host "Deleting row: " $startRow
            $range.Delete()
        }
     
        # Reset Start Row
        $startRow = 2 # First row is title bar
     
        for ($i = 0; $i -le $ssidarray.Length - 1; $i++) {
     
            $sh.Cells.Item($i + 2, $ColumnNumber)                       = $ssidarray[$i].number
            $sh.Cells.Item($i + 2, $ColumnName)                         = $ssidarray[$i].name 
            $sh.Cells.Item($i + 2, $ColumnEnabled)                      = $ssidarray[$i].enabled
            $sh.Cells.Item($i + 2, $ColumnSplashPage)                   = $ssidarray[$i].splashpage
            $sh.Cells.Item($i + 2, $ColumnssidAdminAccessible)          = $ssidarray[$i].ssidadminaccessible 
            $sh.Cells.Item($i + 2, $ColumnAuthMode)                     = $ssidarray[$i].authmode 
            $sh.Cells.Item($i + 2, $ColumnIpAssignmentMode)             = $ssidarray[$i].ipassignmentmode 
            $sh.Cells.Item($i + 2, $ColumnBandSelection)                = $ssidarray[$i].bandselection
            $sh.Cells.Item($i + 2, $ColumnPerClientBandwidthLimitUp)    = $ssidarray[$i].perclientbandwidthlimitup
            $sh.Cells.Item($i + 2, $ColumnPerClientBandwidthLimitDown)  = $ssidarray[$i].perclientbandwidthlimitdown
        } 
     
        write-host "Closing workbook..."
        $wb.close($true)
        Start-Sleep -Seconds 1
    
        write-host "Quit Excel..."
        $excel.quit()
        Start-Sleep -Seconds 1
    
        write-host "Releasing objects..."
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($sh) | out-null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | out-null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | out-null
        Start-Sleep -Seconds 1
    
        write-host "Removing variables..."
        remove-variable sh
        Remove-Variable wb
        Remove-Variable excel
    
        Start-Sleep -Seconds 1
    } 
    
#endregion
######################################################################################################################################################
