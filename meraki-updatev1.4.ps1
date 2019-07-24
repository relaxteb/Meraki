###############################################################################################################################################
#                               MERAKI UPDATE
###############################################################################################################################################
#   This scripts can update hosts & ports in excel and meraki cloud
#
#
##############################################################################################################################################

#region Settings

# SET TLS version
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#endregion

#region Variables

# Import my settings
. C:\Temp\local_variables.ps1 # YOUR MERAKI API KEY !

# Import funtions
. .\meraki_functions_1.4.ps1



# API KEY
$api_key = $Meraki_api_key #ENTER YOUR API KEY HERE
# Test value
if  (!$api_key) {write-host "No API key found" -ForegroundColor Red
                exit
            }

# ENDPOINTS FOR API / ivm 308 redirects
$get_endpoint = 'https://dashboard.meraki.com/api/v0'
$put_endpoint = '' #'https://nXXXXXXXXX.meraki.com/api/v0'
#$get_endpoint = 'https://api.meraki.com/api/v0'
#$put_endpoint = 'https://api.meraki.com/api/v0'

if  (!$put_endpoint) {write-host "No PUT endpoint found" -ForegroundColor Red
                exit
            }

# EXCEL FILENAME
$excelfilename = "MerakiExcel template.xlsx"
$switchportsheetname = "switchports"
$firewallsheetname = "firewallrules"
$devicesheetname = "devices"
$ssidsheetname = "ssid"

if  (!(Test-Path $excelfilename))  {write-host "No EXCEL file found" -ForegroundColor Red
            write-host "Make sure to create the excel file and sheets: " -ForegroundColor Red
            write-host $switchportsheetname -ForegroundColor Yellow
            write-host $firewallsheetname -ForegroundColor Yellow
            Write-Host $devicesheetname -ForegroundColor Yellow
            write-host $ssidsheetname -ForegroundColor Yellow
            pause
            exit
            }


#endregion

#region Functions


      
   
    
  







#endregion

Clear-Host

# Get the Organization
write-host "Getting organizations: " -NoNewline
$organization = Get-MerakiOrganizations -endpointurl $get_endpoint -apikey $api_key
write-host "Organization: "$organization -ForegroundColor Green

# Get the Networks
write-host "Getting networks: " -NoNewline
$networks = Get-MerakiNetworks -organizationid $organization.id -endpointurl $get_endpoint -apikey $api_key 
write-host $networks -ForegroundColor Green

# Get the Devices
write-host "Getting Devices: " -NoNewline
$devices = Get-MerakiDevices -endpointurl $get_endpoint -apikey $api_key -networkid $networks.id
write-host $devices.count -ForegroundColor Green
start-sleep -seconds 1

# Do MENU
do {
    clear-host  
    write-host "--------------------Main Menu--------------------------------------"
    write-host "1. Device"
    write-host "2. Switchport"
    write-host "3. MX"
    write-host "4. WIFI"
    write-host ""
    write-host ""
    write-host "Q. Quit"
    write-host ""
    write-host ""
    write-host "Devices:`t`t`t" $devices.Count
    write-host "File:`t`t`t`t" $excelfilename
    write-host "Network:`t`t`t"$networks.name $networks.id
    write-host "Organization:`t`t`t"$organization
    write-host "-------------------------------------------------------------------"

    $input = Read-Host "Please make a selection"
    switch ($input) {
        '1' {
            # Device
            do {
                clear-host
                write-host "--------------------Device Menu -----------------------------------"
                write-host "1. Read Devices from Excel and update API"
                write-host "2. Read Devices from API and write to excel"
                write-host "3. Get Device info"
                write-host ""
                write-host ""
                write-host ""
                write-host "B. Back"
                write-host ""
                write-host "Sheet:`t`t`t`t"$devicesheetname
                write-host "Devices:`t`t`t" $devices.Count
                write-host "File:`t`t`t`t"$excelfilename
                write-host "Network:`t`t`t"$networks.name $networks.id
                write-host "Organization:`t`t`t"$organization
                write-host "-------------------------------------------------------------------"

                           
                $inputlevel1 = Read-Host "Please make a selection"
                switch ($inputlevel1) {
                    '1' {
                        # Read Devices from Excel and update API
                        Clear-host
                        
                        #$devices = read-MerakiDevicesFromExcel -filename $excelfilename -devicesheetname $devicesheetname
                        #foreach ($device in $devices) {
                            write-host "Updating devices...." -ForegroundColor Blue 
                            Update-MerakiDevices -endpointurl $put_endpoint -apikey $api_key -filename $excelfilename -devicesheetname $devicesheetname -networkid $networks.id 
                            
                    
                            # Max 5 Calls per second
                            # See https://documentation.meraki.com/zGeneral_Administration/Other_Topics/The_Cisco_Meraki_Dashboard_API
                            Start-Sleep -Milliseconds 200
                        #} 
                    pause
                    }

                    '2' {
                        # Read Devices from API and write to excel"
                        $devices = Get-MerakiDevices -endpointurl $get_endpoint -apikey $api_key -networkid $networks.id
                        write-MerakiDevices2Excel -devicearray $devices -filename $excelfilename -devicesheetname $devicesheetname
                        
                    }
                                   
                    '3' {
                        # Get Device info
                        Clear-host
                        $Inputdevice = Read-Host -Prompt 'Enter the devicename'
                        Get-MerakiDevice -endpointurl $get_endpoint -apikey $api_key -switchname $Inputdevice -networkid $networks.id
                        pause
                    }
                }  
            } until ($inputlevel1 -eq 'b')
        }
        '2' {
            # Ports
            do {
                clear-host
                write-host "--------------------Switchport Menu -------------------------------"
                write-host "1. Read Ports from Excel and update API"
                write-host "2. Get Switchports from API and write to excel"
                write-host "3. Get port info"
                write-host ""
                write-host ""
                write-host ""
                write-host "B. Back"
                write-host ""
                write-host "Sheet:`t`t`t`t"$switchportsheetname
                write-host "Devices:`t`t`t" $devices.Count
                write-host "File:`t`t`t`t"$excelfilename
                write-host "Network:`t`t`t"$networks.name $networks.id
                write-host "Organization:`t`t`t"$organization
                write-host "-------------------------------------------------------------------"

                $inputlevel2 = Read-Host "Please make a selection"
                switch ($inputlevel2) {
                    '1' {
                        # Read Ports from Excel and update API
                        Clear-host
                        write-host "Getting ports from excel"
                        $ports = read-MerakiSwitchPortsFromExcel -filename $excelfilename -switchportsheetname $switchportsheetname
                        foreach ($port in $ports) {
                            if ($port.Switch) {   
                                foreach ($device in $devices) {
                                    if ($port.switch -eq $device.name) { 
                                        $serial = $device.serial  
                                        
                                        write-host "Updating port:" $device.name $port.number $serial `
                                            "name:"$port.name  `
                                            "tags:"$port.tags `
                                            "enabled:"$port.enabled `
                                            "poeEnabled:"$port.poeEnabled `
                                            "type:"$port.type `
                                            "vlan:"$port.vlan `
                                            "voicevlan:"$port.voiceVlan `
                                            "allowedVlans:"$port.allowedVlans `
                                            "isolationEnabled:"$port.isolationEnabled `
                                            "rstpEnabled:"$port.rstpEnabled `
                                            "stpGuard:"$port.stpGuard -ForegroundColor Green
                    
                                        Update-MerakiSwitchPort -endpointurl $put_endpoint `
                                                                -apikey $api_key `
                                                                -port $port.number `
                                                                -serial $serial `
                                                                -name $port.name  `
                                                                -tags $port.tags `
                                                                -enabled $port.enabled `
                                                                -poeEnabled $port.poeEnabled `
                                                                -type $port.type `
                                                                -vlan $port.vlan `
                                                                -voicevlan $port.voiceVlan `
                                                                -allowedVlans $port.allowedVlans `
                                                                -isolationEnabled $port.isolationEnabled `
                                                                -rstpEnabled $port.rstpEnabled `
                                                                -stpGuard $port.stpGuard
                                        
                                        # Max 5 Calls per second
                                        # See https://documentation.meraki.com/zGeneral_Administration/Other_Topics/The_Cisco_Meraki_Dashboard_API
                                        write-host "Waiting 100 milli seconds" -foregroundcolor yellow
                                        Start-Sleep -Milliseconds 100
                                    }
                                }
                            }
                        }
                        pause
                    }

                    '2' {
                        # Read Ports from API and update Excel
                        Clear-Host
                        write-host "Getting devices from API"
                       
                        $devices = Get-MerakiDevices -endpointurl $get_endpoint -apikey $api_key -networkid $networks.id
                        write-host "Getting switchports from devices"
                        $switchports = @();
                        foreach ($device in $devices) {
                            if ($device.Model -like '*MS*') {
                                write-host "Found Device: " $device.name -ForegroundColor Green
                                $switchports += Get-MerakiSwitchPorts -endpointurl $get_endpoint -apikey $api_key -switchserial $device.serial -switchname $device.name

                            }
                        }
                        $switchports  = $switchports | Sort-Object -Property @{Expression = {$_.switch}; Ascending = $false}, number

                        write-MerakiSwitchports2Excel -filename $excelfilename -switchportsarray $switchports -switchportsheetname $switchportsheetname
                        
                        pause
                    }
                                   
                    '3' {
                        # Get port info
                        Clear-host
                        $Inputdevice = Read-Host -Prompt 'Enter the devicename'
                        $Inputport = Read-host -prompt 'Enter Switchport'
                        Get-MerakiSwitchPort -endpointurl $get_endpoint -apikey $api_key -switchname $Inputport -port $Inputport
                        pause
                    }
                }

            } until ($inputlevel2 -eq 'b')
        } 
                       
        '3' {
            # MX
            do {
                clear-host
                write-host "--------------------Firewall Menu -------------------------------------"
                write-host "1. Get Firewall Rules from API and write to excel"
                write-host "2. Get Firewall Rules from excel and write to API"
                write-host ""
                write-host ""
                write-host ""
                write-host ""
                write-host "B. Back"
                write-host ""
                write-host "Sheet:`t`t`t`t"$firewallsheetname
                write-host "Devices:`t`t`t" $devices.Count
                write-host "File:`t`t`t`t"$excelfilename
                write-host "Network:`t`t`t"$networks.name $networks.id
                write-host "Organization:`t`t`t"$organization
                write-host "-------------------------------------------------------------------"
                $inputlevel3 = Read-Host "Please make a selection"
                switch ($inputlevel3) {
                    '1' {
                        # Get Firewall Rules from API and write to excel
                        Clear-host
                        $firewallrules = Get-MerakiMXFirewallRules  -endpointurl $get_endpoint `
                                                                    -apikey $api_key `
                                                                    -networkid $networks.id

                        write-MerakiMXFirewallRules2Excel   -filename $excelfilename `
                                                            -firewallrulearray $firewallrules `
                                                            -firewallsheetname $firewallsheetname
                        pause
                    }

                    '2' {
                        # Get Firewall Rules from excel and write to API
                        Clear-Host
                        Update-MerakiFirewallrules  -endpointurl $put_endpoint `
                                                    -apikey $api_key `
                                                    -filename $excelfilename `
                                                    -firewallsheetname $firewallsheetname `
                                                    -networkid $networks.id `
                            pause
                        }

                    '3' {
                        # MX option 3
                        pause
                    }
                }
            } until ($inputlevel3 -eq 'b')
        }
        '4' {
            # WIFI
            do {
                clear-host
                write-host "--------------------WIFI Menu -------------------------------------"
                write-host "1. Get SSIDs from API and write to excel"
                write-host "2. Get SSIDs from excel and write to API"
                write-host ""
                write-host ""
                write-host ""
                write-host ""
                write-host "B. Back"
                write-host ""
                write-host "Sheet:`t`t`t`t"$ssidsheetname
                write-host "Devices:`t`t`t" $devices.Count
                write-host "File:`t`t`t`t"$excelfilename
                write-host "Network:`t`t`t"$networks.name $networks.id
                write-host "Organization:`t`t`t"$organization
                write-host "-------------------------------------------------------------------"
                $inputlevel4 = Read-Host "Please make a selection"
                switch ($inputlevel4) {
                    '1' {
                        # Get SSID from API and write to excel
                        Clear-host
                        $ssids = Get-MerakiWifiSSIDs  -endpointurl $get_endpoint `
                                                                    -apikey $api_key `
                                                                    -networkid $networks.id

                        write-MerakiWifiSSIDs2Excel   -filename $excelfilename `
                                                            -ssidarray $ssids `
                                                            -ssidsheetname $ssidsheetname
                        pause
                    }

                    '2' {
                        # Get SSID from excel and write to API
                        Clear-Host
                        Update-MerakiWifiSSIDs  -endpointurl $put_endpoint `
                                                    -apikey $api_key `
                                                    -filename $excelfilename `
                                                    -ssidsheetname $ssidsheetname `
                                                    -networkid $networks.id
                            pause
                        }

                    '3' {
                        # SSID option 3
                        pause
                    }
                }
            } until ($inputlevel4 -eq 'b')
        }
    }
                
} until ($input -eq 'q')


# Do garbage collection
[System.GC]::Collect()


