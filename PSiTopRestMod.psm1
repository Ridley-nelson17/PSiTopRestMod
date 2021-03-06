<#
 Tool:    PSiTopRestMod.psm1
 Author:  Johann Enquist
 Email:   administrator@boomandfreeze.com
 NOTES:   Powershell module to interact with iTop Web API
#>

function Get-iTopObject {
<#
.SYNOPSIS
  Generic function to query iTop server with specific OQL query.
.DESCRIPTION
  Sends a core/get operation to the iTop REST API.
.PARAMETER ServerAddress
  FQDN of the iTop server you are running against.
.PARAMETER Protocol
  Whether you are connecting to the iTop instance over HTTP or HTTPS. Default is HTTPS.
.PARAMETER Credential
  The credentials that you are going to authenticate against the iTop REST API.
.PARAMETER Class
  The value to be passed into the class property of the JSON.
.PARAMETER OQLFilter
  Custom OQL query to be used.
.NOTES
.EXAMPLE
  Get-iTopObject -ServerAddress $itopserver -Credential $apiuser -Protocol http -Class VirtualMachine -OQLFilter "SELECT VirtualMachine WHERE osfamily_name = 'Linux'"
  Retrieve all VMs that have Linux as their osfamily_name property.
.LINK
  https://github.com/jenquist/PSiTopRestMod
#> 
[CmdletBinding()]
param(    
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$Class,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$OQLFilter
)
  # Creating header with credentials being used for authentication
  [string]$username = $Credential.UserName
  [string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
  $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  $headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

  # Creating in-line JSON to be sent within URI
  $sendJSON = @{
               operation = 'core/get'
               class = "$Class"
               key = ("$OQLFilter")
               output_fields = '*'
               } | ConvertTo-Json -Compress
  Write-Verbose "Sending JSON..."
  Write-Verbose "$sendJSON"

  # Generate REST URI
  $uri = "$($Protocol)://$ServerAddress/webservices/rest.php?version=1.1&json_data=$sendJSON"
  Write-Verbose "REST URI: $uri"

  # Execute command and store returned JSON
  $returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'
  Write-Verbose "Server returned: 
  $returnedJSON"

  # Break out of function with warning message if no results returned
  if (!$returnedJSON.objects) {
    Write-Warning "Search has returned 0 results."
    break
  }

  # Convert returned JSON into easily consumable PowerShell objects by
  # Parsing server response to build a non-nested object
  $objData = @()
  foreach ($name in ($returnedJSON.objects | Get-Member -MemberType Properties).Name){
    $objData += $returnedJSON.objects.$name.fields
  }

  # Run where block for specific query
  $objData
}

function Get-iTopVirtualMachine {
<#
.SYNOPSIS
  Query iTop server for all available VMs and select a VM if one is supplied.
.DESCRIPTION
  Sends a core/get operation to the iTop REST api. If no VM guest name is supplied, will return all VMs. If one is supplied will apply: 
  
  '"SELECT Brand WHERE name = " + "'" +$iTopVM + "'"'
.PARAMETER ServerAddress
  FQDN of the iTop server you are running against.
.PARAMETER Protocol
  Whether you are connecting to the iTop instance over HTTP or HTTPS. Default is HTTPS.
.PARAMETER Credential
  The credentials that you are going to authenticate against the iTop REST API.
.PARAMETER OSFamily
  The OS flavor you want to filter on. Only Linux and Windows have been tested.
.PARAMETER OQLFilter
  Custom additional OQL query commands to be passed. By default, if this isn't used, the query is just "SELECT VirtualMachine" or "SELECT VirtualMachine WHERE osfamily_name = '$OS'"
  If the OQLFilter is used, it is the same as doing "SELECT VirtualMachine AND $OQLFilter" or "SELECT VirtualMachine WHERE osfamily_name = '$OS' AND $OQLFilter"
.NOTES
.EXAMPLE
  Get-iTopVirtualMachine -ServerAddress "itop.foo.com" -Protocol https -Credential $Credentials -OSFamily Linux
  Search for all Linux VMs.
.EXAMPLE
  C:\PS>$RegExp = '^server00.*$'
  C:\PS>Get-iTopVirtualMachine -ServerAddress "itop.foo.com" -Protocol https -Credential $Credentials -OSFamily Windows -OQLFilter "name REGEXP '$REGEXP'"
  
  Search for Windows VMs starting with server00 in their 'name' property.
.LINK
  https://github.com/jenquist/PSiTopRestMod
#> 
[CmdletBinding()]
param(    
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$OSFamily,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$OQLFilter
)
  # Creating header with credentials being used for authentication
  [string]$username = $Credential.UserName
  [string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
  $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  $headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

  $key = "SELECT VirtualMachine"
  if ($OSFamily) {
    # Custom OSFamily, if used
    $key = "$key WHERE osfamily_name = '$OS'"
  }
  if ($OQLFilter) {
    # Custom Filter is based off of iTop OQL query language
    $key = "$key AND $OQLFilter"
  }

  # Creating in-line JSON to be sent within URI
  $sendJSON = @{
               operation = 'core/get'
               class = 'VirtualMachine'
               key = ("$key")
               output_fields = '*'
               } | ConvertTo-Json -Compress
  Write-Verbose "Sending JSON..."
  Write-Verbose "$sendJSON"

  # Generate REST URI
  $uri = "$($Protocol)://$ServerAddress/webservices/rest.php?version=1.1&json_data=$sendJSON"
  Write-Verbose "REST URI: $uri"

  # Execute command and store returned JSON
  $returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'
  Write-Verbose "Server returned: 
  $returnedJSON"

  # Break out of function with warning message if no results returned
  if (!$returnedJSON.objects) {
    Write-Warning "Search has returned 0 results."
    break
  }

  # Convert returned JSON into easily consumable PowerShell objects by
  # Parsing server response to build a non-nested object
  $objData = @()
  foreach ($name in ($returnedJSON.objects | Get-Member -MemberType Properties).Name){
    $objData += $returnedJSON.objects.$name.fields
  }

  # Run where block for specific query
  $objData
}

function Get-iTopBrand {
<#
.SYNOPSIS
  Query iTop server for all available Brands and select a brand if one is supplied.
.DESCRIPTION
  Sends a core/get operation to the iTop REST api. If no brand is supplied, will return all brands. If one is supplied will apply: 
  
  '"SELECT Brand WHERE name = " + "'" +$iTopBrand + "'"'
.NOTES
.EXAMPLE
  Get-iTopBrand -ServerAddress "itop.foo.com" -Protocol "https" -Credential (get-credential) -itop_Brand "Cisco"
.LINK
  https://github.com/jenquist/PSiTopRestMod
#> 
[CmdletBinding()]
param(    
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$iTopBrand = "*"
)
  # Creating header with credentials being used for authentication
  [string]$username = $Credential.UserName
  [string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
  $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  $headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

  # Creating in-line JSON to be sent within URI
  $sendJSON = @{
               operation = 'core/get'
               class = 'Brand'
               key = ("SELECT Brand")
               output_fields= 'name,finalclass'
               } | ConvertTo-Json -Compress
  Write-Verbose "Sending JSON..."
  Write-Verbose "$sendJSON"

  # Generate REST URI
  $uri = "$($Protocol)://$ServerAddress/webservices/rest.php?version=1.1&json_data=$sendJSON"
  Write-Verbose "REST URI: $uri"

  # Execute command and store returned JSON
  $returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'
  Write-Verbose "Server returned: 
  $returnedJSON"

  # Break out of function with warning message if no results returned
  if (!$returnedJSON.objects) {
    Write-Warning "Search has returned 0 results."
    break
  }

  # Convert returned JSON into easily consumable PowerShell objects by
  # Parsing server response to build a non-nested object
  $objData = @()
  foreach ($name in ($returnedJSON.objects | Get-Member -MemberType Properties).Name){
    $objData += [PSCustomObject]@{'name'=$returnedJSON.objects.$name.fields.name
                                  'class'=$returnedJSON.objects.$name.fields.finalclass
                                  'key'=$returnedJSON.objects.$name.key}
  }

  # Run where block for specific query
  $objData | where {$_.name -like "*$iTopBrand*"}
}

function Get-iTopLocation {
<#
.SYNOPSIS
  Query iTop server for all available Locations and select a Location if one is supplied.
.DESCRIPTION
  Sends a core/get operation to the iTop REST api. If no Location is supplied, will return all Locations. If one is supplied will apply: 
  
  '| where {$_.name -like "*SuppliedLocation*"}'
.NOTES
.EXAMPLE
  Get-iTopLocation -ServerAddress "itop.foo.com" -Protocol "https" -Credential (get-credential) -itop_Location "DataCenter1"
.LINK
  https://github.com/jenquist/PSiTopRestMod
#>  
[CmdletBinding()]
param( 
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_Name = "*"
)         
  # Creating header with credentials being used for authentication
  [string]$username = $Credential.UserName
  [string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
  $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  $headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

  # Creating in-line JSON to be sent within URI
  $sendJSON = @{
               operation = 'core/get';
               class = 'Location';
               key = ("SELECT Location");
               output_fields= '*';
               } | ConvertTo-Json -Compress
  Write-Verbose "Sending JSON..."
  Write-Verbose "$sendJSON"

  # Generate REST URI
  $uri = "$($Protocol)://$ServerAddress/webservices/rest.php?version=1.1&json_data=$sendJSON"
  Write-Verbose "REST URI: $uri"

  # Execute command and store returned JSON
  $returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'
  Write-Verbose "Server returned: 
  $returnedJSON"

  # Break out of function with warning message if no results returned
  if (!$returnedJSON.objects) {
    Write-Warning "Search has returned 0 results."
    break
  }

  # Convert returned JSON into easily consumable PowerShell objects by
  # Parsing server response to build a non-nested object
  $objData = @()
  foreach ($name in (($returnedJSON.objects | Get-Member -MemberType NoteProperty).Name)){
    $objData += [PSCustomObject]@{'name'=$returnedJSON.objects.$name.fields.name
                                  'status'=$returnedJSON.objects.$name.fields.status
                                  'org_id'=$returnedJSON.objects.$name.fields.org_id
                                  'org_name'=$returnedJSON.objects.$name.fields.org_name
                                  'address'=$returnedJSON.objects.$name.fields.address
                                  'postal_code'=$returnedJSON.objects.$name.fields.postal_code
                                  'city'=$returnedJSON.objects.$name.fields.city
                                  'country'=$returnedJSON.objects.$name.fields.country
                                  'physicaldevice_list'=$returnedJSON.objects.$name.fields.physicaldevice_list
                                  'person_list'=$returnedJSON.objects.$name.fields.person_list
                                  'friendlyname'=$returnedJSON.objects.$name.fields.friendlyname
                                  'org_id_friendlyname'=$returnedJSON.objects.$name.fields.org_id_friendlyname
                                  'key'=$returnedJSON.objects.$name.key}

  }

  # Run where block for specific query
  $objData | where {$_.name -like "*$itop_Name*"}
}

function Get-iTopOrganization {  
<#
.SYNOPSIS
  Query iTop server for all available Organizations and select a Organization if one is supplied.
.DESCRIPTION
  Sends a core/get operation to the iTop REST api. If no Organization is supplied, will return all Organizations. If one is supplied, will apply:
  
  '| where {$_.name -like "*SuppliedOrganization*"}'
.NOTES
.EXAMPLE
  Get-iTopOrganization -ServerAddress "itop.foo.com" -Protocol "https" -Credential (get-credential) -itop_name "My Company/Department"
.LINK
  https://github.com/jenquist/PSiTopRestMod
#>  
[CmdletBinding()]
param( 
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_org = "*"
)
  # Creating header with credentials being used for authentication
  [string]$username = $Credential.UserName
  [string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
  $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  $headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

  # Creating in-line JSON to be sent within URI
  $sendJSON = @{
               operation = 'core/get';
               class = 'Organization';
               key = ("SELECT Organization");;
               output_fields= '*';
               } | ConvertTo-Json -Compress
  Write-Verbose "Sending JSON..."
  Write-Verbose "$sendJSON"

  # Generate REST URI
  $uri = "$($Protocol)://$ServerAddress/webservices/rest.php?version=1.1&json_data=$sendJSON"
  Write-Verbose "REST URI: $uri"

  # Execute command and store returned JSON
  $returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'
  Write-Verbose "Server returned: 
  $returnedJSON"

  # Break out of function with warning message if no results returned
  if (!$returnedJSON.objects) {
    Write-Warning "Search has returned 0 results."
    break
  }

  # Convert returned JSON into easily consumable PowerShell objects by
  # Parsing server response to build a non-nested object
  $objData = @()
  foreach ($name in (($returnedJSON.objects | Get-Member -MemberType NoteProperty).Name)){
    $objData += [PSCustomObject]@{'name'=$returnedJSON.objects.$name.fields.name
                                  'code'=$returnedJSON.objects.$name.fields.code
                                  'parent_id'=$returnedJSON.objects.$name.fields.parent_id
                                  'status'=$returnedJSON.objects.$name.fields.status
                                  'parent_name'=$returnedJSON.objects.$name.fields.parent_name
                                  'deliverymodel_id'=$returnedJSON.objects.$name.fields.deliverymodel_id
                                  'deliverymodel_name'=$returnedJSON.objects.$name.fields.deliverymodel_name
                                  'friendlyname'=$returnedJSON.objects.$name.fields.friendlyname
                                  'parent_id_friendlyname'=$returnedJSON.objects.$name.fields.parent_id_friendlyname
                                  'deliverymodel_id_friendlyname'=$returnedJSON.objects.$name.fields.deliverymodel_id_friendlyname
                                  'key'=$returnedJSON.objects.$name.key}
  }
  
  # Run where block for specific query
  $objData | where {$_.name -like "*$itop_org*"}
}

function Get-iTopModel { 
<#
.SYNOPSIS
  Query iTop server for all available Models and select a Model if one is supplied.
.DESCRIPTION
  Sends a core/get operation to the iTop REST api. If no Model is supplied, will return all Models. If one is supplied, will apply: 
  
  '| where {$_.name -like "*SuppliedModel*"}'
.NOTES
.EXAMPLE
  Get-iTopModel -ServerAddress "itop.foo.com" -Protocol "https" -Credential (get-credential) -itop_name "WS-C2960C"
.LINK
  https://github.com/jenquist/PSiTopRestMod
#>  
[CmdletBinding()]
param(
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_name = "*"
)
  # Creating header with credentials being used for authentication
  [string]$username = $Credential.UserName
  [string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
  $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  $headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

  # Creating in-line JSON to be sent within URI
  $sendJSON = @{
               operation = 'core/get';
               class = 'Model';
               key = 'SELECT Model';
               output_fields= 'name,brand_id,brand_name,type';
               } | ConvertTo-Json -Compress
  Write-Verbose "Sending JSON..."
  Write-Verbose "$sendJSON"

  # Generate REST URI
  $uri = "$($Protocol)://$ServerAddress/webservices/rest.php?version=1.1&json_data=$sendJSON"
  Write-Verbose "REST URI: $uri"

  # Execute command and store returned JSON
  $returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'
  Write-Verbose "Server returned: 
  $returnedJSON"

  # Break out of function with warning message if no results returned
  if (!$returnedJSON.objects) {
    Write-Warning "Search has returned 0 results."
    break
  }

  # Convert returned JSON into easily consumable PowerShell objects by
  # Parsing server response to build a non-nested object
  $objData = @()
  foreach ($name in (($returnedJSON.objects | Get-Member -MemberType NoteProperty).Name)){
    $objData += [PSCustomObject]@{'name'=$returnedJSON.objects.$name.fields.name
                                  'brand_id'=$returnedJSON.objects.$name.fields.brand_id
                                  'brand_name'=$returnedJSON.objects.$name.fields.brand_name
                                  'type'=$returnedJSON.objects.$name.fields.type                               
                                  'key'=$returnedJSON.objects.$name.key}
  }

  # Run where block for specific query
  $objData | where {$_.name -like "*$itop_name*"}
}

function Get-iTopIOSVersion {  
<#
.SYNOPSIS
  Query iTop server for all available IOSVersion and select a IOSVersions if one is supplied.
.DESCRIPTION
  Sends a core/get operation to the iTop REST api. If no IOSVersion is supplied, will return all IOSVersions. If one is supplied, will apply: 
  
  '| where {$_.name -like "*SuppliedIOSVersion*"}'
.NOTES
.EXAMPLE
  Get-iTopIOSVersion -ServerAddress "itop.foo.com" -Protocol "https" -Credential (get-credential) -itop_name "Version 15.0(2)EX5"
.LINK
  https://github.com/jenquist/PSiTopRestMod
#>  
[CmdletBinding()]
param(
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_name = "*"
)
  # Creating header with credentials being used for authentication
  [string]$username = $Credential.UserName
  [string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
  $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  $headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

  # Creating in-line JSON to be sent within URI
  $sendJSON = @{
               operation = 'core/get';
               class = 'IOSVersion';
               key = 'SELECT IOSVersion';
               output_fields= '*';
               } | ConvertTo-Json -Compress
  Write-Verbose "Sending JSON..."
  Write-Verbose "$sendJSON"

  # Generate REST URI
  $uri = "$($Protocol)://$ServerAddress/webservices/rest.php?version=1.1&json_data=$sendJSON"
  Write-Verbose "REST URI: $uri"

  # Execute command and store returned JSON
  $returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'
  Write-Verbose "Server returned: 
  $returnedJSON"

  # Break out of function with warning message if no results returned
  if (!$returnedJSON.objects) {
    Write-Warning "Search has returned 0 results."
    break
  }

  # Convert returned JSON into easily consumable PowerShell objects by
  # Parsing server response to build a non-nested object
  $objData = @()
  foreach ($name in (($returnedJSON.objects | Get-Member -MemberType NoteProperty).Name)){  
    $objData += [PSCustomObject]@{'name'=$returnedJSON.objects.$name.fields.name
                                  'brand_id'=$returnedJSON.objects.$name.fields.brand_id
                                  'brand_name'=$returnedJSON.objects.$name.fields.brand_name
                                  'finalclass'=$returnedJSON.objects.$name.fields.finalclass                               
                                  'key'=$returnedJSON.objects.$name.key}
  }

  # Run where block for specific query
  $objData | where {$_.name -like "*$itop_name*"}
}

function Get-iTopOSFamily {
<#
.SYNOPSIS
  Query iTop server for all OSFamilies and select a OSFamily if one is supplied.
.DESCRIPTION
  Sends a core/get operation to the iTop REST api. If no OSFamily name is supplied, will return all OSFamilies. If one is supplied will apply: 
  
  "SELECT OSFamily WHERE name = " + "'" +$itop_name
.PARAMETER ServerAddress
  FQDN of the iTop server you are running against.
.PARAMETER Protocol
  Whether you are connecting to the iTop instance over HTTP or HTTPS. Default is HTTPS.
.PARAMETER Credential
  The credentials that you are going to authenticate against the iTop REST API.
.PARAMETER itop_name
  The iTop name of the OSFamily.
.NOTES
.EXAMPLE
  C:\PS>function Get-iTopOSFamily -ServerAddress "itop.foo.com" -Protocol https -Credential $Credentials -itop_name "Linux"
  
.LINK
  https://github.com/jenquist/PSiTopRestMod
#> 
[CmdletBinding()]
param(    
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$iTop_name
)
  # Creating header with credentials being used for authentication
  [string]$username = $Credential.UserName
  [string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
  $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  $headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

  $key = "SELECT OSFamily"
  if ($iTop_name) {
    $key = "$key WHERE location_name = '$iTop_name'"
  }


  # Creating in-line JSON to be sent within URI
  $sendJSON = @{
               operation = 'core/get'
               class = 'OSFamily'
               key = ("$key")
               output_fields = 'name'
               } | ConvertTo-Json -Compress
  Write-Verbose "Sending JSON..."
  Write-Verbose "$sendJSON"

  # Generate REST URI
  $uri = "$($Protocol)://$ServerAddress/webservices/rest.php?version=1.1&json_data=$sendJSON"
  Write-Verbose "REST URI: $uri"

  # Execute command and store returned JSON
  $returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'
  Write-Verbose "Server returned: 
  $returnedJSON"

  # Break out of function with warning message if no results returned
  if (!$returnedJSON.objects) {
    Write-Warning "Search has returned 0 results."
    break
  }

  # Convert returned JSON into easily consumable PowerShell objects by
  # Parsing server response to build a non-nested object
  $objData = @()
  foreach ($name in ($returnedJSON.objects | Get-Member -MemberType Properties).Name){
    $objData += [PSCustomObject]@{'name'=$returnedJSON.objects.$name.fields.name
                                  'Key'=$returnedJSON.objects.$name.key}
  }

  # Run where block for specific query
  $objData
}

function Get-iTopOSVersion {
<#
.SYNOPSIS
  Query iTop server for all OSVersions and select a OSVersion if one is supplied.
.DESCRIPTION
  Sends a core/get operation to the iTop REST api. If no OSVersion name is supplied, will return all OSVersions. If one is supplied will apply: 
  
  "SELECT OSVersion WHERE name = " + "'" +$itop_name
.PARAMETER ServerAddress
  FQDN of the iTop server you are running against.
.PARAMETER Protocol
  Whether you are connecting to the iTop instance over HTTP or HTTPS. Default is HTTPS.
.PARAMETER Credential
  The credentials that you are going to authenticate against the iTop REST API.
.PARAMETER itop_name
  The iTop name of the OSVersion.
.NOTES
.EXAMPLE
  C:\PS>function Get-iTopOSVersion -ServerAddress "itop.foo.com" -Protocol https -Credential $Credentials -itop_name "Linux"
  
.LINK
  https://github.com/jenquist/PSiTopRestMod
#> 
[CmdletBinding()]
param(    
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$iTop_name
)
  # Creating header with credentials being used for authentication
  [string]$username = $Credential.UserName
  [string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
  $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  $headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

  $key = "SELECT OSVersion"
  if ($iTop_name) {
    $key = "$key WHERE location_name = '$iTop_name'"
  }


  # Creating in-line JSON to be sent within URI
  $sendJSON = @{
               operation = 'core/get'
               class = 'OSVersion'
               key = ("$key")
               output_fields = '*'
               } | ConvertTo-Json -Compress
  Write-Verbose "Sending JSON..."
  Write-Verbose "$sendJSON"

  # Generate REST URI
  $uri = "$($Protocol)://$ServerAddress/webservices/rest.php?version=1.1&json_data=$sendJSON"
  Write-Verbose "REST URI: $uri"

  # Execute command and store returned JSON
  $returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'
  Write-Verbose "Server returned: 
  $returnedJSON"

  # Break out of function with warning message if no results returned
  if (!$returnedJSON.objects) {
    Write-Warning "Search has returned 0 results."
    break
  }

  # Convert returned JSON into easily consumable PowerShell objects by
  # Parsing server response to build a non-nested object
  $objData = @()
  foreach ($name in ($returnedJSON.objects | Get-Member -MemberType Properties).Name){
    $objData += [PSCustomObject]@{'name'=$returnedJSON.objects.$name.fields.name
                                  'osfamily_name'=$returnedJSON.objects.$name.fields.osfamily_name
                                  'osfamily_id'=$returnedJSON.objects.$name.fields.osfamily_id
                                  'Key'=$returnedJSON.objects.$name.key}
  }

  # Run where block for specific query
  $objData
}

function Get-iTopRack {
<#
.SYNOPSIS
  Query iTop server for all Racks and select a Rack if one is supplied.
.DESCRIPTION
  Sends a core/get operation to the iTop REST api. If no Rack name or Location is supplied, will return all Racks. If one is supplied will apply: 
  
  '"SELECT Rack WHERE name = " + "'" +$itop_name + "'" AND location_name = " + "'" +$Location + "'"
.PARAMETER ServerAddress
  FQDN of the iTop server you are running against.
.PARAMETER Protocol
  Whether you are connecting to the iTop instance over HTTP or HTTPS. Default is HTTPS.
.PARAMETER Credential
  The credentials that you are going to authenticate against the iTop REST API.
.PARAMETER Location
  The iTop Location where the rack resides.
.PARAMETER itop_name
  The iTop name of the Rack.
.NOTES
.EXAMPLE
  Get-iTopVirtualMachine -ServerAddress "itop.foo.com" -Protocol https -Credential $Credentials -OSFamily Linux
  Search for all Linux VMs.
.EXAMPLE
  C:\PS>function Get-iTopRack -ServerAddress "itop.foo.com" -Protocol https -Credential $Credentials -name TestRack -Location "DataCenter1"
  
.LINK
  https://github.com/jenquist/PSiTopRestMod
#> 
[CmdletBinding()]
param(    
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$Location,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$iTop_name
)
  # Creating header with credentials being used for authentication
  [string]$username = $Credential.UserName
  [string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
  $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  $headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

  $key = "SELECT Rack"
  if (($Location) -and ($iTop_name)) {
    # Custom OSFamily, if used
    $key = "$key WHERE location_name = '$Location' AND name = '$iTop_name'"
  } elseif ($Location) {
    $key = "$key WHERE location_name = '$Location'"
  } elseif ($iTop_name) {
    $key = "$key WHERE name = '$iTop_name'"
  }


  # Creating in-line JSON to be sent within URI
  $sendJSON = @{
               operation = 'core/get'
               class = 'Rack'
               key = ("$key")
               output_fields = '*'
               } | ConvertTo-Json -Compress
  Write-Verbose "Sending JSON..."
  Write-Verbose "$sendJSON"

  # Generate REST URI
  $uri = "$($Protocol)://$ServerAddress/webservices/rest.php?version=1.1&json_data=$sendJSON"
  Write-Verbose "REST URI: $uri"

  # Execute command and store returned JSON
  $returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'
  Write-Verbose "Server returned: 
  $returnedJSON"

  # Break out of function with warning message if no results returned
  if (!$returnedJSON.objects) {
    Write-Warning "Search has returned 0 results."
    break
  }

  # Convert returned JSON into easily consumable PowerShell objects by
  # Parsing server response to build a non-nested object
  $objData = @()
  foreach ($name in ($returnedJSON.objects | Get-Member -MemberType Properties).Name){
    $objData += [PSCustomObject]@{'name'=$returnedJSON.objects.$name.fields.name
                                  'organization_name'=$returnedJSON.objects.$name.fields.organization_name
                                  'org_id'=$returnedJSON.objects.$name.fields.org_id
                                  'location_name'=$returnedJSON.objects.$name.fields.location_name
                                  'location_id'=$returnedJSON.objects.$name.fields.location_id
                                  'brand_name'=$returnedJSON.objects.$name.fields.brand_name
                                  'brand_id'=$returnedJSON.objects.$name.fields.brand_id
                                  'model_name'=$returnedJSON.objects.$name.fields.model_name
                                  'RU'=$returnedJSON.objects.$name.fields.nb_u
                                  'Key'=$returnedJSON.objects.$name.key}
  }

  # Run where block for specific query
  $objData
}

function Get-iTopServer {
<#
.SYNOPSIS
  Query iTop server for all Servers and select a server if one is supplied.
.DESCRIPTION
  Sends a core/get operation to the iTop REST api. If no server name or Location is supplied, will return all servers. If one is supplied will apply: 
  
  '"SELECT Server WHERE name = " + "'" +$itop_name + "'" AND location_name = " + "'" +$Location + "'"
.PARAMETER ServerAddress
  FQDN of the iTop server you are running against.
.PARAMETER Protocol
  Whether you are connecting to the iTop instance over HTTP or HTTPS. Default is HTTPS.
.PARAMETER Credential
  The credentials that you are going to authenticate against the iTop REST API.
.PARAMETER Location
  The iTop Location where the server resides.
.PARAMETER itop_name
  The iTop name of the server.
.NOTES
.EXAMPLE
  C:\PS>function Get-iTopserver -ServerAddress "itop.foo.com" -Protocol https -Credential $Credentials -name TestServer -Location "DataCenter1"
  
.LINK
  https://github.com/jenquist/PSiTopRestMod
#> 
[CmdletBinding()]
param(    
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$Location,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$iTop_name
)
  # Creating header with credentials being used for authentication
  [string]$username = $Credential.UserName
  [string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
  $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  $headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

  $key = "SELECT Server"
  if (($Location) -and ($iTop_name)) {
    # Custom OSFamily, if used
    $key = "$key WHERE location_name = '$Location' AND name = '$iTop_name'"
  } elseif ($Location) {
    $key = "$key WHERE location_name = '$Location'"
  } elseif ($iTop_name) {
    $key = "$key WHERE name = '$iTop_name'"
  }


  # Creating in-line JSON to be sent within URI
  $sendJSON = @{
               operation = 'core/get'
               class = 'Server'
               key = ("$key")
               output_fields = '*'
               } | ConvertTo-Json -Compress
  Write-Verbose "Sending JSON..."
  Write-Verbose "$sendJSON"

  # Generate REST URI
  $uri = "$($Protocol)://$ServerAddress/webservices/rest.php?version=1.1&json_data=$sendJSON"
  Write-Verbose "REST URI: $uri"

  # Execute command and store returned JSON
  $returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'
  Write-Verbose "Server returned: 
  $returnedJSON"

  # Break out of function with warning message if no results returned
  if (!$returnedJSON.objects) {
    Write-Warning "Search has returned 0 results."
    break
  }

  # Convert returned JSON into easily consumable PowerShell objects by
  # Parsing server response to build a non-nested object
  $objData = @()
  foreach ($name in ($returnedJSON.objects | Get-Member -MemberType Properties).Name){
    $objData += [PSCustomObject]@{'name'=$returnedJSON.objects.$name.fields.name
                                  'description'=$returnedJSON.objects.$name.fields.description
                                  'organization_name'=$returnedJSON.objects.$name.fields.organization_name
                                  'business_criticity'=$returnedJSON.objects.$name.fields.business_criticity
                                  'move2production'=$returnedJSON.objects.$name.fields.move2production
                                  'serialnumber'=$returnedJSON.objects.$name.fields.serialnumber
                                  'asset_number'=$returnedJSON.objects.$name.fields.asset_number
                                  'status'=$returnedJSON.objects.$name.fields.status
                                  'end_of_warranty'=$returnedJSON.objects.$name.fields.end_of_warranty
                                  'org_id'=$returnedJSON.objects.$name.fields.org_id
                                  'location_name'=$returnedJSON.objects.$name.fields.location_name
                                  'location_id'=$returnedJSON.objects.$name.fields.location_id
                                  'brand_name'=$returnedJSON.objects.$name.fields.brand_name
                                  'brand_id'=$returnedJSON.objects.$name.fields.brand_id
                                  'model_name'=$returnedJSON.objects.$name.fields.model_name
                                  'model_id'=$returnedJSON.objects.$name.fields.model_id
                                  'Key'=$returnedJSON.objects.$name.key}
  }

  # Run where block for specific query
  $objData
}

function New-iTopNetModel {
    <#
.SYNOPSIS
  Post core/create to iTop server for new NetModel and return NetModel name and key.
.DESCRIPTION
  Sends a core/create operation to the iTop REST api. Currently does not check for duplicate name, just creates another object with the info you supply.
  Will lookup brand_ID if brand_name is supplied but brand_id is not.
.NOTES
  
.EXAMPLE 
  New-iTopNetModel -Credential $Credential -ServerAddress itop.isd625.sppsmn.int -Protocol https -itop_friendlyname "AIR-CAP3702I-A-K9" -itop_brand_name "Cisco" -itop_brand_id "2" -itop_type NetworkDevice
.LINK
  https://github.com/jenquist/PSiTopRestMod
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_name,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_brand_name,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_type,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_brand_id,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_friendlyname
) 
  # Creating header with credentials being used for authentication
  [string]$username = $Credential.UserName
  [string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
  $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  $headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))


#cleanup for password
$password=$null

#Get Vendor Key if null

if ([string]::IsNullOrEmpty($itop_brand_id)){

$getBrand = @{
             operation = 'core/get';
             class = 'Brand';
             key = ("SELECT Brand WHERE name = " + "'" +$itop_brand_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$brandURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getBrand"
$brand = Invoke-RestMethod -Uri $brandURI -Headers $headers -Method Post -ContentType 'application/json'

$objBrand = @()

foreach ($name in (($brand.objects | Get-Member -MemberType NoteProperty).Name))
    {
    

    $type = New-Object System.Object 
    $type | Add-Member -type NoteProperty -name Name -Value ($brand.objects.$name.fields.name)
    $type | Add-Member -type NoteProperty -name DeviceTypeKey -Value ($brand.objects.$name.key)
    
    $objBrand += $type

    }


$itop_brand_id = $objBrand[0].Key


}

$CreateModel = @{
   operation='core/create';
   comment='PowershellAPI';
   class= 'Model';
   fields = @{} 
} 

$variables = @((get-help Import-iTopNetModel -Parameter *).name)


#add each parameter function to fields HT that starts with iTop and is not null or empty
foreach ($var in $variables) {
    
    if (($var -like "itop_*") -and (![string]::IsNullOrEmpty($itop_brand_id))){
        
        $name = ($var).Replace('itop_','')
        
        $value = Get-Variable $var -ValueOnly


        $CreateModel.fields.Add("$name","$value")

    }


}


#Generate JSON object
$CreateModel = $CreateModel | ConvertTo-Json -Compress
#$CreateModel

#generate ReST URI
$uri = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$CreateModel"
#$uri




#execute command ans store returned JSON
$returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'
#$returnedJSON


$objData = @()

foreach ($name in (($returnedJSON.objects | Get-Member -MemberType NoteProperty).Name)){
    

    $type = New-Object System.Object 
    $type | Add-Member -type NoteProperty -name name -Value ($returnedJSON.objects.$name.fields.name)
    $type | Add-Member -type NoteProperty -name key -Value ($returnedJSON.objects.$name.key)
    
   $objData += $type

}

return $objData 


#cleanup for headers and base64 var
$base64AuthInfo = $null
$headers = $null
}

function New-iTopIOSversion {
<#
.SYNOPSIS
  Post core/create to iTop server for new IOSVersion and return IOSVersion name and key.
.DESCRIPTION
  Sends a core/create operation to the iTop REST api. Currently does not check for duplicate name, just creates another object with the info you supply.
  Will lookup brand_ID if brand_name is supplied but brand_id is not.
.NOTES
  
.EXAMPLE 
  New-iTopIOSversion -Credential $Credential -ServerAddress itop.isd625.sppsmn.int -Protocol https -itop_IOSname "Version 12.2(25)SEE3" -itop_brand_name "Cisco" -itop_brand_id "2"
.LINK
  https://github.com/jenquist/PSiTopRestMod
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_IOSname,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_brand_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_brand_id             
) 
  # Creating header with credentials being used for authentication
  [string]$username = $Credential.UserName
  [string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
  $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  $headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

#cleanup for password
$password=$null

#Get Vendor Key if null


$getIOSVers = @{
             operation = 'core/get';
             class = 'IOSVersion';
             key = "SELECT IOSVersion";
             output_fields= 'friendlyname';
             } | ConvertTo-Json -Compress


$iosURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getIOSVers"
$ios = Invoke-RestMethod -Uri $iosURI -Headers $headers -Method Post -ContentType 'application/json'

$objIosVer = @{}

foreach ($name in (($ios.objects | Get-Member -MemberType NoteProperty).Name))
    {
    

    $fname = ($ios.objects.$name.fields.friendlyname)
    $key = ($ios.objects.$name.key)
    
    $objIosVer["$fname"] = $key

    }

$modifiedName = "Cisco " + $itop_IOSname
if(![string]::IsNullOrEmpty($objIosVer["$modifiedName"])){

    "IOS Version exists in iTop"
    
} else {

    if ([string]::IsNullOrEmpty($itop_brand_id)){

    $getBrand = @{
                 operation = 'core/get';
                 class = 'Brand';
                 key = ("SELECT Brand WHERE name = " + "'" +$itop_brand_name + "'");
                 output_fields= 'name';
                 } | ConvertTo-Json -Compress


    $brandURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getBrand"
    $brand = Invoke-RestMethod -Uri $brandURI -Headers $headers -Method Post -ContentType 'application/json'

    $objBrand = @()

    foreach ($name in (($brand.objects | Get-Member -MemberType NoteProperty).Name))
        {
        $objBrand += [PSCustomObject]@{'name'=$brand.objects.$name.fields.name                                                            
                                      'key'=$brand.objects.$name.key}    
        

        }


    $itop_brand_id = $objBrand[0].key


    }


    $CreateIOS = @{
       operation='core/create';
       comment='PowershellAPI';
       class= 'IOSVersion';
       fields = @{
                 name = "$itop_IOSname";
                 friendlyname = "$itop_IOSname";
                 brand_name = "$itop_brand_name";
                 brand_id = "$itop_brand_id"
                 } 
    }  | ConvertTo-Json -Compress




    #generate ReST URI
    $uri = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$CreateIOS"
    #$uri




    #execute command ans store returned JSON
    $returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'
    #$returnedJSON


    $objData = @()

    foreach ($name in (($returnedJSON.objects | Get-Member -MemberType NoteProperty).Name)){
        
        $objData += [PSCustomObject]@{'friendlyname'=$returnedJSON.objects.$name.fields.name                                                            
                                      'key'=$returnedJSON.objects.$name.key}    


    }

    return $objData 


    #cleanup for headers and base64 var
    $base64AuthInfo = $null
    $headers = $null
    }

}

function New-iTopLocation {
<#
.SYNOPSIS
  Post core/create to iTop server for new Location and return all Location fields.
.DESCRIPTION
  Sends a core/create operation to the iTop REST api. Currently does not check for duplicate names, just creates another object with the info you supply.
  Will lookup org_ID if organization_name is supplied but org_id is not.
.NOTES
  
.EXAMPLE 
  New-iTopLocation -Credential $Credential -ServerAddress itop.isd625.sppsmn.int -Protocol https -itop_IOSname "LocationName" -itop_status "active" -itop_org_id "2" -itop_org_name "My Company/Department" -itop_address "123 Fake St." -$itop_postal_code "123456" -itop_city "Minneapolis" -itop_country "United States"
  
.LINK
  https://github.com/jenquist/PSiTopRestMod
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_status = "active",
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_org_id,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_org_name,   
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_address,    
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_postal_code,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_city,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_country
) 
  # Creating header with credentials being used for authentication
  [string]$username = $Credential.UserName
  [string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
  $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

  $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
  $headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))


#cleanup for password
$password=$null

#find ORG ID from org name
if ([string]::IsNullOrEmpty($itop_org_id)){

$getOrg = @{
             operation = 'core/get';
             class = 'Organization';
             key = ("SELECT Organization WHERE name = " + "'" +$itop_organization_name + "'");
             output_fields= '*';
             } | ConvertTo-Json -Compress



$orgURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getOrg"

$org = Invoke-RestMethod -Uri $orgURI -Headers $headers -Method Post -ContentType 'application/json'

$objOrg = @()

foreach ($name in (($org.objects | Get-Member -MemberType NoteProperty).Name))
    {
    
    $objOrg += [PSCustomObject]@{'name'=$org.objects.$name.fields.name                            
                                  'key'=$org.objects.$name.key}


    }


$itop_org_id = $objOrg[0].Key


}

$CreateLocation = @{
   operation='core/create';
   comment='PowershellAPI';
   class= 'Location';
   output_fields= '*';
   fields = @{} 
} 

$variables = @((get-help New-iTopLocation -Parameter *).name)


#add each parameter function to fields HT that starts with iTop and is not null or empty
foreach ($var in $variables) {
    
    if (($var -like "itop_*") -and (![string]::IsNullOrEmpty((Get-Variable $var -ValueOnly)))){
        
        $name = ($var).Replace('itop_','')
        
        $value = Get-Variable $var -ValueOnly


        $CreateLocation.fields.Add("$name","$value")

    }


}

$CreateLocation = $CreateLocation | ConvertTo-Json -Compress



    #generate ReST URI
    $uri = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$CreateLocation"
    #$uri




    #execute command ans store returned JSON
    $returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'
    #$returnedJSON


    $objData = @()

    foreach ($name in (($returnedJSON.objects | Get-Member -MemberType NoteProperty).Name)){
    
    $objData += [PSCustomObject]@{'name'=$returnedJSON.objects.$name.fields.name    
                                  'status'=$returnedJSON.objects.$name.fields.status
                                  'org_id'=$returnedJSON.objects.$name.fields.org_id
                                  'org_name'=$returnedJSON.objects.$name.fields.org_name
                                  'address'=$returnedJSON.objects.$name.fields.address
                                  'postal_code'=$returnedJSON.objects.$name.fields.postal_code
                                  'city'=$returnedJSON.objects.$name.fields.city
                                  'country'=$returnedJSON.objects.$name.fields.country
                                  'physicaldevice_list'=$returnedJSON.objects.$name.fields.physicaldevice_list
                                  'person_list'=$returnedJSON.objects.$name.fields.person_list  
                                  'friendlyname'=$returnedJSON.objects.$name.fields.friendlyname
                                  'org_id_friendlyname'=$returnedJSON.objects.$name.fields.org_id_friendlyname                      
                                  'key'=$returnedJSON.objects.$name.key}

    }

    return $objData 


    #cleanup for headers and base64 var
    $base64AuthInfo = $null
    $headers = $null
   

}

function New-iTopNetDevice { 
[CmdletBinding()]
param(
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_description,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_org_id,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_organization_name,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_brand_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('high','medium','low')]
  [string]$itop_business_criticity = "low",
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_brand_id,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_friendlyname = $itop_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_serialnumber,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_location_id = 0,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_location_name,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [ValidateSet('production','implementation','obsolete','stock')]
  [string]$itop_status,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_model_id,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_model_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_asset_number,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_iosversion_id,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_iosversion_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ipaddress]$itop_managementip,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_networkdevicetype_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_networkdevicetype_id,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [datetime]$itop_purchase_date,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [datetime]$itop_move2production,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [datetime]$itop_end_of_warranty
  
) 

#Convert Valid IP Back to string
$itop_managementip = $itop_managementip.IPAddressToString

#Convert datetimes into correct string formatting for iTop dates
if ($itop_purchase_date){
  [string]$itop_purchase_date = "{0:yyyy-MM-dd}" -f $itop_purchase_date
  }
if ($itop_end_of_warranty){
  [string]$itop_end_of_warranty = "{0:yyyy-MM-dd}" -f $itop_end_of_warranty
  }
if ($itop_move2production){
  [string]$itop_move2production = "{0:yyyy-MM-dd}" -f $itop_move2production
  }

# Creating header with credentials being used for authentication
[string]$username = $Credential.UserName
[string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

#cleanup for password
$password=$null


#Sanitize Description of special Characters
if (![string]::IsNullOrEmpty($itop_description)){
    $itop_description = [System.Text.RegularExpressions.Regex]::Replace($itop_description,"[^0-9a-zA-Z_]"," ")
    }

#Sanitize Assetnumber of special Characters
if (![string]::IsNullOrEmpty($itop_description)){
    $itop_asset_number = [System.Text.RegularExpressions.Regex]::Replace($itop_asset_number,"[^0-9a-zA-Z_]"," ")
    }



#Get Brand_ID key if null

if ([string]::IsNullOrEmpty($itop_brand_id)){

$getBrand = @{
             operation = 'core/get';
             class = 'Brand';
             key = ("SELECT Brand WHERE name = " + "'" +$itop_brand_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$brandURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getBrand"
$brand = Invoke-RestMethod -Uri $brandURI -Headers $headers -Method Post -ContentType 'application/json'

$objBrand = @()

foreach ($name in (($brand.objects | Get-Member -MemberType NoteProperty).Name))
    {
    
    $objBrand += [PSCustomObject]@{'name'=$brand.objects.$name.fields.name                  
                                  'key'=$brand.objects.$name.key}

    }


$itop_brand_id = $objBrand[0].key


}

#get IOSVersion_id if IOS Version is populated.


if ((![string]::IsNullOrEmpty($itop_iosversion_name)) -and [string]::IsNullOrEmpty($itop_iosversion_id)){

$getIOS = $getIOSVers = @{
             operation = 'core/get';
             class = 'IOSVersion';
             key = ("SELECT IOSVersion WHERE name = " + "'" +$itop_iosversion_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress



$iosURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getIOS"
$ios = Invoke-RestMethod -Uri $iosURI -Headers $headers -Method Post -ContentType 'application/json'

$objIOS = @()

foreach ($name in (($ios.objects | Get-Member -MemberType NoteProperty).Name))
    {
    
    $objIOS += [PSCustomObject]@{'name'=$ios.objects.$name.fields.name                  
                                  'key'=$ios.objects.$name.key}

    }


$itop_iosversion_id = $objIOS[0].key


}

#find ORG_ID key from org_name, if null
if ([string]::IsNullOrEmpty($itop_org_id)){

$getOrg = @{
             operation = 'core/get';
             class = 'Organization';
             key = ("SELECT Organization WHERE name = " + "'" +$itop_organization_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$orgURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getOrg"
$org = Invoke-RestMethod -Uri $orgURI -Headers $headers -Method Post -ContentType 'application/json'

$objOrg = @()

foreach ($name in (($org.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objOrg += [PSCustomObject]@{'name'=$org.objects.$name.fields.name                  
                                 'key'=$org.objects.$name.key}

    }


$itop_org_id = $objOrg[0].Key


}



if ((![string]::IsNullOrEmpty($itop_model_name)) -and [string]::IsNullOrEmpty($itop_model_id)){

$getModel = @{
             operation = 'core/get';
             class = 'Model';
             key = ("SELECT Model WHERE name = " + "'" +$itop_model_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$modelURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getModel"
$model = Invoke-RestMethod -Uri $modelURI -Headers $headers -Method Post -ContentType 'application/json'

$objModel = @()

foreach ($name in (($model.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objModel += [PSCustomObject]@{'name'=$model.objects.$name.fields.name                  
                                   'key'=$model.objects.$name.key}    

    }


$itop_model_id = $objModel[0].Key


}


#Get NetworkDeviceType_id
if ([string]::IsNullOrEmpty($itop_networkdevicetype_id)){
$getNetDeviceID = @{
             operation = 'core/get';
             class = 'NetworkDeviceType';
             key = ("SELECT NetworkDeviceType WHERE name = " + "'" +$itop_networkdevicetype_name + "'" ); 
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$netDevIDURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getNetDeviceID"
$netDevID = Invoke-RestMethod -Uri $netDevIDURI -Headers $headers -Method Post -ContentType 'application/json'

$objnetDevID = @()

foreach ($name in (($netDevID.objects | Get-Member -MemberType NoteProperty).Name))
    {
    
    $objnetDevID += [PSCustomObject]@{'name'=$netDevID.objects.$name.fields.name                  
                                      'key'=$netDevID.objects.$name.key}  

    }


$itop_networkdevicetype_id = $objnetDevID[0].Key
}

$CreateDevice = @{
   operation='core/create';
   comment='PowershellAPI';
   class= 'NetworkDevice';
   output_fields= 'name';
   fields = @{} 
} 

$variables = @((get-help New-iTopNetDevice -Parameter *).name)


#add each parameter function to fields HT that starts with iTop and is not null or empty
foreach ($var in $variables) {
    
    if (($var -like "itop_*") -and (![string]::IsNullOrEmpty((Get-Variable $var -ValueOnly))))
    {
        
        $name = ($var).Replace('itop_','')
        
        $value = Get-Variable $var -ValueOnly


        $CreateDevice.fields.Add("$name","$value")

    }


}


#Generate JSON object
$CreateDevice = $CreateDevice | ConvertTo-Json -Compress



#generate ReST URI
$uri = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$CreateDevice"
#$uri




#execute command and store returned JSON
$returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'

if($returnedJSON.message -like "Error*"){
    return $returnedJSON
    } else {
    #$returnedJSON


    $objData = @()

    foreach ($name in (($returnedJSON.objects | Get-Member -MemberType NoteProperty).Name)){
        
        $objData += [PSCustomObject]@{'name'=$returnedJSON.objects.$name.fields.name                  
                                      'key'=$returnedJSON.objects.$name.key}

    }

    return $objData 


    

    }
    #cleanup for headers and base64 var
    $base64AuthInfo = $null
    $headers = $null
}

function New-iTopRack { 
[CmdletBinding()]
param(
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_description,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_org_id,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_organization_name,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_brand_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('high','medium','low')]
  [string]$itop_business_criticity = "low",
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_brand_id,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_friendlyname = $itop_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_serialnumber,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_location_id,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_location_name,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [ValidateSet('production','implementation','obsolete','stock')]
  [string]$itop_status = "production",
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_model_id,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_model_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_asset_number,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [datetime]$itop_purchase_date,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [datetime]$itop_move2production,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [datetime]$itop_end_of_warranty,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_nb_u = "42" #This is the Rack Unit Field
) 


#Convert datetimes into correct string formatting for iTop dates
if ($itop_purchase_date){
  [string]$itop_purchase_date = "{0:yyyy-MM-dd}" -f $itop_purchase_date
  }
if ($itop_end_of_warranty){
  [string]$itop_end_of_warranty = "{0:yyyy-MM-dd}" -f $itop_end_of_warranty
  }
if ($itop_move2production){
  [string]$itop_move2production = "{0:yyyy-MM-dd}" -f $itop_move2production
  }


# Creating header with credentials being used for authentication
[string]$username = $Credential.UserName
[string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

#cleanup for password
$password=$null


#Sanitize Description of special Characters
if (![string]::IsNullOrEmpty($itop_description)){
    $itop_description = [System.Text.RegularExpressions.Regex]::Replace($itop_description,"[^0-9a-zA-Z_]"," ")
    }

#Sanitize Assetnumber of special Characters
if (![string]::IsNullOrEmpty($itop_description)){
    $itop_asset_number = [System.Text.RegularExpressions.Regex]::Replace($itop_asset_number,"[^0-9a-zA-Z_]"," ")
    }



#Get Brand Key ID if null

if ([string]::IsNullOrEmpty($itop_brand_id)){

$getBrand = @{
             operation = 'core/get';
             class = 'Brand';
             key = ("SELECT Brand WHERE name = " + "'" +$itop_brand_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$brandURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getBrand"
$brand = Invoke-RestMethod -Uri $brandURI -Headers $headers -Method Post -ContentType 'application/json'

$objBrand = @()

foreach ($name in (($brand.objects | Get-Member -MemberType NoteProperty).Name))
    {
    
    $objBrand += [PSCustomObject]@{'name'=$brand.objects.$name.fields.name                  
                                  'key'=$brand.objects.$name.key}

    }


$itop_brand_id = $objBrand[0].key


}


#find ORG ID from org name
if ([string]::IsNullOrEmpty($itop_org_id)){

$getOrg = @{
             operation = 'core/get';
             class = 'Organization';
             key = ("SELECT Organization WHERE name = " + "'" +$itop_organization_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$orgURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getOrg"
$org = Invoke-RestMethod -Uri $orgURI -Headers $headers -Method Post -ContentType 'application/json'

$objOrg = @()

foreach ($name in (($org.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objOrg += [PSCustomObject]@{'name'=$org.objects.$name.fields.name                  
                                 'key'=$org.objects.$name.key}

    }


$itop_org_id = $objOrg[0].Key


}

#Get location_ID if name but not id is provided
if ((![string]::IsNullOrEmpty($itop_location_name)) -and [string]::IsNullOrEmpty($itop_location_id)){

$getLocation = @{
             operation = 'core/get';
             class = 'Location';
             key = ("SELECT Location WHERE name = " + "'" +$itop_location_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$localURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getLocation"
$local = Invoke-RestMethod -Uri $localURI -Headers $headers -Method Post -ContentType 'application/json'

$objLocal = @()

foreach ($name in (($local.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objlocal += [PSCustomObject]@{'name'=$local.objects.$name.fields.name                  
                                   'key'=$local.objects.$name.key}    

    }


$itop_location_id = $objlocal[0].Key


}


if ((![string]::IsNullOrEmpty($itop_model_name)) -and [string]::IsNullOrEmpty($itop_model_id)){

$getModel = @{
             operation = 'core/get';
             class = 'Model';
             key = ("SELECT Model WHERE name = " + "'" +$itop_model_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$modelURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getModel"
$model = Invoke-RestMethod -Uri $modelURI -Headers $headers -Method Post -ContentType 'application/json'

$objModel = @()

foreach ($name in (($model.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objModel += [PSCustomObject]@{'name'=$model.objects.$name.fields.name                  
                                   'key'=$model.objects.$name.key}    

    }


$itop_model_id = $objModel[0].Key


}


$CreateDevice = @{
   operation='core/create';
   comment='PowershellAPI';
   class= 'Rack';
   output_fields= 'name';
   fields = @{} 
} 

$variables = @((get-help New-iTopRack -Parameter *).name)


#add each parameter function to fields HT that starts with iTop and is not null or empty
foreach ($var in $variables) {
    
    if (($var -like "itop_*") -and (![string]::IsNullOrEmpty((Get-Variable $var -ValueOnly))))
    {
        
        $name = ($var).Replace('itop_','')
        
        $value = Get-Variable $var -ValueOnly


        $CreateDevice.fields.Add("$name","$value")

    }


}


#Generate JSON object
$CreateDevice = $CreateDevice | ConvertTo-Json -Compress



#generate ReST URI
$uri = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$CreateDevice"
#$uri




#execute command and store returned JSON
$returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'

if($returnedJSON.message -like "Error*"){
    return $returnedJSON
    } else {
    #$returnedJSON


    $objData = @()

    foreach ($name in (($returnedJSON.objects | Get-Member -MemberType NoteProperty).Name)){
        
        $objData += [PSCustomObject]@{'name'=$returnedJSON.objects.$name.fields.name                  
                                      'key'=$returnedJSON.objects.$name.key}

    }

    return $objData 


    

    }
    #cleanup for headers and base64 var
    $base64AuthInfo = $null
    $headers = $null
}

function New-iTopServer { 
[CmdletBinding()]
param(
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$ServerAddress,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('https','http')]
  [string]$Protocol='https',
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [PSCredential]$Credential,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_description,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ipaddress]$itop_managementip,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_org_id,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_organization_name,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_brand_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [ValidateSet('high','medium','low')]
  [string]$itop_business_criticity = "low",
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_brand_id,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_friendlyname = $itop_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_serialnumber,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_cpu,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_ram,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_nb_u,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_location_id,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_location_name,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [ValidateSet('production','implementation','obsolete','stock')]
  [string]$itop_status = "production",
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_model_id,
  [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
  [string]$itop_model_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_rack_id,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_rack_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_enclosure_id,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_enclosure_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_powerA_id,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_powerA_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_powerB_id,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_powerB_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_osfamily_id,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_osfamily_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_osversion_id,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_osversion_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_asset_number,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_oslicense_name,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [string]$itop_oslicense_id,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [datetime]$itop_purchase_date,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [datetime]$itop_move2production,
  [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
  [datetime]$itop_end_of_warranty
) 

#Convert Valid IP Back to string
$itop_managementip = $itop_managementip.IPAddressToString

#Convert datetimes into correct string formatting for iTop dates
if ($itop_purchase_date){
  [string]$itop_purchase_date = "{0:yyyy-MM-dd}" -f $itop_purchase_date
  }
if ($itop_end_of_warranty){
  [string]$itop_end_of_warranty = "{0:yyyy-MM-dd}" -f $itop_end_of_warranty
  }
if ($itop_move2production){
  [string]$itop_move2production = "{0:yyyy-MM-dd}" -f $itop_move2production
  }


# Creating header with credentials being used for authentication
[string]$username = $Credential.UserName
[string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

#cleanup for password
$password=$null


#Sanitize Description of special Characters
if (![string]::IsNullOrEmpty($itop_description)){
    $itop_description = [System.Text.RegularExpressions.Regex]::Replace($itop_description,"[^0-9a-zA-Z_]"," ")
    }

#Sanitize Assetnumber of special Characters
if (![string]::IsNullOrEmpty($itop_description)){
    $itop_asset_number = [System.Text.RegularExpressions.Regex]::Replace($itop_asset_number,"[^0-9a-zA-Z_]"," ")
    }



#Get Brand Key ID if null

if ([string]::IsNullOrEmpty($itop_brand_id)){

$getBrand = @{
             operation = 'core/get';
             class = 'Brand';
             key = ("SELECT Brand WHERE name = " + "'" +$itop_brand_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$brandURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getBrand"
$brand = Invoke-RestMethod -Uri $brandURI -Headers $headers -Method Post -ContentType 'application/json'

$objBrand = @()

foreach ($name in (($brand.objects | Get-Member -MemberType NoteProperty).Name))
    {
    
    $objBrand += [PSCustomObject]@{'name'=$brand.objects.$name.fields.name                  
                                  'key'=$brand.objects.$name.key}

    }


$itop_brand_id = $objBrand[0].key


}


#find ORG ID from org name
if ([string]::IsNullOrEmpty($itop_org_id)){

$getOrg = @{
             operation = 'core/get';
             class = 'Organization';
             key = ("SELECT Organization WHERE name = " + "'" +$itop_organization_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$orgURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getOrg"
$org = Invoke-RestMethod -Uri $orgURI -Headers $headers -Method Post -ContentType 'application/json'

$objOrg = @()

foreach ($name in (($org.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objOrg += [PSCustomObject]@{'name'=$org.objects.$name.fields.name                  
                                 'key'=$org.objects.$name.key}

    }


$itop_org_id = $objOrg[0].Key


}

#Get RAck_ID if name but not id is provided
if ((![string]::IsNullOrEmpty($itop_rack_name)) -and [string]::IsNullOrEmpty($itop_rack_id)){

$getrack = @{
             operation = 'core/get';
             class = 'Rack';
             key = ("SELECT Rack WHERE name = " + "'" +$itop_rack_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$rackURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getrack"
$rack = Invoke-RestMethod -Uri $rackURI -Headers $headers -Method Post -ContentType 'application/json'

$objRack = @()

foreach ($name in (($rack.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objRack += [PSCustomObject]@{'name'=$rack.objects.$name.fields.name                  
                                   'key'=$rack.objects.$name.key}    

    }


$itop_rack_id = $objRack[0].Key


}


#Get enclosure_ID if name but not id is provided
if ((![string]::IsNullOrEmpty($itop_enclosure_name)) -and [string]::IsNullOrEmpty($itop_enclosure_id)){

$getEnclosure = @{
             operation = 'core/get';
             class = 'Enclosure';
             key = ("SELECT Enclosure WHERE name = " + "'" +$itop_enclosure_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$EnclosureURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getEnclosure"
$enclosure = Invoke-RestMethod -Uri $EnclosureURI -Headers $headers -Method Post -ContentType 'application/json'

$objEnclosure = @()

foreach ($name in (($enclosure.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objEnclosure += [PSCustomObject]@{'name'=$enclosure.objects.$name.fields.name                  
                                   'key'=$enclosure.objects.$name.key}    

    }


$itop_enclosure_id = $objEnclosure[0].Key


}

#Get powerA_ID if name but not id is provided
if ((![string]::IsNullOrEmpty($itop_powerA_name)) -and [string]::IsNullOrEmpty($itop_powerA_id)){

$getpowerA = @{
             operation = 'core/get';
             class = 'PowerA';
             key = ("SELECT powerA WHERE name = " + "'" +$itop_powerA_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$powerAURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getpowerA"
$powerA = Invoke-RestMethod -Uri $powerAURI -Headers $headers -Method Post -ContentType 'application/json'

$objpowerA = @()

foreach ($name in (($powerA.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objpowerA += [PSCustomObject]@{'name'=$powerA.objects.$name.fields.name                  
                                   'key'=$powerA.objects.$name.key}    

    }


$itop_powerA_id = $objpowerA[0].Key


}

#Get powerB_ID if name but not id is provided
if ((![string]::IsNullOrEmpty($itop_powerB_name)) -and [string]::IsNullOrEmpty($itop_powerB_id)){

$getpowerB = @{
             operation = 'core/get';
             class = 'PowerA';
             key = ("SELECT powerA WHERE name = " + "'" +$itop_powerB_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$powerBURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getpowerB"
$powerB = Invoke-RestMethod -Uri $powerBURI -Headers $headers -Method Post -ContentType 'application/json'

$objpowerB = @()

foreach ($name in (($powerB.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objpowerB += [PSCustomObject]@{'name'=$powerB.objects.$name.fields.name                  
                                   'key'=$powerB.objects.$name.key}    

    }


$itop_powerB_id = $objpowerB[0].Key


}


#Get powerB_ID if name but not id is provided
if ((![string]::IsNullOrEmpty($itop_powerB_name)) -and [string]::IsNullOrEmpty($itop_powerB_id)){

$getpowerB = @{
             operation = 'core/get';
             class = 'PowerB';
             key = ("SELECT PowerB WHERE name = " + "'" +$itop_powerB_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$powerBURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getpowerB"
$powerB = Invoke-RestMethod -Uri $powerBURI -Headers $headers -Method Post -ContentType 'application/json'

$objpowerB = @()

foreach ($name in (($powerB.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objpowerB += [PSCustomObject]@{'name'=$powerB.objects.$name.fields.name                  
                                   'key'=$powerB.objects.$name.key}    

    }


$itop_powerB_id = $objpowerB[0].Key


}


#Get osfamily_ID if name but not id is provided
if ((![string]::IsNullOrEmpty($itop_osfamily_name)) -and [string]::IsNullOrEmpty($itop_osfamily_id)){

$getosfamily = @{
             operation = 'core/get';
             class = 'osfamily';
             key = ("SELECT osfamily WHERE name = " + "'" +$itop_osfamily_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$osfamilyURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getosfamily"
$osfamily = Invoke-RestMethod -Uri $osfamilyURI -Headers $headers -Method Post -ContentType 'application/json'

$objosfamily = @()

foreach ($name in (($osfamily.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objosfamily += [PSCustomObject]@{'name'=$osfamily.objects.$name.fields.name            
                                      'key'=$osfamily.objects.$name.key}    

    }


$itop_osfamily_id = $objosfamily[0].key


}


#Get osversion_ID if name but not id is provided
if ((![string]::IsNullOrEmpty($itop_osversion_name)) -and [string]::IsNullOrEmpty($itop_osversion_id)){

$getosversion = @{
             operation = 'core/get';
             class = 'osversion';
             key = ("SELECT osversion WHERE name = " + "'" +$itop_osversion_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$osversionURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getosversion"
$osversion = Invoke-RestMethod -Uri $osversionURI -Headers $headers -Method Post -ContentType 'application/json'

$objosfamily = @()

foreach ($name in (($osversion.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objosversion += [PSCustomObject]@{'name'=$osversion.objects.$name.fields.name                  
                                   'key'=$osversion.objects.$name.key}    

    }


$itop_osversion_id = $objosversion[0].Key


}


#Get oslicence_id if name but not id is provided
if ((![string]::IsNullOrEmpty($itop_oslicence_name)) -and [string]::IsNullOrEmpty($itop_oslicence_id)){

$getoslicence = @{
             operation = 'core/get';
             class = 'oslicence';
             key = ("SELECT oslicence WHERE name = " + "'" +$itop_oslicence_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$oslicenceURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getoslicence"
$oslicence = Invoke-RestMethod -Uri $oslicenceURI -Headers $headers -Method Post -ContentType 'application/json'

$objoslicence = @()

foreach ($name in (($oslicence.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objoslicence += [PSCustomObject]@{'name'=$oslicence.objects.$name.fields.name                  
                                   'key'=$oslicence.objects.$name.key}    

    }


$itop_oslicence_id = $objoslicence[0].Key


}

#Get location_ID if name but not id is provided
if ((![string]::IsNullOrEmpty($itop_location_name)) -and [string]::IsNullOrEmpty($itop_location_id)){

$getLocation = @{
             operation = 'core/get';
             class = 'Location';
             key = ("SELECT Location WHERE name = " + "'" +$itop_location_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$localURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getModel"
$local = Invoke-RestMethod -Uri $modelURI -Headers $headers -Method Post -ContentType 'application/json'

$objLocal = @()

foreach ($name in (($local.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objlocal += [PSCustomObject]@{'name'=$local.objects.$name.fields.name                  
                                   'key'=$local.objects.$name.key}    

    }


$itop_location_id = $objlocal[0].Key


}


if ((![string]::IsNullOrEmpty($itop_model_name)) -and [string]::IsNullOrEmpty($itop_model_id)){

$getModel = @{
             operation = 'core/get';
             class = 'Model';
             key = ("SELECT Model WHERE name = " + "'" +$itop_model_name + "'");
             output_fields= 'name';
             } | ConvertTo-Json -Compress


$modelURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getModel"
$model = Invoke-RestMethod -Uri $modelURI -Headers $headers -Method Post -ContentType 'application/json'

$objModel = @()

foreach ($name in (($model.objects | Get-Member -MemberType NoteProperty).Name))
    {
    $objModel += [PSCustomObject]@{'name'=$model.objects.$name.fields.name                  
                                   'key'=$model.objects.$name.key}    

    }


$itop_model_id = $objModel[0].Key


}


$CreateDevice = @{
   operation='core/create';
   comment='PowershellAPI';
   class= 'Server';
   output_fields= 'name';
   fields = @{} 
} 

$variables = @((get-help New-iTopServer -Parameter *).name)


#add each parameter function to fields HT that starts with iTop and is not null or empty
foreach ($var in $variables) {
    
    if (($var -like "itop_*") -and (![string]::IsNullOrEmpty((Get-Variable $var -ValueOnly))))
    {
        
        $name = ($var).Replace('itop_','')
        
        $value = Get-Variable $var -ValueOnly


        $CreateDevice.fields.Add("$name","$value")

    }


}


#Generate JSON object
$CreateDevice = $CreateDevice | ConvertTo-Json -Compress



#generate ReST URI
$uri = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$CreateDevice"
#$uri




#execute command and store returned JSON
$returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'

if($returnedJSON.message -like "Error*"){
    return $returnedJSON
    } else {
    #$returnedJSON


    $objData = @()

    foreach ($name in (($returnedJSON.objects | Get-Member -MemberType NoteProperty).Name)){
        
        $objData += [PSCustomObject]@{'name'=$returnedJSON.objects.$name.fields.name                  
                                      'key'=$returnedJSON.objects.$name.key}

    }

    return $objData 


    

    }
    #cleanup for headers and base64 var
    $base64AuthInfo = $null
    $headers = $null
}
