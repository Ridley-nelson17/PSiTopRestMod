<#
 Tool:    PSiTopRestMod.psm1
 Author:  Johann Enquist
 Email:   administrator@boomandfreeze.com
 NOTES:   Powershell module to interact with iTop Web API
#>

function Get-iTopBrand {
<#
.SYNOPSIS
  Query iTop server for all available Brands and select a brand if one is supplied.

.DESCRIPTION
  Sends a core/get operation to the iTop REST api. If no brand is supplied, will return all brands. If one is supplied will apply: 
  
  '| where {$_.name -like "*SuppliedBrand*"}'

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
  [string]$Protocol="https",
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
  $uri = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.1&json_data=$sendJSON"
  Write-Verbose "REST URI: $uri"

  # Execute command and store returned JSON
  $returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'

  # Convert returned JSON into easily consumable PowerShell objects
  $objData = @()
  foreach ($name in ($returnedJSON.objects | Get-Member -MemberType Properties).Name){
    $objData += [PSCustomObject]@{'name'=$returnedJSON.objects.$name.fields.name
                                  'finalclass'=$returnedJSON.objects.$name.fields.finalclass
                                  'key'=$returnedJSON.objects.$name.key}
  }

  # Run where block for specific query
  # Should have a proper JSON query doing the filter for us on the API end in future
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
             
             #Path to Tab Delimited user import file.
             [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
             [string]$ServerAddress,
             [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
             [string]$Protocol="https",
             [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
             [PSCredential]$Credential,
             [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
             [string]$itop_Name = "*"

             )
             



[string]$username = $Credential.UserName
[string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

$sendJSON = @{
             operation = 'core/get';
             class = 'Location';
             key = ("SELECT Location");
             output_fields= '*';
             } | ConvertTo-Json -Compress


#generate ReST URI
$uri = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$sendJSON"
#$uri




#execute command ans store returned JSON
$returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'

$objData = @()

foreach ($name in (($returnedJSON.objects | Get-Member -MemberType NoteProperty).Name)){
    

    $type = New-Object System.Object 
    $type | Add-Member -type NoteProperty -name name -Value ($returnedJSON.objects.$name.fields.name)
    $type | Add-Member -type NoteProperty -name status -Value ($returnedJSON.objects.$name.fields.status)
    $type | Add-Member -type NoteProperty -name org_id -Value ($returnedJSON.objects.$name.fields.org_id)
    $type | Add-Member -type NoteProperty -name org_name -Value ($returnedJSON.objects.$name.fields.org_name)
    $type | Add-Member -type NoteProperty -name address -Value ($returnedJSON.objects.$name.fields.address)
    $type | Add-Member -type NoteProperty -name postal_code -Value ($returnedJSON.objects.$name.fields.postal_code)
    $type | Add-Member -type NoteProperty -name city -Value ($returnedJSON.objects.$name.fields.city)
    $type | Add-Member -type NoteProperty -name country -Value ($returnedJSON.objects.$name.fields.country)
    $type | Add-Member -type NoteProperty -name physicaldevice_list -Value ($returnedJSON.objects.$name.fields.physicaldevice_list)
    $type | Add-Member -type NoteProperty -name person_list -Value ($returnedJSON.objects.$name.fields.person_list)
    $type | Add-Member -type NoteProperty -name friendlyname -Value ($returnedJSON.objects.$name.fields.friendlyname)
    $type | Add-Member -type NoteProperty -name org_id_friendlyname -Value ($returnedJSON.objects.$name.fields.org_id_friendlyname)
    $type | Add-Member -type NoteProperty -name Key -Value ($returnedJSON.objects.$name.key)
    
   $objData += $type

}

return $objData | where {$_.name -like "*$itop_name*"}

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
             
             #Path to Tab Delimited user import file.
             [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
             [string]$ServerAddress,
             [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
             [string]$Protocol="https",
             [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
             [PSCredential]$Credential,
             [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
             [string]$itop_org = "*"

             )
             



[string]$username = $Credential.UserName
[string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

$sendJSON = @{
             operation = 'core/get';
             class = 'Organization';
             key = 'SELECT Organization';
             output_fields= '*';
             } | ConvertTo-Json -Compress


#generate ReST URI
$uri = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$sendJSON"
#$uri




#execute command ans store returned JSON
$returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'

$objData = @()

foreach ($name in (($returnedJSON.objects | Get-Member -MemberType NoteProperty).Name)){
    

    $type = New-Object System.Object 
    $type | Add-Member -type NoteProperty -name name -Value ($returnedJSON.objects.$name.fields.name)
    $type | Add-Member -type NoteProperty -name code -Value ($returnedJSON.objects.$name.fields.code)
    $type | Add-Member -type NoteProperty -name parent_id -Value ($returnedJSON.objects.$name.fields.parent_id)
    $type | Add-Member -type NoteProperty -name status -Value ($returnedJSON.objects.$name.fields.status)
    $type | Add-Member -type NoteProperty -name parent_name -Value ($returnedJSON.objects.$name.fields.parent_name)
    $type | Add-Member -type NoteProperty -name deliverymodel_id -Value ($returnedJSON.objects.$name.fields.deliverymodel_id)
    $type | Add-Member -type NoteProperty -name deliverymodel_name -Value ($returnedJSON.objects.$name.fields.deliverymodel_name)
    $type | Add-Member -type NoteProperty -name friendlyname -Value ($returnedJSON.objects.$name.fields.friendlyname)
    $type | Add-Member -type NoteProperty -name parent_id_friendlyname -Value ($returnedJSON.objects.$name.fields.parent_id_friendlyname)
    $type | Add-Member -type NoteProperty -name deliverymodel_id_friendlyname -Value ($returnedJSON.objects.$name.fields.deliverymodel_id_friendlyname)
    $type | Add-Member -type NoteProperty -name Key -Value ($returnedJSON.objects.$name.key)
    
   $objData += $type

}

return $objData | where {$_.name -like "*$itop_name*"}

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
             
             #Path to Tab Delimited user import file.
             [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
             [string]$ServerAddress,
             [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
             [string]$Protocol="https",
             [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
             [PSCredential]$Credential,
             [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
             [string]$itop_name = "*"

             )
             



[string]$username = $Credential.UserName
[string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

$sendJSON = @{
             operation = 'core/get';
             class = 'Model';
             key = 'SELECT Model';
             output_fields= 'name,brand_id,brand_name,type';
             } | ConvertTo-Json -Compress


#generate ReST URI
$uri = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$sendJSON"
#$uri




#execute command ans store returned JSON
$returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'

$objData = @()

foreach ($name in (($returnedJSON.objects | Get-Member -MemberType NoteProperty).Name)){
    

    $type = New-Object System.Object 
    $type | Add-Member -type NoteProperty -name name -Value ($returnedJSON.objects.$name.fields.name)
    $type | Add-Member -type NoteProperty -name brand_id -Value ($returnedJSON.objects.$name.fields.brand_id)
    $type | Add-Member -type NoteProperty -name brand_name -Value ($returnedJSON.objects.$name.fields.brand_name)
    $type | Add-Member -type NoteProperty -name type -Value ($returnedJSON.objects.$name.fields.type)
    $type | Add-Member -type NoteProperty -name key -Value ($returnedJSON.objects.$name.key)

   $objData += $type

}

return $objData | where {$_.name -like "*$itop_name*"}

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
             
             #Path to Tab Delimited user import file.
             [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
             [string]$ServerAddress,
             [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
             [string]$Protocol="https",
             [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
             [PSCredential]$Credential,
             [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
             [string]$itop_name = "*"

             )
             



[string]$username = $Credential.UserName
[string]$password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Credential.Password))
$base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f "$username","$password")))

$headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$headers.Add("Authorization",("Basic {0}" -f $base64AuthInfo))

$sendJSON = @{
             operation = 'core/get';
             class = 'IOSVersion';
             key = 'SELECT IOSVersion';
             output_fields= '*';
             } | ConvertTo-Json -Compress


#generate ReST URI
$uri = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$sendJSON"
#$uri




#execute command ans store returned JSON
$returnedJSON = Invoke-RestMethod -Uri $uri -Headers $headers -Method Post -ContentType 'application/json'

$objData = @()

foreach ($name in (($returnedJSON.objects | Get-Member -MemberType NoteProperty).Name)){
    

    $type = New-Object System.Object 
    $type | Add-Member -type NoteProperty -name name -Value ($returnedJSON.objects.$name.fields.name)
    $type | Add-Member -type NoteProperty -name brand_id -Value ($returnedJSON.objects.$name.fields.brand_id)
    $type | Add-Member -type NoteProperty -name brand_name -Value ($returnedJSON.objects.$name.fields.brand_name)
    $type | Add-Member -type NoteProperty -name finalclass -Value ($returnedJSON.objects.$name.fields.finalclass)
    $type | Add-Member -type NoteProperty -name key -Value ($returnedJSON.objects.$name.key)

   $objData += $type

}

return $objData | where {$_.name -like "*$itop_name*"}

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
             
             #Path to Tab Delimited user import file.
             [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
             [string]$ServerAddress,
             [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
             [string]$Protocol="https",
             [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
             [PSCredential]$Credential,
             [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
             [string]$itop_IOSname,
             [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
             [string]$itop_brand_name,
             [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
             [string]$itop_brand_id             
              ) 



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
             output_fields= '*';
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
                 output_fields= '*';
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


    $itop_brand_id = $objBrand[0].DeviceTypeKey


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
    

        $type = New-Object System.Object 
        $type | Add-Member -type NoteProperty -name friendlyname -Value ($returnedJSON.objects.$name.fields.friendlyname)
        $type | Add-Member -type NoteProperty -name key -Value ($returnedJSON.objects.$name.key)
    
       $objData += $type

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
             
             #Path to Tab Delimited user import file.
             [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
             [string]$ServerAddress,
             [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
             [string]$Protocol="https",
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
             key = ("SELECT Organization WHERE name = " + "'" +$itop_org_name + "'");
             output_fields= '*';
             } | ConvertTo-Json -Compress



$orgURI = "$Protocol" + "://$ServerAddress/webservices/rest.php?version=1.0&json_data=$getOrg"

$org = Invoke-RestMethod -Uri $orgURI -Headers $headers -Method Post -ContentType 'application/json'

$objOrg = @()

foreach ($name in (($org.objects | Get-Member -MemberType NoteProperty).Name))
    {
    

    $type = New-Object System.Object 
    $type | Add-Member -type NoteProperty -name Name -Value ($org.objects.$name.fields.name)
    $type | Add-Member -type NoteProperty -name Key -Value ($org.objects.$name.key)
    
    $objOrg += $type

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
    

        $type = New-Object System.Object 
        $type | Add-Member -type NoteProperty -name name -Value ($returnedJSON.objects.$name.fields.name)
        $type | Add-Member -type NoteProperty -name status -Value ($returnedJSON.objects.$name.fields.status)
        $type | Add-Member -type NoteProperty -name org_id -Value ($returnedJSON.objects.$name.fields.org_id)
        $type | Add-Member -type NoteProperty -name org_name -Value ($returnedJSON.objects.$name.fields.org_name)
        $type | Add-Member -type NoteProperty -name address -Value ($returnedJSON.objects.$name.fields.address)
        $type | Add-Member -type NoteProperty -name postal_code -Value ($returnedJSON.objects.$name.fields.postal_code)
        $type | Add-Member -type NoteProperty -name city -Value ($returnedJSON.objects.$name.fields.city)
        $type | Add-Member -type NoteProperty -name country -Value ($returnedJSON.objects.$name.fields.country)
        $type | Add-Member -type NoteProperty -name physicaldevice_list -Value ($returnedJSON.objects.$name.fields.physicaldevice_list)
        $type | Add-Member -type NoteProperty -name person_list -Value ($returnedJSON.objects.$name.fields.person_list)
        $type | Add-Member -type NoteProperty -name friendlyname -Value ($returnedJSON.objects.$name.fields.friendlyname)
        $type | Add-Member -type NoteProperty -name org_id_friendlyname -Value ($returnedJSON.objects.$name.fields.org_id_friendlyname)
        $type | Add-Member -type NoteProperty -name Key -Value ($returnedJSON.objects.$name.key)
    
       $objData += $type

    }

    return $objData 


    #cleanup for headers and base64 var
    $base64AuthInfo = $null
    $headers = $null
   

}

function New-iTopNetDevice {
    
        [CmdletBinding()]
         param(
             
             #Path to Tab Delimited user import file.
             [Parameter(Mandatory=$true,ValueFromPipeline=$False)]
             [string]$ServerAddress,
             [Parameter(Mandatory=$false,ValueFromPipeline=$False)]
             [string]$Protocol="https",
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
             [string]$itop_networkdevicetype_id

              ) 

<# Possible fields

name                              : 
description                       : 
org_id                            : 
organization_name                 : 
business_criticity                : low|medium|high
move2production                   : 
contacts_list                     : {}
documents_list                    : {}
applicationsolution_list          : {}
providercontracts_list            : {}
services_list                     : {}
softwares_list                    : {}
tickets_list                      : {}
serialnumber                      : string
location_id                       : int
location_name                     : 
status                            : production|implementation|stock|obsolete
brand_id                          : 
brand_name                        : 
model_id                          : 
model_name                        :
asset_number                      : 
purchase_date                     : 
end_of_warranty                   : 
networkdevice_list                : {}
physicalinterface_list            : {}
rack_id                           : 0
rack_name                         : 
enclosure_id                      : 0
enclosure_name                    : 
nb_u                              : 
managementip                      : 
powerA_id                         : 0
powerA_name                       : 
powerB_id                         : 0
powerB_name                       : 
fiberinterfacelist_list           : {}
san_list                          : {}
networkdevicetype_id              : 
networkdevicetype_name            : 
connectablecis_list               : {}
iosversion_id                     : 
iosversion_name                   : 
ram                               : 
finalclass                        : NetworkDevice
friendlyname                      : 
org_id_friendlyname               : 
location_id_friendlyname          : 
brand_id_friendlyname             : 
model_id_friendlyname             : 
rack_id_friendlyname              : 
enclosure_id_friendlyname         : 
powerA_id_friendlyname            : 
powerA_id_finalclass_recall       : 
powerB_id_friendlyname            : 
powerB_id_finalclass_recall       : 
networkdevicetype_id_friendlyname : 
iosversion_id_friendlyname        : 


#>

#Convert Valid IP Back to string
$itop_managementip = $itop_managementip.IPAddressToString




#Build headers for invoke-restmethod authentication
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
    

    $type = New-Object System.Object 
    $type | Add-Member -type NoteProperty -name Name -Value ($brand.objects.$name.fields.name)
    $type | Add-Member -type NoteProperty -name DeviceTypeKey -Value ($brand.objects.$name.key)
    
    $objBrand += $type

    }


$itop_brand_id = $objBrand[0].DeviceTypeKey


}

#get IOSVersion_id if IOS Vestion is populated.
#Get Brand Key ID if null

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
    

    $type = New-Object System.Object 
    $type | Add-Member -type NoteProperty -name Name -Value ($ios.objects.$name.fields.name)
    $type | Add-Member -type NoteProperty -name DeviceTypeKey -Value ($ios.objects.$name.key)
    
    $objIOS += $type

    }


$itop_iosversion_id = $objIOS[0].DeviceTypeKey


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
    

    $type = New-Object System.Object 
    $type | Add-Member -type NoteProperty -name Name -Value ($org.objects.$name.fields.name)
    $type | Add-Member -type NoteProperty -name Key -Value ($org.objects.$name.key)
    
    $objOrg += $type

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
    

    $type = New-Object System.Object 
    $type | Add-Member -type NoteProperty -name Name -Value ($model.objects.$name.fields.name)
    $type | Add-Member -type NoteProperty -name Key -Value ($model.objects.$name.key)
    
    $objModel += $type

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
    

    $type = New-Object System.Object 
    $type | Add-Member -type NoteProperty -name Name -Value ($netDevID.objects.$name.fields.name)
    $type | Add-Member -type NoteProperty -name Key -Value ($netDevID.objects.$name.key)
    
    $objnetDevID += $type

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
    
    if (($var -like "itop_*") -and (![string]::IsNullOrEmpty((Get-Variable $var -ValueOnly)))){
        
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
    

        $type = New-Object System.Object 
        $type | Add-Member -type NoteProperty -name name -Value ($returnedJSON.objects.$name.fields.name)
        $type | Add-Member -type NoteProperty -name key -Value ($returnedJSON.objects.$name.key)
    
       $objData += $type

    }

    return $objData 


    

    }
    #cleanup for headers and base64 var
    $base64AuthInfo = $null
    $headers = $null
}

