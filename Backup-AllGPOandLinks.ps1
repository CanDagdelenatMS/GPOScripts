#requires -runasadministrator
. .\get-gplink_V1.3.ps1
$gpolocation = "c:\gpo\"
$currentdate = (get-date).ToString("ddMMyyyy")
$gporetentiondays= 7

Write-Host "Creating new folder for $currentdate under $gpolocation if necessary..." -ForegroundColor Yellow
if (-not (Get-Item "$gpolocation\$currentdate" -ErrorAction SilentlyContinue)) { New-Item -Path $gpolocation -ItemType Directory -Name $currentdate}
    else {
            Write-Host "$gpolocation already exits. Deleting the current folder structure..."
            Remove-Item "$gpolocation\$currentdate"  -Recurse -Force
            New-Item -Path $gpolocation -ItemType Directory -Name $currentdate
        }
if (Get-Item "$gpolocation\$currentdate" -ErrorAction SilentlyContinue) { New-Item -Path "$gpolocation\$currentdate" -ItemType Directory -Name GPOPermissions}
    else {Throw 'Error: folder structure was not created'}


Start-Transcript -Path "$gpolocation\$currentdate\GPOBackup_$currentdate.log" -Force

#Delete folders older than $gporetentionday days
Write-Host "Deleting folders older than $gporetentiondays if necessary..." -ForegroundColor Yellow
foreach ($folder in (Get-ChildItem -Path "$gpolocation" -Directory)) {
if ($folder.CreationTimeUtc.Date -lt (Get-Date).Date.AddDays(-7)) {
    Write-Host "Deleting folders older $folder..." -ForegroundColor Yellow
    Remove-Item $folder.FullName -Recurse -Force -Confirm:$false}
} 


Write-Host "Exporting GP Links..." -ForegroundColor Yellow
#Export All GP Links
Get-ADOrganizationalUnit -Filter * | foreach {Get-Gplink -path $_.DistinguishedName} | Export-csv -Path "$gpolocation\$currentdate\gplink.csv" -Encoding Unicode -Delimiter ";"

Write-Host "Exporting All GPOs..." -ForegroundColor Yellow
#Export All GPOs
$allgpos = Get-GPO -All
foreach ($gpo in $allgpos) {
Backup-GPO -Name $gpo.displayname -Path "$gpolocation\$currentdate" -Comment "$currentdate"
(get-acl  "AD:\$($gpo.Path)").Access | where {$_.IsInherited -eq $false -and $_.AccessControlType -eq "Deny"} `
        | Export-Csv -LiteralPath "$gpolocation\$currentdate\GPOPermissions\$($GPO.Id).csv" #Filter for only deny permissions
}

<# Default Identities on a GPO:
CREATOR OWNER                             
NT AUTHORITY\ENTERPRISE DOMAIN CONTROLLERS
NT AUTHORITY\Authenticated Users          
NT AUTHORITY\SYSTEM                       
RICHCAN\Domain Admins                     
RICHCAN\Enterprise Admins                 
NT AUTHORITY\Authenticated User or Domain Computers --> by default has ApplyGroupPolicy permissions

'edacfd8f-ffb3-11d1-b41d-00a0c968f939' --> ApplyGroupPolicy GUID
#>



Stop-Transcript

