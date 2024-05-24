
. .\get-gplink_V1.3.ps1
$gpolocation = "c:\gpo\"
$currentdate = (get-date).ToString("ddMMyyyy")
$gporetentiondays= 7

Write-Host "Creating new folder for $currentdate under $gpolocation if necessary..." -ForegroundColor Yellow
if (-not (Get-Item "$gpolocation\$currentdate" -ErrorAction SilentlyContinue)) { New-Item -Path $gpolocation -ItemType Directory -Name $currentdate}

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
Get-ADOrganizationalUnit -Filter * | foreach {Get-Gplink -path $_.DistinguishedName} | Export-csv -Path "$gpolocation\$currentdate\gplink.csv" -Encoding Unicode

Write-Host "Exporting All GPOs..." -ForegroundColor Yellow
#Export All GPOs
$allgpos = Get-GPO -All
foreach ($gpo in $allgpos) {
Backup-GPO -Name $gpo.displayname -Path "$gpolocation\$currentdate" -Comment "$currentdate"
}






