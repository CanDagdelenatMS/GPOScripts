<#
----------------------------------------
File: get-gplink.ps1
Version: 1.3
Author: Thomas Bouchereau
-----------------------------------------

Disclaimer:
This sample script is not supported under any Microsoft standard support program or service. 
The sample script is provided AS IS without warranty of any kind. Microsoft further disclaims 
all implied warranties including, without limitation, any implied warranties of merchantability 
or of fitness for a particular purpose. The entire risk arising out of the use or performance of 
the sample scripts and documentation remains with you. In no event shall Microsoft, its authors, 
or anyone else involved in the creation, production, or delivery of the scripts be liable for any 
damages whatsoever (including, without limitation, damages for loss of business profits, business 
interruption, loss of business information, or other pecuniary loss) arising out of the use of or 
inability to use the sample scripts or documentation, even if Microsoft has been advised of the 
possibility of such damages 
 #>


function Get-Gplink  {

<# 
.SYNOPSIS
return and decode content of gplink attribut

.DESCRIPTION
This function purpose is to list the gplink attribut of an OU, Site or DomainDNS.
It will return the following information for each linked GPOs:

    - Target: DN of the targeted object
    - GPOID: GUID of the GPO
    - GPOName: Friendly Name of the GPO
    - GPODomain: Originating domain of the GPO
    - Enforced: <Yes|No>
    - Enabled:  <Yes|No>
    - Order: Link order of the GPO on the OU,Site,DomainDNS (does not report inherited order)

.PARAMETER Path: Give de Distinguished Name of object you want to list the gplink


.INPUTS
DN of the object with GLINK attribut

.OUTPUTS
Target: DC=fourthcoffee,DC=com
GPOID: 31B2F340-016D-11D2-945F-00C04FB984F9
GPOName: Default Domain Policy
GPODomain: fourthcoffee.com
Enforced: <YES - NO>
Enabled: <YES - NO>
Order: 1

.EXAMPLE
get-gplink -path "dc=fourthcoffee,dc=com"

This command will list the GPOs that are linked to the DomainDNS object "dc=fourthcoffee,dc=com" 

.EXAMPLE
get-gplink -path "dc=child,dc=fourthcoffee,dc=com" -server childdc.child.fourthcoffee.com

This command will list the GPOs that are linked to the DomainDNS object "dc=child,dc=fourthcoffee,dc=com". You need to specify a
target DC of the domain child.fourthcoffee.com in order for the command to work.


.EXAMPLE
Get-Gplink -site "CN=Default-First-Site-Name,CN=Sites,CN=Configuration,DC=fourthcoffee,DC=com"

This command will list the GPOs that are linked to site "Default-First-Site-Name"

.EXAMPLE
get-gplink -path "dc=fourthcoffee,dc=com" | export-csv "gplink.csv"

This command will list the GPOs that are linked to the DomainDNS object "dc=fourthcoffee,dc=com" and export them to a csv file.
The csv file can be used as an input to the cmdlet new-gplink 


.EXAMPLE
get-adobject -filter {(objectclass -eq "DomainDNS") -or (objectclass -eq "OrganizationalUnit")} | foreach {get-gplink -path $_.distinguishedname} | export-csv "gplinksall.csv"

This command will list all objects of type "DomainDNS" and "OrganizationalUnit" that have GPOs linked and will list those GPOs, their status and link order.

#>

[cmdletBinding()]
param ([string]$path,[string]$server,[string]$site)

#Import AD and GPO modules
Import-Module activedirectory
Import-Module grouppolicy
# get the DN to te configuration partition
$configpart=(Get-ADRootDSE).configurationNamingContext

#get content of attribut gplink on site object or OU
if ($site)
    {
        $gplink=Get-ADObject -Filter {distinguishedname -eq $site} -searchbase $configpart -Properties gplink
        $target=$site
    }
elseif ($path)
    {
        switch ($server)
            {
             "" {$gplink=Get-ADObject -Filter {distinguishedname -eq $path} -Properties gplink}
             default {$gplink=Get-ADObject -Filter {distinguishedname -eq $path} -Properties gplink -server $server     
                      }
            }
    $target=$path
    }

    

#if DN is not valid return" Invalide DN" error

if ($gplink -eq $null)
    {
        write-host "Either Invalide DN in the current domain, specify a DC of the target DN domain or no GPOlinked to this DN"
    }

# test if glink is not null or only containes white space before continuing. 

if (!((($gplink.gplink) -like "") -or (($gplink.gplink) -like " ")))
    {

        #set variale $o to define link order
        $o=0    
    
        #we split the gplink string in order to seperate the diffent GPO linked

        $split=$gplink.gplink.split("]")
    
        #we need to do a reverse for to get the proper link order

         for ($s=$split.count-1;$s -gt -1;$s--)
            {
          
                #since the last character in the gplink string is a "]" the last split is empty we need to ignore it
           
                if ($split[$s].length -gt 0)
                    {
                        $o++
                        $order=$o            
                        $gpoguid=$split[$s].substring(12,36)
			            $gpodomainDN=($split[$s].substring(72)).split(";")
			            $domain=($gpodomaindn[0].substring(3)).replace(",DC=",".")
			   			$checkdc=(get-addomaincontroller -domainname $domain -discover).name
            			                     
                        #we test if the $gpoguid is a valid GUID in the domain if not we return a "Oprhaned GpLink or External GPO" in the $gponname
                        
			
			
				        $mygpo=get-gpo -guid $gpoguid -domain $domain -server "$($checkdc).$($domain)" 2> $null
                    
				        if ($mygpo -ne $null )
					              {
					               $gponame=$MyGPO.displayname
					               $gpodomain=$domain	
					              }	
    				              
    				        else
    					          {
					              $gponame="Orphaned GPLink" 
                        	      $gpodomain=$domain   
    					          }
    				
			
                        #we test the last 2 charaters of the split do determine the status of the GPO link
           
    
                        if (($split[$s].endswith(";0")))
                            {
                                $enforced= "No"
                                $enabled= "Yes"
                            }
                        elseif (($split[$s].endswith(";1")))
                            {
                                $enabled= "No"
                                $enforced="No"
                            }
                        elseif (($split[$s].endswith(";2")))
                            {
                                $enabled="Yes"
                                $enforced="Yes"
                            }
                        elseif (($split[$s].endswith(";3")))
                            {
                                $enabled="No"
                                $enforced="Yes"
                            }

                         #we create an object representing each GPOs, its links status and link order

                        $return = New-Object psobject 
                        $return | Add-Member -membertype NoteProperty -Name "Target" -Value $target 
                        $return | Add-Member -membertype NoteProperty -Name "GPOID" -Value $gpoguid
                        $return | Add-Member -membertype NoteProperty -Name "DisplayName" -Value $gponame 
			            $return | Add-Member -membertype NoteProperty -Name "Domain" -Value $gpodomain 
                        $return | Add-Member -membertype NoteProperty -Name "Enforced" -Value $enforced
                        $return | Add-Member -membertype NoteProperty -Name "Enabled" -Value $enabled
                        $return | Add-Member -membertype NoteProperty -Name "Order" -Value $order
                        $return
                    }
         
            }
         
    }
   
    

 }
    
  