# 
# This script gathers info and creates a new AD user that is mail enabled in M365 then forces a sync
# Written by Aaron Wolfrom (awolfrom@keystoneind.com)
# Must be run in an Admin context 
# Can be run from any system that has the Exchange tools installed and access to the Entra sync server
#

#Ensure the Exchange PS SnapIn is loaded
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

#Prompt to set variables for the specific user
$firstName = read-host "Enter the new users first name: "
$lastName = read-host "Enter the new users last name: "
$fullName = $firstName + " " + $lastName
$firstInit = $firstName[0]
$UPN = $firstInit + $lastName + "@keystoneind.com"
$description = read-host  "Enter the description of the user: "
$department = read-host "Enter thier department: "
write-host "Which Office should the user be assigned to?" -ForegroundColor Black -BackgroundColor Magenta
write-host "1. IT" -ForegroundColor Yellow -BackgroundColor Magenta
write-host "2. Cherry Hill" -ForegroundColor Yellow -BackgroundColor Magenta
write-host "3. Cinnaminson" -ForegroundColor Yellow -BackgroundColor Magenta
write-host "4. Gibbstown" -ForegroundColor Yellow -BackgroundColor Magenta
write-host "5. Myerstown" -ForegroundColor Yellow -BackgroundColor Magenta
write-host "6. NL-OSS" -ForegroundColor Yellow -BackgroundColor Magenta
write-host "7. CA-Westminster" -ForegroundColor Yellow -BackgroundColor Magenta
write-host "8. China" -ForegroundColor Yellow -BackgroundColor Magenta
write-host "9. France" -ForegroundColor Yellow -BackgroundColor Magenta
write-host "10. Germany" -ForegroundColor Yellow -BackgroundColor Magenta
$ouNum = read-host "Enter the number of the office they are associated with: "

#Check which number was entered in order to set the proper OU for the user
switch ($ouNum){
    "1" {$OU = "keystoneind.com/_USERS (Secure) - Windows 10 Image/End Users (Azure AD Sync)/_IT Department"
         $office = "NJ-Gibbstown"
        }
    "2" {$OU = "keystoneind.com/_USERS (Secure) - Windows 10 Image/End Users (Azure AD Sync)/Cherry Hill"
         $office = "NJ-Cherry Hill"
        }
    "3" {$OU = "keystoneind.com/_USERS (Secure) - Windows 10 Image/End Users (Azure AD Sync)/Cinnaminson"
         $office = "NJ-Cinnaminson"
        }
    "4" {$OU = "keystoneind.com/_USERS (Secure) - Windows 10 Image/End Users (Azure AD Sync)/Gibbstown"
         $office = "NJ-Gibbstown"
        }
    "5" {$OU = "keystoneind.com/_USERS (Secure) - Windows 10 Image/End Users (Azure AD Sync)/Myerstown"
         $office = "PA-Myerstown"
        }
    "6" {$OU = "keystoneind.com/_USERS (Secure) - Windows 10 Image/End Users (Azure AD Sync)/NL-Oss"
         $office = "NL-Oss"
        }
    "7" {$OU = "keystoneind.com/_USERS (Secure) - Windows 10 Image/End Users (Azure AD Sync)/Remote"
         $office = "CA-Westminster"
        }
    "8" {$OU = "keystoneind.com/_USERS (Secure) - Windows 10 Image/End Users (Azure AD Sync)/Remote"
         $office = "CN-Guangzhou"
        }
    "9" {$OU = "keystoneind.com/_USERS (Secure) - Windows 10 Image/End Users (Azure AD Sync)/Remote"
         $office = "FR-France"
        }
    "10" {$OU = "keystoneind.com/_USERS (Secure) - Windows 10 Image/End Users (Azure AD Sync)/Remote"
          $office = "DE-Singen"
         }
    default {$OU = "none"
             $office = "none"
            }
}

#Check for a vaild OU then either prompt for a password, create the user, and force a sync, or notify the Admin that a vaild OU was not supplied.
if ($OU -ne "none"){
        New-RemoteMailbox -FirstName $firstName -LastName $lastName -Name $fullName -Password (Read-Host "Enter their password: " -AsSecureString) -UserPrincipalName $UPN -OnPremisesOrganizationalUnit $OU
        #Update the newly created user to add additonal AD information
        Set-ADUser -Identity $UPN -Company $office -Department $department -Description $description -Manager $manager -Office $office -Tite $description -Replace @{physicalDeliveryOfficeName = $office}
        Enter-PSSession kgsvazure01
        Start-ADSyncSyncCycle -PolicyType delta
        Exit-PSSession
        write-host "The user " $fullName " has been created."
}
else {
    write-host "No vaild Office entered please re-run the script."
}