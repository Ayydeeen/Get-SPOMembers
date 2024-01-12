#Connect to SPO and MSOL services
Connect-SPOService
Connect-MSOLService

#Add user as admin for all associated sites
foreach ($i in get-content('C:\links.csv')) { ####CHANGEME####
Set-SPOUser -Site $i -LoginName admin@domain.com -IsSiteCollectionAdmin $True ####CHANGEME####
}

# Function to remove "_o" suffix from owner login names
function Remove-OOwnerSuffix($loginName) {
    # Check if the login name ends with "_o" and remove it
    if ($loginName -match '(.*)_o$') {
        return $Matches[1]
    }
    return $loginName
}

#Function to translate OID for Owner user objects to regular names
function Get-ObjectName($objectId) {
    try {
        $user = get-msoluser -ObjectID $objectId -ErrorAction Stop
        return $user.DisplayName
    } catch {
        #na
    }

    try {
        $group = Get-MSolGroup -ObjectID $objectID -ErrorAction Stop
        return $group.DisplayName
    } catch {
        #na
    }
    return "Object not found"
}

#Create main info array
$allUserInfo = @()

#--------------------------------

foreach ($siteURL in get-content('C:\links1.csv')) { ####CHANGEME####

    #Set all variables to null
    $CALoginName = ""
    $SiteOwners = ""
    $siteAdmins = ""
    $siteMemebers = ""
    $siteGuests = ""

    # Get site owners
    $CALoginName = Get-SPOSite -Identity $siteUrl | Select-Object -ExpandProperty Owner

    $SiteOwners = $CALoginName

    # Get site admins
    $siteAdmins = Get-SPOUser -Site $siteUrl -Limit All | Where-Object { $_.IsSiteAdmin -eq $true }

    # Get site members (users with contribute permissions or higher)
    $siteMembers = Get-SPOUser -Site $siteUrl -Limit All | Where-Object { $_.IsSiteAdmin -eq $false } 

    # Get site guests (external users)
    $siteGuests = Get-SPOExternalUser -Site $siteUrl

    #------------------------------------


    # Create an array to store user information
    $userInfo = @()

    # Add site owners to the array
    $siteOwners | ForEach-Object {
        if ($_ -ne "") {
            if ($_ -match '_o$') {
                $userInfo += [PSCustomObject]@{
                    Level = "Owner"
                    DisplayName = Get-ObjectName(Remove-OOwnerSuffix($_))
                    LoginName = "Owner"
                    Type = "Owner"
                    SiteURL = $siteURL
                }
            }
            else {
                $userInfo += [PSCustomObject]@{
                    Level = "Owner"
                    DisplayName = $_
                    LoginName = "Owner"
                    Type = "Owner"
                    SiteURL = $siteURL
                } 
            }
        }
    }

    # Add site owners to the array
    $siteAdmins | ForEach-Object {
        if ($_ -ne "") {
            if ($_.Groups -ne "{}") {
                $userInfo += [PSCustomObject]@{
                    Level = "Admin"
                    DisplayName = $_.DisplayName
                    LoginName = "Group"
                    Type = "Group"
                    SiteURL = $siteURL
                }
            }
            else {
                if ($_.DisplayName -ne "Admin Account Display Name") {   ####CHANGEME####
                    $userInfo += [PSCustomObject]@{
                        Level = "Admin"
                        DisplayName = $_.DisplayName
                        LoginName = $_.loginname
                        Type = "User"
                        SiteURL = $siteURL
                    }
                }
            }
        }
    }

    # Add site members to the array
    $siteMembers | ForEach-Object {
        if ($_ -ne "") {
            if (($_.DisplayName -ne "System Account") -and ($_.DisplayName -ne "Sharepoint App") -and ($_.DisplayName -ne "NT Service\spsearch")) {
                if ($_.Groups -ne "{}" -or $_.loginname -notmatch "@") {
                    $userInfo += [PSCustomObject]@{
                        Level = "Member"
                        DisplayName = $_.DisplayName
                        LoginName = "n/a"
                        Type = "Group"
                        SiteURL = $siteURL
                    }
                }
                else {
                    $userInfo += [PSCustomObject]@{
                        Level = "Member"
                        DisplayName = $_.DisplayName
                        LoginName = $_.loginname
                        Type = "User"
                        SiteURL = $siteURL
                    }
                }
            }
        }
    }

    # Add site guests to the array
    $siteGuests | ForEach-Object {
        if ($_ -ne "") {
            $userInfo += [PSCustomObject]@{
                Level = "Guest (External User)"
                DisplayName = $_.Email
                LoginName = "n/a"
                Type = "ExternalEmail"
                SiteURL = $siteURL
            }
        }
    }


    #Add site information to main array
    $allUserInfo += $userInfo
    echo "Successfully Processed $SiteURL"
}

#Export array to CSV
$allUserInfo | Export-Csv -Path "C:\combined_user_info.csv" -NoTypeInformation

#Remove user from Admin for all associated sites in CSV
foreach ($i in get-content('C:\links.csv')) { ####CHANGEME####
Set-SPOUser -Site $i -LoginName admin@domain.com -IsSiteCollectionAdmin $False ####CHANGEME####
}
