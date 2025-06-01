# Version 1.0

# functions
function Initialize-ColorScheme
{
    Set-Variable -Name "successColor" -Value "Green" -Scope "Script" -Option "Constant"
    Set-Variable -Name "infoColor" -Value "DarkCyan" -Scope "Script" -Option "Constant"
    Set-Variable -Name "warningColor" -Value "Yellow" -Scope "Script" -Option "Constant"
    Set-Variable -Name "failColor" -Value "Red" -Scope "Script" -Option "Constant"
}

function Show-Introduction
{
    Write-Host "This script revokes an M365 license from a list of users." -ForegroundColor $script:infoColor
    Write-Host "NOTE: Script does not track users who did not have the license assigned. It will continue without error in that case." -ForegroundColor $script:infoColor
    Read-Host "Press Enter to continue"
}

function Use-Module($moduleName)
{    
    $keepGoing = -not(Test-ModuleInstalled $moduleName)
    while ($keepGoing)
    {
        Prompt-InstallModule $moduleName
        Test-SessionPrivileges
        Install-Module $moduleName

        if ((Test-ModuleInstalled $moduleName) -eq $true)
        {
            Write-Host "Importing module..." -ForegroundColor $infoColor
            Import-Module $moduleName
            $keepGoing = $false
        }
    }
}

function Test-ModuleInstalled($moduleName)
{    
    $module = Get-Module -Name $moduleName -ListAvailable
    return ($null -ne $module)
}

function TryConnect-MgGraph($scopes)
{
    $connected = Test-ConnectedToMgGraph
    while (-not($connected))
    {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor $infoColor

        if ($null -ne $scopes)
        {
            Connect-MgGraph -Scopes $scopes -ErrorAction SilentlyContinue | Out-Null
        }
        else
        {
            Connect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }

        $connected = Test-ConnectedToMgGraph
        if (-not($connected))
        {
            Read-Host "Failed to connect to Microsoft Graph. Press Enter to try again"
        }
        else
        {
            Write-Host "Successfully connected!" -ForegroundColor $successColor
        }
    }    
}

function Test-ConnectedToMgGraph
{
    return $null -ne (Get-MgContext)
}

function Import-UserCsv
{
    $csvPath = Read-Host "Enter path to user CSV (must be .csv)"
    $csvPath = $csvPath.Trim('"')
    return Import-Csv -Path $csvPath
}

function Confirm-CSVHasCorrectHeaders($importedCSV)
{
    $firstRecord = $importedCSV | Select-Object -First 1
    $validCSV = $true

    if (-not($firstRecord | Get-Member -MemberType NoteProperty -Name "UserPrincipalName"))
    {
        Write-Warning "This CSV file is missing a header called 'UserPrincipalName'."
        $validCSV = $false
    }

    if (-not($validCSV))
    {
        Write-Host "Please make corrections to the CSV."
        Read-Host "Press Enter to exit"
        Exit
    }
}

function Get-AvailableLicenses
{
    $licenseLookupTable = @{
        "8f0c5670-4e56-4892-b06d-91c085d7004f" = "App Connect IW"
        "4b9405b0-7788-4568-add1-99614e613b69" = "Exchange Online (Plan 1)"
        "19ec0d23-8335-4cbd-94ac-6050e30712fa" = "Exchange Online (Plan 2)"
        "efccb6f7-5641-4e0e-bd10-b4976e1bf68e" = "Enterprise Mobility + Security E3"
        "b05e124f-c7cc-45a0-a6aa-8cf78c946968" = "Enterprise Mobility + Security E5"
        "c2273bd0-dff7-4215-9ef5-2c7bcfb06425" = "Microsoft 365 Apps for Enterprise"
        "3b555118-da6a-4418-894f-7df1e2096870" = "Microsoft 365 Business Basic"
        "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46" = "Microsoft 365 Business Premium"
        "f245ecc8-75af-4f8e-b61f-27d8114de5f3" = "Microsoft 365 Business Standard"
        "05e9a617-0261-4cee-bb44-138d3ef5d965" = "Microsoft 365 E3"
        "dcf0408c-aaec-446c-afd4-43e3683943ea" = "Microsoft 365 E3 (no Teams)"
        "06ebc4ee-1bb5-47dd-8120-11324bc54e06" = "Microsoft 365 E5"
        "18a4bd3f-0b5b-4887-b04f-61dd0ee15f5e" = "Microsoft 365 E5 (no Teams)"
        "44575883-256e-4a79-9da4-ebe9acabe2b2" = "Microsoft 365 F1"
        "66b55226-6b4f-492c-910c-a3b7a3c9d993" = "Microsoft 365 F3"
        "4ef96642-f096-40de-a3e9-d83fb2f90211" = "Microsoft Defender for Office 365 (Plan 1)"
        "3dd6cf57-d688-4eed-ba52-9e40b5468c3e" = "Microsoft Defender for Office 365 (Plan 2)"
        "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235" = "Microsoft Fabric (Free)"
        "dcb1a3ae-b33f-4487-846a-a640262fadf4" = "Microsoft Power Apps Plan 2 Trial"
        "f30db892-07e9-47e9-837c-80727f46fd3d" = "Microsoft Power Automate Free"
        "5b631642-bd26-49fe-bd20-1daaa972ef80" = "Microsoft PowerApps for Developer"
        "1f2f344a-700d-42c9-9427-5cea1d5d7ba6" = "Microsoft Stream"
        "7e31c0d9-9551-471d-836f-32ee72be4a01" = "Microsoft Teams Enterprise"
        "3ab6abff-666f-4424-bfb7-f0bc274ec7bc" = "Microsoft Teams Essentials"    
        "36a0f3b3-adb5-49ea-bf66-762134cf063a" = "Microsoft Teams Premium"
        "4cde982a-ede4-4409-9ae6-b003453c8ea6" = "Microsoft Teams Rooms Pro"
        "18181a46-0d4e-45cd-891e-60aabd171b4e" = "Office 365 E1"
        "f8ced641-8e17-4dc5-b014-f5a2d53f6ac8" = "Office 365 E1 (no Teams)"
        "6fd2c87f-b296-42f0-b197-1e91e994b900" = "Office 365 E3"
        "46c3a859-c90d-40b3-9551-6178a48d5c18" = "Office 365 E3 (no Teams)"
        "c7df2760-2c81-4ef7-b578-5b5392b571df" = "Office 365 E5"
        "7b26f5ab-a763-4c00-a1ac-f6c4b5506945" = "Power BI Premium P1"
        "f8a1db68-be16-40ed-86d5-cb42ce701560" = "Power BI Pro"
        "52ea0e27-ae73-4983-a08f-13561ebdb823" = "Teams Premium (for Departments)"
        "6470687e-a428-4b7a-bef2-8a291ad947c9" = "Windows Store for Business"
    }
    
    $uri = "https://graph.microsoft.com/v1.0/subscribedSkus"
    try
    {
        $licenses = Invoke-MgGraphRequest -Method "Get" -Uri $uri -ErrorAction "Stop"
    }    
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue getting available licenses." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
        return
    }

    $licenseTable = [System.Collections.Generic.List[object]]::new(30)
    foreach ($license in $licenses.value)
    {
        $name = $licenseLookupTable[$license.skuId]
        if ($null -eq $name ) { $name = $license.skuPartNumber }
        $amountPurchased = $license.prepaidUnits.enabled
        $amountAvailable = $amountPurchased - $license.consumedUnits

        $licenseInfo = [PSCustomObject]@{
            "Name"      = $name
            "Available" = $amountAvailable
            "Purchased" = $amountPurchased
            "SkuId"     = $license.skuId        
        }
        $licenseTable.Add($licenseInfo)
    }

    return Write-Output $licenseTable -NoEnumerate
}

function Prompt-LicenseToRevoke($availableLicenses)
{   
    # Display available licenses with an option number next to each.
    $option = 0    
    $availableLicenses | Sort-Object -Property "Name" | ForEach-Object { $option++; $_ | Add-Member -NotePropertyName "Option" -NotePropertyValue $option }
    $availableLicenses | Sort-Object -Property "Option" | Format-Table -Property @("Option", "Name", "Available", "Purchased") | Out-Host
    $selection = (Read-Host "Select an option (1-$option)") -As [int]

    do
    {        
        # Check that selection is a number between 1 and option count. (Avoids use of regex because that's not great for matching multi-digit number ranges.)
        $validSelection = ($selection -is [int]) -and (($selection -ge 1) -and ($selection -le $option))
        if (-not($validSelection)) 
        {
            Write-Host "Please enter 1-$option." -ForegroundColor $warningColor
            $selection = (Read-Host "Select an option (1-$option)") -As [int]
        }
    }
    while (-not($validSelection))

    foreach ($license in $availableLicenses)
    {
        if ($license.option -eq [int]$selection)
        {
            return $license
        }
    }
}

function Revoke-LicenseFromUsers($userCsv, $license)
{
    Write-Host "Revoking license: $($license.name)" -ForegroundColor $script:infoColor
    foreach ($user in $userCsv)
    {
        Write-Progress -Activity "Revoking license from users..." -Status $user.UserPrincipalName
        Revoke-License -User $user -License $license
    }
}

function Revoke-License($user, $license)
{
    if ($null -eq $user.UserPrincipalName) { continue }
    $upn = $user.UserPrincipalName.Trim()

    try
    {
        Set-MgUserLicense -UserId $upn -AddLicenses @() -RemoveLicenses @($license.SkuId) -ErrorAction "Stop" | Out-Null
    }
    catch
    {
        $errorRecord = $_
        Log-Warning "There was an issue revoking M365 license: $($license.name) from $upn`n$errorRecord"
    }
}

function Log-Warning($message, $logPath = "$PSScriptRoot\logs.txt")
{
    $message = "[$(Get-Date -Format 'yyyy-MM-dd hh:mm tt') W] $message"
    Write-Output $message | Tee-Object -FilePath $logPath -Append | Write-Host -ForegroundColor $script:warningColor
}

# main
Initialize-ColorScheme
Show-Introduction
Use-Module "Microsoft.Graph.Users"
TryConnect-MgGraph -Scopes @("User.ReadWrite.All", "Organization.Read.All")
$userCsv = Import-UserCsv
Confirm-CSVHasCorrectHeaders $userCsv
$availableLicenses = Get-AvailableLicenses
$licenseToRevoke = Prompt-LicenseToRevoke $availableLicenses
Revoke-LicenseFromUsers -UserCsv $userCsv -License $licenseToRevoke
Write-Host "All done!" -ForegroundColor $script:successColor
Read-Host "Press Enter to exit"