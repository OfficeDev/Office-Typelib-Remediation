# Copyright (c) Microsoft Corporation. All rights reserved.
#
# Detection and remediation for orphaned TypeLib registry keys
#

# Log actions to %windir%\temp\typelibfix.log
function Log
{
    param(
        [Parameter(Mandatory=$true)][string]$logMessage
    )

    $LogFile = "$env:windir\Temp\typelibfix.log"
    $LogDate = get-date -format "MM/dd/yyyy HH:mm:ss"
    $LogLine = "$LogDate $logMessage"
    Add-Content -Path $LogFile -Value $LogLine -ErrorAction SilentlyContinue
}

# Clean up stale TypeLib entries
function ProcessTypelibHive
{
    param (
        [string] $sHive
    )
    ForEach ($tl in $arrTypeLibs) 
    {
        $sKey = $sHive + $tl

        # Check keys for confirmed TypeLibs
        if (Test-Path -Path $sKey)
        {
            Log("Found registration for typelib $tl")
            Write-Verbose "Found registration for typelib $tl"

            # Get Versioned subkey
            $versions = @(Get-ChildItem -Path $sKey -Name)

            ForEach ($version in $versions)
            {
                $sWin32Key = $sKey + "\" + $version + "\0\Win32"

                if (Test-Path -Path $sWin32Key)
                {
                    Log("Found win32 registration for typelibe $tl version $version")
                    Write-Verbose "Found win32 registration for typelibe $tl version $version"

                    $libraryPath = Get-ItemPropertyValue -Path $sWin32Key -Name "(default)"
                    $resourceId = 0

                    if ($libraryPath -match "(.*)\\(\d+)")
                    {
                        $libraryPath = $Matches[1]
                        $resourceId = [int]$Matches[2]
                    }
                
                    Log("Found library path at $libraryPath" + $(if ([int]$resourceId -gt 0) { " (resourceId $resourceId)"}))
                    Write-Verbose ("Found library path at $libraryPath" + $(if ([int]$resourceId -gt 0) { " (resourceId $resourceId)"}))
                
                    if (!(Test-Path -Path $libraryPath))
                    {
                        Log("Found corrupt win32 typelib registration for $tl referencing non-existant file $libraryPath")
                        Write-Output "Found corrupt win32 typelib registration for $tl referencing non-existant file $libraryPath"

                        Log("Removing key $swin32Key")
                        Write-Output "Removing key $swin32Key"
                        Remove-Item $sWin32Key
                    }
                }
            }
        }
    }   
}

# List of Office TypeLibs
$arrTypeLibs = @(
    '{000204EF-0000-0000-C000-000000000046}',
    '{000204EF-0000-0000-C000-000000000046}',
    '{00020802-0000-0000-C000-000000000046}',
    '{00020813-0000-0000-C000-000000000046}',
    '{00020905-0000-0000-C000-000000000046}',
    '{0002123C-0000-0000-C000-000000000046}',
    '{00024517-0000-0000-C000-000000000046}',
    '{0002E157-0000-0000-C000-000000000046}',
    '{00062FFF-0000-0000-C000-000000000046}',
    '{0006F062-0000-0000-C000-000000000046}',
    '{0006F080-0000-0000-C000-000000000046}',
    '{012F24C1-35B0-11D0-BF2D-0000E8D0D146}',
    '{06CA6721-CB57-449E-8097-E65B9F543A1A}',
    '{07B06096-5687-4D13-9E32-12B4259C9813}',
    '{0A2F2FC4-26E1-457B-83EC-671B8FC4C86D}',
    '{0AF7F3BE-8EA9-4816-889E-3ED22871FE05}',
    '{0D452EE1-E08F-101A-852E-02608C4D0BB4}',
    '{0EA692EE-BB50-4E3C-AEF0-356D91732725}',
    '{1F8E79BA-9268-4889-ADF3-6D2AABB3C32C}',
    '{2374F0B1-3220-4c71-B702-AF799F31ABB4}',
    '{238AA1AC-786F-4C17-BAAB-253670B449B9}',
    '{28DD2950-2D4A-42B5-ABBF-500AA42E7EC1}',
    '{2A59CA0A-4F1B-44DF-A216-CB2C831E5870}',
    '{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}',
    '{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}',
    '{2F7FC181-292B-11D2-A795-DFAA798E9148}',
    '{3120BA9F-4FC8-4A4F-AE1E-02114F421D0A}',
    '{31411197-A502-11D2-BBCA-00C04F8EC294}',
    '{3B514091-5A69-4650-87A3-607C4004C8F2}',
    '{47730B06-C23C-4FCA-8E86-42A6A1BC74F4}',
    '{49C40DDF-1B04-4868-B3B5-E49F120E4BFA}',
    '{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}',
    '{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}',
    '{4D95030A-A3A9-4C38-ACA8-D323A2267698}',
    '{55A108B0-73BB-43db-8C03-1BEF4E3D2FE4}',
    '{56D04F5D-964F-4DBF-8D23-B97989E53418}',
    '{5B87B6F0-17C8-11D0-AD41-00A0C90DC8D9}',
    '{66CDD37F-D313-4E81-8C31-4198F3E42C3C}',
    '{6911FD67-B842-4E78-80C3-2D48597C2ED0}',
    '{698BB59C-38F1-4CEF-92F9-7E3986E708D3}',
    '{6DDCE504-C0DC-4398-8BDB-11545AAA33EF}',
    '{6EFF1177-6974-4ED1-99AB-82905F931B87}',
    '{73720002-33A0-11E4-9B9A-00155D152105}',
    '{759EF423-2E8F-4200-ADF0-5B6177224BEE}',
    '{76F6F3F5-9937-11D2-93BB-00105A994D2C}',
    '{773F1B9A-35B9-4E95-83A0-A210F2DE3B37}',
    '{7D868ACD-1A5D-4A47-A247-F39741353012}',
    '{7E36E7CB-14FB-4F9E-B597-693CE6305ADC}',
    '{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}',
    '{8404DD0E-7A27-4399-B1D9-6492B7DD7F7F}',
    '{8405D0DF-9FDD-4829-AEAD-8E2B0A18FEA4}',
    '{859D8CF5-7ADE-4DAB-8F7D-AF171643B934}',
    '{8E47F3A2-81A4-468E-A401-E1DEBBAE2D8D}',
    '{91493440-5A91-11CF-8700-00AA0060263B}',
    '{9A8120F2-2782-47DF-9B62-54F672075EA1}',
    '{9B7C3E2E-25D5-4898-9D85-71CEA8B2B6DD}',
    '{9B92EB61-CBC1-11D3-8C2D-00A0CC37B591}',
    '{9D58B963-654A-4625-86AC-345062F53232}',
    '{9DCE1FC0-58D3-471B-B069-653CE02DCE88}',
    '{A4D51C5D-F8BF-46CC-92CC-2B34D2D89716}',
    '{A717753E-C3A6-4650-9F60-472EB56A7061}',
    '{AA53E405-C36D-478A-BBFF-F359DF962E6D}',
    '{AAB9C2AA-6036-4AE1-A41C-A40AB7F39520}',
    '{AB54A09E-1604-4438-9AC7-04BE3E6B0320}',
    '{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}',
    '{AC2DE821-36A2-11CF-8053-00AA006009FA}',
    '{B30CDC65-4456-4FAA-93E3-F8A79E21891C}',
    '{B8812619-BDB3-11D0-B19E-00A0C91E29D8}',
    '{B9164592-D558-4EE7-8B41-F1C9F66D683A}',
    '{B9AA1F11-F480-4054-A84E-B5D9277E40A8}',
    '{BA35B84E-A623-471B-8B09-6D72DD072F25}',
    '{BDEADE33-C265-11D0-BCED-00A0C90AB50F}',
    '{BDEADEF0-C265-11D0-BCED-00A0C90AB50F}',
    '{BDEADEF0-C265-11D0-BCED-00A0C90AB50F}',
    '{C04E4E5E-89E6-43C0-92BD-D3F2C7FBA5C4}',
    '{C3D19104-7A67-4EB0-B459-D5B2E734D430}',
    '{C78F486B-F679-4af5-9166-4E4D7EA1CEFC}',
    '{CA973FCA-E9C3-4B24-B864-7218FC1DA7BA}',
    '{CBA4EBC4-0C04-468d-9F69-EF3FEED03236}',
    '{CBBC4772-C9A4-4FE8-B34B-5EFBD68F8E27}',
    '{CD2194AA-11BE-4EFD-97A6-74C39C6508FF}',
    '{E0B12BAE-FC67-446C-AAE8-4FA1F00153A7}',
    '{E985809A-84A6-4F35-86D6-9B52119AB9D7}',
    '{ECD5307E-4419-43CF-8BDA-C9946AC375CF}',
    '{EDCD5812-6A06-43C3-AFAC-46EF5D14E22C}',
    '{EDCD5812-6A06-43C3-AFAC-46EF5D14E22C}',
    '{EDCD5812-6A06-43C3-AFAC-46EF5D14E22C}',
    '{EDDCFF16-3AEE-4883-BD91-0F3978640DFB}',
    '{EE9CFA8C-F997-4221-BE2F-85A5F603218F}',
    '{F2A7EE29-8BF6-4a6d-83F1-098E366C709C}',
    '{F3685D71-1FC6-4CBD-B244-E60D8C89990B}'
)

# Ensure process is running as 64-bit
if (![Environment]::Is64BitProcess)
{
    Log("This script is designed to run in 64-bit PowerShell. Exiting.")
    Write-Error "This script is designed to run in 64-bit PowerShell. Exiting."
    return
}

Log("*** Script Start ***")

# Process Wow64 typelib node
ProcessTypelibHive("HKLM:\Software\WOW6432Node\Classes\TypeLib\")

# Process native typelib node
ProcessTypelibHive("HKLM:\Software\Classes\TypeLib\")

Log("*** Script End ***")
