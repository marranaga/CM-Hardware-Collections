$ErrorActionPreference = 'Stop'

if (-not (Get-Module -ListAvailable 'ConfigurationManager')) {
    try {
        Import-Module (Join-Path (Split-Path $ENV:SMS_ADMIN_UI_PATH -Parent) 'ConfigurationManager.psd1') -ErrorAction Stop
    } catch {
        Throw [System.Management.Automation.ItemNotFoundException] 'Failed to locate the ConfigurationManager.psd1 file'
    }
}

if (-not ($Settings = Get-Content "$PSScriptRoot\Settings.json" | ConvertFrom-Json)) {
    Throw [System.Management.Automation.ItemNotFoundException] 'Failed to locate the Settings.json file'
}

if (-not (Get-PSDrive -Name $Settings.SiteCode -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name $Settings.SiteCode -PSProvider 'CMSite' -Root $Settings.ServerAddress -Description "SCCM Site" -ErrorAction Stop
}


Push-Location "$($Settings.SiteCode):"

$RefreshSchedule = New-CMSchedule -RecurInterval Days -Recurcount 1 -Start (Get-Date)


$FolderRoot = "$($Settings.SiteCode):\DeviceCollection\$($Settings.RootFolderPath)".Trim('\')
Write-Verbose ('Folder Root: {0}' -f $FolderRoot)

if ($Settings.Computers) {
    # Getting a list of all Manufacturers
    $Manufacturers = (Invoke-CMWmiQuery -Query "select distinct SMS_G_System_PC_BIOS.Manufacturer from SMS_G_System_PC_BIOS").Manufacturer
    
    # Sanitizing HP, since they can't decide who they really are consistantly....
    $Manufacturers = $Manufacturers | Where-Object { ($_ -ne 'HP') -and $_ }
    Write-Verbose ('Manufacturers: {0}' -f ($Manufacturers | Out-String))
    
    foreach ($Manufacturer in $Manufacturers) {
        Write-Verbose '=================================================='
        Write-Verbose ('Evaluating Manufacturer: {0}' -f $Manufacturer)
        Write-Verbose '=================================================='
        
        $ManufacturerWhere = if ($Manufacturer -eq 'Hewlett-Packard') {
            foreach ($Man in @('Hewlett-Packard', 'HP')) {
                Write-Output 'SMS_G_System_COMPUTER_SYSTEM.Manufacturer = "{0}"' -f $Man
            }
        } else {
            foreach ($Man in $Manufacturer) {
                Write-Output 'SMS_G_System_COMPUTER_SYSTEM.Manufacturer = "{0}"' -f $Man
            }
        }

        $ManufacturerWhere = $ManufacturerWhere -join ' or '
        Write-Verbose ('Manufacturer Where Clause: {0}' -f $ManufacturerWhere)

        $Query = 'select distinct SMS_G_System_COMPUTER_SYSTEM.Model from SMS_G_System_COMPUTER_SYSTEM where ({0})' -f $ManufacturerWhere
        Write-Verbose ('Model Query: {0}' -f $Query)

        $Models = (Invoke-CMWmiQuery -Query $Query).Model
        Write-Verbose ('Found Models: {0}' -f ($Models -join ', '))

        if ($Models.Count -ge 1) {
            $Manufacturer = if ($Manufacturer -contains 'HP') { 'Hewlett-Packard' } else { $Manufacturer }
            Write-Verbose ('Manufacturer again {0}' -f $Manufacturer)

            $ManufacturerFolder = Join-Path $FolderRoot $Manufacturer
            Write-Verbose ('Manufacturer folder path {0}' -f $ManufacturerFolder)

            if (-not (Test-Path $ManufacturerFolder)) {
                Write-Verbose ('Manufacturer folder does not exist. Creating {0}' -f $ManufacturerFolder)
                New-Item -Path $ManufacturerFolder -ItemType Directory -Force
            }

            $ManufacturerCollectionName = $Manufacturer

            if ($Settings.Prefix) {
                $ManufacturerCollectionName = '{0}-{1}' -f $Settings.Prefix, $ManufacturerCollectionName
            }

            if (-not ($ManufacturerDeviceCollection = Get-CMDeviceCollection -Name $ManufacturerCollectionName)) {
                Write-Verbose 'Manufacturer Collection does not exist. Creating...'

                $NewCMDeviceCollection = @{
                    Name                   = $ManufacturerCollectionName
                    Comment                = ('All {0} systems' -f $Manufacturer)
                    LimitingCollectionName = 'All Systems'
                    RefreshSchedule        = $RefreshSchedule
                }

                Write-Verbose ('New-CMDeviceCollection: {0}' -f ($NewCMDeviceCollection | Out-String))
                $ManufacturerDeviceCollection = New-CMDeviceCollection @NewCMDeviceCollection
                $ManufacturerDeviceCollection | Move-CMObject -FolderPath $ManufacturerFolder
            }

            foreach ($Model in $Models) {
                $ModelCollectionName = "$Manufacturer - $Model"

                if ($Settings.Prefix) {
                    $ModelCollectionName = '{0}-{1}' -f $Settings.Prefix, $ModelCollectionName
                }

                if (-not ($ModelDeviceCollection = Get-CMDeviceCollection -Name $ModelCollectionName)) {
                    Write-Verbose 'Model Collection does not exist. Creating...'
                    $NewCMDeviceCollection = @{
                        Name                   = $ModelCollectionName
                        Comment                = ('All {0} systems of model {1}' -f $Manufacturer, $Model)
                        LimitingCollectionName = 'All Systems'
                        RefreshSchedule        = $RefreshSchedule
                    }

                    Write-Verbose ('New-CMDeviceCollection: {0}' -f ($NewCMDeviceCollection | Out-String))
                    $ModelDeviceCollection = New-CMDeviceCollection @NewCMDeviceCollection
                    $ModelDeviceCollection | Move-CMObject -FolderPath $ManufacturerFolder

                    $AddCMDeviceCollectionQueryMembershipRule = @{
                        Collection = $ModelDeviceCollection
                        QueryExpression = ('select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_COMPUTER_SYSTEM on SMS_G_System_COMPUTER_SYSTEM.ResourceID = SMS_R_System.ResourceId where ({0}) and SMS_G_System_COMPUTER_SYSTEM.Model = "{1}"' -f $ManufacturerWhere, $Model)
                        RuleName = "$Manufacturer - $Model"
                    }

                    Write-Verbose ('Add-CMDeviceCollectionQueryMembershipRule: {0}' -f ($AddCMDeviceCollectionQueryMembershipRule | Out-String))
                    Add-CMDeviceCollectionQueryMembershipRule @AddCMDeviceCollectionQueryMembershipRule

                    $AddCMDeviceCollectionIncludeMembershipRule = @{
                        CollectionName        = $ManufacturerCollectionName
                        IncludeCollectionName = $ModelCollectionName
                    }
                    Write-Verbose ('Add-CMDeviceCollectionIncludeMembershipRule: {0}' -f ($AddCMDeviceCollectionIncludeMembershipRule | Out-String))
                    Add-CMDeviceCollectionIncludeMembershipRule @AddCMDeviceCollectionIncludeMembershipRule
                }
            }
        }
    }
}

if ($Settings.VideoCards) {
    # Get all graphics card manufacturers 
    $VideoCards = (Invoke-CMWmiQuery -Query "select distinct SMS_G_SYSTEM_VIDEO_CONTROLLER.AdapterCompatibility from SMS_G_SYSTEM_VIDEO_CONTROLLER").AdapterCompatibility
    
    # Get Nvidia and AMD (previously ATI).
    $VideoCards = $VideoCards | Where-Object { ($_ -eq 'nvidia') -or ($_ -like 'ATI*') -or ($_ -like 'Advanced Micro*') -or ($_ -like 'Intel*') }
    
    foreach ($Manufacturer in $VideoCards) {
        Write-Verbose '=================================================='
        Write-Verbose ('Evaluating VideoCard Manufacturer: {0}' -f $Manufacturer)
        Write-Verbose '=================================================='

        $ManufacturerWhere = 'SMS_G_System_Video_Controller.AdapterCompatibility = "{0}"' -f $Manufacturer
        Write-Verbose ('Manufacturer Where Clause: {0}' -f $ManufacturerWhere)

        $Query = 'select distinct SMS_G_System_Video_Controller.Name from SMS_G_System_Video_Controller where ({0})' -f $ManufacturerWhere
        Write-Verbose ('Model Query: {0}' -f $Query)

        $Models = (Invoke-CMWmiQuery -Query $Query).Name
        Write-Verbose ('Found Models: {0}' -f ($Models -join ', '))

        if ($Models.Count -ge 1 ) {
            $ManufacturerFolder = Join-Path $FolderRoot $Manufacturer
            Write-Verbose ('Manufacturer folder path {0}' -f $ManufacturerFolder)

            if (-not (Test-Path $ManufacturerFolder)) {
                Write-Verbose ('Manufacturer folder does not exist. Creating {0}' -f $ManufacturerFolder)
                New-Item -Path $ManufacturerFolder -ItemType Directory -Force
            }

            $ManufacturerCollectionName = $Manufacturer

            if ($Settings.Prefix) {
                $ManufacturerCollectionName = '{0}-{1}' -f $Settings.Prefix, $ManufacturerCollectionName
            }

            if (-not ($ManufacturerDeviceCollection = Get-CMDeviceCollection -Name $ManufacturerCollectionName)) {
                Write-Verbose 'Manufacturer Collection does not exist. Creating...'

                $NewCMDeviceCollection = @{
                    Name                   = $ManufacturerCollectionName
                    Comment                = ('All {0} graphics cards' -f $Manufacturer)
                    LimitingCollectionName = 'All Systems'
                    RefreshSchedule        = $RefreshSchedule
                }

                Write-Verbose ('New-CMDeviceCollection: {0}' -f ($NewCMDeviceCollection | Out-String))
                $ManufacturerDeviceCollection = New-CMDeviceCollection @NewCMDeviceCollection
                $ManufacturerDeviceCollection | Move-CMObject -FolderPath $ManufacturerFolder
            }

            foreach ($Model in $Models) {
                # Trim manufacturer off the model for collection name.
                switch -Wildcard ($Manufacturer) {
                    'NVIDIA'    { $CModel = $Model.Trim("$Manufacturer "); Break }
                    'ATI*'      { $CModel = $Model.Trim("ATI "); Break}
                    'Advanced*' {
                        if ($Model.Contains("AMD")) {
                            $CModel = $Model.Trim("AMD ")
                            Break
                        } elseif ($Model.Contains("ATI")) {
                            $CModel = $Model.Trim("ATI ")
                            Break
                        }else {
                            $CModel = $Model
                            Break
                        }
                    }
                    'Intel*' {
                        if ($Model -match '(?i)^(Intel\(R\)\s)(.*)$'){
                            $CModel = $Model -replace '(?i)^(Intel\(R\)\s)(.*)$', '$2'
                            Break
                        }
                        if ($Model -match '(?i)^(Mobile\s)(Intel\(R\)\s)(.*)$'){
                            $CModel = $Model -replace '(?i)^(Mobile\s)(Intel\(R\)\s)(.*)$', '$1$3'
                        }
                        if ($Model -match '(?i)^(Intel\s)(.*)$') {
                            $CModel = $Model -replace '(?i)^(Intel\s)(.*)$', '$2'
                        }
                    }
                    Default { Write-Verbose "Uh-oh! You broke it now!"}
                }

                $ModelCollectionName = "$Manufacturer - $CModel"

                if ($Settings.Prefix) {
                    $ModelCollectionName = '{0}-{1}' -f $Settings.Prefix, $ModelCollectionName
                }

                if (-not ($ModelDeviceCollection = Get-CMDeviceCollection -Name $ModelCollectionName)) {
                    Write-Verbose 'Model Collection does not exist. Creating...'
                    $NewCMDeviceCollection = @{
                        Name                   = $ModelCollectionName
                        Comment                = ('All {0} systems with a {1}' -f $Manufacturer, $Model)
                        LimitingCollectionName = 'All Systems'
                        RefreshSchedule        = $RefreshSchedule
                    }

                    Write-Verbose ('New-CMDeviceCollection: {0}' -f ($NewCMDeviceCollection | Out-String))
                    $ModelDeviceCollection = New-CMDeviceCollection @NewCMDeviceCollection
                    $ModelDeviceCollection | Move-CMObject -FolderPath $ManufacturerFolder

                    $AddCMDeviceCollectionQueryMembershipRule = @{
                        Collection = $ModelDeviceCollection
                        QueryExpression = ('select SMS_R_SYSTEM.ResourceID,SMS_R_SYSTEM.ResourceType,SMS_R_SYSTEM.Name,SMS_R_SYSTEM.SMSUniqueIdentifier,SMS_R_SYSTEM.ResourceDomainORWorkgroup,SMS_R_SYSTEM.Client from SMS_R_System inner join SMS_G_System_VIDEO_CONTROLLER on SMS_G_System_VIDEO_CONTROLLER.ResourceId = SMS_R_System.ResourceId where ({0}) and SMS_G_System_VIDEO_CONTROLLER.Name = "{1}"' -f $ManufacturerWhere, $Model)
                        RuleName = "$Manufacturer - $Model"
                    }

                    Write-Verbose ('Add-CMDeviceCollectionQueryMembershipRule: {0}' -f ($AddCMDeviceCollectionQueryMembershipRule | Out-String))
                    Add-CMDeviceCollectionQueryMembershipRule @AddCMDeviceCollectionQueryMembershipRule

                    $AddCMDeviceCollectionIncludeMembershipRule = @{
                        CollectionName        = $ManufacturerCollectionName
                        IncludeCollectionName = $ModelCollectionName
                    }
                    Write-Verbose ('Add-CMDeviceCollectionIncludeMembershipRule: {0}' -f ($AddCMDeviceCollectionIncludeMembershipRule | Out-String))
                    Add-CMDeviceCollectionIncludeMembershipRule @AddCMDeviceCollectionIncludeMembershipRule
                }
            }
        }
    }
}

Pop-Location