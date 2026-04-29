using namespace System.Net

param($Request, $TriggerMetadata)

$body = $Request.Body

$HubName = $body.HubName
$BusinessArea = $body.BusinessArea
$HubOwner = $body.HubOwner
$SensitivityLabel = $body.SensitivityLabel
$SelectedPages = $body.SelectedPages
$SelectedLibraries = $body.SelectedLibraries
$Theme = $body.Theme
$CustomPagesJson = $body.CustomPagesJson

$clientId = $env:ClientId
$tenantId = $env:TenantId
$tenantUrl = "https://onesharepoint.sharepoint.com"
$certPath = "C:\home\site\wwwroot\Provision-Hub\onesharepoint.pfx"
$certPassword = $env:CertificatePassword

$validThemes = @("CobaltTeal", "MossPistachio", "RosewoodPoppy", "GoldMustard")

$siteAlias = $HubName -replace '[^a-zA-Z0-9\s]', '' -replace '\s+', '-'
$siteAlias = $siteAlias.ToLower().Trim('-')
$siteAlias = $siteAlias.ToLower().Trim('-')
$siteUrl = "$tenantUrl/sites/$siteAlias"

if ($Theme -notin $validThemes) { $Theme = "CobaltTeal" }

# --- CONNECT ---
Write-Host "Connecting to SharePoint..."
Connect-PnPOnline -Url $tenantUrl -ClientId $clientId -Tenant "$tenantId" -CertificatePath $certPath -CertificatePassword (ConvertTo-SecureString $certPassword -AsPlainText -Force)

# --- CREATE SITE ---
Write-Host "Creating site: $HubName..."
New-PnPSite -Type CommunicationSite -Title $HubName -Url $siteUrl -Owner $HubOwner
Start-Sleep -Seconds 15

Write-Host "Reconnecting to new site..."
Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant "$tenantId" -CertificatePath $certPath -CertificatePassword (ConvertTo-SecureString $certPassword -AsPlainText -Force)
Write-Host "Site created: $siteUrl"

# --- APPLY THEME ---
Set-PnPWebTheme -Theme $Theme
Write-Host "Theme applied: $Theme"

# --- PAGE DEFINITIONS ---
$pageDefinitions = @{
    "About" = @{ SubPages = @("Our Team", "Our Mission"); Description = "Welcome to the About section. This area provides an overview of the department, its people, and its purpose." }
    "Resources" = @{ SubPages = @("Key Documents", "Useful Links"); Description = "Welcome to Resources. Find key documents, reference materials, and useful links for your team." }
    "News" = @{ SubPages = @(); Description = "Welcome to the News page. Department news and announcements will be posted here." }
    "Contact" = @{ SubPages = @(); Description = "Welcome to the Contact page. Find contact details and enquiry information for the department." }
    "Projects" = @{ SubPages = @("Current Projects", "Completed Projects"); Description = "Welcome to Projects. Track and showcase current and completed department projects here." }
}

$subPageDescriptions = @{
    "Our Team" = "This page will introduce the team members within the department."
    "Our Mission" = "This page outlines the department's mission, vision, and values."
    "Key Documents" = "This page provides access to key documents for the department."
    "Useful Links" = "This page contains useful links and external resources."
    "Current Projects" = "This page tracks all active and ongoing projects."
    "Completed Projects" = "This page showcases completed projects and their outcomes."
}

$libraryDefinitions = @{
    "General Documents" = @{ Folders = @("Drafts", "Final", "Archive") }
    "Policies & Procedures" = @{ Folders = @("Active Policies", "Under Review", "Retired") }
    "Templates" = @{ Folders = @() }
    "Meeting Notes" = @{ Folders = @("Weekly Standups", "Leadership Meetings") }
    "Training & Learning" = @{ Folders = @("How-To Guides", "Reference Materials") }
    "Project Files" = @{ Folders = @() }
}

# --- CREATE PREDEFINED PAGES ---
Write-Host "Creating predefined pages..."
$selectedPageList = $SelectedPages -split ','
foreach ($pageName in $selectedPageList) {
    $pageName = $pageName.Trim()
    if ($pageDefinitions.ContainsKey($pageName)) {
        Add-PnPPage -Name $pageName -LayoutType Article
        Set-PnPPage -Identity "$pageName.aspx" -Title $pageName
        Add-PnPPageTextPart -Page "$pageName.aspx" -Text "<p>$($pageDefinitions[$pageName].Description)</p>"
        Set-PnPPage -Identity "$pageName.aspx" -Publish
        Write-Host "  Created page: $pageName"
        foreach ($subPage in $pageDefinitions[$pageName].SubPages) {
            Add-PnPPage -Name $subPage -LayoutType Article
            if ($subPageDescriptions.ContainsKey($subPage)) {
                Set-PnPPage -Identity "$subPage.aspx" -Title $subPage
                Add-PnPPageTextPart -Page "$subPage.aspx" -Text "<p>$($subPageDescriptions[$subPage])</p>"
                Set-PnPPage -Identity "$subPage.aspx" -Publish
            }
            Write-Host "    Created sub-page: $subPage"
        }
    }
}

# --- CREATE CUSTOM PAGES ---
if ($CustomPagesJson -and $CustomPagesJson -ne "") {
    Write-Host "Creating custom pages..."
    try {
        $customPages = $CustomPagesJson | ConvertFrom-Json

        foreach ($customPage in $customPages) {
            $cpName = $customPage.name.Trim()
            $cpNameClean = $cpName -replace '[^a-zA-Z0-9\s-]', ''

            if ($cpNameClean -ne "") {
                Add-PnPPage -Name $cpNameClean -LayoutType Article
                Set-PnPPage -Identity "$cpNameClean.aspx" -Title $cpNameClean
                Add-PnPPageTextPart -Page "$cpNameClean.aspx" -Text "<p>Welcome to the $cpNameClean page. This page was created as a custom addition to the hub.</p>"
                Set-PnPPage -Identity "$cpNameClean.aspx" -Publish
                Write-Host "  Created custom page: $cpNameClean"

                if ($customPage.subPages -and $customPage.subPages.Count -gt 0) {
                    foreach ($csp in $customPage.subPages) {
                        $cspClean = $csp.Trim() -replace '[^a-zA-Z0-9\s-]', ''
                        if ($cspClean -ne "") {
                            Add-PnPPage -Name $cspClean -LayoutType Article
                            Set-PnPPage -Identity "$cspClean.aspx" -Title $cspClean
                            Add-PnPPageTextPart -Page "$cspClean.aspx" -Text "<p>Welcome to the $cspClean page.</p>"
                            Set-PnPPage -Identity "$cspClean.aspx" -Publish
                            Write-Host "    Created custom sub-page: $cspClean"
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-Host "Warning: Could not parse custom pages JSON. Skipping. Error: $($_.Exception.Message)"
    }
}

# --- CREATE LIBRARIES ---
Write-Host "Creating document libraries..."
$selectedLibList = $SelectedLibraries -split ','
$allLibraries = @("General Documents") + $selectedLibList
$createdLibraryInfo = @()

foreach ($libName in $allLibraries) {
    $libName = $libName.Trim()
    if ($libraryDefinitions.ContainsKey($libName)) {
        New-PnPList -Title $libName -Template DocumentLibrary
        $libServerRelUrl = (Get-PnPList -Identity $libName).RootFolder.ServerRelativeUrl
        $createdLibraryInfo += @{ Name = $libName; Url = $libServerRelUrl }
        foreach ($folder in $libraryDefinitions[$libName].Folders) {
            Add-PnPFolder -Name $folder -Folder $libServerRelUrl
            Write-Host "    Created folder: $folder"
        }
        Write-Host "  Created library: $libName"
    }
}

# --- BUILD NAVIGATION ---
Write-Host "Building navigation..."
Get-PnPNavigationNode -Location QuickLaunch | Remove-PnPNavigationNode -Force
Add-PnPNavigationNode -Title "Home" -Url "/sites/$siteAlias/SitePages/Home.aspx" -Location QuickLaunch

foreach ($pageName in $selectedPageList) {
    $pageName = $pageName.Trim()
    if ($pageDefinitions.ContainsKey($pageName)) {
        $navNode = Add-PnPNavigationNode -Title $pageName -Url "/sites/$siteAlias/SitePages/$pageName.aspx" -Location QuickLaunch
        foreach ($subPage in $pageDefinitions[$pageName].SubPages) {
            $subPageUrl = $subPage -replace '\s+', '%20'
            Add-PnPNavigationNode -Title $subPage -Url "/sites/$siteAlias/SitePages/$subPageUrl.aspx" -Location QuickLaunch -Parent $navNode.Id
        }
    }
}

# Add custom pages to navigation
if ($CustomPagesJson -and $CustomPagesJson -ne "") {
    try {
        $customPages = $CustomPagesJson | ConvertFrom-Json

        foreach ($customPage in $customPages) {
            $cpName = $customPage.name.Trim()
            $cpNameClean = $cpName -replace '[^a-zA-Z0-9\s-]', ''

            if ($cpNameClean -ne "") {
                $cpUrl = $cpNameClean -replace '\s+', '%20'
                $navNode = Add-PnPNavigationNode -Title $cpNameClean -Url "/sites/$siteAlias/SitePages/$cpUrl.aspx" -Location QuickLaunch

                if ($customPage.subPages -and $customPage.subPages.Count -gt 0) {
                    foreach ($csp in $customPage.subPages) {
                        $cspClean = $csp.Trim() -replace '[^a-zA-Z0-9\s-]', ''
                        if ($cspClean -ne "") {
                            $cspUrl = $cspClean -replace '\s+', '%20'
                            Add-PnPNavigationNode -Title $cspClean -Url "/sites/$siteAlias/SitePages/$cspUrl.aspx" -Location QuickLaunch -Parent $navNode.Id
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-Host "Warning: Could not parse custom pages JSON for navigation."
    }
}

# Documents navigation group
$docsNavNode = Add-PnPNavigationNode -Title "Documents" -Url "/sites/$siteAlias/_layouts/15/viewlsts.aspx?view=14" -Location QuickLaunch
foreach ($libInfo in $createdLibraryInfo) {
    Add-PnPNavigationNode -Title $libInfo.Name -Url $libInfo.Url -Location QuickLaunch -Parent $docsNavNode.Id
}

# --- RETURN RESPONSE ---
$responseBody = @{
    status = "success"
    siteUrl = $siteUrl
    hubName = $HubName
    theme = $Theme
    owner = $HubOwner
} | ConvertTo-Json

Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $responseBody
    ContentType = "application/json"
})