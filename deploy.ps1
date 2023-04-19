# TODO: Move variables to config file
# TODO: Add error handling
# TODO: Fix error: Invalid JWT token. The token is expired 

# ENV variables
$domain = ""
$clientId = ""

# TODO: Improve config structure and make it more configurable
# Config
$hubSitesNumber = 2
$assotiatedSitesNumber = 25
$sitePagesNumber = 200
$subWebsNumber = 2
$subWebPagesNumber = 5

$hubSiteBaseName = "Hub Site"
$hubSiteBaseUrlName = "HubSiteDev"
$assSiteBaseName = "Associated Site"
$assSiteBaseUrlName = "AssociatedSiteDev"
$sitePageBaseName = "Site Page"
$subWebBaseName = "Sub Web"
$subWebBaseUrlName = "SubWebDev"
$subWebPageBaseName = "Sub Web Page"

# TODO: Enable ability to create sub webs(disabled by default)
$enableSubWebCreation = $false

function Get-Site {
    param ([string] $siteUrl, [string] $token)
    
    Connect-PnPOnline $siteUrl -AccessToken $token
    $site = Get-PnPSite -Includes Id, Url
    return $site
}

if ([string]::IsNullOrEmpty($clientId)) {
    $adApp = Register-PnPAzureADApp `
        -ApplicationName "PnP PowerShell SP Page Generator" `
        -Tenant "$($domain).onmicrosoft.com" `
        -SharePointDelegatePermissions Sites.FullControl.All `
        -DeviceLogin

    $clientId = $adApp.'AzureAppId/ClientId'
}


Write-Host "Start authentification"
Connect-PnPOnline "https://$domain.sharepoint.com" -Tenant "$domain.onmicrosoft.com" -ClientId $clientId -DeviceLogin
$accessToken = Get-PnPAppAuthAccessToken
Connect-PnPOnline "https://$domain-admin.sharepoint.com" -Tenant "$domain.onmicrosoft.com" -ClientId $clientId -DeviceLogin
$adminAccessToken = Get-PnPAppAuthAccessToken
Write-Host "Authentificated successfully"

$PSStyle.Progress.View = 'Classic'
for ($i = 0; $i -lt $hubSitesNumber; $i++) {
    $percentCompleteHubSites = $i / $hubSitesNumber * 100
    Write-Progress -Activity "Creating hub sites..." -Status "$([math]::floor($percentCompleteHubSites))% Complete:" -PercentComplete $percentCompleteHubSites -CurrentOperation "Creating hub sites" -ID 1

    $hubSiteUrl = "https://$domain.sharepoint.com/sites/$hubSiteBaseUrlName-$($i + 1)"
    $hubSiteName = "$hubSiteBaseName $($i + 1)"
    Connect-PnPOnline "https://$($domain)-admin.sharepoint.com" -AccessToken $adminAccessToken
    # You can add -Wait parameter to wait until all resources will be provisioned
    New-PnPSite -Type CommunicationSite -Title $hubSiteName -Url $hubSiteUrl | Out-Null
    Register-PnPHubSite -Site $hubSiteUrl | Out-Null
    $hubSite = Get-Site -siteUrl $hubSiteUrl -token $accessToken

    for ($j = 0; $j -lt $assotiatedSitesNumber; $j++) {
        $percentCompleteAssSites = $j / $assotiatedSitesNumber * 100
        Write-Progress -Activity "Creating associated sites..." -Status "$([math]::floor($percentCompleteAssSites))% Complete:" -PercentComplete $percentCompleteAssSites -CurrentOperation "Creating associated sites..." -ID 2

        $assSiteUrl = "https://$domain.sharepoint.com/sites/$assSiteBaseUrlName-$($j + 1)"
        $assSiteName = "$assSiteBaseName $($j + 1)"
        # You can add -Wait parameter to wait until all resources will be provisioned
        New-PnPSite -Type CommunicationSite -Title $assSiteName -Url $assSiteUrl -HubSiteId $hubSite.Id | Out-Null
        $site = Get-Site -siteUrl $assSiteUrl -token $accessToken
        
        if ($enableSubWebCreation) {
            for ($k = 0; $k -lt $subWebsNumber; $k++) {
                $percentCompleteSubWebs = $k / $subWebsNumber * 100
                Write-Progress -Activity "Creating sub webs..." -Status "$([math]::floor($percentCompleteSubWebs))% Complete:" -PercentComplete $percentCompleteSubWebs -CurrentOperation "Creating sub webs..." -ID 3
    
                $subWebUrl = "$($site.Url)/$subWebBaseUrlName-$($k + 1)"
                $subWebName = "$subWebBaseName $($k + 1)"
                $subWeb = New-PnPWeb -Title $subWebName -Url $subWebUrl -Template "STS#0"
    
                for ($f = 0; $f -lt $subWebPagesNumber; $f++) {
                    $percentCompleteSubWebPages = $f / $subWebPagesNumber * 100
                    Write-Progress -Activity "Creating sub web pages..." -Status "$([math]::floor($percentCompleteSubWebPages))% Complete:" -PercentComplete $percentCompleteSubWebPages -CurrentOperation "Creating sub web pages..." -ID 4
    
                    Connect-PnPOnline -Url $subWeb.'WebUrl' -AccessToken $accessToken
                    $subWebPageName = "$subWebPageBaseName $($f + 1)"
                    Add-PnPPage -Name $subWebPageName -CommentsEnabled -Publish | Out-Null
                }
            }
        }

        for ($l = 0; $l -lt $sitePagesNumber; $l++) {
            $percentCompletSitePages = $l / $sitePagesNumber * 100
            Write-Progress -Activity "Creating site pages..." -Status "$([math]::floor($percentCompletSitePages))% Complete:" -PercentComplete $percentCompletSitePages -CurrentOperation "Creating site pages..." -ID 5

            Connect-PnPOnline -Url $site.Url -AccessToken $accessToken
            $sitePageName = "$sitePageBaseName $($l + 1)"
            Add-PnPPage -Name $sitePageName -CommentsEnabled -Publish | Out-Null
        }
    }
}
