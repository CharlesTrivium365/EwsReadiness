<#
.SYNOPSIS
    Diagnostic de préparation à la fin d'Exchange Web Services (EWS) dans Exchange Online.

.DESCRIPTION
    Outil en LECTURE SEULE : il ne modifie aucun paramètre du locataire.

    Il rassemble trois sources d'information :
      1. La configuration du locataire : EwsEnabled et la liste d'autorisation EWSAllowedAppIDs.
      2. Le rapport « Utilisation d'EWS » exporté en CSV depuis le Centre d'administration Microsoft 365
         (Rapports > Utilisation > Exchange > onglet EWS). Facultatif mais recommandé.
      3. Les applications Entra qui détiennent une permission EWS sur Office 365 Exchange Online
         (full_access_as_app en application, EWS.AccessAsUser.All en délégué).

    Il produit un tableau, un CSV et un rapport HTML qui classent chaque application :
      - À AJOUTER : utilisée mais absente de la liste, elle sera bloquée.
      - À VALIDER : autorisée mais non vue dans le rapport d'utilisation.
      - PERMISSION SEULEMENT : détient une permission EWS sans usage ni autorisation.
      - OK : utilisée et autorisée.

.PARAMETER UsageReportCsv
    Chemin de l'export CSV du rapport d'utilisation d'EWS.

.PARAMETER OutputFolder
    Dossier où écrire le CSV et le rapport HTML. Par défaut : le dossier courant.

.PARAMETER SkipConnect
    N'ouvre pas de nouvelles connexions (utile si Connect-ExchangeOnline et Connect-MgGraph sont déjà faits).

.EXAMPLE
    .\EwsReadiness.ps1 -UsageReportCsv .\EWSUsage.csv

.NOTES
    Prérequis : modules ExchangeOnlineManagement, Microsoft.Graph.Applications et
    Microsoft.Graph.Identity.SignIns ; un rôle Exchange en lecture ; les autorisations Graph
    Application.Read.All et Directory.Read.All.
    Projet : https://github.com/CharlesTrivium365/EwsReadiness
#>
[CmdletBinding()]
param(
    [ValidateScript({ Test-Path -Path $_ -PathType Leaf })]
    [string]$UsageReportCsv,

    [ValidateScript({ Test-Path -Path $_ -PathType Container })]
    [string]$OutputFolder = (Get-Location).Path,

    [switch]$SkipConnect
)

#Requires -Modules ExchangeOnlineManagement, Microsoft.Graph.Applications, Microsoft.Graph.Identity.SignIns

$ExchangeOnlineAppId = '00000002-0000-0ff1-ce00-000000000000'

function Get-AppIdList {
    param([string]$Text)
    # Seuls les GUID valides sont conservés : évite les valeurs vides et protège les filtres Graph.
    $guid = [guid]::Empty
    @($Text -split '[,;\s]+' | Where-Object { [guid]::TryParse($_, [ref]$guid) } | ForEach-Object { $_.ToLower() } | Sort-Object -Unique)
}

function ConvertTo-HtmlText {
    param([string]$Text)
    [System.Net.WebUtility]::HtmlEncode($Text)
}

if (-not $SkipConnect) {
    Connect-ExchangeOnline -ShowBanner:$false
    Connect-MgGraph -Scopes 'Application.Read.All', 'Directory.Read.All' -NoWelcome
}

# 1. Configuration du locataire
Write-Verbose 'Lecture de la configuration EWS du locataire'
$org = Get-OrganizationConfig
$allowed = Get-AppIdList -Text ((Get-OrganizationConfig -RetrieveEwsOperationAccessPolicy).EwsAllowedAppIDs -join ',')

$diagnostic = switch ($org.EwsEnabled) {
    $null { 'EwsEnabled est vide (valeur par défaut). Microsoft le passera à False lors de la phase 2, avec un avertissement 7 jours avant dans le Centre de messages.' }
    $false { 'EwsEnabled vaut False : tout appel EWS est bloqué.' }
    $true {
        if ($allowed.Count -eq 0) { 'EwsEnabled vaut True SANS liste d''autorisation : depuis le 10 octobre 2026, tout EWS est bloqué.' }
        else { "EwsEnabled vaut True avec $($allowed.Count) application(s) autorisée(s)." }
    }
}
Write-Warning $diagnostic

# 2. Rapport d'utilisation
$used = @()
if ($UsageReportCsv) {
    $rows = @(Import-Csv -Path $UsageReportCsv)
    if ($rows.Count -gt 0) {
        # Le nom de la colonne varie selon la langue de l'export : on cherche celle qui contient l'AppID.
        $column = $rows[0].PSObject.Properties.Name | Where-Object { $_ -match 'app.*id|id.*app' } | Select-Object -First 1
        if (-not $column) { throw "Colonne d'AppID introuvable dans $UsageReportCsv." }
        $used = Get-AppIdList -Text (($rows | ForEach-Object { $_.$column }) -join ',')
    }
}

# 3. Applications qui détiennent une permission EWS
Write-Verbose 'Recherche des permissions EWS accordées'
$exchange = Get-MgServicePrincipal -Filter "appId eq '$ExchangeOnlineAppId'"
$fullAccessRole = $exchange.AppRoles | Where-Object { $_.Value -eq 'full_access_as_app' }
$withPermission = @{}

Get-MgServicePrincipalAppRoleAssignedTo -ServicePrincipalId $exchange.Id -All |
    Where-Object { $_.AppRoleId -eq $fullAccessRole.Id } |
    ForEach-Object { $withPermission[$_.PrincipalId] = 'full_access_as_app (application)' }

Get-MgOauth2PermissionGrant -Filter "resourceId eq '$($exchange.Id)'" -All |
    Where-Object { $_.Scope -split ' ' -contains 'EWS.AccessAsUser.All' } |
    ForEach-Object { if (-not $withPermission[$_.ClientId]) { $withPermission[$_.ClientId] = 'EWS.AccessAsUser.All (délégué)' } }

$permissionAppIds = @{}
foreach ($objectId in $withPermission.Keys) {
    $sp = Get-MgServicePrincipal -ServicePrincipalId $objectId -Property AppId -ErrorAction SilentlyContinue
    if ($sp) { $permissionAppIds[$sp.AppId.ToLower()] = $withPermission[$objectId] }
}

# 4. Croisement
$allIds = @($allowed + $used + @($permissionAppIds.Keys)) | Sort-Object -Unique
$results = foreach ($appId in $allIds) {
    $sp = Get-MgServicePrincipal -Filter "appId eq '$appId'" -Property DisplayName, PublisherName -ErrorAction SilentlyContinue
    $inList = $allowed -contains $appId
    $inUse = $used -contains $appId

    $status = if ($inUse -and -not $inList) { 'À AJOUTER' }
              elseif ($inList -and -not $inUse -and $UsageReportCsv) { 'À VALIDER' }
              elseif (-not $inList -and -not $inUse) { 'PERMISSION SEULEMENT' }
              else { 'OK' }

    [pscustomobject]@{
        Statut      = $status
        Application = if ($sp) { $sp.DisplayName } else { '(inconnue dans le locataire)' }
        Editeur     = $sp.PublisherName
        AppId       = $appId
        Autorisee   = $inList
        Utilisee    = $inUse
        Permission  = $permissionAppIds[$appId]
    }
}
$results = @($results | Sort-Object Statut, Application)

# 5. Sorties
$stamp = Get-Date -Format 'yyyyMMdd-HHmm'
$csvPath = Join-Path $OutputFolder "EwsReadiness-$stamp.csv"
$htmlPath = Join-Path $OutputFolder "EwsReadiness-$stamp.html"
$results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

$colors = @{ 'À AJOUTER' = '#c62828'; 'À VALIDER' = '#ef6c00'; 'PERMISSION SEULEMENT' = '#6a1b9a'; 'OK' = '#2e7d32' }
$tableRows = foreach ($r in $results) {
    "<tr><td style='color:$($colors[$r.Statut]);font-weight:600'>$($r.Statut)</td><td>$(ConvertTo-HtmlText $r.Application)</td><td>$(ConvertTo-HtmlText $r.Editeur)</td><td><code>$($r.AppId)</code></td><td>$($r.Autorisee)</td><td>$($r.Utilisee)</td><td>$(ConvertTo-HtmlText $r.Permission)</td></tr>"
}
$counts = foreach ($s in $colors.Keys) { "<li><strong>$s</strong> : $(@($results | Where-Object Statut -eq $s).Count)</li>" }

@"
<!doctype html>
<html lang="fr"><head><meta charset="utf-8"><title>Diagnostic EWS</title>
<style>body{font-family:Segoe UI,Arial,sans-serif;margin:2rem;color:#1a1a2e}table{border-collapse:collapse;width:100%}th,td{border:1px solid #ddd;padding:6px 10px;text-align:left;font-size:14px}th{background:#1a1446;color:#fff}.diag{background:#fff4e5;border-left:4px solid #ef6c00;padding:10px 14px}</style>
</head><body>
<h1>Diagnostic de préparation à la fin d'EWS</h1>
<p>Locataire : <strong>$(ConvertTo-HtmlText $org.DisplayName)</strong> · Généré le $(Get-Date -Format 'yyyy-MM-dd HH:mm')</p>
<p class="diag">$(ConvertTo-HtmlText $diagnostic)</p>
<ul>$($counts -join '')</ul>
<table><tr><th>Statut</th><th>Application</th><th>Éditeur</th><th>AppId</th><th>Autorisée</th><th>Utilisée</th><th>Permission EWS</th></tr>
$($tableRows -join "`n")
</table>
<p>Rapport d'utilisation fourni : $([bool]$UsageReportCsv). Sans ce rapport, les applications utilisées mais non autorisées ne peuvent pas être détectées.</p>
</body></html>
"@ | Set-Content -Path $htmlPath -Encoding UTF8

$results | Format-Table -AutoSize
Write-Information "CSV : $csvPath" -InformationAction Continue
Write-Information "Rapport HTML : $htmlPath" -InformationAction Continue
