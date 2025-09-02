# Install-Module -Name Microsoft.Entra -Verbose -AllowClobber
# Set-ExecutionPolicy -ExecutionPolicy RemoteSigned

# ====================================================================
# Konfiguration
# ====================================================================

$UserIdentifier   = ""      # UPN oder Mailadresse
$CreateHtmlReport = $true   # $true = HTML-Report erzeugen, $false = nur Konsole

Connect-Entra -Scopes "User.Read.All","Group.Read.All","GroupMember.Read.All"

function Write-Log {
    param([string]$Message)
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Write-Host "[$ts] $Message"
}

try {
    Write-Log "Pruefe Entra-Verbindung"
    try {
        [void](Get-EntraUser -Top 1 -ErrorAction Stop)
        Write-Log "Entra-Verbindung ok"
    } catch {
        Write-Log "Nicht mit Entra verbunden. Bitte Connect-Entra ausfuehren und erneut starten."
        throw "Verbindung zu Entra erforderlich."
    }

    Write-Log "Suche Benutzer: $UserIdentifier"
    $user = Get-EntraUser -Filter "userPrincipalName eq '$UserIdentifier'" -ErrorAction SilentlyContinue
    if (-not $user) {
        $user = Get-EntraUser -Filter "mail eq '$UserIdentifier'" -ErrorAction SilentlyContinue
    }
    if (-not $user) {
        throw "Kein Benutzer gefunden fuer $UserIdentifier."
    }
    Write-Log "Gefundener Benutzer: $($user.DisplayName) [$($user.Id)] Typ=$($user.UserType)"

    Write-Log "Lade Gruppenliste"
    $groups = Get-EntraGroup -All -ErrorAction Stop
    Write-Log "Anzahl gefundener Gruppen: $($groups.Count)"

    $foundGroups = New-Object System.Collections.Generic.List[object]

    foreach ($group in $groups) {
        try {
            $members = Get-EntraGroupMember -GroupId $group.Id -All -ErrorAction Stop
            if ($members.Id -contains $user.Id) {
                $foundGroups.Add([pscustomobject]@{
                    GroupDisplayName = $group.DisplayName
                    GroupId          = $group.Id
                    GroupType        = $group.GroupTypes -join ","
                }) | Out-Null
                Write-Log "Treffer: Benutzer ist Mitglied in Gruppe '$($group.DisplayName)'"
            }
        } catch {
            Write-Log "Fehler beim Laden der Mitglieder von Gruppe '$($group.DisplayName)': $($_.Exception.Message)"
        }
    }

    Write-Log "Mitgliedschaften gefunden: $($foundGroups.Count)"

    # Ausgabe in der Konsole
    $foundGroups | Sort-Object GroupDisplayName | Format-Table -AutoSize

    if ($CreateHtmlReport) {
        Write-Log "Erzeuge HTML-Report"
        $desktop = [Environment]::GetFolderPath('Desktop')
        $safeId  = ($UserIdentifier -replace '[^a-zA-Z0-9._-]', '_')
        $stamp   = (Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")
        $file    = Join-Path $desktop ("EntraUserGroups_{0}_{1}.html" -f $safeId, $stamp)

        # HTML Inhalt
        $header = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; font-size: 13px; color: #222; }
h1 { font-size: 18px; margin-bottom: 6px; }
h2 { font-size: 15px; margin-top: 18px; }
.summary { background: #f5f5f5; padding: 10px; border: 1px solid #ddd; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #ddd; padding: 6px; text-align: left; }
th { background: #fafafa; }
small { color: #666; }
</style>
"@

        $summary = [pscustomobject]@{
            Benutzeranzeige   = $user.DisplayName
            BenutzerId        = $user.Id
            UserPrincipalName = $user.UserPrincipalName
            Mail              = $user.Mail
            UserType          = $user.UserType
            Gruppenanzahl     = $foundGroups.Count
            BerichtZeitpunkt  = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        }

        $summaryHtml = $summary | ConvertTo-Html -As List -Fragment
        $groupsHtml  = ($foundGroups | Sort-Object GroupDisplayName) | ConvertTo-Html -Property GroupDisplayName, GroupId, GroupType -Fragment

        $html = @"
<html>
<head>
<meta charset="utf-8">
<title>Entra Gruppenmitgliedschaften - $($user.UserPrincipalName)</title>
$header
</head>
<body>
<h1>Entra Gruppenmitgliedschaften</h1>
<div class="summary">
$summaryHtml
</div>
<h2>Gruppen</h2>
$groupsHtml
<br>
<small>Generiert am $((Get-Date).ToString("yyyy-MM-dd HH:mm:ss")) von $($ENV:USERNAME)</small>
</body>
</html>
"@

        $html | Out-File -FilePath $file -Encoding UTF8
        Write-Log "HTML-Report gespeichert: $file"
        Start-Process $file
    } else {
        Write-Log "HTML-Report deaktiviert (Variable CreateHtmlReport=$CreateHtmlReport)."
    }
}
catch {
    $msg = $_.Exception.Message
    Write-Log "Abbruch: $msg"
    throw
}

Disconnect-Entra
