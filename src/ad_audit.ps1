param(
  [string]$InputPath = ".\data\ad_export.json",
  [string]$OutDir = ".\output",
  [int]$InactiveDays = 30,
  [int]$PasswordOldDays = 90
)

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

$data = Get-Content -Path $InputPath -Raw | ConvertFrom-Json

# Använd export_date som "nu" för stabila resultat
$now = [datetime]$data.export_date

function Safe-Date([string]$s) {
  try { return [datetime]$s } catch { return $null }
}

# Inaktiva användare
$inactiveUsers = @()
foreach ($u in $data.users) {
  $last = Safe-Date $u.lastLogon
  if ($last -ne $null) {
    $days = ($now - $last).Days
    if ($days -ge $InactiveDays) {
      $inactiveUsers += [pscustomobject]@{
        SamAccountName = $u.samAccountName
        DisplayName    = $u.displayName
        Department     = $u.department
        Site           = $u.site
        LastLogon      = $u.lastLogon
        DaysInactive   = $days
        Enabled        = $u.enabled
        AccountExpires = $u.accountExpires
      }
    }
  }
}

# Konton som löper ut inom 30 dagar
$expiringSoon = @()
foreach ($u in $data.users) {
  $exp = Safe-Date $u.accountExpires
  if ($exp -ne $null) {
    $daysLeft = ($exp - $now).Days
    if ($daysLeft -ge 0 -and $daysLeft -le 30) {
      $expiringSoon += [pscustomobject]@{
        SamAccountName = $u.samAccountName
        DisplayName    = $u.displayName
        AccountExpires = $u.accountExpires
        DaysLeft       = $daysLeft
        Enabled        = $u.enabled
      }
    }
  }
}

# Lösenordsålder
$pwdOld = @()
foreach ($u in $data.users) {
  $pls = Safe-Date $u.passwordLastSet
  if ($pls -ne $null -and (-not $u.passwordNeverExpires)) {
    $age = ($now - $pls).Days
    if ($age -ge $PasswordOldDays) {
      $pwdOld += [pscustomobject]@{
        SamAccountName = $u.samAccountName
        DisplayName    = $u.displayName
        Site           = $u.site
        PasswordLastSet= $u.passwordLastSet
        PasswordAgeDays= $age
      }
    }
  }
}

# Datorer ej sedda på 30+ dagar
$inactiveComputers = @()
foreach ($c in $data.computers) {
  $last = Safe-Date $c.lastLogon
  if ($last -ne $null) {
    $days = ($now - $last).Days
    if ($days -ge $InactiveDays) {
      $inactiveComputers += [pscustomobject]@{
        Name          = $c.name
        Site          = $c.site
        OperatingSystem = $c.operatingSystem
        LastLogon     = $c.lastLogon
        DaysInactive  = $days
        Enabled       = $c.enabled
      }
    }
  }
}

# Users per department (enkel loop)
$deptCounts = @{}
foreach ($u in $data.users) {
  $d = $u.department
  if (-not $deptCounts.ContainsKey($d)) { $deptCounts[$d] = 0 }
  $deptCounts[$d]++
}

# Computers per site
$compBySite = $data.computers | Group-Object -Property site | Sort-Object Name

# OS-översikt
$osCounts = $data.computers | Group-Object operatingSystem | Sort-Object Count -Descending

# Export CSV (G-krav)
$inactiveUsers | Sort-Object DaysInactive -Descending |
  Export-Csv -Path (Join-Path $OutDir "inactive_users.csv") -NoTypeInformation -Encoding UTF8

# Extra: computer_status.csv (nyttig, ofta uppskattat)
$computerStatus = @()
foreach ($g in $compBySite) {
  $total = $g.Count
  $win10 = ($g.Group | Where-Object { $_.operatingSystem -like "Windows 10*" }).Count
  $win11 = ($g.Group | Where-Object { $_.operatingSystem -like "Windows 11*" }).Count
  $server = ($g.Group | Where-Object { $_.operatingSystem -like "Windows Server*" }).Count
  $inactive = ($g.Group | ForEach-Object {
    $last = Safe-Date $_.lastLogon
    if ($last -ne $null) { (($now - $last).Days -ge $InactiveDays) } else { $false }
  } | Where-Object { $_ -eq $true }).Count

  $computerStatus += [pscustomobject]@{
    Site = $g.Name
    TotalComputers = $total
    InactiveComputers = $inactive
    Windows10Count = $win10
    Windows11Count = $win11
    WindowsServerCount = $server
  }
}
$computerStatus | Export-Csv -Path (Join-Path $OutDir "computer_status.csv") -NoTypeInformation -Encoding UTF8

# Rapport (VG-stil, men fungerar för G också)
$report = @()
$report += "=" * 80
$report += "ACTIVE DIRECTORY AUDIT REPORT".PadLeft(52).PadRight(80)
$report += "=" * 80
$report += "Generated: $($now.ToString('yyyy-MM-dd HH:mm:ss'))"
$report += "Domain: $($data.domain)"
$report += "Export Date: $($data.export_date)"
$report += ""
$report += "EXECUTIVE SUMMARY"
$report += "-" * 80
$report += "CRITICAL: $($expiringSoon.Count) accounts expiring within 30 days"
$report += "WARNING:  $($inactiveUsers.Count) users haven't logged in for $InactiveDays+ days"
$report += "WARNING:  $($inactiveComputers.Count) computers not seen for $InactiveDays+ days"
$report += "SECURITY: $($pwdOld.Count) users with passwords older than $PasswordOldDays days (where passwordNeverExpires = false)"
$report += ""

$report += "INACTIVE USERS (No login > $InactiveDays days)"
$report += "-" * 80
foreach ($u in ($inactiveUsers | Sort-Object DaysInactive -Descending)) {
  $report += ("{0,-12} {1,-22} {2,-12} {3,-12} {4,4} days  enabled={5}" -f `
    $u.SamAccountName, $u.DisplayName, $u.Department, $u.Site, $u.DaysInactive, $u.Enabled)
}
$report += ""

$report += "USERS PER DEPARTMENT"
$report += "-" * 80
foreach ($k in ($deptCounts.Keys | Sort-Object)) {
  $report += ("{0,-15}: {1,3}" -f $k, $deptCounts[$k])
}
$report += ""

$report += "COMPUTERS BY OPERATING SYSTEM"
$report += "-" * 80
foreach ($g in $osCounts) {
  $report += ("{0,-25}: {1,3}" -f $g.Name, $g.Count)
}
$report += ""

$report += "EXPIRING ACCOUNTS (<=30 days)"
$report += "-" * 80
foreach ($e in ($expiringSoon | Sort-Object DaysLeft)) {
  $report += ("{0,-12} {1,-22} expires={2} daysLeft={3}" -f $e.SamAccountName, $e.DisplayName, $e.AccountExpires, $e.DaysLeft)
}
$report += ""

$report += "=" * 80
$report += "RAPPORT SLUT".PadLeft(44).PadRight(80)
$report += "=" * 80

$report | Out-File -FilePath (Join-Path $OutDir "ad_audit_report.txt") -Encoding UTF8
Write-Host "OK: Skapade output/ad_audit_report.txt + CSV-filer"
