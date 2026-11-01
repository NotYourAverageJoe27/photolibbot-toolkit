<#
pcat-win.ps1 - Windows PowerShell 5.1 companion to pcat.py, for drives the Linux shell cannot mount.
The user launches it with a RUN_*.cmd (double-clicking a .ps1 does not run it).

Modes
  -Mode Inventory   -Roots "E:\PHOTOS|C:\Media\Library" -OutDir <_Catalog\job>     (separate lists with |)
                    Lists every file (path,size,mtime) per root; SHA-1 hashes files whose size occurs
                    more than once across all roots (that is all a duplicate/backup comparison needs).
                    Add -HashAll to hash everything.
  -Mode CopyVerify  -Plan plan.csv -OutDir <_Catalog\job> [-AllowSourceDelete]
                    plan.csv columns: action (copy|move), src, dst. Never overwrites. Every copy is
                    SHA-1 verified source vs destination. 'move' deletes the source only when
                    -AllowSourceDelete is given AND the hash matched (same volume: plain rename).
  -Mode VerifyTree  -A <master> -B <backup> -OutDir <_Catalog\job> [-Exclude "Photo & Video Library\_Others' Phones\Partner|..."] [-NoHash]
                    Path+size comparison, then SHA-1 of every common file. Writes differences.csv
                    (empty = identical).

Rules baked in (each one broke a real run): \\?\ long paths on every .NET call; no function names
that collide with built-in aliases; no PowerShell 7-only syntax; fail fast if the first 50 files fail;
outputs written straight to -OutDir; nothing deleted on failure.
#>
param(
    [Parameter(Mandatory = $true)][ValidateSet('Inventory', 'CopyVerify', 'VerifyTree')][string]$Mode,
    [string[]]$Roots,
    [string]$Plan,
    [string]$A,
    [string]$B,
    [string[]]$Exclude = @(),
    [Parameter(Mandatory = $true)][string]$OutDir,
    [switch]$HashAll,
    [switch]$NoHash,
    [switch]$AllowSourceDelete,
    [int]$ProgressSeconds = 30
)

$ErrorActionPreference = 'Stop'
$script:OnWindows = ([Environment]::OSVersion.Platform -eq 'Win32NT')

function ConvertTo-PcatLongPath([string]$p) {
    if (-not $script:OnWindows) { return $p }
    if ($p.StartsWith('\\?\')) { return $p }
    $full = [IO.Path]::GetFullPath($p)
    if ($full.StartsWith('\\')) { return '\\?\UNC\' + $full.Substring(2) }
    return '\\?\' + $full
}

function ConvertFrom-PcatLongPath([string]$p) {
    if ($p.StartsWith('\\?\UNC\')) { return '\\' + $p.Substring(8) }
    if ($p.StartsWith('\\?\')) { return $p.Substring(4) }
    return $p
}

function Get-PcatSha1([string]$path) {
    $sha = [System.Security.Cryptography.SHA1]::Create()
    $fs = New-Object System.IO.FileStream((ConvertTo-PcatLongPath $path), [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read, 1048576)
    try {
        $bytes = $sha.ComputeHash($fs)
    }
    finally {
        $fs.Dispose(); $sha.Dispose()
    }
    return ([BitConverter]::ToString($bytes) -replace '-', '').ToLowerInvariant()
}

function Get-PcatFiles([string]$root, [string[]]$excludeRel) {
    # returns objects: Rel, Full, Size, MTime ; skips $RECYCLE.BIN / System Volume Information
    $rootFull = [IO.Path]::GetFullPath($root).TrimEnd('\', '/')
    $lpRoot = ConvertTo-PcatLongPath $rootFull
    $list = New-Object System.Collections.Generic.List[object]
    $stack = New-Object System.Collections.Generic.Stack[string]
    $stack.Push($lpRoot)
    $sep = [IO.Path]::DirectorySeparatorChar
    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()
        try {
            foreach ($d in [IO.Directory]::EnumerateDirectories($dir)) {
                $name = [IO.Path]::GetFileName($d)
                if ($name -eq '$RECYCLE.BIN' -or $name -eq 'System Volume Information') { continue }
                $di = New-Object IO.DirectoryInfo($d)
                if (($di.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                $relDir = $d.Substring($lpRoot.Length).TrimStart('\', '/')
                $skip = $false
                foreach ($x in $excludeRel) { if ($relDir -ieq $x.TrimEnd('\', '/')) { $skip = $true } }
                if (-not $skip) { $stack.Push($d) }
            }
            foreach ($f in [IO.Directory]::EnumerateFiles($dir)) {
                $fi = New-Object IO.FileInfo($f)
                $rel = $f.Substring($lpRoot.Length).TrimStart('\', '/')
                $list.Add([pscustomobject]@{ Rel = $rel; Full = (ConvertFrom-PcatLongPath $f); Size = $fi.Length; MTime = $fi.LastWriteTime.ToString('s') })
            }
        }
        catch {
            Write-Host ("  ! cannot list {0}: {1}" -f (ConvertFrom-PcatLongPath $dir), $_.Exception.Message) -ForegroundColor Yellow
        }
    }
    return , $list
}

function Write-PcatProgress([string]$text) {
    Write-Host $text
    try { Set-Content -LiteralPath (Join-Path $OutDir 'progress.txt') -Value ("{0}  {1}" -f (Get-Date -Format s), $text) } catch { }
}

function Format-PcatEta([double]$doneBytes, [double]$totalBytes, [datetime]$t0) {
    $el = ((Get-Date) - $t0).TotalSeconds
    if ($doneBytes -le 0 -or $el -le 0) { return 'ETA ?' }
    $rate = $doneBytes / $el
    $eta = ($totalBytes - $doneBytes) / $rate
    return ('{0:N0} MB/s  ETA {1:N0} min' -f ($rate / 1MB), ($eta / 60))
}

# From a .cmd every list arrives as ONE string: separate multiple values with '|' (never valid in Windows
# names; commas are not safe because folder names such as "Orlando, FL" contain them).
$Roots = @($Roots | Where-Object { $_ } | ForEach-Object { $_ -split '\|' } | Where-Object { $_.Trim() } | ForEach-Object { $_.Trim() })
$Exclude = @($Exclude | Where-Object { $_ } | ForEach-Object { $_ -split '\|' } | Where-Object { $_.Trim() } | ForEach-Object { $_.Trim().Replace('/', [IO.Path]::DirectorySeparatorChar).Replace('\', [IO.Path]::DirectorySeparatorChar) })

New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
$t0 = Get-Date
Write-Host ("pcat-win {0} started {1}" -f $Mode, (Get-Date -Format s))
$failed = 0

if ($Mode -eq 'Inventory') {
    if (-not $Roots) { throw '-Roots is required' }
    $all = New-Object System.Collections.Generic.List[object]
    foreach ($r in $Roots) {
        Write-PcatProgress "listing $r"
        $files = Get-PcatFiles $r @()
        foreach ($f in $files) { $all.Add([pscustomobject]@{ Root = $r; Rel = $f.Rel; Full = $f.Full; Size = $f.Size; MTime = $f.MTime; Sha1 = '' }) }
        Write-PcatProgress ("  {0:N0} files" -f $files.Count)
    }
    $sizeCount = @{}
    foreach ($x in $all) { if ($sizeCount.ContainsKey($x.Size)) { $sizeCount[$x.Size]++ } else { $sizeCount[$x.Size] = 1 } }
    $toHash = @($all | Where-Object { $_.Size -gt 0 -and ($HashAll -or $sizeCount[$_.Size] -gt 1) })
    $total = 0.0; foreach ($x in $toHash) { $total += $x.Size }
    $done = 0.0; $i = 0; $last = Get-Date; $hashT0 = Get-Date
    foreach ($x in $toHash) {
        $i++
        try { $x.Sha1 = Get-PcatSha1 $x.Full } catch { $x.Sha1 = 'ERROR: ' + $_.Exception.Message; $failed++ }
        $done += $x.Size
        if (((Get-Date) - $last).TotalSeconds -ge $ProgressSeconds -or $i -eq $toHash.Count) {
            Write-PcatProgress ("hashed {0:N0}/{1:N0}  {2:N1}/{3:N1} GB  {4}" -f $i, $toHash.Count, ($done / 1GB), ($total / 1GB), (Format-PcatEta $done $total $hashT0))
            $last = Get-Date
        }
        if ($i -eq 50 -and $failed -eq 50) { throw 'first 50 hashes all failed - stopping (fail-fast)' }
    }
    $out = Join-Path $OutDir 'inventory.csv'
    $all | Select-Object Root, Rel, Size, MTime, Sha1 | Export-Csv -LiteralPath $out -NoTypeInformation -Encoding UTF8
    Write-Host ("wrote {0}: {1:N0} files, {2:N0} hashed, {3} hash errors" -f $out, $all.Count, $toHash.Count, $failed)
}

if ($Mode -eq 'CopyVerify') {
    if (-not $Plan) { throw '-Plan is required' }
    $rows = @(Import-Csv -LiteralPath $Plan)
    $logRows = New-Object System.Collections.Generic.List[object]
    $total = 0.0
    foreach ($r in $rows) { try { $total += (New-Object IO.FileInfo((ConvertTo-PcatLongPath $r.src))).Length } catch { } }
    $done = 0.0; $i = 0; $ok = 0; $last = Get-Date
    foreach ($r in $rows) {
        $i++
        $act = ([string]$r.action).Trim().ToLowerInvariant()
        if (-not $act) { $act = 'copy' }
        $src = [string]$r.src; $dst = [string]$r.dst
        $status = 'ok'; $detail = ''; $h1 = ''
        try {
            $lsrc = ConvertTo-PcatLongPath $src; $ldst = ConvertTo-PcatLongPath $dst
            if (-not [IO.File]::Exists($lsrc)) { throw 'source missing' }
            if ([IO.File]::Exists($ldst)) { throw 'destination exists (never overwrite)' }
            if ($act -ne 'copy' -and $act -ne 'move') { throw "unknown action $act" }
            $size = (New-Object IO.FileInfo($lsrc)).Length
            [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($ldst)) | Out-Null
            $sameVol = ([IO.Path]::GetPathRoot($src).ToUpperInvariant() -eq [IO.Path]::GetPathRoot($dst).ToUpperInvariant())
            if ($act -eq 'move' -and $AllowSourceDelete -and $sameVol) {
                [IO.File]::Move($lsrc, $ldst)
                if ((New-Object IO.FileInfo($ldst)).Length -ne $size) { throw 'post-move size check failed' }
                $detail = 'renamed (same volume)'
            }
            else {
                $h1 = Get-PcatSha1 $src
                $tmp = $ldst + '.pcat_tmp'
                [IO.File]::Copy($lsrc, $tmp, $false)
                [IO.File]::SetLastWriteTimeUtc($tmp, [IO.File]::GetLastWriteTimeUtc($lsrc))
                $h2 = Get-PcatSha1 (ConvertFrom-PcatLongPath $tmp)
                if ($h1 -ne $h2) { [IO.File]::Delete($tmp); throw 'SHA-1 mismatch after copy' }
                [IO.File]::Move($tmp, $ldst)
                $detail = 'copied+verified'
                if ($act -eq 'move') {
                    if ($AllowSourceDelete) { [IO.File]::Delete($lsrc); $detail = 'copied+verified, source deleted' }
                    else { $detail = 'copied+verified, source kept (no -AllowSourceDelete)' }
                }
            }
            $done += $size; $ok++
        }
        catch {
            $status = 'FAILED'; $detail = $_.Exception.Message; $failed++
        }
        $logRows.Add([pscustomobject]@{ time = (Get-Date -Format s); action = $act; src = $src; dst = $dst; sha1 = $h1; status = $status; detail = $detail })
        if (((Get-Date) - $last).TotalSeconds -ge $ProgressSeconds -or $i -eq $rows.Count) {
            Write-PcatProgress ("{0:N0}/{1:N0}  ok {2:N0}  failed {3:N0}  {4:N1}/{5:N1} GB  {6}" -f $i, $rows.Count, $ok, $failed, ($done / 1GB), ($total / 1GB), (Format-PcatEta $done $total $t0))
            $last = Get-Date
            $logRows | Export-Csv -LiteralPath (Join-Path $OutDir 'copy_log.csv') -NoTypeInformation -Encoding UTF8
        }
        if ($i -eq 50 -and $failed -eq 50) { Write-Host 'first 50 actions all failed - stopping (fail-fast)' -ForegroundColor Red; break }
    }
    $logRows | Export-Csv -LiteralPath (Join-Path $OutDir 'copy_log.csv') -NoTypeInformation -Encoding UTF8
    Write-Host ("CopyVerify: {0:N0} ok, {1:N0} failed. Log: {2}" -f $ok, $failed, (Join-Path $OutDir 'copy_log.csv'))
}

if ($Mode -eq 'VerifyTree') {
    if (-not $A -or -not $B) { throw '-A and -B are required' }
    Write-PcatProgress "listing $A"
    $la = Get-PcatFiles $A $Exclude
    Write-PcatProgress "listing $B"
    $lb = Get-PcatFiles $B $Exclude
    $mapB = @{}
    foreach ($f in $lb) { $mapB[$f.Rel.ToLowerInvariant()] = $f }
    $mapA = @{}
    foreach ($f in $la) { $mapA[$f.Rel.ToLowerInvariant()] = $f }
    $diff = New-Object System.Collections.Generic.List[object]
    $common = New-Object System.Collections.Generic.List[object]
    foreach ($f in $la) {
        $k = $f.Rel.ToLowerInvariant()
        if (-not $mapB.ContainsKey($k)) { $diff.Add([pscustomobject]@{ path = $f.Rel; status = 'missing_in_B'; size_a = $f.Size; size_b = '' }) }
        elseif ($mapB[$k].Size -ne $f.Size) { $diff.Add([pscustomobject]@{ path = $f.Rel; status = 'size_diff'; size_a = $f.Size; size_b = $mapB[$k].Size }) }
        else { $common.Add(@($f, $mapB[$k])) }
    }
    foreach ($f in $lb) {
        if (-not $mapA.ContainsKey($f.Rel.ToLowerInvariant())) { $diff.Add([pscustomobject]@{ path = $f.Rel; status = 'extra_in_B'; size_a = ''; size_b = $f.Size }) }
    }
    Write-PcatProgress ("A {0:N0} files | B {1:N0} files | same path+size {2:N0} | differences so far {3:N0}" -f $la.Count, $lb.Count, $common.Count, $diff.Count)
    if (-not $NoHash) {
        $total = 0.0; foreach ($p in $common) { $total += $p[0].Size }
        $done = 0.0; $i = 0; $last = Get-Date; $ht0 = Get-Date; $errs = 0
        foreach ($p in $common) {
            $i++
            try {
                if ((Get-PcatSha1 $p[0].Full) -ne (Get-PcatSha1 $p[1].Full)) { $diff.Add([pscustomobject]@{ path = $p[0].Rel; status = 'hash_diff'; size_a = $p[0].Size; size_b = $p[1].Size }) }
            }
            catch { $diff.Add([pscustomobject]@{ path = $p[0].Rel; status = ('read_error: ' + $_.Exception.Message); size_a = $p[0].Size; size_b = '' }); $errs++ }
            $done += $p[0].Size
            if (((Get-Date) - $last).TotalSeconds -ge $ProgressSeconds -or $i -eq $common.Count) {
                Write-PcatProgress ("hashed {0:N0}/{1:N0}  {2:N1}/{3:N1} GB  {4}" -f $i, $common.Count, ($done / 1GB), ($total / 1GB), (Format-PcatEta $done $total $ht0))
                $last = Get-Date
            }
            if ($i -eq 50 -and $errs -eq 50) { throw 'first 50 files all unreadable - stopping (fail-fast)' }
        }
    }
    $out = Join-Path $OutDir 'differences.csv'
    if ($diff.Count -gt 0) { $diff | Export-Csv -LiteralPath $out -NoTypeInformation -Encoding UTF8 }
    else { Set-Content -LiteralPath $out -Value 'path,status,size_a,size_b' -Encoding UTF8 }
    $failed = $diff.Count
    if ($diff.Count -eq 0) { Write-Host ("VerifyTree: IDENTICAL ({0:N0} files{1})" -f $la.Count, $(if ($NoHash) { ', size only' } else { ', SHA-1' })) -ForegroundColor Green }
    else { Write-Host ("VerifyTree: {0:N0} differences - see {1}" -f $diff.Count, $out) -ForegroundColor Red }
}

try { Remove-Item -LiteralPath (Join-Path $OutDir 'progress.txt') -ErrorAction SilentlyContinue } catch { }
Write-Host ("finished {0}  ({1:N1} min)" -f (Get-Date -Format s), ((Get-Date) - $t0).TotalMinutes)
if ($failed -gt 0) { exit 1 } else { exit 0 }
