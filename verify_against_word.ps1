<#
    verify_against_word.ps1 - cross-check docx_wordcount.py against Word itself.

    Opens each .docx invisibly through Word's COM automation, reads the numbers
    from Word's own ComputeStatistics(), and diffs them against the Python
    counter. Use it when a file looks suspicious, or after changing the counter.

    Needs Microsoft Word installed. Much slower than the Python counter
    (seconds per file rather than milliseconds), which is the whole reason the
    Python counter exists.

    Usage:
      .\verify_against_word.ps1 -Path "C:\docs"
      .\verify_against_word.ps1 -Path a.docx,b.docx

    Note: one invisible Word instance is started for the whole run and quit in
    a finally block. Interrupting the script with Ctrl+C can bypass that block
    and leave an invisible WINWORD.EXE behind; if that happens, close it with
    Get-Process WINWORD | Stop-Process.
#>
param(
    [Parameter(Mandatory = $true)][string[]]$Path,
    [string]$Python = "python",
    [string]$Counter = "$PSScriptRoot\docx_wordcount.py"
)

$ErrorActionPreference = 'Stop'

$files = @()
foreach ($p in $Path) {
    if (Test-Path -LiteralPath $p -PathType Container) {
        $files += Get-ChildItem -LiteralPath $p -Filter *.docx -Recurse -File |
                  Where-Object { -not $_.Name.StartsWith('~$') } |
                  ForEach-Object { $_.FullName }
    } elseif (Test-Path -LiteralPath $p) {
        $files += (Resolve-Path -LiteralPath $p).Path
    } else {
        Write-Warning "not found: $p"
    }
}
if ($files.Count -eq 0) { Write-Warning "no .docx files found"; exit 1 }

# --- Word's own numbers -------------------------------------------------
$wdWords = 0; $wdChars = 3; $wdParagraphs = 4; $wdCharsSpaces = 5; $wdFarEast = 6
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$fromWord = @{}
try {
    foreach ($f in $files) {
        $doc = $null
        try {
            $doc = $word.Documents.Open($f, $false, $true)
            $fromWord[$f] = @($doc.ComputeStatistics($wdWords),
                              $doc.ComputeStatistics($wdChars),
                              $doc.ComputeStatistics($wdCharsSpaces),
                              $doc.ComputeStatistics($wdFarEast),
                              $doc.ComputeStatistics($wdParagraphs))
        } catch {
            $fromWord[$f] = $null
            Write-Warning ("Word could not read {0}: {1}" -f (Split-Path $f -Leaf),
                                                             $_.Exception.Message.Split("`n")[0])
        } finally {
            # Must close even when ComputeStatistics threw, or failed documents
            # accumulate open inside the one shared Word instance.
            if ($null -ne $doc) { try { $doc.Close(0) } catch { } }
        }
    }
} finally {
    $word.Quit()
    [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}

# --- the Python counter's numbers ---------------------------------------
# --tsv emits the full path in column 1, so rows are matched by path rather
# than by position - ordering on either side is not something to rely on.
$tsv = & $Python $Counter --tsv @files
if ($LASTEXITCODE -ne 0) { Write-Warning "the Python counter exited with $LASTEXITCODE" }

$fromPython = @{}
foreach ($line in $tsv) {
    $cols = $line -split "`t"
    if ($cols.Count -lt 6 -or $cols[0] -eq 'file') { continue }
    $fromPython[$cols[0]] = @($cols[1..5] | ForEach-Object { [int]$_ })
}

# --- diff ----------------------------------------------------------------
$labels = @('words', 'chars', 'chars+spaces', 'CJK chars', 'paragraphs')
$mismatch = 0
foreach ($f in $files) {
    $w = $fromWord[$f]
    $p = $fromPython[$f]
    $name = Split-Path $f -Leaf
    if ($null -eq $w) {
        Write-Host ("SKIP   {0}" -f $name) -ForegroundColor DarkGray
        continue
    }
    if ($null -eq $p) {
        Write-Host ("NO-PY  {0}  (no row returned for this path)" -f $name) -ForegroundColor Yellow
        $mismatch++
        continue
    }
    $bad = @()
    for ($k = 0; $k -lt 5; $k++) {
        if ($w[$k] -ne $p[$k]) { $bad += ("{0}: Word={1} py={2}" -f $labels[$k], $w[$k], $p[$k]) }
    }
    if ($bad.Count -eq 0) {
        Write-Host ("MATCH  {0,7} words  {1}" -f $w[0], $name) -ForegroundColor Green
    } else {
        $mismatch++
        Write-Host ("DIFF   {0}" -f $name) -ForegroundColor Red
        $bad | ForEach-Object { Write-Host ("         " + $_) -ForegroundColor Red }
    }
}
Write-Host ("`n{0}/{1} files match Word exactly." -f ($files.Count - $mismatch), $files.Count)
if ($mismatch -gt 0) { exit 1 }
