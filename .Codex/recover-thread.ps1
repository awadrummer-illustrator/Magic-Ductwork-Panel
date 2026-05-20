param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$Thread,

    [int]$RecentCount = 5,

    [int]$MaxMessageChars = 4000,

    [switch]$AllAssistant
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
} catch {
}

function Get-ThreadId {
    param([string]$Value)

    $match = [regex]::Match($Value, '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}')
    if (-not $match.Success) {
        throw "No Codex thread id found in '$Value'. Pass a raw id or codex://threads/<id>."
    }

    return $match.Value.ToLowerInvariant()
}

function Get-CodexHomes {
    $candidates = @(
        $env:CODEX_HOME,
        (Join-Path $env:USERPROFILE '.codex'),
        (Join-Path $env:APPDATA 'Codex'),
        (Join-Path $env:LOCALAPPDATA 'Codex'),
        'E:\My Apps\Codex\.codex'
    )

    $seen = @{}
    $homes = @()

    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) {
            continue
        }

        try {
            $resolved = (Resolve-Path -LiteralPath $candidate -ErrorAction Stop).Path
        } catch {
            continue
        }

        $key = $resolved.ToLowerInvariant()
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $homes += $resolved
        }
    }

    return $homes
}

function Find-SessionFile {
    param(
        [string]$ThreadId,
        [string[]]$Homes
    )

    foreach ($codexHome in $Homes) {
        $sessions = Join-Path $codexHome 'sessions'
        if (-not (Test-Path -LiteralPath $sessions)) {
            continue
        }

        $matches = @(Get-ChildItem -LiteralPath $sessions -Recurse -File -Filter "*.jsonl" -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like "*$ThreadId*" } |
            Sort-Object LastWriteTime -Descending)

        if ($matches.Count -gt 0) {
            return $matches[0].FullName
        }
    }

    throw "Couldn't find a local Codex session file for thread '$ThreadId'. Checked: $($Homes -join ', ')"
}

function Get-TextFromContent {
    param($Content)

    $parts = @()
    foreach ($item in @($Content)) {
        if ($null -ne $item.text) {
            $parts += [string]$item.text
        } elseif ($null -ne $item.content) {
            $parts += [string]$item.content
        }
    }

    return ($parts -join "`n")
}

function Add-VisibleMessage {
    param(
        [System.Collections.ArrayList]$Messages,
        [hashtable]$Seen,
        [string]$Timestamp,
        [string]$Role,
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return
    }

    $clean = $Text -replace "`r", ''
    if ($Role -eq 'user' -and (Test-ScaffoldMessage -Text $clean)) {
        return
    }

    $key = $Role + '|' + $clean
    if ($Seen.ContainsKey($key)) {
        return
    }

    $Seen[$key] = $true
    [void]$Messages.Add([pscustomobject]@{
        Timestamp = $Timestamp
        Role = $Role
        Text = $clean.TrimEnd()
    })
}

function Test-ScaffoldMessage {
    param([string]$Text)

    $trimmed = $Text.TrimStart()
    return $trimmed.StartsWith('# AGENTS.md instructions') -or
        $trimmed.StartsWith('<environment_context>') -or
        $trimmed.StartsWith('<INSTRUCTIONS>')
}

function Limit-Text {
    param(
        [string]$Text,
        [int]$Limit
    )

    if ($Text.Length -le $Limit) {
        return $Text
    }

    return $Text.Substring(0, $Limit) + "`n`n[truncated at $Limit characters by recover-thread.ps1]"
}

function Write-MessageBlock {
    param(
        [string]$Title,
        $Message,
        [int]$Limit
    )

    if ($null -eq $Message) {
        return
    }

    Write-Output ""
    Write-Output "## $Title"
    Write-Output ""
    Write-Output "Timestamp: $($Message.Timestamp)"
    Write-Output ""
    Write-Output '```text'
    Write-Output (Limit-Text -Text $Message.Text -Limit $Limit)
    Write-Output '```'
}

$threadId = Get-ThreadId -Value $Thread
$homes = @(Get-CodexHomes)
if ($homes.Count -eq 0) {
    throw "No local Codex data folder found. CODEX_HOME wasn't set and the usual folders don't exist."
}

$sessionFile = Find-SessionFile -ThreadId $threadId -Homes $homes
$messages = New-Object System.Collections.ArrayList
$seenMessages = @{}
$sessionCwd = $null
$sessionName = $null

Get-Content -LiteralPath $sessionFile -Encoding UTF8 | ForEach-Object {
    try {
        $entry = $_ | ConvertFrom-Json
    } catch {
        return
    }

    if ($entry.type -eq 'session_meta') {
        if ($entry.payload.cwd) {
            $sessionCwd = [string]$entry.payload.cwd
        }
        return
    }

    if ($entry.type -eq 'response_item' -and $entry.payload.type -eq 'message') {
        if ($entry.payload.role -ne 'user' -and $entry.payload.role -ne 'assistant') {
            return
        }

        $text = Get-TextFromContent -Content $entry.payload.content
        Add-VisibleMessage -Messages $messages -Seen $seenMessages -Timestamp ([string]$entry.timestamp) -Role ([string]$entry.payload.role) -Text $text
        return
    }

    if ($entry.type -eq 'event_msg') {
        if ($entry.payload.type -eq 'user_message') {
            Add-VisibleMessage -Messages $messages -Seen $seenMessages -Timestamp ([string]$entry.timestamp) -Role 'user' -Text ([string]$entry.payload.message)
            return
        }

        if ($entry.payload.type -eq 'agent_message') {
            Add-VisibleMessage -Messages $messages -Seen $seenMessages -Timestamp ([string]$entry.timestamp) -Role 'assistant' -Text ([string]$entry.payload.message)
            return
        }
    }
}

$indexFile = $null
foreach ($codexHome in $homes) {
    if ($sessionFile.ToLowerInvariant().StartsWith($codexHome.ToLowerInvariant())) {
        $candidateIndex = Join-Path $codexHome 'session_index.jsonl'
        if (Test-Path -LiteralPath $candidateIndex) {
            $indexFile = $candidateIndex
            break
        }
    }
}

if ($indexFile -and (Test-Path -LiteralPath $indexFile)) {
    Get-Content -LiteralPath $indexFile -Encoding UTF8 | ForEach-Object {
        try {
            $entry = $_ | ConvertFrom-Json
        } catch {
            return
        }

        if ($entry.id -eq $threadId -and $entry.thread_name) {
            $sessionName = [string]$entry.thread_name
        }
    }
}

$assistantMessages = @($messages | Where-Object { $_.Role -eq 'assistant' })
$userMessages = @($messages | Where-Object { $_.Role -eq 'user' })
$lastAssistant = $assistantMessages | Select-Object -Last 1
$recentUsers = $userMessages | Select-Object -Last $RecentCount
$recentAssistants = $assistantMessages | Select-Object -Last $RecentCount

$allVisibleText = ($messages | ForEach-Object { $_.Text }) -join "`n"
$hashSource = $allVisibleText -replace '(?i)[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}', ' '
$hashes = @([regex]::Matches($hashSource, '(?i)(?<![0-9a-f])(?=[0-9a-f]{7,64}(?![0-9a-f]))(?=[0-9a-f]*\d)[0-9a-f]{7,64}') |
    ForEach-Object { $_.Value } |
    Sort-Object -Unique)
$gitDirectives = @([regex]::Matches($allVisibleText, '::git-[^\r\n]+') |
    ForEach-Object { $_.Value } |
    Sort-Object -Unique)

Write-Output "# Codex Thread Recovery"
Write-Output ""
Write-Output "Thread: $threadId"
if ($sessionName) {
    Write-Output "Name: $sessionName"
}
Write-Output "Session file: $sessionFile"
if ($sessionCwd) {
    Write-Output "Workspace: $sessionCwd"
}
Write-Output "Visible messages found: $($messages.Count)"

Write-MessageBlock -Title 'Last Assistant Message' -Message $lastAssistant -Limit $MaxMessageChars

if ($recentUsers.Count -gt 0) {
    Write-Output ""
    Write-Output "## Recent User Messages"
    foreach ($message in $recentUsers) {
        $preview = ($message.Text -replace "`n", ' ').Trim()
        $preview = Limit-Text -Text $preview -Limit 500
        Write-Output ""
        Write-Output "- $($message.Timestamp): $preview"
    }
}

if ($AllAssistant -and $recentAssistants.Count -gt 0) {
    Write-Output ""
    Write-Output "## Recent Assistant Messages"
    foreach ($message in $recentAssistants) {
        $preview = ($message.Text -replace "`n", ' ').Trim()
        $preview = Limit-Text -Text $preview -Limit 500
        Write-Output ""
        Write-Output "- $($message.Timestamp): $preview"
    }
}

if ($hashes.Count -gt 0) {
    Write-Output ""
    Write-Output "## Hashes Mentioned"
    foreach ($hash in $hashes) {
        Write-Output "- $hash"
    }
}

if ($gitDirectives.Count -gt 0) {
    Write-Output ""
    Write-Output "## Git Directives Mentioned"
    foreach ($directive in $gitDirectives) {
        Write-Output "- $directive"
    }
}

Write-Output ""
Write-Output "Raw JSONL was parsed but not printed."
