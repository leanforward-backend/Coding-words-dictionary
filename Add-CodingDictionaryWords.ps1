<#
.SYNOPSIS
    Adds a list of common programming keywords and acronyms to Windows' global
    spell-check dictionary and the Microsoft Office custom dictionary, so they
    stop getting flagged as misspellings everywhere on this machine.

.DESCRIPTION
    Windows spell-check (used by Edge, Mail, Teams, UWP apps, etc.) stores
    learned/added words in:
        %AppData%\Microsoft\Spelling\<locale>\default.dic

    Microsoft Office (Word, Outlook, etc.) has its own separate custom
    dictionary at:
        %AppData%\Microsoft\UProof\RoamingCustom.dic

    This script appends words to both files (creating them if missing),
    skipping any word that's already present, so it's safe to re-run.

.NOTES
    - No admin rights required - these are per-user files.
    - Re-run any time you want to add more words; duplicates are skipped.
    - Edit the $words array below to customize the list.
#>

[CmdletBinding()]
param(
    # Locale folder under %AppData%\Microsoft\Spelling - change if you use a different language pack
    [string]$Locale = "en-US"
)

# ---------------------------------------------------------------------------
# Word list - edit/extend this freely
# ---------------------------------------------------------------------------
$words = @(
    # General programming
    "async","await","bool","boolean","enum","struct","typeof","instanceof",
    "namespace","nullable","readonly","subclass","superclass","metaclass",
    "polymorphism","encapsulation","inheritance","singleton","middleware",
    "callback","lambda","closure","iterator","enumerable","serializable",
    "deserialize","serialize","stringify","destructure","destructuring",
    "decorator","annotation","interface","abstraction","refactor","refactoring",
    "linting","linter","minify","minification","transpile","transpiler",
    "polyfill","scaffolding","boilerplate","monorepo","microservice",
    "microservices","containerize","containerization","orchestration",
    "idempotent","mutex","semaphore","threadsafe","multithreading",
    "concurrency","asynchronous","synchronous","throttle","debounce",
    "memoize","memoization","recursion","backtracking","heuristic", "bable",

    # Acronyms - general software/web
    "API","APIs","SDK","SDKs","CLI","GUI","IDE","IDEs","URL","URI","URN",
    "HTTP","HTTPS","REST","RESTful","SOAP","gRPC","JSON","XML","YAML","TOML",
    "CSV","HTML","CSS","SCSS","LESS","DOM","BOM","SPA","SSR","SSG","CSR",
    "CDN","DNS","TCP","UDP","TLS","SSL","SSH","FTP","SFTP","JWT","OAuth",
    "SAML","SSO","CORS","CSRF","XSS","SQL","NoSQL","ORM","ACID","CRUD",
    "MVC","MVVM","MVP","SOLID","DRY","KISS","YAGNI","TDD","BDD","DDD",
    "CI","CD","CICD","DevOps","SRE","IaC","SaaS","PaaS","IaaS","FaaS",
    "VM","VMs","OS","RAM","CPU","GPU","SSD","HDD","IO","I/O",

    # Acronyms - data/algorithms
    "BFS","DFS","FIFO","LIFO","LRU","LRC","CRC","UUID","GUID","RNG","PRNG",
    "ML","AI","NLP","LLM","LLMs","CNN","RNN","GAN","RAG","API",

    # Languages/runtimes/tools (proper nouns spellcheck loves to flag)
    "JavaScript","TypeScript","Node","NodeJS","npm","npx","yarn","webpack",
    "Vite","Babel","ESLint","Prettier","Kubernetes","Docker","Dockerfile",
    "Terraform","Ansible","Jenkins","GitHub","GitLab","Bitbucket","Postgres",
    "PostgreSQL","MySQL","MongoDB","Redis","SQLite","GraphQL","Nginx",
    "Apache","Linux","Ubuntu","Debian","Arch","Homebrew","Conda","venv",
    "PyPI","pip","Django","Flask","FastAPI","Rails","Laravel","Spring",
    "Kotlin","Swift","SwiftUI","Golang","Rust","Cargo","Clang","LLVM",
    "WebAssembly","Wasm","ESP8266","ESP32","Arduino","Raspberry","NRF24L01",

    # File extensions and misc tokens
    "dotfile","gitignore","readme","changelog","monorepo","frontend",
    "backend","fullstack","filesystem","middleware","webhook","webhooks",
    "endpoint","endpoints","payload","payloads","schema","schemas",
    "namespace","keyspace","datastore","datatype","datatypes","metadata",
    "config","configs","env","dotenv","stdout","stderr","stdin"
)

# De-duplicate, case-insensitive, preserve first-seen casing
$words = $words | Select-Object -Unique

function Add-WordsToDictionary {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string[]]$NewWords,
        [Parameter(Mandatory)] [string]$Label
    )

    $dir = Split-Path -Path $Path -Parent
    if (-not (Test-Path -Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $existing = @()
    if (Test-Path -Path $Path) {
        # These .dic files are typically UTF-16 LE; read robustly regardless of encoding
        $existing = Get-Content -Path $Path -ErrorAction SilentlyContinue
    } else {
        New-Item -ItemType File -Path $Path -Force | Out-Null
    }

    $existingSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($e in $existing) { if ($e) { [void]$existingSet.Add($e.Trim()) } }

    $toAdd = $NewWords | Where-Object { $_ -and -not $existingSet.Contains($_) }

    if ($toAdd.Count -eq 0) {
        Write-Host "[$Label] No new words to add - everything's already there." -ForegroundColor Yellow
        return
    }

    # Append using UTF8 (BOM-less) - Windows spelling dictionaries read this fine
    Add-Content -Path $Path -Value $toAdd -Encoding utf8

    Write-Host "[$Label] Added $($toAdd.Count) new word(s) to:`n  $Path" -ForegroundColor Green
}

# ---------------------------------------------------------------------------
# 1. Windows global spell-check dictionary
# ---------------------------------------------------------------------------
$winDicPath = Join-Path $env:APPDATA "Microsoft\Spelling\$Locale\default.dic"
Add-WordsToDictionary -Path $winDicPath -NewWords $words -Label "Windows Global"

# ---------------------------------------------------------------------------
# 2. Microsoft Office custom dictionary
# ---------------------------------------------------------------------------
$officeDicPath = Join-Path $env:APPDATA "Microsoft\UProof\RoamingCustom.dic"
Add-WordsToDictionary -Path $officeDicPath -NewWords $words -Label "Microsoft Office"

# ---------------------------------------------------------------------------
# 3. Obsidian custom dictionary
# ---------------------------------------------------------------------------
function Add-WordsToObsidianDictionary {
    param(
        [Parameter(Mandatory)] [string]$Path,
        [Parameter(Mandatory)] [string[]]$NewWords,
        [Parameter(Mandatory)] [string]$Label
    )

    $dir = Split-Path -Path $Path -Parent
    if (-not (Test-Path -Path $dir)) {
        return # If Obsidian directory doesn't exist, skip silently
    }

    $existing = @()
    if (Test-Path -Path $Path) {
        # Read lines, filtering out the checksum line
        $existing = Get-Content -Path $Path -ErrorAction SilentlyContinue | Where-Object { $_ -and $_ -notlike "checksum_v1 = *" }
    }

    $existingSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($e in $existing) { if ($e.Trim()) { [void]$existingSet.Add($e.Trim()) } }

    $toAdd = $NewWords | Where-Object { $_ -and -not $existingSet.Contains($_) }

    if ($toAdd.Count -eq 0) {
        Write-Host "[$Label] No new words to add - everything's already there." -ForegroundColor Yellow
        return
    }

    $allWords = @()
    foreach ($e in $existing) { if ($e.Trim()) { $allWords += $e.Trim() } }
    foreach ($w in $toAdd) { $allWords += $w }

    # Write all words back using .NET File.WriteAllLines to get UTF-8 without BOM (standard for Chromium)
    [System.IO.File]::WriteAllLines($Path, $allWords)

    Write-Host "[$Label] Added $($toAdd.Count) new word(s) to:`n  $Path" -ForegroundColor Green
}

$obsidianDicPath = Join-Path $env:APPDATA "obsidian\Custom Dictionary.txt"
Add-WordsToObsidianDictionary -Path $obsidianDicPath -NewWords $words -Label "Obsidian"

Write-Host "`nDone. You may need to restart open apps (Edge, Word, Outlook, Obsidian, etc.) for changes to take effect." -ForegroundColor Cyan

