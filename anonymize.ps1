<#
  Convenience wrapper for the anonymizer.
  Runs anonymizer.py with the project's own virtual environment, so you never
  have to type the full interpreter/script paths.

  Usage (from anywhere):
    C:\ClaudeCodeOM\anonimizer\anonymize.ps1 run "tiny.xlsx" -c "anonimizer config.csv" --seed seed1
    C:\ClaudeCodeOM\anonimizer\anonymize.ps1 run "tiny.xlsx" -c "anonimizer config.csv"          # auto seed
    C:\ClaudeCodeOM\anonimizer\anonymize.ps1 -h                                                  # help

  All arguments after the script name are passed straight through to anonymizer.py.
#>
$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
$py  = Join-Path $root '.venv\Scripts\python.exe'
$app = Join-Path $root 'anonymizer.py'

if (-not (Test-Path $py)) {
    Write-Error "venv not found at $py`nCreate it with:`n  python -m venv `"$root\.venv`"`n  & `"$py`" -m pip install openpyxl"
    exit 1
}

& $py $app @args
exit $LASTEXITCODE
