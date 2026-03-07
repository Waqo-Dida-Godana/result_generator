$ErrorActionPreference = "Stop"
$log = ".\run_log.txt"

"=== $(Get-Date) ===" | Out-File $log -Encoding utf8

# Install
try {
    $result = python -m pip install mysql-connector-python --no-cache-dir 2>&1
    $result | Out-File $log -Append -Encoding utf8
    "INSTALL DONE" | Out-File $log -Append -Encoding utf8
} catch {
    "INSTALL ERROR: $_" | Out-File $log -Append -Encoding utf8
}

# Check import
try {
    $ver = python -c "import mysql.connector; print(mysql.connector.__version__)" 2>&1
    "IMPORT OK: $ver" | Out-File $log -Append -Encoding utf8
} catch {
    "IMPORT FAILED: $_" | Out-File $log -Append -Encoding utf8
}
