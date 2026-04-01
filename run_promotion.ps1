<#
.SYNOPSIS
    Student Promotion Task Runner for Windows Task Scheduler

.DESCRIPTION
    This PowerShell script can be scheduled to run automatically using Windows Task Scheduler
    to execute student promotions based on configured dates.

.PARAMETER DryRun
    Simulate promotion without making changes

.PARAMETER ClassName
    Process only a specific class

.PARAMETER Verbose
    Enable verbose output

.PARAMETER Force
    Force promotion even if not due (bypasses date check)

.PARAMETER UserId
    User ID to attribute the promotion to

.EXAMPLE
    .\run_promotion.ps1
    Run automatic promotion check (respects settings)

.EXAMPLE
    .\run_promotion.ps1 -DryRun
    Dry run to see what would happen

.EXAMPLE
    .\run_promotion.ps1 -ClassName "Grade 1"
    Promote specific class only

.EXAMPLE
    .\run_promotion.ps1 -Verbose
    Verbose output
#>

param(
    [switch]$DryRun,
    [string]$ClassName,
    [switch]$Verbose,
    [switch]$Force,
    [string]$UserId
)

# Change to script directory
Set-Location $PSScriptRoot

# Build command arguments
$arguments = @("run_promotion_task.py")

if ($DryRun) {
    $arguments += "--dry-run"
}

if ($ClassName) {
    $arguments += "--class"
    $arguments += $ClassName
}

if ($Verbose) {
    $arguments += "--verbose"
}

if ($Force) {
    $arguments += "--force"
}

if ($UserId) {
    $arguments += "--user"
    $arguments += $UserId
}

# Log file
$logFile = "promotion_task.log"

# Write header to log
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"=" * 60 | Out-File -FilePath $logFile -Append
"Promotion Task Started: $timestamp" | Out-File -FilePath $logFile -Append
"=" * 60 | Out-File -FilePath $logFile -Append

# Run the Python script
try {
    $process = Start-Process -FilePath "python" -ArgumentList $arguments -Wait -PassThru -NoNewWindow -RedirectStandardOutput $logFile -RedirectStandardError $logFile
    
    # Write footer to log
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "=" * 60 | Out-File -FilePath $logFile -Append
    "Promotion Task Completed: $timestamp" | Out-File -FilePath $logFile -Append
    "=" * 60 | Out-File -FilePath $logFile -Append
    
    # Exit with the same code as the Python script
    exit $process.ExitCode
}
catch {
    $errorMessage = "Error running promotion task: $_"
    Write-Error $errorMessage
    $errorMessage | Out-File -FilePath $logFile -Append
    exit 1
}
