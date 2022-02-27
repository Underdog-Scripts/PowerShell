# Turning off UAC: 

param (
	[string]$status="Enable",
	[string]$LoggingPhase = ''
)

$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition

#if (test-path "\\NG00244191\c$\Scripts\PS\Scripts and Modules\InstallAndLoggingTools") { IPMO "\\NG00244191\c$\Scripts\PS\Scripts and Modules\InstallAndLoggingTools" } 
#else { 
IPMO InstallAndLoggingTools 
#}

if (Test-W10) {
	Start-ScriptLogging "UAC - $status.txt" -LoggingPhase $LoggingPhase
		#Import-RegFile "$PSScriptRoot\RegFiles\UAC - $status`.reg"
	Stop-ScriptLogging
} 