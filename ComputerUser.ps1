FUNCTION SCCM-ComputerNameLastUser {
Param([parameter(Mandatory = $true)]$SamAccountName,
    $SiteName="P01",
    $SCCMServer="ALCRMCWSSP01.pmusa.net")
    $SCCMNameSpace="root\sms\site_$SiteName"
    Get-WmiObject -namespace $SCCMNameSpace -computer $SCCMServer -query "select Name from sms_r_system where LastLogonUserName='$SamAccountName'" | select Name
    Write-Host 
}
SCCM-ComputerNameLastUser username