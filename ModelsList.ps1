Invoke-Sqlcmd -QueryTimeout 500 -ConnectionTimeout 500 -Query "use CM_LW2
select distinct V_GS_COMPUTER_SYSTEM.Model0 from  
V_R_System inner join V_GS_COMPUTER_SYSTEM on 
V_GS_COMPUTER_SYSTEM.ResourceID = V_R_System.ResourceId 
where V_GS_COMPUTER_SYSTEM.Model0 like '%'" | Out-GridView
