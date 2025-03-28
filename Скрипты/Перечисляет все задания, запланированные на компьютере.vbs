$strComputer = "." 
 
$colItems = get-wmiobject -class "Win32_ScheduledJob" -namespace "root\CIMV2" ` 
-computername $strComputer 
 
foreach ($objItem in $colItems) { 
      write-host "Caption: " $objItem.Caption 
      write-host "Command: " $objItem.Command 
      write-host "Days Of Month: " $objItem.DaysOfMonth 
      write-host "Days Of Week: " $objItem.DaysOfWeek 
      write-host "Description: " $objItem.Description 
      write-host "Elapsed Time: " $objItem.ElapsedTime 
      write-host "Installation Date: " $objItem.InstallDate 
      write-host "Interact With Desktop: " $objItem.InteractWithDesktop 
      write-host "Job ID: " $objItem.JobId 
      write-host "Job Status: " $objItem.JobStatus 
      write-host "Name: " $objItem.Name 
      write-host "Notify: " $objItem.Notify 
      write-host "Owner: " $objItem.Owner 
      write-host "Priority: " $objItem.Priority 
      write-host "Run Repeatedly: " $objItem.RunRepeatedly 
      write-host "Start Time: " $objItem.StartTime 
      write-host "Status: " $objItem.Status 
      write-host "Time Submitted: " $objItem.TimeSubmitted 
      write-host "Until Time: " $objItem.UntilTime 
      write-host 
} 
