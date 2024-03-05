Connect-ExchangeOnline
[int]$Daybefore = 0
[int]$i = 0
$Report = [System.Collections.Generic.List[Object]]::new(); $Now = Get-Date

do {
    $quarantine = Get-QuarantineMessage -PageSize 1000 -StartReceivedDate (get-date).AddDays($i).ToString("MM/dd/yyyy 00:00:00") -EndReceivedDate (get-date).AddDays($i).ToString("MM/dd/yyyy 23:59:59")
    
    ForEach ($Message in $quarantine) {
   
        $RemainingTime = (New-TimeSpan -Start $Now -End $Message.Expires)
        $Remaining = $RemainingTime.Days.toString() + " days " + $RemainingTime.Hours.toString() + " hours"
        [String]$Recipient = $Null; $c = 0
        ForEach ($Address in $Message.RecipientAddress) {
            If ($c -eq 0) {
               $c++
               $Recipient = $Address} 
            Else 
               {$Recipient = "; " + $Address }
       }
        
        $ReportLine = [PSCustomObject]@{  #Update with details of what we have done
                Identity         = $Message.Identity
                Received         = Get-Date($Message.ReceivedTime) -format g
                Recipient        = $Recipient
                Sender           = $Message.SenderAddress
                Subject          = $Message.Subject
                SenderDomain     = $Message.SenderAddress.Split("@")[1]
                Type             = $Message.QuarantineTypes
                Expires          = Get-Date($Message.Expires) -format g
                "Time Remaining" = $Remaining 
                ReleaseStatus = $Message.ReleaseStatus
                ReceivedTimeTest = $Message.ReceivedTime
                MessageId = $Message.MessageId} 
         $Report.Add($ReportLine) 
         $ReportLine | Export-Csv -path "C:\temp\Quarantine.csv" -NoClobber -NoTypeInformation -Append  -Encoding UTF8 -Delimiter ","
        
         If($quarantine.count -gt 900)
         {
            $Zaehler = $quarantine.Count
            Write-Host "zwischencount $Zaehler"
            $quarantine = Get-QuarantineMessage -PageSize 1000 -StartReceivedDate (get-date).AddDays($i).ToString("MM/dd/yyyy 00:00:00") -EndReceivedDate (get-date).AddDays($i).ToString("MM/dd/yyyy 23:59:59") -Page 2
            
            ForEach ($Message in $quarantine) {
   
                $RemainingTime = (New-TimeSpan -Start $Now -End $Message.Expires)
                $Remaining = $RemainingTime.Days.toString() + " days " + $RemainingTime.Hours.toString() + " hours"
                [String]$Recipient = $Null; 
                $Recipient = $Address
                
                $ReportLine = [PSCustomObject]@{  #Update with details of what we have done
                        Identity         = $Message.Identity
                        Received         = Get-Date($Message.ReceivedTime) -format g
                        Recipient        = $Recipient
                        Sender           = $Message.SenderAddress
                        Subject          = $Message.Subject
                        SenderDomain     = $Message.SenderAddress.Split("@")[1]
                        Type             = $Message.QuarantineTypes
                        Expires          = Get-Date($Message.Expires) -format g
                        "Time Remaining" = $Remaining 
                        ReleaseStatus = $Message.ReleaseStatus
                        ReceivedTimeTest = $Message.ReceivedTime
                        MessageId = $Message.MessageId} 
                 $Report.Add($ReportLine) 
                 $ReportLine | Export-Csv -path "C:\temp\Quarantine.csv" -NoClobber -NoTypeInformation -Append  -Encoding UTF8 -Delimiter ","}
        }
   }

    $Zaehler = $quarantine.Count
    Write-Output "Index: $i Q-count:$Zaehler"
    $i=$i -1
    $Daybefore=$Daybefore -1
    
    
} until ($i -eq -34)

$Report.Count
