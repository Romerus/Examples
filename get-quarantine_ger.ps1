Connect-ExchangeOnline
[int]$Daybefore = 0
[int]$i = 0
$Report = [System.Collections.Generic.List[Object]]::new(); $Now = Get-Date

function Get-AllQuarantine {
    param (
        [int]$pagecount
    )
    
    $quarantine = Get-QuarantineMessage -PageSize 1000 -StartReceivedDate (Get-Date).AddDays($i).ToString("dd/MM/yyyy 00:00:00") -EndReceivedDate (Get-Date).AddDays($i).ToString("dd/MM/yyyy 23:59:59") -Page $pagecount
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
    }
    return $quarantine
}

do {
    $quarantine = Get-AllQuarantine -pagecount 1   
    If($quarantine.count -gt 900) {$quarantine= get-AllQuarantine -pagecount 2}
    
    $i=$i -1
    $Daybefore=$Daybefore -1  
    
} until ($i -eq -31)

$Report.Count
