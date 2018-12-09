# param example
# *.ps1 -auto
# *.ps1 -csv Y
# *.ps1 -startFromDate '11/10/2018'   ## MM/dd/yyyy 

param(
    [switch]$Auto,
    $startFromDate,
    $csv = 'n',
    $TransportService = 'mail02',
    $SmtpServer = 'mail02.bsg-ua.com'
)
if (!$startFromDate) {
    [datetime]$fromDate = (get-date).AddDays(-1)
}
else {
    $fromDate = get-date $startFromDate    
}

if ($csv -eq 'Y' -or $csv -eq ('y')) {
    $exportFlag = $true
    $tempArray = New-Object System.Collections.ArrayList
}
else {
    $exportFlag = $false
}

if ($Auto) {
    if (!$startFromDate) {[datetime]$fromDate = (get-date).AddDays(-1)}
    $exportFlag = $true
    $tempArray = New-Object System.Collections.ArrayList
}


function SendStatToParser ($exportPath) {
    send-mailmessage -Attachments $exportPath -To "queue@bsg-ua.com" -From "mail.stat@bsg-ua.com" -Subject "Daily Email Statistics" -SmtpServer $SmtpServer
}


function CsvExport($tempArray) {
    $listRows = @()
    for ($item = 0; $item -le $tempArray.Count - 5; $item += 5) {
        $oneRow = New-Object -TypeName psobject
        Add-Member -InputObject $oneRow -MemberType NoteProperty -Name 'Date' -Value $tempArray[$item]
        Add-Member -InputObject $oneRow -MemberType NoteProperty -Name 'Sent' -Value $tempArray[$item + 1]
        Add-Member -InputObject $oneRow -MemberType NoteProperty -Name 'SentSize' -Value $tempArray[$item + 2]
        Add-Member -InputObject $oneRow -MemberType NoteProperty -Name 'Recieve' -Value $tempArray[$item + 3]
        Add-Member -InputObject $oneRow -MemberType NoteProperty -Name 'RecieveSize' -Value $tempArray[$item + 4]
        $listRows += $oneRow
    }
    $exportPath = $PSScriptRoot + '\MailStat.csv'
    $listRows | Export-Csv -LiteralPath $exportPath -NoTypeInformation -Delimiter ';'
    if ($Auto) {
        SendStatToParser($exportPath)
    }
}


function GetEmalStat {
    $To = $fromDate.AddDays(1) 
    [Int64] $intSent = $intRec = 0
    [Int64] $intSentSize = $intRecSize = 0
    [String] $strEmails = $null 
    Write-Host('Start Date From:' + $fromDate, 'End Date To:' + $To)
    Write-Host "DayOfWeek,Date,Sent,Sent Size (MB),Received,Received Size (MB)" -ForegroundColor Yellow 
    Do { 
        $strEmails = "$($fromDate.DayOfWeek),$($fromDate.ToShortDateString())," 
        if ($exportFlag -eq $true) {$tempArray.Add($strEmails)}
        $intSent = $intRec = 0 
        (Get-TransportService -Identity $TransportService) | Get-MessageTrackingLog -ResultSize Unlimited -Start $fromDate -End $To | ForEach-Object { 
            # Sent E-mails 
            If ($_.EventId -eq "RECEIVE" -and $_.Source -eq "STOREDRIVER") {
                $intSent++
                $intSentSize += $_.TotalBytes
            }
         
            # Received E-mails 
            If ($_.EventId -eq "DELIVER") {
                $intRec += $_.RecipientCount
                $intRecSize += $_.TotalBytes
            }
        } 
        $intSentSize = [Math]::Round($intSentSize / 1MB, 0)
        $intRecSize = [Math]::Round($intRecSize / 1MB, 0)
        $strEmails += "$intSent,$intSentSize,$intRec,$intRecSize" 
        if ($exportFlag -eq $true) {
            $tempArray.Add($intSent)
            $tempArray.Add($intSentSize)
            $tempArray.Add($intRec)
            $tempArray.Add($intRecSize)
        }
        $fromDate = $fromDate.AddDays(1) 
        $To = $fromDate.AddDays(1) 
    } 
    While ($To -lt (Get-Date))
    if ($exportFlag -eq $true) {
        CsvExport($tempArray)
    }
    else {
        $strEmails
    }
}

try {
    Write-Host('Auto flag is:', $Auto, 'From date:', $fromDate, 'CSV export:', $exportFlag) -ForegroundColor Yellow 
    GetEmalStat
}
catch [Exception] {Write-Host $_.Exception}