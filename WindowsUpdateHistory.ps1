$UpdateSession = New-Object -ComObject 'Microsoft.Update.Session'
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
$HistoryCount = $UpdateSearcher.GetTotalHistoryCount()
$History = $UpdateSearcher.QueryHistory(0, $HistoryCount)

$CleanHistory = foreach($Entry in $History){
    [PSCustomObject]@{
        Title = $Entry.Title
        Date = $Entry.Date.ToShortDateString()
        Result = if ($Entry.ResultCode -eq 2) {"Installed"} else {"Not Installed"}
    }
}

$CleanHistory | Format-Table -AutoSize

