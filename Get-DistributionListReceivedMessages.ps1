<#
    .SYNOPSIS
    Identifies number of messages delivered to Distribution Lists in specified timeframe

    .PARAMETER PreviousDays
    Number of days to search back in the Message Tracking logs (Maximum of 10 days). Default value if not specified is 7 days.

    .PARAMETER Path
    Filepath to export report of received message data to

    .EXAMPLE
    Get-DistributionListReceivedMessages -PreviousDays 7

    Returns count of any inbound messages per recipient in the past 7 days

    .EXAMPLE
    Get-DistributionListReceivedMessages -PreviousDays 7 -Path "C:\DLReceivedMessagesReports\"

    Returns count of any inbound messages per recipient in the past 7 days and saves output to a csv file in a specified location
#>

function Get-DistributionListReceivedMessages {
    param (
        [int]$PreviousDays = 7,
        [System.IO.FileInfo]$Path
    )

    #Connect to Exchange Online
    Connect-ExchangeOnline

    #Get all DL primary SMTP addresses
    $distributionLists = Get-DistributionGroup -ResultSize Unlimited | Select-Object -ExpandProperty PrimarySmtpAddress

    #Create empty ArrayList to store data from Message Trace in
    $arrMessageData = New-Object System.Collections.ArrayList

    #Run Message Trace for each DL
    foreach ($distributionList in $distributionLists) {

        #Strongly typed so that message count reports correctly if == 1
        [array]$messages = @()
        $messages = Get-MessageTrace -RecipientAddress $distributionList -Status Expanded -StartDate (Get-Date).AddDays(-$PreviousDays) -EndDate (Get-Date)
        $messageCount = $messages.Count

        #Store data in PSCustomObject
        $psObjMessageData = [PSCustomObject]@{
            RecipientAddress = $distributionList
            MessageCount     = $messageCount
        }

        #Add PSCustomObject to previously defined ArrayList. Cast to void to avoid returning index value in console.
        [void]$arrMessageData.Add($psObjMessageData)
    }

    #Output data in tabular format in console
    $arrMessageData | Format-Table -AutoSize

    #If -Path parameter is specified, export a CSV to that filepath
    if ($Path) {

        #Get date for use in export filename
        $date = Get-Date -Format MMddyyyy

        #Generate filename
        $exportPath = "$($Path)DLReceivedMessages_$date.csv"

        #Export data as a CSV
        $arrMessageData | Export-Csv -Path $exportPath -NoTypeInformation
    }

    #Disconnect from ExchangeOnline
    Disconnect-ExchangeOnline -Confirm:$false
}