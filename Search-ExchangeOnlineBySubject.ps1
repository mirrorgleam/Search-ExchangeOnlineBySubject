<#
Title: Search-ExchangeOnlineBySubject.ps1
Author: Melvin Baughman
Date: 06/28/17

Function: Allows performing a regex search in Exchange 365 on an email subject
            and provide an estimated time remaining to completion of search.

# Notes: 
    > Can only search past 7 days
        > Limitation of Get-MessageTrace 
    > Requires permissions to run Get-MessageTrace cmdlet
	> StartDate and EndDate are converted to UTC since that is the time that is used by Get-MessageTrace
	> Search is not case sensitive
	> There is overlap on Get-MessageTrace captured to avoid missed findings
	> Search takes between 2-4 mins / hour of time searched
		> Sometimes higher if the Exchange server is busy
#>

function Search-ExchangeOnlineBySubject
{
    Param([DateTime]$StartDate = (Get-Date).addhours(-2),
          [DateTime]$EndDate = (Get-Date),
		  [string]$TargetDirectory = "$([Environment]::GetFolderPath("MyDocuments"))\EmailSearchResults",
          [Parameter(Mandatory=$true)]
		  [string]$Subject)
    

	#region Define Function Parameters 
    
    # Get UTC Offset
    [int]$UTCOffset = (Get-Date ([datetime]::UtcNow) -UFormat %H) - (get-date -UFormat %H)

	  # Count number of regex matches found
	  [int]$matchCount = 0

    # Variable to store the EndDate
    [datetime]$LastDatePulled = $EndDate.AddHours($UTCOffset)

    # Variable to store the StartDate
    [datetime]$DateToSearchBackTo = $StartDate.AddHours($UTCOffset)

    # Create a unique name for each file
    $DateTimeCmdStarted = Get-Date -Format "yy-MM-dd HH.mm.ss"

	#endregion


    # Verify date range not outside available data
    if ($DateToSearchBackTo -lt ((Get-Date).Date.AddDays(-7)) -or $EndDate -gt (Get-Date) -or $DateToSearchBackTo -gt $LastDatePulled)
    {
        Write-Host "`nDate Range must be between $((Get-Date).Date.addDays(-7).AddHours(-$UTCOffset)) and $(Get-Date)"
        break
    }
    
	#region Check if Microsoft Exchange session already exists; if not, create one and log in
    
	if (!(Get-PSSession |where {$_.Availability -eq "Available" -and $_.ConfigurationName -eq "Microsoft.Exchange"}))
    {
        $UserCredential = Get-Credential 
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.protection.outlook.com/powershell-liveid/ `
	        -Credential $UserCredential -Authentication Basic -AllowRedirection
	        Import-PSSession $Session 
    }

	#endregion

	#region If destination folder doesn't exist create it

	if (!(Test-Path -PathType Container $TargetDirectory))
	{
		New-Item -ItemType Directory -Path $TargetDirectory
	}

	#endregion


	# Display to user time frame to be searched
    $titleOfProgress = "Searching Timeframe   From: $($DateToSearchBackTo) UTC    To: $($LastDatePulled) UTC"

    Write-progress -Activity $titleOfProgress

    # $MC is used to collect time taken to perform all Get-MessageTrace loops
    $MC = Measure-Command {

        # Buld loop to continue pulling more results until the date range is met
        do {
            
            [TimeSpan]$searchSpaceRemaining = $LastDatePulled - $DateToSearchBackTo
 
            # Retain EndDate value for current Get-MessageTrace query
            $previousLastDatePulled = $LastDatePulled

            # Get TimeSpan taken to perform last loop
            Measure-Command -OutVariable timeToRunLastLoop {
                
                # Write to variable 5000 results
                $tempStore = Get-MessageTrace -StartDate $DateToSearchBackTo -EndDate $LastDatePulled -PageSize 5000
        
                # Perform regex query on previously pulled email messages and write output to csv file
                $matchCountTemp = $tempStore |where {$_.Subject -match $Subject} 
				
				$matchCountTemp | Export-Csv -Path "$($TargetDirectory)\EmailSearch_$($DateTimeCmdStarted).csv" -NoClobber -NoTypeInformation -Append

                # Increment the displayed match count
				$matchCount += $($matchCountTemp |measure).Count
            
            } # END Inner Loop Measure-Command


            # Stop writing lines when query is finished
            if ($tempStore.Count -eq 5000) {
                
                # Set DateTime for next Get-MessageTrace loop
                $LastDatePulled = $tempStore.Item($tempStore.count - 1).Received
        
                # Get TimeSpan for search space covered during last loop
                [TimeSpan]$timeCoveredDuringLastLoop = $previousLastDatePulled - $LastDatePulled

				# Caluculate seconds for approximate time remaining in query
                [int]$secondsRemaining = $($searchSpaceRemaining.TotalSeconds) / $($timeCoveredDuringLastLoop.TotalSeconds) * $($timeToRunLastLoop.TotalSeconds)
	    		
                Write-Progress -Activity $titleOfProgress -Status "Matches Found: $matchCount     Approximately" -SecondsRemaining $secondsRemaining
            }
			

        } while ($tempStore.Count -eq 5000)
    
    } # End external loop Measure-Command

    # Build and format total run time string
	[string]$MCTimeString = "Run Time:`n{0:dd} {0:hh}:{0:mm}:{0:ss}" -f $MC
    # Write out total time to run
    Write-Host "Matches Found: $matchCount`n$MCTimeString"
}
