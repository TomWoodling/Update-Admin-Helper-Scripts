echo "Starting Script"


#Get contents of Servers_Control spreadsheet (on Google Drive) and get the server details
$shejs = "https://spreadsheets.google.com/feeds/list/1ajcHPKp4R3YZ7GckuO6e9cjj6eqPPaBy_5sRN9wbOQg/od6/public/values?alt=json"
$sozz = Invoke-WebRequest -Uri $shejs | ConvertFrom-Json
$dtsers = $sozz.feed.entry | select -expandproperty 'content'
$outhed = "Name, Day, Time, Update, NagName, Excluded_Ups, SS, Notify, FQDN, Patch_Name"

#Make empty arrays
$jsoncol = New-Object System.Collections.ArrayList
$poots = New-Object System.Collections.ArrayList
$pouts = New-Object System.Collections.ArrayList
$pouts.Add("$outhed") > null

#Daylight Saving Time offset
$daysa = (Get-Date).IsDaylightSavingTime()
[int]$difo = 3600
If ($daysa -eq $true) {$difo = 7200}
    Else {$difo = 3600}

#Makes each server work as an individual entry
function indiv {
foreach ($dtser in $dtsers) {
    $sting = ($dtser.'$t' | Out-String)
    $jsoncol.Add("$sting")
        }
    }
indiv > null
echo "step 1 complete"

#Works on formatting the entries - strips field names
function stripper {
foreach ($js in $jsoncol) {
    $hoots = New-Object System.String("")
    $pecs = $js.split(",",[System.StringSplitOptions]::RemoveEmptyEntries)
    foreach ($pec in $pecs) {
        $pic = $pec -replace ".*:"
        $hoots += ",$pic"
        }
    $poots.Add($hoots)
        }
    }
stripper > null
echo "step 2 complete"

#Removes leading comma and appends to the last array
function commago {
foreach ($poot in $poots) {
    $pout = $poot.TrimStart(", ")
    $pouts.Add("$pout")
        }
    }
commago > null
echo "all 3 steps done"

#Export csv file then re-import it because that's how you get a table!
$pouts > "$env:systemroot\TEMP\grah.csv"
$RAR = import-csv -path "$env:systemroot\TEMP\grah.csv"

#The update plan can now be manipulated
#$RAR | Out-GridView

echo "Read details - move to booking downtime"

#Gets Wsus Servers and their patching windows
$Wsusers = $RAR | select * | where 'Update' -eq 'Wsus' 
$timoos = $Wsusers | select 'NagName', 'Day', 'Time' | where 'NagName' -ne 'unmonitored'

echo "Take the text that appears in notepad, copy to nagios server, chmod and run it to book mass downtime"
#Create a temp file to hold the downtime script
$wowo = (Get-Date).day 
$dtloc = "$env:systemroot\TEMP\windt$wowo.txt"
If (test-path $dtloc) {Remove-Item $dtloc}
Else {}

#Populates the header for temp file
Write-output "#!/bin/sh
# This is a simple script to pass SCHEDULE_HOST_DOWNTIME and SCHEDULE_HOST_SVC_DOWNTIME commands
# to Nagios.  Built from powershell script at:
now=``date +%s``
commandfile='/usr/local/nagios/var/rw/nagios.cmd'

" | out-file $dtloc

#Sections of the Nagios commands
$aaaa = '/usr/bin/printf "[%lu] SCHEDULE_HOST_DOWNTIME;'
$bbbb = ';1;0;7200;'
$dddd = ';Scheduled Updates\n" $now > $commandfile'
$cccc = '/usr/bin/printf "[%lu] SCHEDULE_HOST_SVC_DOWNTIME;'

#Builds each individual command in the temp file
$serbles = $timoos
ForEach ($serble in $serbles) {
$tyger = $serble.Time
$grn = ':00'
$booker = @(@(0..7) | % {$(Get-Date $tyger$grn).AddDays($_)} | ? {($_ -gt $(Get-Date)) -and ($_.DayOfWeek -ieq $serble.Day)})[0]
[int]$starti = get-date -Uformat %s $booker
[int]$starto = $starti-$difo
[int]$endr = $starto
Write-output "$($aaaa)$($serble.nagname);$($starto);$($endr+7200)$($bbbb)$env:USERNAME$($dddd)" | out-file $dtloc -append
Write-output "$($cccc)$($serble.nagname);$($starto);$($endr+7200)$($bbbb)$env:USERNAME$($dddd)" | out-file $dtloc -append
}

#Launched for copying
notepad $dtloc
