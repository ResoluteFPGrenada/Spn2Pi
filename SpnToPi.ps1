cls

. "$PSScriptRoot\import-excel.ps1"
$ConfigJSON = get-content ".\config.txt" | out-string | ConvertFrom-Json 
$CopyFileSource = $ConfigJSON.CopySourceFile
$File = "$PSScriptRoot\import.xlsm"
$CSVFile = "$PSSCriptRoot\spn2pi.csv"

    #COPY file to another location#
copy-item -Path $CopyFileSource -Destination $File -force

$Date = (get-date ((get-date).AddDays(-1)) -format "yy/MM/dd") + " 00:00:00"
write-host "Yesterday's Date = $Date"
$Worksheet = "OUTGENOEE_Daily"

$SPNData = import-SPNexcel -File $File -Worksheet $Worksheet -Date $Date

#Values#
$NumSets = $SPNData.NumSets
$WRChTimeDay = $SPNData.WRChTimeDay
$WdrProdDay = $SPNData.WdrProdDay
$PMProdDay = $SPNData.PMProdDay
$UWLengthDay = $SPNData.UWLengthDay
$UWChTimeDay = $SPNData.UWChTimeDay
$RunTimePercent = $SPNData.RunTimePercent
$LORBLengthDay = $SPNData.LORBLengthDay
$NumShtBrk100Sets = $SPNData.NumShtBrk100Sets

#Pi Tags#
$NumSetsTag = "SPN:NumSets.NM"
$WRChTimeDayTag = "SPN:WRChTimeDay.TM"
$PMProdDayTag = "SPN:PMProDay.MPM"
$WdrProdDayTag = "SPN:WdrProdDay.MPM"
$UWLengthDayTag = "SPN:UWLength.MM"
$UWChTimeDayTag = "SPN:UWChTimeDay.TM"
$RunTimePercentTag = "SPN:RunTimePercent.TM"
$LORBLengthDayTag = "SPN:LORBLength.MM"
$NumShtBrk100SetsTag = "SPN:NumShtBrk100Sets.NM"

$PiTimeStamp = (get-date ((get-date).AddDays(-1)) -format "MM-dd-yy") + " 00:00:00"

new-item -path $CSVFile -ItemType file -Force
set-content $CSVFile -Value "$NumSetsTag, $NumSets, $PiTimeStamp" -force
add-content $CSVFile -Value "$WRChTimeDayTag, $WRChTimeDay, $PiTimeStamp" -force
add-content $CSVFile -Value "$WdrProdDayTag, $WdrProdDay, $PiTimeStamp" -force
add-content $CSVFile -Value "$PMProdDayTag, $PMProdDay, $PiTimeStamp" -force
add-content $CSVFile -Value "$UWChTimeDayTag, $UWChTimeDay, $PiTimeStamp" -force
add-content $CSVFile -Value "$RunTimePercentTag, $RunTimePercent, $PiTimeStamp" -force
add-content $CSVFile -Value "$LORBLengthDayTag, $LORBLengthDay, $PiTimeStamp" -force
add-content $CSVFile -Value "$NumShtBrk100SetsTag, $NumShtBrk100Sets, $PiTimeStamp" -force
add-content $CSVFile -Value "$UWLengthDayTag, $UWLengthDay, $PiTimeStamp" -force

#output values#
#write-host "Num Sets: $NumSets"
#write-host "WRChTimeDay: $WRChTimeDay"
#write-host "WdrProdDay: $WdrProdDay"
#write-host "PMProdDay: $PMProdDay"
#write-host "UWLengthDay: $UWLengthDay"
#write-host "UWChTimeDay: $UWChTimeDay"
#write-host "RunTimePercent: $RunTimePercent"
#write-host "LORBLengthDay: $LORBLengthDay"
#write-host "NumShtBrk100Sets: $NumShtBrk100Sets"

