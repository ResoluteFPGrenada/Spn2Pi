# import parameters #
$ImportPath = "$PSScriptRoot\importShift.txt"
$ConfigJSON = get-content ".\config.txt" | out-string | ConvertFrom-Json 
$ReportPath = $ConfigJSON.CopySourceFile
$File = "$PSScriptRoot\import.xlsm"
$LogPath = "$PSScriptRoot\log.txt"
$params = get-content $ImportPath
$Date = get-date
$SpreadSheet = "Calc_Shifts"

$imports = @()
foreach($param in $params){
    $ColName, $PiTags, $ColNum = $param -split("~")
    $prop = @{
        PiTag = "$PiTags"
        ColNumber = [int]$ColNum
        Value = $Null
        TimeStamp = $Null
    }#END_HASHTABLE#

    $obj = new-object -TypeName psobject -Property $prop
    
    $imports += $obj
}#END_FOREACH#


# check path to SPN report #
$TP = test-path $ReportPath

if($TP){
    add-content $LogPath -Value "Shift: Found report file, $Date" -force

    # copy SPN report to this directory #
    try{
        copy-item -Path $ReportPath -Destination $File -force -errorAction Stop

        add-content $LogPath -Value "Shift: Report was copied to local directory, $Date" -force
    }catch{
        add-content $LogPath -Value "Shift: there was an issue copying the file to this directory, $Date" -force
    }

    # Run read report function #
    $Data = read-report -SpreadSheet "$Spreadsheet" -FileName "$File" -InputObject $imports

    # store column value with Pi tag name and timestamp in csv #!!!!!!!!!
    write-output $Data

}else{
    add-content $LogPath -Value "Shift: Can not find report file, $Date" -force
}#END IF/ELSE#

function read-report{
    [cmdletbinding()]
    param(
        [string]$SpreadSheet,
        [string]$FileName,
        $InputObject
    )PROCESS{
        #Set Variables#
        $TimeStampCol = 3
        
        # Open Spreadsheet #
        $objExcel = New-Object -ComObject Excel.Application
        $objExcel.visible = $false
        $WorkBook = $objExcel.Workbooks.Open("$FileName")
        $Sheet = $Workbook.sheets.Item("$SpreadSheet")
        $RowCount =  ($Sheet.UsedRange.Rows).count

        # Get yesterday and today's date #
        $Today = get-date -format "%M/%d/yyyy"
        $Yesterday = get-date ((get-date).AddDays(-1)) -format "%M/%d/yyyy"
        $Time = get-date -format "HH:00"

        # read last line. select column based on import object. Save value #
        foreach($o in $InputObject){
            $o.value = $Sheet.cells.Item($RowCount, $o.ColNumber).text
            $o.TimeStamp = $Sheet.cells.Item($RowCount, $TimeStampCol).text
        }# END_FOREACH #

        # reverse loop through rows until dates needed are found #
        #$Objs = @()
        #for($i = $RowCount; $i -ge 0; $i--){
            
        #}#END_FOR#

        
        $Workbook.close($False)
        $objExcel.quit()
        write-output $InputObject

    }# END_PROCESS #
}#END_FUNCTION#
