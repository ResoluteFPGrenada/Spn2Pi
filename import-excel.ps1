function import-SPNexcel{
    [cmdletbinding()]
    param(
        [string]$File,
        [string]$Worksheet,
        [string]$Date
    )
    PROCESS{

        $LogFile = "$PSScriptRoot\log.txt"
        $LogDate = get-date -format "MM/dd/yyyy hh:mm"
        try{
            $TP = test-path $File 
            if($TP -eq $True){
                write-host "$File is found"
                add-content $LogFile -Value "$File is found, $LogDate"

                $objExcel = New-Object -ComObject Excel.Application
                $objExcel.visible = $false
                $WorkBook = $objExcel.Workbooks.Open("$File")
                $Sheet = $Workbook.sheets.Item("$WorkSheet")
                $RowCount =  ($Sheet.UsedRange.Rows).count
                $Objs = @()
                for($i = $RowCount; $i -ge 0; $i--){
                    $TestValue = $Sheet.cells.Item($i, 1).text
                    if($Sheet.cells.Item($i, 1).text -eq "$Date"){
                        $props = @{
                            EventTime = $Sheet.cells.Item($i, 1).text
                            WdrProdDay = $Sheet.cells.Item($i, 6).text
                            PMProdDay = $Sheet.cells.Item($i, 7).text
                            RunTimePercent = $Sheet.cells.Item($i, 8).text
                            LORBLengthDay = $Sheet.cells.Item($i, 9).text
                            UWLengthDay = $Sheet.cells.Item($i, 10).text
                            NumSets = $Sheet.cells.Item($i, 12).text
                            NumShtBrk100Sets = $Sheet.cells.Item($i, 13).text
                            WRChTimeDay = $Sheet.cells.Item($i, 14).text
                            UWChTimeDay = $Sheet.cells.Item($i, 15).text

                        }#END_HASHTABLE#

                        $obj = New-Object -TypeName PSObject -Property $props
                        $Objs += $obj
                        #ADD Logging#
                        write-host "values stored"
                        add-content $LogFile -Value "Values Stored, $LogDate"

                        break
                    }#END_IF#
                     
                }#END_FOR#
                
                $Workbook.close($False)
                $objExcel.quit()
                write-output $Objs
            }else{
                write-host "No Path was found at: $File"
                add-content "No Path was found at :$File, $LogDate"
            }#END_IF/ELSE#

            
        }catch{
            write-host "An error has occured"
            add-content "An error has occured, $LogDate"

        }#END_TRY/CATCH#
    }#END_PROCESS#
}#END_FUNCTION#