$scriptPath = Get-Location
$excelFileName = "FixData.xlsm"
$excelPath = Join-Path $scriptPath $excelFileName

function CallExcelFunction{
    $excelFunctionName = $args[0]
    $writeString = $excelFileName + "の関数" + $excelFunctionName + "を呼び出しています."
    Write-Output $writeString

    # Excelオブジェクトを取得
    $excel = New-Object -ComObject Excel.Application
    try
    {
        # ExcelファイルをOPEN
        $book = $excel.Workbooks.Open($excelPath)
        # プロシージャを実行
        $excel.Run($excelFunctionName)
        # ExcelファイルをCLOSE
        $book.Close()
    }
    catch
    {
        $ws = New-Object -ComObject Wscript.Shell
        $ws.popup("エラー : " + $PSItem)
    }
    finally
    {
        # Excelを終了
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null

        $writeString = $excelFileName + "の関数" + $excelFunctionName + "を正常終了."
        Write-Output $writeString
    }
}
function ConvertCore {
    $writeString = "シート名" + $args[0] + "指定の内容のコンバートを実施."
    Write-Output $writeString

    $inputFileName = $args[0] + ".csv"
    $outputDirPath = $args[1]
    $inputFilePath = Join-Path $scriptPath $inputFileName
    $allText = Get-Content $inputFilePath
    $lineTextArray = $allText.split("`n")

    $outputInGameIDText = "";
    $outputInGameIDFileName = $args[0] + ".csv"
    $outputInGameIDFilePath = Join-Path $outputDirPath $outputInGameIDFileName 

    foreach($lineText in $lineTextArray){
        $elementTextArray = $lineText.split(",");

        $str_id = $elementTextArray[0]
        $str_inGameId = $elementTextArray[1]
        $str_sourceRelativePath = $elementTextArray[2]

        $sourceFilePath = Join-Path $scriptPath $str_sourceRelativePath
        $outputFileName = $str_inGameId + ".mp3"
        $outputFilePath = Join-Path $outputDirPath $outputFileName

        $writeString = "ID " +  $str_id + " コンバートを実施."
        Write-Output $writeString
        $writeString = "　入力元は次の通り"
        Write-Output $writeString
        $writeString = "　　" + $sourceFilePath
        Write-Output $writeString
        $writeString = "　出力先は次の通り"
        Write-Output $writeString
        $writeString = "　　" + $outputFilePath
        Write-Output $writeString
        Copy-Item $sourceFilePath -Destination $outputFilePath -Force

        $outputInGameIDText = $outputInGameIDText + $str_inGameId + "`n"
    }

    Write-Output $outputInGameIDText | Out-File $outputInGameIDFilePath -Encoding UTF8
}

if(Test-Path $excelPath){
    CallExcelFunction "OutputFixData"
    
    #成果物をAsset以下に移動.
    $projectRoot = [System.Environment]::GetEnvironmentVariable("CONVERT_AUDIO_ROOT", "Machine")
    $destDirPath = Join-Path $projectRoot "Assets\Resources\Audio"
    if(Test-Path $destDirPath){
        ConvertCore "BGM" $destDirPath
        ConvertCore "SE" $destDirPath
    }else{
        $writeString = $destDirPath + "がありません. ERROR!!!!!"
        Write-Error $writeString
        Read-Host "Enterキーで終了"
    }

}else{
    $writeString = $excelPath + "がありません. ERROR!!!!!"
    Write-Error $writeString
    Read-Host "Enterキーで終了"
}
