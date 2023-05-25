$scriptPath = Get-Location
$excelFileName = "FixData.xlsm"
$excelPath = Join-Path $scriptPath $excelFileName

function CallExcelFunction{
    $excelFunctionName = $args[0]
    $writeString = $excelFileName + "�̊֐�" + $excelFunctionName + "���Ăяo���Ă��܂�."
    Write-Output $writeString

    # Excel�I�u�W�F�N�g���擾
    $excel = New-Object -ComObject Excel.Application
    try
    {
        # Excel�t�@�C����OPEN
        $book = $excel.Workbooks.Open($excelPath)
        # �v���V�[�W�������s
        $excel.Run($excelFunctionName)
        # Excel�t�@�C����CLOSE
        $book.Close()
    }
    catch
    {
        $ws = New-Object -ComObject Wscript.Shell
        $ws.popup("�G���[ : " + $PSItem)
    }
    finally
    {
        # Excel���I��
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null

        $writeString = $excelFileName + "�̊֐�" + $excelFunctionName + "�𐳏�I��."
        Write-Output $writeString
    }
}
function ConvertCore {
    $writeString = "�V�[�g��" + $args[0] + "�w��̓��e�̃R���o�[�g�����{."
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

        $writeString = "ID " +  $str_id + " �R���o�[�g�����{."
        Write-Output $writeString
        $writeString = "�@���͌��͎��̒ʂ�"
        Write-Output $writeString
        $writeString = "�@�@" + $sourceFilePath
        Write-Output $writeString
        $writeString = "�@�o�͐�͎��̒ʂ�"
        Write-Output $writeString
        $writeString = "�@�@" + $outputFilePath
        Write-Output $writeString
        Copy-Item $sourceFilePath -Destination $outputFilePath -Force

        $outputInGameIDText = $outputInGameIDText + $str_inGameId + "`n"
    }

    Write-Output $outputInGameIDText | Out-File $outputInGameIDFilePath -Encoding UTF8
}

if(Test-Path $excelPath){
    CallExcelFunction "OutputFixData"
    
    #���ʕ���Asset�ȉ��Ɉړ�.
    $projectRoot = [System.Environment]::GetEnvironmentVariable("CONVERT_AUDIO_ROOT", "Machine")
    $destDirPath = Join-Path $projectRoot "Assets\Resources\Audio"
    if(Test-Path $destDirPath){
        ConvertCore "BGM" $destDirPath
        ConvertCore "SE" $destDirPath
    }else{
        $writeString = $destDirPath + "������܂���. ERROR!!!!!"
        Write-Error $writeString
        Read-Host "Enter�L�[�ŏI��"
    }

}else{
    $writeString = $excelPath + "������܂���. ERROR!!!!!"
    Write-Error $writeString
    Read-Host "Enter�L�[�ŏI��"
}
