 Function Get-FileName($initialDirectory)
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
 Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "All files (*.*)| *.*"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
} #end function Get-FileName

function Save-File([string] $initialDirectory ) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowDialog() |  Out-Null
	
	$nameWithExtension = "$($OpenFileDialog.filename).csv"
	return $nameWithExtension
}

#Open a file dialog window to get the source file
$serverList = Get-Content -Path (Get-FileName -initialDirectory "C:\")

#open a file dialog window to save the output
$fileName = Save-File $fileName

$array =@() 
$keyname = 'SOFTWARE\\Microsoft\\Internet Explorer'

$i = 0

foreach ($server in $serverList){
	$i++
	Write-Progress -activity "checking IE version on server $i of $($serverList.count)" -percentComplete ($i / $serverList.Count*100)
	$obj = New-Object PSObject
	if(Test-Connection -computer $server -count 1 -quiet){
		try{
			$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $server)
			$key = $reg.OpenSubKey($keyname)
			$value = $key.GetValue('Version')
		}
		Catch{
			$value = $_.Exception.Message
		}
	}
	else{
	$value = 'offline'	
	}
	$holder = $value.split(".")
	if($holder[0] -eq "9"){
		if($holder[1] -eq "10")
		{
		$newVal = "IE10 v."+$value
		$value = $newVal
		}
		if($holder[1] -eq "11"){
		$newVal = "IE11 v."+$value
		$value = $newVal
		}
	}
	
	$obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -value $server
	$obj | Add-Member -MemberType NoteProperty -Name "IEVersion" -value $value
	$array += $obj
}
$array | select ComputerName,IEVersion | Export-csv $fileName
