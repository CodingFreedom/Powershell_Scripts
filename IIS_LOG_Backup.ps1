param(
[Int]$start_year,
[string]$Backupfolder
)

import-module WebAdministration

if ((!$start_year) -or (!$Backupfolder))
{
	Write-Host "Non hai fornito l'anno di inizio o la cartella di destinazione per il salvataggio dei LOG!"  
	Write-Host "Esempio di uso in Console PS: .\IIS_LOG_Backup.ps1 2011 C:\IIS_LOG_BUP"
	Write-Host "Esempio di uso in Console DOS: powershell -c "".\IIS_LOG_Backup.ps1 2011 C:\IIS_LOG_BUP"""
	Write-Host "Questo eseguibile comprime i LOG IIS a partire dall'anno di inizio fornito e prosegue fino all'anno precedente a quello corrente."
}
else 
{
	if (-not (test-path "$env:ProgramFiles\7-Zip\7z.exe")) {
		Write-Host "Il programma 7-Zip è richiesto per il corretto funzionamento di questo eseguibile"
		Write-Host "Potete scaricarlo sul sito ufficiale del progetto: http://www.7-zip.org/"
	}
	set-alias sz "$env:ProgramFiles\7-Zip\7z.exe"
	$EndYear = (get-date).year - 1
	$IISLogFolder = [System.Environment]::ExpandEnvironmentVariables((get-item 'IIS://sites/default web site').logfile.directory)
	set-location $IISLogFolder
	for ($i=$start_year; $i -le $EndYear; $i++){
		Get-ChildItem -Recurse | ?{ $_.PSIsContainer } | ForEach-Object {
			$IISLogSubFolder = $_
			$Files = Get-Childitem $IISLogSubFolder | Where {$_.LastWriteTime.year -eq $i}
			$FilesCount = ($Files | Measure-Object).count
			if ( $FilesCount -gt 0 ) {
				set-location $IISLogSubFolder
				$Archivio = "$Backupfolder\IIS_LOG_BUP_$IISLogSubFolder-$i.zip"
				sz a -tzip -mx9 $Archivio $Files | out-null
				if($LASTEXITCODE -eq 0){
					write-host "$Archivio Creato per l'anno $i. Files contenuti: $FilesCount"
					sz t $Archivio | out-null
					if($LASTEXITCODE -eq 0){
						write-host "$Archivio Verificato"
						remove-item $Files
					} else {
						write-host "$Archivio non Verificato. Verificare la situazione. Errore ($LASTEXITCODE)"
					}
					if($LASTEXITCODE -eq 0){
						write-host "Rimossi da $IISLogFolder\$IISLogSubFolder i log precedentementi compressi in $Archivio"
					} else {
						write-host "Non sono stati rimossi da $IISLogFolder\$IISLogSubFolder i log precedentementi compressi in $Archivio. Controllare la situazione. Errore ($LASTEXITCODE)"
					}
				} else {
					write-host "$Archivio non creato. Verificare la situazione. Errore ($LASTEXITCODE)"
				}
				set-location $IISLogFolder
			} else {
				write-host "Non sono stati rilevati file da comprimere per l'anno $i."
			}
		}
	}
}