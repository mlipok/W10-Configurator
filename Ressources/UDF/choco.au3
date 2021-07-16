; INSTALL & CONFIGURE CHOCOLATEY AS DESIRED SOFTWARE INSTALLER


Global $cho = 'powershell Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol' & _
'= [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ' & _
"((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))"
Func choco()

	$Powershell = Run($cho, @SystemDir, @SW_HIDE, $STDERR_MERGED)

	While 1

		$sStreamOut = StdoutRead($Powershell)

		If @error Then
			ExitLoop
		Else
			$ReplaceStream = StringRegExpReplace(StringRegExpReplace($sStreamOut, '[\s]*$|([\n\r]){2,}', '$1'), "[^0-9a-zA-Z:,\/-_%&+#@!*;(){}." & '"' & "'" & @LF & " ]+","")
;~ $ReplaceStream = StringRegExpReplace($sStreamOut, '[\s]*$|([\n\r]){2,}', '$1')
			If StringLen($ReplaceStream) > 1 Then c($ReplaceStream)
		EndIf
	WEnd

	Local $Pshell = Run("powershell -Command " & "choco" & " list -r | Out-String -Stream | foreach {$_.trimend(|)} | Where {$_ -ne ''}", @SystemDir, @SW_HIDE, $STDERR_MERGED)
	c("Collecting list of available softwares... may take few minutes.")
	ProcessWaitClose($Pshell)
	$sRZCatalog = StringSplit(StdoutRead($Pshell), @CRLF, 1)
	If $sRZCatalog=Null Then c(StderrRead($Pshell))

	Local $ColCount = UBound($sRZCatalog, 1) - 1
	_ArrayDelete($sRZCatalog, $ColCount) ;Removes the blank line at the end

	If Not $sRZCatalog = Null Or IsArray($sRZCatalog) = True Then

;~ 		$RZUpdated = True
		c("Retrieve complete, " & $sRZCatalog[0] & " softwares has been found.")
		_ArrayDisplay($sRZCatalog)
		AdlibRegister("AutoComplete")

	Else

		c("Unable to retrieve catalog...")
	EndIf
EndFunc
