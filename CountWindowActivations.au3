#Include <Array.au3>
#Include <File.au3> ; For _PathSplit to determine script name
#Include <Date.au3> ; For _NowCalc()

Opt("TrayOnEventMode",1)
Opt("TrayAutoPause",0)
Opt("TrayMenuMode",1) ; Default tray menu items (Script Paused/Exit) will not be shown.
Dim $ModalPopup ; Declare $ModalPopup, use of a modal popup will be read from .ini file later
Dim $MonitoredWindowTitle, $WindowTitleToCheckFor, $WindowCounter, $startTime,$oneToLog ; Global variables because of logging

$windowtitle="Count Window Activations 1.41"
TraySetToolTip($windowtitle)

; Read scriptname (of executable)
Dim $FP[4]
_PathSplit(@ScriptFullPath, $FP[0], $FP[1], $FP[2], $FP[3])
$inifile=@UserProfileDir & "\" & $FP[2] & ".ini" ; ini filename: defaults to %userprofile%\CountWindowActivations.ini
$logfile=@UserProfileDir & "\" & $FP[2] & "-log.csv" ; ini filename: defaults to %userprofile%\CountWindowActivations.ini

; Create a first line with headers for the logfile
If NOT FileExists ($logfile) Then
	$fileHandle = FileOpen($logfile, 1) ; Open in Write mode (append to end of file)
	If $fileHandle <> -1 Then ; Check if file opened for writing OK
		FileWriteLine($fileHandle, "Date;Time;MonitoredWindowTitle;ActualWindowTitle;TimesOpened;SecondsOpened")
		FileClose($fileHandle)
	EndIf
EndIf

; If no ini file is present offer to create the default one.
$success = 0
Do
	$WindowTitles = IniReadSection($inifile,"WindowTitles")
	If @error Then 
		$createfile=MsgBox(4096+16+4, $windowtitle, "The file: " & $inifile & @LF & "could not found or does not contain WindowTitles to monitor." & @LF & "Create file with default settings?")
		If $createfile=7 Then Exit
		If $createfile=6 Then CreateIni()
	Else
		$success = 1
	EndIf
Until $success = 1

; Read whether the user wants a modal popup
If IniRead($inifile,"Settings", "ModalPopup","") = "True" Then $ModalPopup = "True"

; Set the key values in the array to 0 (use as a counter)
For $i = 1 To $WindowTitles[0][0]
	$WindowTitles[$i][0] = 0
Next

; Create the menu
TraySetClick(1+4+8+32) ; Show tray menu at any mouseclick
$infoitem = TrayCreateItem("Open counter popup")
TrayItemSetOnEvent(-1,"WindowTitlesCounter")
$infoitem = TrayCreateItem("Open log file")
TrayItemSetOnEvent(-1,"OpenLog")
TrayCreateItem("")
$infoitem = TrayCreateItem("Open settings file")
TrayItemSetOnEvent(-1,"OpenIni")
$infoitem = TrayCreateItem("About")
TrayItemSetOnEvent(-1,"ShowInfo")
$exititem = TrayCreateItem("Exit")
TrayItemSetOnEvent(-1,"ExitScript")
TraySetState() ; To show the tray icon

;Not used (yet). Code to check for todays date from TOPlunch Subscriber (to reset counters each day instead of each run)
;If IniRead($inifile,"TOPlunch", "LastDecided","0") == _NowCalcDate() Then Exit ; Exit if question is already answered today
;IniWrite ($inifile,"TOPlunch", "LastDecided", _NowCalcDate()) ; Remember when 'No' was decided to prevent repeating the question today

While 1 ; Main program loop
	$handle = WinGetHandle ("") ; Internal Windows handle of currently active Window
	$ActiveWindowTitle = WinGetTitle ($handle) ; WindowTitle of currently active Window

	; To do: loop through all Windows in .ini file
	For $i = 1 To $WindowTitles[0][0]
		$WindowTitleToCheckFor = $WindowTitles[$i][1] ; Value to check for from array of Windows from .ini file
		If StringInStr ($ActiveWindowTitle, $WindowTitleToCheckFor, 0, 1) Then ; If the ActiveWindowTitle contains the WindowTitleToCheckFor (case INsensitive)
			$startTime = _NowCalc() ; Start time measure
			$oneToLog=1 ; This is a window to log
			If $ModalPopup = "True" Then
				$yesno = MsgBox(4+32+256+4096,$windowtitle, "Do you really want to spend time in " & $WindowTitleToCheckFor & "?") ; Yes/No, Question-marked Modal message box
				If $yesno = 7 Then WinSetState ($handle, "", @SW_MINIMIZE) ; Minimize Window if user answers "No"
			EndIf
			$WindowTitles[$i][0] = $WindowTitles[$i][0] + 1 ; Increase counter by one
			$WindowCounter = $WindowTitles[$i][0] ; for logging later
			$MonitoredWindowTitle = $WindowTitleToCheckFor
		EndIf
	Next
	WinWaitNotActive ($handle) ; Wait until the currently active Window changes to go at it again
	If $oneToLog=1 Then
		$currentTime=_NowCalc()
		
		; Split the date of "now"
		Dim $timeNowArray, $dateNowArray
		_DateTimeSplit($currentTime,$dateNowArray,$timeNowArray)
		$dateNow = $dateNowArray[1] & "-" & $dateNowArray[2] & "-" & $dateNowArray[3]
		$timeNow = $timeNowArray[1] & ":" & $timeNowArray[2] & ":" & $timeNowArray[3]

		; Split the date of "started"
		Dim $timeStartedArray, $dateStartedArray
		_DateTimeSplit($startTime,$dateStartedArray,$timeStartedArray)
		$dateStarted = $dateStartedArray[1] & "-" & $dateStartedArray[2] & "-" & $dateStartedArray[3]
		$timeStarted = $timeStartedArray[1] & ":" & $timeStartedArray[2] & ":" & $timeStartedArray[3]

		; Check if the date changed at 0:00
		If $dateNow = $dateStarted Then
			$duration = _DateDiff( 's',$startTime,$currentTime)
			$fileHandle = FileOpen($logfile, 1) ; Open in Write mode (append to end of file)
			If $fileHandle <> -1 Then ; Check if file opened for writing OK
				FileWriteLine($fileHandle, "" & $dateNow & ";" & $timeNow & ";" & $MonitoredWindowTitle & ";" & $ActiveWindowTitle & ";" & $WindowCounter & ";" & $duration)
				FileClose($fileHandle)
			EndIf
		Else ; If the date changed, split log
			$durationTill12 = _DateDiff( 's',$startTime,$dateStarted & " 23:59:59")
			$durationFrom12 = _DateDiff( 's',$dateNow & " 00:00:00",$currentTime)
			$fileHandle = FileOpen($logfile, 1) ; Open in Write mode (append to end of file)
			If $fileHandle <> -1 Then ; Check if file opened for writing OK
				FileWriteLine($fileHandle, "" & $dateStarted & ";" & $timeStarted & ";" & $MonitoredWindowTitle & ";" & $ActiveWindowTitle & ";" & $WindowCounter & ";" & $durationTill12)
				FileWriteLine($fileHandle, "" & $dateNow & ";" & $timeNow & ";" & $MonitoredWindowTitle & ";" & $ActiveWindowTitle & ";" & $WindowCounter & ";" & $durationFrom12)
				FileClose($fileHandle)
			EndIf
			For $i = 1 To $WindowTitles[0][0] ; Loop through counters to reset
				If $WindowTitleToCheckFor = $WindowTitles[$i][1] Then ; If the WindowTitleToCheckFor
					$WindowTitles[$i][0] = 1 ; Set counter to 1
				Else
					$WindowTitles[$i][0] = 0 ; Reset counter to zero
				EndIf
			Next
		EndIf
		
		$oneToLog=0
	EndIf
WEnd

; Menu functions
Func OpenIni()
    ShellExecute ($inifile)
EndFunc
Func OpenLog()
    ShellExecute ($logfile)
EndFunc
Func ShowInfo()
    Msgbox(0,"About",$windowtitle & @LF & "By Patrick Mackaaij" & @LF & @LF & "http://www.eenmanierom.nl/" & @LF & "http://twitter.com/mackaaij/")
EndFunc
Func ExitScript()
    Exit
EndFunc
Func WindowTitlesCounter()
	; Output all WindowTitles plus counter
	$TrayTip = ""
	For $i = 1 To $WindowTitles[0][0]
		$TrayTip = $TrayTip & $WindowTitles[$i][1] & ": " & @TAB & $WindowTitles[$i][0] & @LF
	Next
		MsgBox(64,$windowtitle,"Counted the following activations:" & @LF & @LF & $TrayTip)
EndFunc
	
Func CreateIni()
	; Read return value of IniWrite to display success or failure
	$var = IniWrite($inifile, "WindowTitles", "Window1", "Gmail - Inbox")
	If $var = 0 Then
		MsgBox(4096+16, $windowtitle, "Could not create file " & $inifile " in folder """ & @UserProfileDir & """." & @LF & "Please check file access rights. Program will exit.")
	ElseIf $var = 1 Then
		IniWrite($inifile, "WindowTitles", "Window1", "Inbox - Microsoft Outlook")
		IniWrite($inifile, "WindowTitles", "Window2", "Gmail - Inbox")
		IniWrite($inifile, "WindowTitles", "Window3", "Google Reader")
		IniWrite($inifile, "WindowTitles", "Window4", "FeedDemon")
		IniWrite($inifile, "WindowTitles", "Window5", "Twitter")
		IniWrite($inifile, "WindowTitles", "Window6", "TweetDeck")
		IniWrite($inifile, "WindowTitles", "Window7", "Yammer")
		IniWrite($inifile, "WindowTitles", "Window8", "nu.nl")
		IniWrite($inifile, "Settings", "ModalPopup", "False")
		MsgBox(4096+64, $windowtitle, "File created: " & $inifile & "." & @LF & "File contains WindowTitles for Gmail, Outlook and Twitter.")
	EndIf	
EndFunc	