#RequireAdmin
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#Include <WinAPI.au3>
#Include <SendMessage.au3>
#include <FileConstants.au3>
#include <Array.au3>
#include <StringConstants.au3>
#include <Constants.au3>

Opt("WinTitleMatchMode", 2)
Global $sServer_Title = 'AutoIt Server'

Func HideServer()
	 Local $hEcel = WinWait($sServer_Title, "", 2)
	 ;Local $hEcel = WinGetHandle("[CLASS:XLMAIN]")
	if $hEcel Then WinSetState($hEcel, "", @SW_MINIMIZE)
EndFunc

Func StartServer()
	Local $aList = WinList($sServer_Title) ;Winlist all hos os
	;_ArrayDisplay($aList)
	Local $aList_Length = $aList[0][0]
	If $aList_Length > 0 Then
				For $i = 1 To $aList_Length
					$sTitle = $aList[$i][0]
					$hHandle = $aList[$i][1]
						If $sTitle <> "" And BitAND(WinGetState($hHandle), 2) Then
							Sleep(300)
							ExitLoop
						EndIf
				Next
	Else
	   Local $CMD2 =  'cd '&@ScriptDir&'\server\ && ' & _
        'npm start'
	   Run('"' & @ComSpec & '" /k ' & $CMD2)
	EndIf
EndFunc

Func StopServer()
	Local $aList = WinList($sServer_Title) ;Winlist all hos os
	;_ArrayDisplay($aList)
	Local $aList_Length = $aList[0][0]
	 If $aList_Length > 0 Then
				For $i = 1 To $aList_Length
					$sTitle = $aList[$i][0]
					$hHandle = $aList[$i][1]
						If $sTitle <> "" And BitAND(WinGetState($hHandle), 2) Then
							WinKill($sTitle)
							Sleep(300)
							ExitLoop
						EndIf
				Next
	EndIf
    Run(@ComSpec & ' /c taskkill /IM node.exe /F', "", @SW_HIDE)
    ConsoleWrite("Node.js server stopped by taskkill" & @CRLF)
EndFunc

Func MsgStopServer()
	Local $iAnswer = MsgBox($MB_OKCANCEL + $MB_ICONQUESTION, _
    "ยืนยัน", "คุณต้องการปิด server หรือไม่?")
	If $iAnswer = $IDOK Then
		StopServer()
	Else
		MsgBox($MB_ICONWARNING, "ยกเลิก", "กรุณาปิด server ด้วยตนเอง")
	EndIf
EndFunc

StartServer()
Sleep(100)
HideServer()
MsgStopServer()
Exit(0)