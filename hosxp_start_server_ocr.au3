#RequireAdmin
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
Opt("WinTitleMatchMode", 2)
Global $sServer_Title = 'AutoIt Server'
Global $sOCR_Title = 'AutoIt OCR'
Global Const $sOcrUrl = "http://127.0.0.1:3078"
Global Const $gOcrStartupMs = 180000
Global $gOcrHttpActive = False, $gOcrHttpFailed = False
Global $gOcrNodePID = 0, $gOcrPythonPID = 0
Global $gOcrCacheKey = ""
Global $oErr = ObjEvent("AutoIt.Error", "_ComErrHandler")

Func _ComErrHandler($oError)
    If $gOcrHttpActive Then $gOcrHttpFailed = True
    ; เวลามี COM error (เช่น timeout) ให้เก็บ error message ไว้
    ConsoleWrite("COM Error: " & $oError.description & @CRLF)
    ; คืนค่าเฉย ๆ เพื่อไม่ให้ AutoIt หยุดทำงาน
EndFunc

Func _OCRHttp($sMethod, $sUrl, $bBody, $iTimeoutMs, ByRef $iStatus)
    $iStatus = 0
    $gOcrHttpFailed = False
    $gOcrHttpActive = True
    Local $oHttp = ObjCreate("WinHttp.WinHttpRequest.5.1")
    If Not IsObj($oHttp) Then
        $gOcrHttpActive = False
        Return SetError(3, 0, "")
    EndIf
    $oHttp.SetProxy(1)
    $oHttp.SetTimeouts(2000, 2000, 5000, $iTimeoutMs)
    $oHttp.Open($sMethod, $sUrl, True)
    If $sMethod = "POST" Then
        $oHttp.SetRequestHeader("Content-Type", "image/png")
        $oHttp.Send($bBody)
    Else
        $oHttp.Send()
    EndIf
    Local $hWait = TimerInit(), $iFailure = 0
    While Not $gOcrHttpFailed
        If $oHttp.WaitForResponse(0) Then ExitLoop
        If TimerDiff($hWait) >= $iTimeoutMs Then
            $iFailure = 5
            ExitLoop
        EndIf
        Sleep(50)
    WEnd
    If $gOcrHttpFailed Then $iFailure = 3
    If $iFailure Then
        $oHttp.Abort()
        $gOcrHttpActive = False
        Return SetError($iFailure, 0, "")
    EndIf
    $iStatus = $oHttp.Status
    Local $sText = $oHttp.ResponseText
    Local $bFailed = $gOcrHttpFailed
    $gOcrHttpActive = False
    If $bFailed Then Return SetError(3, 0, "")
    Return $sText
EndFunc

Func StartServer()
    ;_ArrayDisplay($aList)
    ;_ArrayDisplay($aList)
    Local $sNodeDir = @ScriptDir & "\server"
    Local $sPythonDir = @ScriptDir & "\OCR"
    If Not FileExists($sNodeDir & "\package.json") Or Not FileExists($sPythonDir & "\.venv\Scripts\python.exe") Or Not FileExists($sPythonDir & "\server.py") Then
        MsgBox($MB_ICONERROR, "Start server", "Missing server/package.json or OCR/.venv/Scripts/python.exe or OCR/server.py. Run install.bat first.")
        Return False
    EndIf
    Local $hNode = WinGetHandle($sServer_Title)
    If $hNode Then
        $gOcrNodePID = WinGetProcess($hNode)
    Else
        ; Working directory passed to Run handles spaces and drive changes.
        $gOcrNodePID = Run('"' & @ComSpec & '" /d /c "title ' & $sServer_Title & ' & npm start"', $sNodeDir, @SW_MINIMIZE)
        If @error Or $gOcrNodePID = 0 Then Return False
    EndIf
    Local $iStatus = 0
    Local $sHealth = _OCRHttp("GET", $sOcrUrl & "/health", Binary(""), 2000, $iStatus)
    If $iStatus = 200 And StringInStr($sHealth, "READY") Then
        $gOcrPythonPID = _OCRFindPythonPID()
        Return True
    EndIf
    $gOcrPythonPID = _OCRFindPythonPID()
    If $gOcrPythonPID = 0 Then
        EnvSet("PYTHONUTF8", "1")
        $gOcrPythonPID = Run('"' & $sPythonDir & '\.venv\Scripts\python.exe" "' & $sPythonDir & '\server.py"', $sPythonDir, @SW_MINIMIZE)
        If @error Or $gOcrPythonPID = 0 Then Return False
    EndIf
    Local $hStartup = TimerInit()
    While TimerDiff($hStartup) < $gOcrStartupMs
        If Not ProcessExists($gOcrPythonPID) Then ExitLoop
        $sHealth = _OCRHttp("GET", $sOcrUrl & "/health", Binary(""), 2000, $iStatus)
        If $iStatus = 200 And StringInStr($sHealth, "READY") Then Return True
        Sleep(500)
    WEnd
    MsgBox($MB_ICONERROR, "OCR", "OCR is not READY on port 3078. Check the Python terminal. First model download may take longer than 180 seconds.")
    Return False
EndFunc

Func _OCRFindPythonPID()
    Local $oWmi = ObjGet("winmgmts:\\.\root\cimv2")
    If Not IsObj($oWmi) Then Return 0
    Local $oProcesses = $oWmi.ExecQuery("SELECT ProcessId, ExecutablePath, CommandLine FROM Win32_Process WHERE Name='python.exe'")
    If Not IsObj($oProcesses) Then Return 0
    For $oProc In $oProcesses
        If StringLower($oProc.ExecutablePath) = StringLower(@ScriptDir & "\OCR\.venv\Scripts\python.exe") And StringInStr($oProc.CommandLine, "server.py") Then Return Number($oProc.ProcessId)
    Next
    Return 0
EndFunc

Func StopServer()
    ; ปิดหน้าต่าง Node.js server ตาม title เดิม
    Local $aList = WinList($sServer_Title) ;Winlist all hos os
    ;_ArrayDisplay($aList)
    Local $aList_Length = $aList[0][0]

    For $i = 1 To $aList_Length
        Local $sTitle = $aList[$i][0]
        Local $hHandle = $aList[$i][1]

        If $sTitle <> "" Then
            WinKill($hHandle)
            Sleep(300)
        EndIf
    Next

    ; ปิด Node.js และรอให้คำสั่งทำงานเสร็จ
    Local $iNodeExit = RunWait( _
            '"' & @SystemDir & '\taskkill.exe" /IM node.exe /T /F', _
            "", @SW_HIDE)

    ConsoleWrite("Node.js taskkill exit code=" & $iNodeExit & @CRLF)

    ; ปิด OCR server ตาม PID
    If $gOcrPythonPID = 0 Or Not ProcessExists($gOcrPythonPID) Then
        $gOcrPythonPID = _OCRFindPythonPID()
    EndIf

    If $gOcrPythonPID > 0 Then
        Local $iOcrExit = RunWait( _
                '"' & @SystemDir & '\taskkill.exe" /PID ' & _
                Int($gOcrPythonPID) & ' /T /F', _
                "", @SW_HIDE)

        ConsoleWrite("OCR taskkill exit code=" & $iOcrExit & @CRLF)

        If Not ProcessExists($gOcrPythonPID) Then
            $gOcrPythonPID = 0
        EndIf
    Else
        ConsoleWrite("OCR server ไม่ได้ทำงาน" & @CRLF)
    EndIf

    $gOcrNodePID = 0
    $gOcrCacheKey = ""
    FileDelete(@ScriptDir & "\TempOCR\hosxp_ocr_" & @AutoItPID & ".png")
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


Func HideServer()
    Local $hNode = WinGetHandle($sServer_Title)
    Local $hOCR = WinGetHandle($sOCR_Title)
    If $hNode Then WinSetState($hNode, "", @SW_MINIMIZE)
    If $hOCR Then WinSetState($hOCR, "", @SW_MINIMIZE)
EndFunc


Local $bStarted = StartServer()
If $bStarted Then
    Sleep(300)
    HideServer()
Else
    MsgBox($MB_ICONWARNING, "Start server", "เปิด server ไม่ครบ กรุณาตรวจหน้าต่าง Node.js และ OCR")
EndIf
; ให้เลือกปิด server ได้ แม้เริ่มต้นสำเร็จเพียงบางส่วน
MsgStopServer()
If Not $bStarted Then Exit 1
Exit 0
