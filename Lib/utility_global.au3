#ifndef UTILITY_GLOBAL_AU3
#define UTILITY_GLOBAL_AU3

#include "autoIt_libs.au3"

;=============================== ROOT PATH =================================

Global Const $gRootDir = @ScriptDir

Global Const $gLibDir = $gRootDir & "\Lib"
Global Const $gLogDir = $gRootDir & "\BotLog"
Global Const $gServerDir = $gRootDir & "\server"
Global Const $gMatchPicDir = $gRootDir & "\Match"

;=============================== FILE PATH =================================



;===================== start log path config ==================================================================
Global $logStudent =        $gLogDir &"\log_telemed.txt"
Global $logStudenProgress = $gLogDir &"\log_telemed_progress.txt"
Global $logPrice =          $gLogDir &"\log_telemed_price.txt"
;===================== end log path config ==================================================================



;=============================== start Global 1 ========================================================================
Global Const $iPos1 = -8,  $iPos2 = -8,  $iSize1 = 1552, $iSize2 = 840
Global Const $aXpPos = [-8, -8, 1552, 840]
Global Const $sHosXp_Title = "BMS-HOSxP XE 4.0"
Global Const $sServer_Title = 'AutoIt Server'
Global Const $sOCR_Title = 'AutoIt OCR'
Global $iSleepAfterLoad = 1500
Global $sSchool = ''

Global $hWndXp
Global $oExcel
Global $oLogStudent
Global $oLogStudentProgress
Global $oLogPrice
Global $oWorkbook
Global $iMaxRecTime = 4*60*1000 ; 4 min
;=============================== end Global 1 ========================================================================


;=== start telegram and database request config ===============================================
Global $gEnvFile = $gServerDir & "\.env"
Global $gEnvCache = 0        ; เก็บ cache (Dictionary ของ KEY=VALUE)
Global $gEnvTimestamp = ""  ; เวลาล่าสุดของไฟล์
Global $gTelegramToken = ""
Global $gTelegramChatID = ""
Global $gHTTPTelegram = ObjCreate("WinHttp.WinHttpRequest.5.1")
Global $gHTTPNode = ObjCreate("WinHttp.WinHttpRequest.5.1")
Global $gHTTPError = False
Global $sReqPath = "query_telemed"
Global $oErr = ObjEvent("AutoIt.Error", "_ComErrHandler")
;=== end telegram  and database request config ===============================================


; start OCR v2: confidence 0-100, independent of OpenCV tolerance. ===================================
Global $gOcrHttpActive = False, $gOcrHttpFailed = False
Global $gOcrNodePID = 0, $gOcrPythonPID = 0
Global $gOcrRowTimer = 0, $gOcrRowLog = 0, $gOcrRowNumber = 0
Global Const $gOcrRequestMs = 60000
Global Const $gOcrStartupMs = 180000
Global $gOcrCacheImage = Binary(""), $gOcrCacheKey = "", $gOcrCacheResult = ""
Global $gOcrLastStatus = ""
Global Const $sOcrUrl = "http://127.0.0.1:3078"  ; node server use 3074
;=== end OCR config ===============================================



;=== start element class and name config ================================================
Global Const $sLockPtForm = "[CLASS:THOSxPVisitNumberLockForm]"  ; pop up other user lock
Global Const $sBtnSaveCC = "[CLASS:TcxButton; INSTANCE:52]"  ; ปุ่มบันทึก CC
Global Const $sBtnAddOperation = "[CLASS:TcxButton; INSTANCE:31]"    ; ปุ่มเพิ่มหัตถการ
Global Const $sOperationBoxTitle = "HOSxPDentalOperationEntryForm" ; ฟอร์มบันทึกหัตถการ THOSxPDentalOperationEntryForm
Global Const $sBtnSaveOperation = "[CLASS:TcxButton; INSTANCE:10]"  ; ปุ่มบันทึกหัตถการ

;=== end element class and name config ================================================



;=== start picture config ==============================================================
Global Const $kScreen = 1 ; เครื่องพัฒนา display screen 125% = 1.25 / เครื่องจริง 100% = 1

Global Const $sFoundVisit = $gMatchPicDir & "\xp_hn_found_visit.png"   ;รูปชายหญิง
Global Const $aFvFirstPos = [0, 81,75,169] ;start loop position
Global Const $aFvPos = [256, 18, 683, 374]  ;end loop position

Global Const $sLoadPtSuccess2 = $gMatchPicDir & "\xp_load_pt_success2.png" ;รูป icon พยาบาลกากบาทสีน้ำเงิน
Global Const $aLoadPtPos = [159, 147,1880, 432]

Global Const $sCC =   $gMatchPicDir & "\xp_cc.png"  ;ข้อความ CC
Global Const $aCcPos = [253, 365, 948, 562]

Global Const $sRec_allergy_iden =  $gMatchPicDir & "\rec_allergy_iden.png"  ;select box การแพ้ยา
Global Const $aRecAllergyPos = [166, 186, 931, 390]

Global Const $sEditPrice =   $gMatchPicDir & "\xp_edit_price.png"   ; ข้อความ Inv Setting
Global Const $aEditPricePos = [154, 638, 1515, 941]


Global Const $sFinalSave2 =  $gMatchPicDir & "\xp_final_save2.png"  ;ภาพติ๊กถูกในวงกลมสีเขียวตอนบันทึกครั้งสุดท้าย
Global Const $aFinalS2Pos = [496, 197, 1081, 534]
Global Const $sLastOK =   $gMatchPicDir & "\xp_lastOK.png"
Global Const $aLastOkPos = [728, 341, 1272, 701]
;=== end picture config =================================================================


;=============================== start utility function 1 ========================================================================
Func Hkey()
	AdlibUnRegister("ContXp1")
	FileClose($oLogStudent)
	FileClose($oLogStudentProgress) ;
	FileClose($oLogPrice)
	ShowExcel()
    _Excel_Close($oExcel)
	ShowSciTE()
	_OpenCV_Shutdown();Closes DLLs
	_CppDllClose()
	 MsgBox(0, "Hkey", "Stop Process")
	Exit 0
EndFunc

Func CtrlSend($hWnd,$sClass,$iVal)
	ControlSend($hWnd, "", $sClass,"{DELETE}")
	Sleep(200)
	ControlSend($hWnd, "", $sClass, "{ESC}")
	Sleep(200)
	ControlSend($hWnd, "", $sClass, $iVal)
	Sleep(200)
EndFunc

Func CtrlSendDt($hWnd,$sClass,$iVal)
	ControlClick($hWnd, "left", $sClass)
	Sleep(300)
	ControlSend($hWnd, "", $sClass,"{DELETE}")
	Sleep(200)
	ControlSend($hWnd, "", $sClass, "{ESC}")
	Sleep(200)
	ControlSend($hWnd, "", $sClass, $iVal)
	Sleep(200)
EndFunc

;=============================== end utility function 1 ========================================================================



;=============================== start match pic function ==================================================================

 Func FindToClick($picPath = @ScriptDir&"\Match\dad_name.png", $x1 = 0, $y1 = 370, $x2 = 51, $y2 = 400, $tol = 0.75)
	Local $Match = _MatchPicture("", $picPath, $x1*$kScreen, $y1*$kScreen, $x2*$kScreen, $y2*$kScreen, $tol)
	If Not IsArray($Match) Then Return False
    If UBound($Match) < 6 Then Return False
    If $Match[0] <= 0 Then Return False

    Local $realX = Round($Match[4] / $kScreen)
    Local $realY = Round($Match[5] / $kScreen)

    MouseClick($MOUSE_CLICK_LEFT, $realX, $realY, 1, 2)
    Return True
 EndFunc

 Func FindToClickRt($picPath = @ScriptDir&"\Match\dad_name.png", $x1 = 0, $y1 = 370, $x2 = 51, $y2 = 400, $tol = 0.75)
	Local $Match = _MatchPicture("", $picPath, $x1*$kScreen, $y1*$kScreen, $x2*$kScreen, $y2*$kScreen, $tol)
	If Not IsArray($Match) Then Return False
    If UBound($Match) < 6 Then Return False
    If $Match[0] <= 0 Then Return False

    Local $realX = Round($Match[4] / $kScreen)
    Local $realY = Round($Match[5] / $kScreen)

    MouseClick($MOUSE_CLICK_RIGHT, $realX, $realY, 1, 2)
    Return True

EndFunc

Func FindToCon($picPath = @ScriptDir&"\Match\dad_name.png", $x1 = 0, $y1 = 370, $x2 = 51, $y2 = 400, $tol = 0.75)
 Local $Match = _MatchPicture("", $picPath, $x1*$kScreen, $y1*$kScreen, $x2*$kScreen, $y2*$kScreen, $tol)
 ;_ArrayDisplay($Match1)
 If $Match[0] > 0 Then
	 ;ConsoleWrite("result = match")
	 Return True
 Else
	 ;ConsoleWrite("result = not match")
	 Return False
 EndIf
EndFunc
;=============================== end match pic function ==================================================================



; ==================== start OCR functions share the original template's DPI conversion. ==================================================
; Always capture current pixels; reuse OCR only when PNG bytes AND query match.
; No Found popup/tooltip. @error: 1 input, 2 capture, 3 HTTP, 4 protocol, 5 timeout.
Func FindToClickOCR($sTarget, $x1, $y1, $x2, $y2, $minConfidence = 60)
    Local $aHit = _FindOCR($sTarget, $x1, $y1, $x2, $y2, $minConfidence)
    Local $iError = @error
    If Not IsArray($aHit) Then Return SetError($iError, 0, False)
    _OCRCheckDeadline()
    MouseClick($MOUSE_CLICK_LEFT, $aHit[0], $aHit[1], 1, 2)
    Return True
EndFunc

Func FindToClickRtOCR($sTarget, $x1, $y1, $x2, $y2, $minConfidence = 60)
    Local $aHit = _FindOCR($sTarget, $x1, $y1, $x2, $y2, $minConfidence)
    Local $iError = @error
    If Not IsArray($aHit) Then Return SetError($iError, 0, False)
    _OCRCheckDeadline()
    MouseClick($MOUSE_CLICK_RIGHT, $aHit[0], $aHit[1], 1, 2)
    Return True
EndFunc

Func FindToConOCR($sTarget, $x1, $y1, $x2, $y2, $minConfidence = 60)
    Local $aHit = _FindOCR($sTarget, $x1, $y1, $x2, $y2, $minConfidence)
    Local $iError = @error
    Return SetError($iError, 0, IsArray($aHit))
EndFunc

Func _OCRSetRow($hStartTime, $oLog, $iRow)
    $gOcrRowTimer = $hStartTime
    $gOcrRowLog = $oLog
    $gOcrRowNumber = $iRow
EndFunc

Func _OCRCheckDeadline()
    If $gOcrRowTimer <> 0 Then ExitMaxTime($gOcrRowTimer, $gOcrRowLog, $gOcrRowNumber)
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
        _OCRCheckDeadline()
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
    _OCRCheckDeadline()
    If $bFailed Then Return SetError(3, 0, "")
    Return $sText
EndFunc

Func _FindOCR($sTarget, $x1, $y1, $x2, $y2, $minConfidence)
    _OCRCheckDeadline()
    $gOcrLastStatus = "ERROR"

    If $sTarget = "" Or $x2 <= $x1 Or $y2 <= $y1 Or _
            $minConfidence < 0 Or $minConfidence > 100 Then
        Return SetError(1, 0, False)
    EndIf

    Local $iLeft = Round($x1 * $kScreen)
    Local $iTop = Round($y1 * $kScreen)
    Local $iRight = Round($x2 * $kScreen)
    Local $iBottom = Round($y2 * $kScreen)

    Local $sOcrDir = @ScriptDir & "\TempOCR"

	If Not FileExists($sOcrDir) Then
		If Not DirCreate($sOcrDir) Then
			ConsoleWrite("OCR: สร้างโฟลเดอร์ไม่สำเร็จ: " & $sOcrDir & @CRLF)
			Return SetError(2, 0, False)
		EndIf
	EndIf

    Local $sImage = $sOcrDir & "\hosxp_ocr_" & @AutoItPID & ".png"

    ; จับภาพปัจจุบันก่อนตรวจ cache
    FileDelete($sImage)
    _ScreenCapture_Capture($sImage, $iLeft, $iTop, $iRight, $iBottom, False)

    If Not FileExists($sImage) Then Return SetError(2, 0, False)

    Local $hFile = FileOpen($sImage, 16)
    If $hFile = -1 Then Return SetError(2, 0, False)

    Local $bImage = FileRead($hFile)
    FileClose($hFile)

    If BinaryLen($bImage) = 0 Then Return SetError(2, 0, False)

    Local $sKey = $sTarget & "|" & $minConfidence & "|" & _
            $iLeft & "|" & $iTop & "|" & $iRight & "|" & $iBottom

    Local $sResponse = "", $iStatus = 0

    If $sKey == $gOcrCacheKey And $bImage = $gOcrCacheImage Then
        ; ภาพ คำค้น confidence และ area เหมือนเดิม ใช้ผลสำเร็จที่จำไว้
        $sResponse = $gOcrCacheResult
    Else
        _OCRCheckDeadline()

        ; ส่งคำค้นจริงครั้งเดียว Python จะตรวจ exact ก่อน fuzzy
        Local $sHex = StringTrimLeft(String(StringToBinary($sTarget, 4)), 2)

        $sResponse = _OCRHttp("POST", _
                $sOcrUrl & "/ocr?target=" & $sHex & _
                "&confidence=" & $minConfidence, _
                $bImage, $gOcrRequestMs, $iStatus)

        Local $iError = @error

        ; ล้าง cache เดิมเมื่อมีการค้นหาใหม่
        $gOcrCacheKey = ""

        If $iError Then Return SetError($iError, 0, False)

        If $iStatus <> 200 Then
            ConsoleWrite("OCR HTTP=" & $iStatus & @CRLF)
            Return SetError(3, $iStatus, False)
        EndIf
    EndIf

    ; Python: status, x, y, confidence, text, similarity, match_method
    Local $aResult = StringSplit(StringStripWS($sResponse, 2), @TAB, 2)
    If UBound($aResult) < 7 Then Return SetError(4, 0, False)

    $gOcrLastStatus = $aResult[0]

    ; ทั้ง NO_MATCH และ AMBIGUOUS จะไม่คืนพิกัดสำหรับคลิก
    If $aResult[0] <> "MATCH" Then Return False

    If Not StringRegExp($aResult[1], "^\d+(?:\.\d+)?$") Or _
            Not StringRegExp($aResult[2], "^\d+(?:\.\d+)?$") Or _
            Not StringRegExp($aResult[3], "^\d+(?:\.\d+)?$") Or _
            Not StringRegExp($aResult[5], "^\d+(?:\.\d+)?$") Then
        $gOcrLastStatus = "ERROR"
        Return SetError(4, 0, False)
    EndIf

    Local $fX = Number($aResult[1])
    Local $fY = Number($aResult[2])
    Local $fConfidence = Number($aResult[3])
    Local $fSimilarity = Number($aResult[5])
    Local $sMethod = $aResult[6]

    If $fX < 0 Or $fX > $iRight - $iLeft Or _
            $fY < 0 Or $fY > $iBottom - $iTop Or _
            $fConfidence > 100 Or $fSimilarity > 100 Then
        $gOcrLastStatus = "ERROR"
        Return SetError(4, 0, False)
    EndIf

    If $sMethod <> "exact" And $sMethod <> "fuzzy" Then
        $gOcrLastStatus = "ERROR"
        Return SetError(4, 0, False)
    EndIf

    If $fConfidence < $minConfidence Then Return False

    ; เกณฑ์ similarity ตรวจโดย Python ตาม FUZZY_MINIMUM
    _OCRCheckDeadline()

    ; เก็บ cache หลังตรวจสอบ response สำเร็จแล้วเท่านั้น
    $gOcrCacheKey = $sKey
    $gOcrCacheImage = $bImage
    $gOcrCacheResult = $sResponse

    ConsoleWrite("OCR MATCH | method=" & $sMethod & _
            " | confidence=" & $fConfidence & _
            " | similarity=" & $fSimilarity & @CRLF)

    ; พิกัดภายในภาพ → พิกัดหน้าจอ ตาม scaling เดิม
    Local $aHit[2] = [ _
            Round(($iLeft + $fX) / $kScreen), _
            Round(($iTop + $fY) / $kScreen)]

    Return $aHit
EndFunc
; ==================== end OCR functions share the original template's DPI conversion. ==================================================



;=============================== start utility function 2 ========================================================================
Func ContXp1()
	Local $hErrPopUp = WinGetHandle("[CLASS:madExceptWndClass]")  ;madExceptWndClass
	if $hErrPopUp Then
		ConsoleWrite("kill error occurred"&@CRLF)
	    WinKill($hErrPopUp)
	EndIf
EndFunc

Func ContXp2()
	Local $hErrPopUp
	 While 1
		    Sleep(100)
		    $hErrPopUp = WinGetHandle("[CLASS:madExceptWndClass]")
			Sleep(100)
			if $hErrPopUp Then
				ConsoleWrite("kill error occurred"&@CRLF)
	            WinKill($hErrPopUp)
			EndIf
			if Not $hErrPopUp Then ExitLoop
	  WEnd
EndFunc

Func ExitMaxTime($hTime, $oFile, $iRow)
	if TimerDiff($hTime) > $iMaxRecTime Then
		;SwapFacebook("Error Time Exit R= "&$iRow)
		FileWrite($oFile, "Error Time Exit R= "&$iRow&", ")
		SendTeleGram("Error Time Exit R= "&$iRow)
		Sleep(200)
		Exit(0)
	EndIf
EndFunc
;=============================== end utility function 2 ========================================================================



;=============== send request telegram func ==========================================
Func LoadEnv($sFile = ".env")
    Local $aLines, $sLine, $aPart

    If Not FileExists($sFile) Then
        Return SetError(1, 0, 0)
    EndIf

    $aLines = FileReadToArray($sFile)
    If @error Then Return SetError(2, 0, 0)

    ; สร้าง Dictionary สำหรับเก็บ key/value
    Local $oMap = ObjCreate("Scripting.Dictionary")

    For $i = 0 To UBound($aLines) - 1
        $sLine = StringStripWS($aLines[$i], 3)
        If $sLine = "" Or StringLeft($sLine, 1) = "#" Then ContinueLoop

        $aPart = StringSplit($sLine, "=", 2)
        If UBound($aPart) = 2 Then
            $oMap($aPart[0]) = $aPart[1]
        EndIf
    Next

    $gEnvCache = $oMap
    $gEnvFile = $sFile
    $gEnvTimestamp = FileGetTime($sFile, 0, 1) ; บันทึก timestamp

    Return $oMap
EndFunc

Func GetEnv($sKey)
    ; --- โหลด cache ถ้ายังไม่ถูกสร้าง ---
    If Not IsObj($gEnvCache) Then
        $gEnvCache = LoadEnv($gEnvFile)
    EndIf

    ; --- ตรวจสอบไฟล์ถูกแก้ไขใหม่หรือไม่ ---
    Local $sCurrentTime = FileGetTime($gEnvFile, 0, 1)
    If $sCurrentTime <> $gEnvTimestamp Then
        $gEnvCache = LoadEnv($gEnvFile)
    EndIf

    ; --- คืนค่า value ---
    If $gEnvCache.Exists($sKey) Then
        Return $gEnvCache($sKey)
    EndIf

    Return ""
EndFunc

Func SendTeleGram($sMessage)
    If $gTelegramToken = "" Or $gTelegramChatID = "" Then Return

    Local $sURL = "https://api.telegram.org/bot" & $gTelegramToken & "/sendMessage?chat_id=" & $gTelegramChatID & "&text=" & $sMessage

    $gHTTPTelegram.Open("GET", $sURL, False)
    $gHTTPTelegram.SetTimeouts(5000, 5000, 5000, 10000) ; timeout: resolve, connect, send, receive
    $gHTTPTelegram.Send()
    ;ConsoleWrite("Telegram response: " & $gHTTPTelegram.ResponseText & @CRLF)
EndFunc
;=============== send request telegram func ==========================================



;=============== start send request nodejs localhost func ======================================
Func QueryPostgres2($sHn)
    If StringLen($sHn) <> 9 Or Not StringRegExp($sHn, "^[0-9]{9}$") Then
        Return '{"error":"INVALID_HN"}'
    EndIf

    Local $sData = '{"hn":"' & $sHn & '"}'
    Local $sUrl = "http://127.0.0.1:3074/" & $sReqPath

    $gHTTPError = False

    $gHTTPNode.SetTimeouts(5000, 5000, 5000, 10000)
    If $gHTTPError Then Return '{"error":"API_TIMEOUT_SETUP_FAILED"}'

    $gHTTPNode.Open("POST", $sUrl, False)
    If $gHTTPError Then Return '{"error":"API_OPEN_FAILED"}'

    $gHTTPNode.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
    If $gHTTPError Then Return '{"error":"API_HEADER_FAILED"}'

    $gHTTPNode.Send($sData)
    If $gHTTPError Then Return '{"error":"API_SEND_FAILED"}'

    Local $iStatus = $gHTTPNode.Status
    If $gHTTPError Then Return '{"error":"API_STATUS_FAILED"}'

    Local $sResponse = $gHTTPNode.ResponseText
    If $gHTTPError Then Return '{"error":"API_RESPONSE_FAILED"}'

    If $iStatus <> 200 Then
        Return '{"error":"API_HTTP_' & $iStatus & '"}'
    EndIf

    If $sResponse = "" Then
        Return '{"error":"API_EMPTY_RESPONSE"}'
    EndIf

    Return $sResponse
EndFunc

Func GetSitFromDb2($iHnVal, $oFile, $iRow)
    ; Never convert HN to a number, pad it, or silently remove invalid characters.
    Local $sHn = String($iHnVal)
    If StringLen($sHn) <> 9 Or Not StringRegExp($sHn, "^[0-9]{9}$") Then
        FileWrite($oFile, "Invalid HN R=" & $iRow & @CRLF)
        MsgBox(16, "Invalid HN", "HN must contain exactly 9 digits. Excel row: " & $iRow)
        Exit(1)
    EndIf
    Local $result = QueryPostgres2($sHn)
    If StringInStr($result, '"error"') Then
        SendTeleGram("Error Server Connect R= " & $iRow)
        FileWrite($oFile, "Error Server Connect R= " & $iRow & ", ")
        MsgBox(16, "API Error", "Cannot read data from local API:" & @CRLF & $result)
        Exit(1)
    EndIf
    ; Parse only the fixed three-string array produced by this endpoint.
    Local $aResult = StringRegExp($result, '^\s*\[\s*"([0-9]+[dno]?)"\s*,\s*"(-?[0-9]+)"\s*,\s*"([0-9]+)"\s*\]\s*$', 1)
    If @error Then
        FileWrite($oFile, "Invalid API response R=" & $iRow & @CRLF)
        MsgBox(16, "API Error", "Invalid API response. Excel row: " & $iRow)
        Exit(1)
    EndIf
    Return $aResult
EndFunc


Func _ComErrHandler($oError)
    ; เวลามี COM error (เช่น timeout) ให้เก็บ error message ไว้
    ConsoleWrite("COM Error: " & $oError.description & @CRLF)
    ; คืนค่าเฉย ๆ เพื่อไม่ให้ AutoIt หยุดทำงาน
EndFunc

;=============== end send request nodejs localhost func ======================================



;=============================== start preload function ======================================================================
Func HideExcel()
	 Local $hEcel = WinWait("[CLASS:XLMAIN]", "", 10)
	 ;Local $hEcel = WinGetHandle("[CLASS:XLMAIN]")
	if $hEcel Then WinSetState($hEcel, "", @SW_MINIMIZE)
EndFunc

Func ShowExcel()
	Local $hEcel = WinWait("[CLASS:XLMAIN]", "", 10)
	 ;Local $hEcel = WinGetHandle("[CLASS:XLMAIN]")
	if $hEcel Then WinSetState($hEcel, "", @SW_SHOW )
EndFunc

Func HideSciTE()
	 Local $hEcel = WinWait("[CLASS:SciTEWindow]", "", 2)
	 ;Local $hEcel = WinGetHandle("[CLASS:XLMAIN]")
	if $hEcel Then WinSetState($hEcel, "", @SW_MINIMIZE)
EndFunc

Func ShowSciTE()
	Local $hEcel = WinWait("[CLASS:SciTEWindow]", "", 2)
	 ;Local $hEcel = WinGetHandle("[CLASS:XLMAIN]")
	if $hEcel Then WinSetState($hEcel, "", @SW_SHOW )
EndFunc

Func SwapXP()
		;_ArrayDisplay($aList)
	Local $aList = WinList($sHosXp_Title) ;Winlist all hos os
	Local $aList_Length = $aList[0][0]

	If $aList_Length > 0 Then
				For $i = 1 To $aList_Length
					Local $sTitle = $aList[$i][0]
					Local $hHandle = $aList[$i][1]
						If $sTitle <> "" And BitAND(WinGetState($hHandle), 2) Then
							;MsgBox($MB_SYSTEMMODAL, "", "Title: " & $aList[$i][0] & @CRLF & "Handle: " & $aList[$i][1])
							WinActivate($sTitle)
							Sleep(1000)
							;WinMove($sTitle, "", $iPos1, $iPos2, $iSize1, $iSize2)
							ExitLoop
						EndIf
				Next
	Else
			AdlibUnRegister("ContXp1")
			FileClose($oLogStudent)
			FileClose($oLogStudentProgress)
			FileClose($oLogPrice)
			ShowExcel()
			_Excel_Close($oExcel)
			ShowSciTE()
			_OpenCV_Shutdown();Closes DLLs
			_CppDllClose()
			;SendTeleGram("Not Found Hos XP")
		    MsgBox($MB_SYSTEMMODAL, "Warning!","ไม่พบโปรแกรม Hos XP",10)
		    Exit
    EndIf
	;ConsoleWrite("SwapHosOS"& @CRLF & "Success")
EndFunc

Func SetHnBox($hWndXp)
    ; เลือกเฉพาะ TcxGroupBox Instance 1
    Local $hCtrl = ControlGetHandle($hWndXp, "", "[CLASS:TcxGroupBox; INSTANCE:1]") ;ด้านขวาของ list คนไข้ เส้นแนวตั้งสีฟ้าข้างๆ ไม่ใช่ตาราง grid
    Local $aPos = ControlGetPos($hWndXp, "", $hCtrl)
    ;MsgBox(0, "HN Box Pos", "x=" & $aPos[0] & ", y=" & $aPos[1] & ", w=" & $aPos[2] & ", h=" & $aPos[3])
    If $aPos[0] <> 0 Or $aPos[1] <> 125 Or $aPos[2] <> 400 Or $aPos[3] <> 676 Then
        ControlMove($hWndXp, "", $hCtrl, 0, 125, 400, 676)
		Sleep(1000)
    EndIf
EndFunc

;=============================== end preload function ======================================================================



;=============================== start after load exit function =====================================================================

Func KillFinance2()
    Local $aList = WinList()

    For $i = 1 To $aList[0][0]
        Local $hWnd = $aList[$i][1]
        If $hWnd = 0 Then ContinueLoop

        ; --- 1) ต้องเป็น dialog class ---
        If _WinAPI_GetClassName($hWnd) <> "#32770" Then ContinueLoop

        ; --- 2) ต้อง visible + enabled ---
        Local $iState = WinGetState($hWnd)
        If Not BitAND($iState, 2) Then ContinueLoop

        ; --- 3) ต้องมีปุ่มแรกเป็น "OK" ---
        Local $hOK = ControlGetHandle($hWnd, "", "[CLASS:Button; INSTANCE:1]")
        If $hOK = 0 Then ContinueLoop
        If StringStripWS(ControlGetText($hWnd, "", $hOK), 8) <> "OK" Then ContinueLoop

        ; --- 4) ต้องมีปุ่มเดียวเท่านั้น (กัน dialog Yes/No, OK/Cancel) ---
        If ControlGetHandle($hWnd, "", "[CLASS:Button; INSTANCE:2]") <> 0 Then ContinueLoop

        ; --- 5) title ต้องเป็นของ HOSxP ---
        If Not StringInStr(WinGetTitle($hWnd), "BMS-HOSxP") Then ContinueLoop

        ConsoleWrite("Found popup -> " & $hWnd & @CRLF)

        ; --- 6) ยิงปุ่ม OK แบบไล่ระดับ ---
        _ClickOK($hWnd, $hOK)
        Return True
    Next

    Return False
EndFunc

Func _ClickOK($hWnd, $hOK)
    ; วิธี 1: BM_CLICK ตรงเข้า handle ปุ่ม (ทำงานแม้หน้าต่างไม่ active)
    _SendMessage($hOK, $BM_CLICK, 0, 0)
    If _WaitClose($hWnd, 800) Then Return

    ; วิธี 2: WM_COMMAND + IDOK ไปที่ dialog
    _SendMessage($hWnd, $WM_COMMAND, 1, $hOK)
    If _WaitClose($hWnd, 800) Then Return

    ; วิธี 3: ControlClick แบบปกติ
    WinActivate($hWnd)
    WinWaitActive($hWnd, "", 1)
    ControlClick($hWnd, "", $hOK)
    If _WaitClose($hWnd, 800) Then Return

    ; วิธี 4: คลิกเมาส์จริงที่พิกัดปุ่ม
    Local $aPos = ControlGetPos($hWnd, "", $hOK)
    Local $aWin = WinGetPos($hWnd)
    If IsArray($aPos) And IsArray($aWin) Then
        MouseClick("left", $aWin[0] + $aPos[0] + Int($aPos[2] / 2), _
                           $aWin[1] + $aPos[1] + Int($aPos[3] / 2), 1, 0)
    EndIf
    If _WaitClose($hWnd, 800) Then Return

    ; วิธี 5: Enter / Esc
    WinActivate($hWnd)
    Send("{ENTER}")
    If _WaitClose($hWnd, 800) Then Return
    Send("{ESC}")
EndFunc

Func _WaitClose($hWnd, $iMs)
    Local $t = TimerInit()
    While TimerDiff($t) < $iMs
        If Not WinExists($hWnd) Then Return True
        Sleep(30)
    WEnd
    Return False
EndFunc



Func FinanceLockExit($hWndXp, $hStartTime, $oLogStudent, $iRow)
	 Sleep(200)
	 ;Send("{ENTER}")
    Local $hPtNote
	While True
		Sleep(1000)
		$hPtNote = WinGetHandle("PatientNoteViewDisplayForm")
		Sleep(200)
		if $hPtNote Then WinKill($hPtNote)
		Sleep(1000)
		If Findtocon($sLoadPtSuccess2, $aLoadPtPos[0], $aLoadPtPos[1], $aLoadPtPos[2], $aLoadPtPos[3], 0.80)  Then ExitLoop   ;refresh blue pic
		 ExitMaxTime($hStartTime,$oLogStudent,$iRow)
	WEnd
    Sleep($iSleepAfterLoad)
	Local $clk
	While True
		$clk = ControlClick($hWndXp, "left", "TcxButton8") ;close pt btn
		Sleep(1000)
		 ExitMaxTime($hStartTime,$oLogStudent,$iRow)
		 If $clk = 1 Then ExitLoop
	WEnd

	Local $hFeeSch
	Local $hSpsh
	 While True
		Sleep(300)
		$hFeeSch = WinGetHandle("OvstFeeScheduleEntryForm")
		Sleep(100)
	    if $hFeeSch Then WinKill($hFeeSch)
		Sleep(400)

		$hSpsh = WinGetHandle("HOSxPNHSOConfirmPrivilegeForm")
		Sleep(100)
		if $hSpsh Then WinKill($hSpsh)
		Sleep(400)
		if Findtocon($sFoundVisit, $aFvPos[0], $aFvPos[1], $aFvPos[2], $aFvPos[3], 0.80)   Then ExitLoop   ;รูปคนตำแหน่งกลางค่อนมาทางซ้าย
	 WEnd
	 Sleep(1000)
	 Return True
EndFunc

Func PtExit($hWndXp, $hStartTime, $oLogStudent, $iRow)  ; code same as FinanceLockExit
	Sleep(200)

	Local $hPtNote
	Local $clk
	Local $hFeeSch
	Local $hSpsh

	; 1. รอให้ฟอร์มผู้ป่วยโหลดเสร็จ
	While True
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)

		Sleep(1000)
		$hPtNote = WinGetHandle("PatientNoteViewDisplayForm")
		Sleep(200)
		If $hPtNote Then WinKill($hPtNote)
		Sleep(500)

		If Findtocon($sLoadPtSuccess2, $aLoadPtPos[0], $aLoadPtPos[1], $aLoadPtPos[2], $aLoadPtPos[3], 0.80)  Then ExitLoop   ;refresh blue pic
	WEnd

	Sleep($iSleepAfterLoad)

	; 2. กดปิดหน้าผู้ป่วย
	While True
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)

		$clk = ControlClick($hWndXp, "", "TcxButton8", "left")
		Sleep(1000)

		If $clk = 1 Then ExitLoop
	WEnd

	; 3. ปิด popup หลังออก และรอกลับหน้ารายการ
	While True
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)

		Sleep(300)
		$hFeeSch = WinGetHandle("OvstFeeScheduleEntryForm")
		Sleep(100)
		If $hFeeSch Then WinKill($hFeeSch)
		Sleep(400)

		$hSpsh = WinGetHandle("HOSxPNHSOConfirmPrivilegeForm")
		Sleep(100)
		If $hSpsh Then WinKill($hSpsh)
		Sleep(400)

		if Findtocon($sFoundVisit, $aFvPos[0], $aFvPos[1], $aFvPos[2], $aFvPos[3], 0.80)   Then ExitLoop   ;รูปคนตำแหน่งกลางค่อนมาทางซ้าย
	WEnd

	Sleep(1000)
	Return True
EndFunc
;=============================== end after load exit function ======================================================================



;=============================== start record Hx function ========================================================================
Func ChiefComp($vDetail, $hStartTime, $oLogStudent, $hWndXp, $iRow)
	Local $ccclick =  "TcxCustomInnerTextEdit25"  ;TcxCustomInnerTextEdit25
	Local $ccadd = $sBtnSaveCC

	Local $cctextadd =  ""  ; $sSchool
	Local $cup =  ""   ;$aDetail[0]
	Local $date_exam = ""  ;$aDetail[1]

    ;MouseClick("left",417, 426,1,1)
	If IsArray($vDetail) Then
		If UBound($vDetail) > 2 Then
			$cctextadd = $vDetail[13] & " " & $sSchool   ;CC for student is class and school
		Else
			; elder $vDetail length = 2
			Local $cup = $vDetail[0]
			Local $date_exam = $vDetail[1]
			$cctextadd = $sSchool &" "& $cup &" "& $date_exam
		ElseIf
	Else
		$cctextadd = $sSchool  ;CC for telemed is school only config $vDetail = ""
	EndIf

    ClipPut($cctextadd)
	Sleep(100)
	ControlClick($hWndXp, "left", $ccclick)
	Sleep(500)
	_Cpp_ctrlV()
	Sleep(500)

	Local $cctext = ""
	Local $ccstatus = False

	For $a = 1 To 4
		; อ่านค่าใหม่ทุกรอบ เพื่อรอให้ข้อความแสดงผล
		$cctext = ControlGetText($hWndXp, "", "TcxDBTextEdit4")

		If StringLen($cctext) > 10 Then
			$ccstatus = True
			ExitLoop
		EndIf
		Sleep(500)
	Next

    If $ccstatus Then
		 ControlClick($hWndXp, "left", $ccadd)
	Else
		Send("{ENTER}") ;for close not found CC text popup
		Sleep(700)
		ControlClick($hWndXp, "left", $ccclick)
		ClipPut($cctextadd)
		Sleep(100)
		Send("{ESC}") ;_Cpp_ctrlA()
	    Sleep(700)
        _Cpp_ctrlV()
	    Sleep(700)
		; Send("{ENTER}")
		Sleep(800)
		ControlClick($hWndXp, "left",  $ccadd)
	EndIf
	Sleep(500)
	ExitMaxTime($hStartTime,$oLogStudent,$iRow)
	Return True
EndFunc

Func Allergy()
  Sleep(200)
  Local $null_allergy = False
  Local $a = 1
  While $a < 5
	 Sleep(50)
	 if  FindToClick($sRec_allergy_iden, $aRecAllergyPos[0], $aRecAllergyPos[1], $aRecAllergyPos[2], $aRecAllergyPos[3], 0.75) Then
		  $null_allergy = True
		  Sleep(200)
		 ExitLoop
	 EndIf
	  $a = $a + 1
  WEnd

  if $null_allergy Then
	    Send("{DOWN}")
		Sleep(300)
		Send("{DOWN}")
		Sleep(300)
		Send("{DOWN}")
		Sleep(300)
  EndIf

EndFunc

Func ClickSaveOperation($hWndItem)
    If Not WinExists($hWndItem) Then
        ConsoleWrite("ERROR: ไม่พบฟอร์มบันทึกหัตถการ" & @CRLF)
        Return False
    EndIf

    If Not ControlClick($hWndItem, "", $sBtnSaveOperation) Then
        ConsoleWrite("ERROR: ControlClick ปุ่มบันทึกไม่สำเร็จ" & @CRLF)
        Return False
    EndIf

    Return True
EndFunc
;===============================  end record Hx function ========================================================================



;=============================== start save function ============================================================================
Func FinalSave1($hStartTime, $oLogStudent, $hWndXp, $iRow)
    Local $clk
	Local $hControl
	Sleep(500)
	While 1
		Sleep(300)
		$hControl = ControlGetHandle($hWndXp, "", "TcxButton7")
		If $hControl Then ExitLoop
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)
    WEnd
	While True
		Sleep(300)
		 $clk = ControlClick($hWndXp, "left",$hControl)
		 ExitMaxTime($hStartTime,$oLogStudent,$iRow)
		 If $clk = 1 Then ExitLoop
	WEnd
	Return True
EndFunc

Func Select021($hStartTime, $oLogStudent, $iRow)
	Local $tt = "OPDSignDoctorEntryForm"
	Local $hWnd = WinWait($tt, "", 5)
	If Not WinActive($hWnd) Then WinActivate($hWnd)
	; WinActivate($hWnd)
	Local $sendCl = "TcxDBLookupComboBox2"
	ControlClick($hWnd, "left", $sendCl)
	Sleep(700)
	ControlSend($hWnd, "", $sendCl,"{DELETE}")
	Sleep(200)
	ControlSend($hWnd, "", $sendCl, "{ESC}")
	Sleep(200)
	ControlSend($hWnd,  "", $sendCl, "021")
	Sleep(300)
	Send("{ENTER}")
	Local $sSend = ""
	   Do
	      $sSend = ControlGetText($hWnd, "", $sendCl)
		  Sleep(200)
		  ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	   Until  $sSend = "021 ห้องการเงิน"
	  ;MsgBox(0, "สิทธิ", $sSend,1)
EndFunc

Func FinalSave2($hStartTime, $oLogStudent, $iRow)
	 Local $hFeeSch
	 Local $hSpsh
	 Local $hNoAuthen
	While True
		Sleep(1000)
		ExitMaxTime($hStartTime,$oLogStudent,$iRow)
		if FindToClick($sFinalSave2, $aFinalS2Pos[0], $aFinalS2Pos[1], $aFinalS2Pos[2], $aFinalS2Pos[3], 0.75)  Then  ExitLoop
	WEnd

	While True
		Sleep(100)
		$hSpsh = WinGetHandle("HOSxPNHSOConfirmPrivilegeForm")
		Sleep(100)
		if $hSpsh Then WinKill($hSpsh)
		Sleep(300)

		$hNoAuthen = WinGetHandle("BMSTextMessageDialogForm")
		Sleep(100)
		if $hNoAuthen Then WinKill($hNoAuthen)
		Sleep(300)

		$hFeeSch = WinGetHandle("OvstFeeScheduleEntryForm")  ;appear first but delay more than  200 ms
		Sleep(100)
		if $hFeeSch Then WinKill($hFeeSch)
		Sleep(300)

        FindToClick($sLastOK, $aLastOkPos[0], $aLastOkPos[1], $aLastOkPos[2], $aLastOkPos[3], 0.75)
		Sleep(200)
		if Findtocon($sFoundVisit, $aFvPos[0], $aFvPos[1], $aFvPos[2], $aFvPos[3], 0.75)  Then ExitLoop  ;รูปคนที่ตำแหน่งค่อนมาตรงกลาง
		ExitMaxTime($hStartTime,$oLogStudent,$iRow)
	WEnd
	Return True
EndFunc

Func FinalSave2Krg($hStartTime,$oLogStudent,$iRow)
	 Local $hFeeSch
	 Local $hSpsh
	 Local $hNoAuthen

	While True
		Sleep(300)
		if Findtocon($sFinalSave2, $aFinalS2Pos[0], $aFinalS2Pos[1], $aFinalS2Pos[2], $aFinalS2Pos[3], 0.75)  Then ExitLoop
		ExitMaxTime($hStartTime,$oLogStudent,$iRow)
	WEnd

    Sleep(300)
    Select021($hStartTime, $oLogStudent, $iRow)
	;ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	FindToClick($sFinalSave2, $aFinalS2Pos[0], $aFinalS2Pos[1], $aFinalS2Pos[2], $aFinalS2Pos[3], 0.75)
	While True
		Sleep(100)
		$hSpsh = WinGetHandle("HOSxPNHSOConfirmPrivilegeForm")
		Sleep(100)
		if $hSpsh Then WinKill($hSpsh)
		Sleep(300)

		$hNoAuthen = WinGetHandle("BMSTextMessageDialogForm")
		Sleep(100)
		if $hNoAuthen Then WinKill($hNoAuthen)
		Sleep(300)

		$hFeeSch = WinGetHandle("OvstFeeScheduleEntryForm") ;appear first but delay more than  200 ms
		Sleep(100)
		if $hFeeSch Then WinKill($hFeeSch)
		Sleep(300)

        FindToClick($sLastOK, $aLastOkPos[0], $aLastOkPos[1], $aLastOkPos[2], $aLastOkPos[3], 0.75)
		Sleep(200)
		if Findtocon($sFoundVisit, $aFvPos[0], $aFvPos[1], $aFvPos[2], $aFvPos[3], 0.75)  Then ExitLoop  ;รูปคนที่ตำแหน่งค่อนมาตรงกลาง
		ExitMaxTime($hStartTime,$oLogStudent,$iRow)
	WEnd
	Return True
EndFunc
;===============================  end save function ============================================================================



;=============================== short cut F1 F4 function ========================================================================
Func Click_F1($hWndXp, $hStartTime, $oLogStudent, $iRow)
	If Not WinExists($hWndXp) Then
		MsgBox($MB_ICONERROR, "Error", "HOSxP window was not found.")
		Return False
	EndIf
	;Local $hStartTime = TimerInit()
	; Activate once if HOSxP is not already active
	If Not WinActive($hWndXp) Then WinActivate($hWndXp)
	While Not WinActive($hWndXp)
		Sleep(200)
		ExitMaxTime($hStartTime,$oLogStudent,$iRow)
	WEnd
	Sleep(300)
	If Not WinActive($hWndXp) Then
		MsgBox($MB_ICONERROR, "Error", "F1 cancelled: the active window changed.")
		Return False
	EndIf
	Send("{F1}")
	Sleep(200)
	Return True
EndFunc

Func Click_F4($hWndXp, $hStartTime, $oLogStudent, $iRow)
	If Not WinExists($hWndXp) Then
		MsgBox($MB_ICONERROR, "Error", "HOSxP window was not found.")
		Return False
	EndIf
	;Local $hStartTime = TimerInit()
	; Activate once if HOSxP is not already active
	If Not WinActive($hWndXp) Then WinActivate($hWndXp)
	While Not WinActive($hWndXp)
		Sleep(200)
		ExitMaxTime($hStartTime,$oLogStudent,$iRow)
	WEnd
	Sleep(300)
	If Not WinActive($hWndXp) Then
		MsgBox($MB_ICONERROR, "Error", "F1 cancelled: the active window changed.")
		Return False
	EndIf
	Send("{F4}")
	Sleep(1000)
	Return True
EndFunc

Func Ctrl_F2($hWndXp, $hStartTime, $oLogStudent, $iRow)
	If Not WinExists($hWndXp) Then
		MsgBox($MB_ICONERROR, "Error", "HOSxP window was not found.")
		Return False
	EndIf
	; Activate once if HOSxP is not already active
	If Not WinActive($hWndXp) Then WinActivate($hWndXp)
	While Not WinActive($hWndXp)
		Sleep(200)
		ExitMaxTime($hStartTime,$oLogStudent,$iRow)
	WEnd
	Sleep(300)
	If Not WinActive($hWndXp) Then
		MsgBox($MB_ICONERROR, "Error", "Ctrl F2 cancelled: the active window changed.")
		Return False
	EndIf
	_Cpp_ctrlF2()
	Sleep(500)
	Return True
EndFunc


Func ClickAddOperation($hWndXp, $hStartTime, $oLogStudent, $iRow)
    Local $sFormTitle = $sOperationBoxTitle
    Local $sButton = $sBtnAddOperation

    If Not WinExists($hWndXp) Then
        MsgBox($MB_ICONERROR, "Error", "HOSxP window was not found.")
        Return False
    EndIf

    ; ป้องกันฟอร์มเดิมค้างอยู่
    If WinExists($sFormTitle) Then
        MsgBox($MB_ICONWARNING, "Test", "Close the operation entry form before testing.")
        Return False
    EndIf

    Local $hButton = ControlGetHandle($hWndXp, "", $sButton)

    If @error Or $hButton = 0 Then
        MsgBox($MB_ICONERROR, "Error", "The Add Operation button was not found.")
        Return False
    EndIf

    If ControlGetText($hWndXp, "", $hButton) <> "เพิ่มหัตถการ" Then
        MsgBox($MB_ICONERROR, "Error", "The button text does not match.")
        Return False
    EndIf

    If ControlClick($hWndXp, "", $hButton, "left", 1) <> 1 Then
        MsgBox($MB_ICONERROR, "Error", "ControlClick failed.")
        Return False
    EndIf

    ; รอจนกว่าจะพบฟอร์ม หรือ ExitMaxTime สั่งปิดสคริปต์
    While True
        Sleep(500)

        Local $hForm = WinGetHandle($sFormTitle)

        If $hForm <> 0 Then
            If _WinAPI_IsWindowVisible($hForm) Then
                Return True
            EndIf
        EndIf

        ExitMaxTime($hStartTime, $oLogStudent, $iRow)
    WEnd

EndFunc
;=============================== end short cut F1 F4 function ========================================================================

#endif