#RequireAdmin
#include <ScreenCapture.au3>
#include <MsgBoxConstants.au3>
Opt("MouseCoordMode", 1)
Global Const $kScreen = 1 ; เครื่องพัฒนา 125% = 1.25 / เครื่องจริง 100% = 1
Global Const $sOcrUrl = "http://127.0.0.1:3078"
Global $aAddItemBtnPos = [170, 118, 546, 241]
Global $bTesting = False
Global $gHttpFailed = False
Global $aItemTxtPos = [241, 252, 719, 461]  ;item txt position
Global $aOpdPricePos = [198, 547, 1083, 988]
Global $aEditPricePos = [154, 638, 1515, 941]
Global $aAddHPIPos = [92, 101, 420, 249]

HotKeySet("{F2}", "RunTest")
HotKeySet("{ESC}", "StopTest")
ToolTip("OCR Test: F1 = test / ESC = exit", 10, 10)

While True
    Sleep(100)
WEnd

Func RunTest()
    If $bTesting Then Return
    $bTesting = True
    ToolTip("")
    ; TestToClickOCR("TELEMED อายุ 40 ปีขึ้นไป",$aItemTxtPos[0],  $aItemTxtPos[1],  $aItemTxtPos[2],  $aItemTxtPos[3], 60)
	 TestToClickOCR("TELEMED อายุน้อยกว่า 40 ปี",$aItemTxtPos[0],  $aItemTxtPos[1],  $aItemTxtPos[2],  $aItemTxtPos[3], 60)
	; TestToClickOCR("ค่าธรรมเนียมบัตรสุขภาพถ้วนหน้า", $aOpdPricePos[0], $aOpdPricePos[1], $aOpdPricePos[2], $aOpdPricePos[3], 60)
	; TestToClickOCR("Inv Setting", $aEditPricePos[0], $aEditPricePos[1], $aEditPricePos[2], $aEditPricePos[3], 60)
	; TestToClickOCR("ประวัติ [Ctrl+F2]", $aAddHPIPos[0], $aAddHPIPos[1], $aAddHPIPos[2], $aAddHPIPos[3], 60)
    $bTesting = False
EndFunc
; $minConfidence เป็นค่า OCR 0-100 ไม่ใช่ค่า tolerance ของ OpenCV
; คืน True เมื่อพบหนึ่งตำแหน่งและเลื่อนเมาส์แล้วเท่านั้น ไม่มีการคลิก
Func TestToClickOCR($sTarget = "เพิ่มหัตถการ", $x1 = 170, $y1 = 118, $x2 = 546, $y2 = 241, $minConfidence = 60)
    If $x2 <= $x1 Or $y2 <= $y1 Then Return SetError(1, 0, False)

    Local $sOcrDir = @ScriptDir & "\TempOCR"

	If Not FileExists($sOcrDir) Then
		If Not DirCreate($sOcrDir) Then
			ConsoleWrite("OCR: สร้างโฟลเดอร์ไม่สำเร็จ: " & $sOcrDir & @CRLF)
			Return SetError(2, 0, False)
		EndIf
	EndIf

    Local $sImage = $sOcrDir & "\hosxp_ocr_" & @AutoItPID & ".png"
    Local $iLeft = Round($x1 * $kScreen), $iTop = Round($y1 * $kScreen)
    Local $iRight = Round($x2 * $kScreen), $iBottom = Round($y2 * $kScreen)
    FileDelete($sImage)
    _ScreenCapture_Capture($sImage, $iLeft, $iTop, $iRight, $iBottom, False)
    If Not FileExists($sImage) Then
        MsgBox($MB_ICONERROR, "OCR Test", "จับภาพไม่สำเร็จ")
        Return SetError(2, 0, False)
    EndIf
    Local $hFile = FileOpen($sImage, 16)
    If $hFile = -1 Then Return SetError(2, 0, False)
    Local $bImage = FileRead($hFile)
    FileClose($hFile)
    ; ส่งคำภาษาไทยเป็น UTF-8 hex เพื่อไม่ต้องใช้ JSON UDF
    Local $sHex = StringTrimLeft(String(StringToBinary($sTarget, 4)), 2)
    $gHttpFailed = False
    Local $oHttp = ObjCreate("WinHttp.WinHttpRequest.5.1")
    If Not IsObj($oHttp) Then Return SetError(3, 0, False)
    Local $oError = ObjEvent("AutoIt.Error", "HttpError")
    $oHttp.SetTimeouts(3000, 3000, 10000, 180000)
    $oHttp.Open("POST", $sOcrUrl & "/ocr?target=" & $sHex & "&confidence=" & $minConfidence, False)
    $oHttp.SetRequestHeader("Content-Type", "image/png")
    $oHttp.Send($bImage)
    If $gHttpFailed Then
        MsgBox($MB_ICONERROR, "OCR Test", "ติดต่อ Python OCR ไม่สำเร็จ กรุณาตรวจสอบหน้าต่าง server.py")
        Return SetError(3, 0, False)
    EndIf
    Local $iStatus = $oHttp.Status
    If $iStatus <> 200 Then
        MsgBox($MB_ICONERROR, "OCR Test", "ติดต่อ OCR ไม่สำเร็จ HTTP=" & $iStatus & @CRLF & "ตรวจสอบว่า Python OCR แสดง READY แล้ว")
        Return SetError(3, 0, False)
    EndIf
    Local $sResponse = $oHttp.ResponseText
    Local $aResult = StringSplit($sResponse, @TAB, 2)
    If UBound($aResult) < 5 Then Return SetError(4, 0, False)
    ConsoleWrite($sResponse & @CRLF)
    If $aResult[0] <> "MATCH" Then
        MsgBox($MB_ICONWARNING, "OCR Test: " & $aResult[0], "คำเป้าหมาย: " & $sTarget & @CRLF & "OCR อ่านได้: " & $aResult[4] & @CRLF & "ภาพทดสอบ: " & $sImage)
        Return False
    EndIf
    ; OCR คืนพิกัดภายในภาพ: บวกต้น area ก่อนหาร scaling ตาม template เดิม
    Local $realX = Round(($iLeft + Number($aResult[1])) / $kScreen)
    Local $realY = Round(($iTop + Number($aResult[2])) / $kScreen)
    MouseMove($realX, $realY, 10)
    ; ToolTip("Found: " & $sTarget & " | confidence=" & Round(Number($aResult[3]), 1) & " | x=" & $realX & " y=" & $realY)
    Return True
EndFunc

Func HttpError($oError)
    $gHttpFailed = True
    ConsoleWrite("HTTP COM error: " & $oError.description & @CRLF)
EndFunc

Func StopTest()
    ToolTip("")
    Exit
EndFunc
