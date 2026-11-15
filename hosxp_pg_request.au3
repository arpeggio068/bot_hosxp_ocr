#include <MsgBoxConstants.au3>

Global $gHTTPNode = ObjCreate("WinHttp.WinHttpRequest.5.1")
Global $oErr = ObjEvent("AutoIt.Error", "_ComErrHandler")
Global $sReqPath = "query_telemed"  ; query_student  query_telemed
Global $gHTTPError = False

Func _ComErrHandler($oError)
    $gHTTPError = True
    ConsoleWrite("COM Error: " & $oError.description & @CRLF)
EndFunc

; ส่งเฉพาะ HN ไปยัง local API
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

; รับผลลัพธ์เป็น array: [pttype, bw, oqueue]
Func GetSitFromDb2($iHnVal)
    Local $sHn = String($iHnVal)

    If StringLen($sHn) <> 9 Or Not StringRegExp($sHn, "^[0-9]{9}$") Then
        MsgBox($MB_ICONERROR, "Invalid HN", "HN ต้องเป็นตัวเลข 9 หลัก")
        Exit(1)
    EndIf

    Local $sResult = QueryPostgres2($sHn)

    If StringInStr($sResult, '"error"') Then
        MsgBox($MB_ICONERROR, "API Error", _
            "ไม่สามารถอ่านข้อมูลจาก local API ได้:" & @CRLF & $sResult)
        Exit(1)
    EndIf

    Local $aResult = StringRegExp($sResult, _
        '^\s*\[\s*"([0-9]+[dno]?)"\s*,\s*"(-?[0-9]+)"\s*,\s*"([0-9]+)"\s*\]\s*$', 1)

    If @error Then
        MsgBox($MB_ICONERROR, "API Error", _
            "รูปแบบผลลัพธ์จาก API ไม่ถูกต้อง:" & @CRLF & $sResult)
        Exit(1)
    EndIf

    Return $aResult
EndFunc

; ====== ตัวอย่างใช้งาน ======
Local $sHn = "000036802"
Local $aSit = GetSitFromDb2($sHn)

ConsoleWrite("sit : " & $aSit[0] & @CRLF)
ConsoleWrite("bw  : " & $aSit[1] & @CRLF)
ConsoleWrite("Q   : " & $aSit[2] & @CRLF)

Exit(0)