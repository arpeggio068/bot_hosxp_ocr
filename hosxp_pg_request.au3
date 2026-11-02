#include <MsgBoxConstants.au3>

Global $gHTTPNode = ObjCreate("WinHttp.WinHttpRequest.5.1")
Global $oErr = ObjEvent("AutoIt.Error", "_ComErrHandler")

Func _ComErrHandler($oError)
    ; เวลามี COM error (เช่น timeout) ให้เก็บ error message ไว้
    ConsoleWrite("COM Error: " & $oError.description & @CRLF)
    ; คืนค่าเฉย ๆ เพื่อไม่ให้ AutoIt หยุดทำงาน
EndFunc

Func QueryPostgres1($sSQL)
    ;Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")
    Local $sUrl = "http://localhost:3074/query"
    Local $sData = '{"sql":"' & StringReplace($sSQL, '"', '\"') & '"}'
	; เปิด connection
    $gHTTPNode.Open("POST", $sUrl, False)
    $gHTTPNode.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
	; ส่ง request
    $gHTTPNode.Send($sData)
    ; ใช้ ResponseBody (Binary) → ADODB.Stream → UTF-8 → Unicode
    Local $bResponse = $gHTTPNode.ResponseBody
    Local $oStream = ObjCreate("ADODB.Stream")
    $oStream.Type = 1          ; adTypeBinary
    $oStream.Open
    $oStream.Write($bResponse)
    $oStream.Position = 0
    $oStream.Type = 2          ; adTypeText
    $oStream.Charset = "utf-8" ; decode UTF-8
    Local $sResponse = $oStream.ReadText
    $oStream.Close
    ConsoleWrite("qpg1 : "&$sResponse&@CRLF)
    Return $sResponse
EndFunc

Func QueryPostgres2($sSQL)
    ;Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")
    Local $sUrl = "http://localhost:3074/query"
    Local $sData = '{"sql":"' & StringReplace($sSQL, '"', '\"') & '"}'

    ; ตั้ง timeout: resolve=5s, connect=5s, send=5s, receive=10s
    $gHTTPNode.SetTimeouts(5000, 5000, 5000, 5000)
    Local $sResponse = ""

    ; ลองส่ง request ถ้า COM error จะไปเข้าที่ _ComErrHandler
    $gHTTPNode.Open("POST", $sUrl, False)
    $gHTTPNode.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
    $gHTTPNode.Send($sData)

    ; ถ้าไม่ error → อ่าน ResponseText
    If Not @error Then
        $sResponse = $gHTTPNode.ResponseText
    EndIf

    ; ถ้า response ยังว่าง → ถือว่า timeout/เชื่อมต่อไม่ได้
    If $sResponse = "" Then
        $sResponse = '{"error":"error time out"}'
    EndIf

    ConsoleWrite("qpg2 : " & $sResponse & @CRLF)
    Return $sResponse
EndFunc

Func GetSitFromDb1($iHnVal)
	Local $sSQL = "SELECT ovst.vn, ovst.oqueue, ovst.pttype, " & _
			  "CASE WHEN COALESCE(ofs.nhso_fee_schedule_type_id, 999) = 101 THEN 'd'  " & _
			  "WHEN COALESCE(ofs.nhso_fee_schedule_type_id, 999) = 102 THEN 'o'  " & _
			  "ELSE 'n' END AS lst,  " & _
			  "COALESCE(pbw.bw,0)::int AS bw  " & _
              "FROM ovst " & _
			  "LEFT JOIN ovst_fee_schedule ofs ON ofs.vn = ovst.vn " & _
			  "LEFT JOIN ( " & _
			  "SELECT opdscreen.vn, opdscreen.hn, opdscreen.bw, " & _
			  "CASE WHEN opdscreen.vn = MAX(opdscreen.vn) OVER (PARTITION BY opdscreen.hn) THEN '1' ELSE '0' END AS flag_last_vn " & _
			  "FROM opdscreen " & _
			  "WHERE opdscreen.hn = '" & $iHnVal & "' " & _
			  "AND opdscreen.bw IS NOT NULL ORDER BY vn " & _
			  ") AS pbw ON pbw.hn = ovst.hn AND pbw.flag_last_vn = '1' " & _
              "WHERE ovst.hn = '" & $iHnVal & "' " & _
              "AND ovst.vstdate = CURRENT_DATE " & _
			  "ORDER BY ovst.vn DESC " & _
              "LIMIT 1;"
	Local $result = QueryPostgres1($sSQL)
	Local $pttype = StringRegExpReplace($result, '.*"pttype":"(.*?)".*', '\1')
	Local $lst    = StringRegExpReplace($result, '.*"lst":"(.*?)".*', '\1')
	Local $bw   = StringRegExpReplace($result, '.*"bw":("?)(\d+)\1.*', '\2')
	Local $oqueue   = StringRegExpReplace($result, '.*"oqueue":("?)(\d+)\1.*', '\2')

	If $result = "[]" Or $pttype = "" Or $pttype = $result Then $pttype = "888"
	If ($pttype = "25" Or $pttype = "47") And _
	   ($lst = "d" Or $lst = "n" Or $lst = "o") Then
		$pttype &= $lst
	EndIf

	Local $ArrayDb[3]
	$ArrayDb[0] = $pttype
	$ArrayDb[1] = $bw
	$ArrayDb[2] = $oqueue

	Return $ArrayDb
EndFunc

;~ Func GetSitFromDb2($iHnVal)
;~     Local $sSQL = "SELECT ovst.vn, ovst.oqueue, ovst.pttype, " & _
;~                   "CASE WHEN COALESCE(ofs.nhso_fee_schedule_type_id, 999) = 101 THEN 'd'  " & _
;~                   "WHEN COALESCE(ofs.nhso_fee_schedule_type_id, 999) = 102 THEN 'o'  " & _
;~                   "ELSE 'n' END AS lst,  " & _
;~                   "COALESCE(pbw.bw,0)::int AS bw  " & _
;~                   "FROM ovst " & _
;~                   "LEFT JOIN ovst_fee_schedule ofs ON ofs.vn = ovst.vn " & _
;~                   "LEFT JOIN ( " & _
;~                   "SELECT opdscreen.vn, opdscreen.hn, opdscreen.bw, " & _
;~                   "CASE WHEN opdscreen.vn = MAX(opdscreen.vn) OVER (PARTITION BY opdscreen.hn) THEN '1' ELSE '0' END AS flag_last_vn " & _
;~                   "FROM opdscreen " & _
;~                   "WHERE opdscreen.hn = '" & $iHnVal & "' " & _
;~                   "AND opdscreen.bw IS NOT NULL ORDER BY vn " & _
;~                   ") AS pbw ON pbw.hn = ovst.hn AND pbw.flag_last_vn = '1' " & _
;~                   "WHERE ovst.hn = '" & $iHnVal & "' " & _
;~                   "AND ovst.vstdate = CURRENT_DATE " & _
;~ 				  "ORDER BY ovst.vn DESC " & _
;~                   "LIMIT 1;"

;~     Local $result = QueryPostgres2($sSQL)

;~     ; --- ตรวจสอบ response error จาก database / API ---
;~     If StringInStr($result, '"error"') Then
;~         MsgBox(16, "Error", "API ไม่สามารถเชื่อมต่อฐานข้อมูลได้:" & @CRLF & $result)
;~         exit(0)
;~     EndIf

;~     ; --- ปกติ ดึงค่าจาก response ---
;~     Local $pttype = StringRegExpReplace($result, '.*"pttype":"(.*?)".*', '\1')
;~     Local $lst    = StringRegExpReplace($result, '.*"lst":"(.*?)".*', '\1')
;~     Local $bw     = StringRegExpReplace($result, '.*"bw":("?)(\d+)\1.*', '\2')
;~ 	Local $oqueue   = StringRegExpReplace($result, '.*"oqueue":("?)(\d+)\1.*', '\2')

;~     ; --- ตั้งค่า default ถ้า response ว่าง ---
;~     If $result = "[]" Or $pttype = "" Or $pttype = $result Then $pttype = "888"

;~     ; --- ต่อท้าย lst ถ้าเป็น 25/47 และ lst มีค่า d/n/o ---
;~     If ($pttype = "25" Or $pttype = "47") And _
;~        ($lst = "d" Or $lst = "n" Or $lst = "o") Then
;~         $pttype &= $lst
;~     EndIf

;~     Local $ArrayDb[3]
;~     $ArrayDb[0] = $pttype
;~     $ArrayDb[1] = $bw
;~ 	$ArrayDb[2] = $oqueue

;~     Return $ArrayDb
;~ EndFunc

Func GetSitFromDb2($iHnVal)
    Local $sSQL = "SELECT ovst.vn, COALESCE(ovst.oqueue, 0) AS oqueue, ovst.pttype, " & _
                  "CASE WHEN COALESCE(ofs.nhso_fee_schedule_type_id, 999) = 101 THEN 'd'  " & _
                  "WHEN COALESCE(ofs.nhso_fee_schedule_type_id, 999) = 102 THEN 'o'  " & _
                  "ELSE 'n' END AS lst, " & _
                  "COALESCE(pbw.bw,0)::int AS bw " & _
                  "FROM ovst " & _
                  "LEFT JOIN ovst_fee_schedule ofs ON ofs.vn = ovst.vn " & _
                  "LEFT JOIN ( " & _
                  "SELECT opdscreen.vn, opdscreen.hn, opdscreen.bw, " & _
                  "CASE WHEN opdscreen.vn = MAX(opdscreen.vn) OVER (PARTITION BY opdscreen.hn) THEN '1' ELSE '0' END AS flag_last_vn " & _
                  "FROM opdscreen " & _
                  "WHERE opdscreen.hn = '" & $iHnVal & "' " & _
                  "AND opdscreen.bw IS NOT NULL ORDER BY vn " & _
                  ") AS pbw ON pbw.hn = ovst.hn AND pbw.flag_last_vn = '1' " & _
                  "WHERE ovst.hn = '" & $iHnVal & "' " & _
                  "AND ovst.vstdate = CURRENT_DATE " & _
                  "ORDER BY ovst.vn DESC " & _
                  "LIMIT 1;"

    Local $result = QueryPostgres2($sSQL)
    If StringInStr($result, '"error"') Then
        MsgBox(16, "Error", "API ไม่สามารถเชื่อมต่อฐานข้อมูลได้:" & @CRLF & $result)
        Exit(0)
    EndIf

    If $result = "[]" Then
        Local $ArrayDb[3]
        $ArrayDb[0] = "888"
        $ArrayDb[1] = "0"
        $ArrayDb[2] = "0"
        Return $ArrayDb
    EndIf

    Local $pttype  = StringRegExpReplace($result, '.*"pttype":"(.*?)".*', '\1')
    Local $lst     = StringRegExpReplace($result, '.*"lst":"(.*?)".*', '\1')
    Local $bw      = StringRegExpReplace($result, '.*"bw":("?)(\d+)\1.*', '\2')
    Local $oqueue  = StringRegExpReplace($result, '.*"oqueue":("?)(\d+)\1.*', '\2')

    If $pttype = "" Or $pttype = $result Then $pttype = "888"

    If ($pttype = "25" Or $pttype = "47") And _
       ($lst = "d" Or $lst = "n" Or $lst = "o") Then
        $pttype &= $lst
    EndIf

    Local $ArrayDb[3]
    $ArrayDb[0] = $pttype
    $ArrayDb[1] = $bw
    $ArrayDb[2] = $oqueue

    Return $ArrayDb
EndFunc


; ====== ตัวอย่างใช้งาน ======
Local $iHnVal = "000218214"

Local $sit = GetSitFromDb2($iHnVal)
ConsoleWrite("sit : "&$sit[0]&@CRLF)
ConsoleWrite("bw : "&$sit[1]&@CRLF)
ConsoleWrite("Q : "&$sit[2]&@CRLF)

Exit(0)

