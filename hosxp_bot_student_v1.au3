#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Match\dentist_Exp_icon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <Excel.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#Include <WinAPI.au3>
#include <WinAPISys.au3>
#Include <SendMessage.au3>
#include <FileConstants.au3>
#include <Array.au3>
#include <StringConstants.au3>
#include "OpenCV-Match_UDF_Mod.au3"
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <CppKeySend.au3>



Opt("MouseCoordMode", 1)
Opt("WinTitleMatchMode", 2)
HotKeySet("{ESC}", "Hkey")

;=============================== start GUI config ========================================================================
Global $sConfigFile = @ScriptDir &"\Config\student_config.txt"
Global $iDefaultValue1 = "D:\schooldatasheet\sum2567\Dental_LR_2567.xlsx"
Global $iDefaultValue2 = "data"
Global $iDefaultValue3 =  "ออกหน่วยตรวจฟัน โรงเรียนลูกรักเชียงของ 23 กรกฎาคม 2567"
Global $iDefaultValue4 = 1
Global $iDefaultValue5 = 2
Global $iDefaultValue6 = 27
Global $iDefaultValue7 = 2000
Global $iDefaultValue8 = 2

If FileExists($sConfigFile) Then
    Local $aConfig = FileReadToArray($sConfigFile)
	;_ArrayDisplay($aConfig)
    If IsArray($aConfig) Then
        If UBound($aConfig) >= 8 Then
            $iDefaultValue1 = $aConfig[0]
            $iDefaultValue2 = $aConfig[1]
			$iDefaultValue3 = $aConfig[2]
			$iDefaultValue4 = Number($aConfig[3])
			$iDefaultValue5 = Number($aConfig[4])
			$iDefaultValue6 = Number($aConfig[5])
			$iDefaultValue7 = Number($aConfig[6])
			$iDefaultValue8 = Number($aConfig[7])
        EndIf
    EndIf
EndIf

#Region ### START Koda GUI section ### Form=
Global $Form1 = GUICreate("Student Bot", 634, 502, 192, 124)
Global $Input1 = GUICtrlCreateInput($iDefaultValue1, 24, 32, 577, 21)
Global $Input2 = GUICtrlCreateInput($iDefaultValue2, 24, 88, 289, 21)
Global $Input3 = GUICtrlCreateInput($iDefaultValue3, 24, 144, 577, 21)
$Label1 = GUICtrlCreateLabel("Excel file path", 24, 8, 70, 17)
$Label2 = GUICtrlCreateLabel("Sheet name", 24, 64, 61, 17)
$Label3 = GUICtrlCreateLabel("CC", 24, 120, 18, 17)
Global $Input4 = GUICtrlCreateInput($iDefaultValue4, 24, 216, 113, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
Global $Input8 = GUICtrlCreateInput($iDefaultValue8, 240, 216, 113, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
Global $Input5 = GUICtrlCreateInput($iDefaultValue5, 24, 280, 113, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
Global $Input6 = GUICtrlCreateInput($iDefaultValue6, 240, 280, 113, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
Global $Input7 = GUICtrlCreateInput($iDefaultValue7, 24, 344, 113, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
$Label4 = GUICtrlCreateLabel("School Type 1 = รัฐบาล 2 = เอกชน", 24, 192, 168, 17)
$Label8 = GUICtrlCreateLabel("Fluoride 1 = เคลือบ 2 = ไม่เคลือบ", 240, 192, 168, 17)
$Label5 = GUICtrlCreateLabel("Start Row", 56, 256, 51, 17)
$Label6 = GUICtrlCreateLabel("End Row", 275, 256, 48, 17)
$Label7 = GUICtrlCreateLabel("Delay", 67, 320, 31, 17)
Global $botButton = GUICtrlCreateButton("Start Bot", 24, 384, 113, 25)
;GUICtrlSetBkColor($Button1, 0xf7f7d4 ) ; Set red color for alerts
; Create the alert label, initially hidden
Global $AlertLabel = GUICtrlCreateLabel("", 24, 416, 244, 17)
GUICtrlSetColor($AlertLabel, 0xFF0000) ; Set red color for alerts
GUICtrlSetState($AlertLabel, $GUI_HIDE)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

Global $sWorkbook = $iDefaultValue1
Global $sSheet = $iDefaultValue2
Global $sSchool = $iDefaultValue3
Global $iSchoolType = $iDefaultValue4 ; รัฐบาล = 1 เอกชน = 2
Global $iStartRow = $iDefaultValue5
Global $iEndRow = $iDefaultValue6
Global $iSleepAfterLoad = $iDefaultValue7
Global $iAddFluoride = $iDefaultValue8
;=============================== end GUI config ========================================================================


Global $aXpPos = [-8, -8, 1552, 840]
Global $sHosXp_Title = "BMS-HOSxP XE 4.0"
Global $sServer_Title = 'AutoIt Server'
Global $sGoogle_Title = 'Google Chrome'
Global $logStudent = @ScriptDir &"\BotLog\log_student.txt"
Global $logStudenProgress = @ScriptDir &"\BotLog\log_student_progress.txt"
Global $logPrice = @ScriptDir &"\BotLog\log_student_price.txt"

Global $hWndXp
Global $oExcel
Global $oLogStudent
Global $oLogStudentProgress
Global $oLogPrice
Global $oWorkbook
Global $iMaxRecTime = 4*60*1000 ; 4 min

;=== start telegram and database request config ===============================================
Global $gEnvCache = 0       ; เก็บ cache (Dictionary ของ KEY=VALUE)
Global $gEnvFile = @ScriptDir & "\server\.env"
Global $gEnvTimestamp = ""  ; เวลาล่าสุดของไฟล์
Global $gTelegramToken = ""
Global $gTelegramChatID = ""
Global $gHTTPTelegram = ObjCreate("WinHttp.WinHttpRequest.5.1")
Global $gHTTPNode = ObjCreate("WinHttp.WinHttpRequest.5.1")
Global $oErr = ObjEvent("AutoIt.Error", "_ComErrHandler")
;=== end telegram  and database request config ===============================================

;=== start picture config ==============================================================
;Global $sDtMenu = "\Match\xp_dental_menu.png"  ;ไม่ใช้
Global $sFoundVisit = "\Match\xp_hn_found_visit.png"   ;รูปชายหญิง
Global $aFvFirstPos = [0, 81,75,169] ;start loop position
Global $aFvPos = [256, 18, 683, 374]   ;end loop position
Global $sFiLock =  "\Match\xp_finance_lock.png"  ;ภาพกล่องข้อความ finance lock no use
Global $aFiLockPos = [558, 293, 1412, 767]
;Global $sClosePt = "\Match\xp_close_pt.png"  ;ไม่ใช้
Global $sLoadPtSuccess2 = "\Match\xp_load_pt_success2.png" ;รูป icon พยาบาลกากบาทสีน้ำเงิน
Global $aLoadPtPos = [159, 147,1880, 432]
;Global $sDentalCare =  "\Match\xp_dental_care.png" ;ไม่ใช้

Global $sF1 =   "\Match\xp_f1_click.png"  ;ปุ่มซักประวัติ no use
Global $aF1Pos = [ 8, 652,246, 872]
Global $sCC =   "\Match\xp_cc.png"  ;ข้อความ CC
Global $aCcPos = [253, 365, 436, 485]
;Global $sAddCC =   "\Match\xp_add_cc.png"  ;ไม่ใช้
Global $sRec_allergy_iden =   "\Match\rec_allergy_iden.png"  ;select box การแพ้ยา
Global $aRecAllergyPos = [166, 186, 931, 390]

Global $sF4 =   "\Match\xp_f4_click.png"  ;ปุ่มหัตถการ no use
Global $aF4Pos = [8, 652,246, 872]
Global $sAddItem =   "\Match\xp_add_item.png"  ;ปุ่มเครืองหมายบวกเพิ่มหัตถการ no use
Global $aAddItemBtnPos = [177, 190, 629, 489]
Global $sPratomItem =   "\Match\xp_item_found_pratom.png"  ;ข้อความนักเรียน ประถม will test
Global $sFluorideItem =   "\Match\xp_item_found_fluoride.png"  ;ข้อความ ฟลูออไรด์วาร์นิช will test
Global $aItemTxtPos = [417, 239, 1245, 658]  ;item txt position

Global $sSaveItem =   "\Match\xp_item_save.png"   ;ปุ่มติ๊กถูกในวงกลมสีเขียว สำหรับบันทึกหัตถการ no use
Global $aSaveItemBtnPos = [1012, 580, 1763, 994]  ; save item btn position
Global $sTask =   "\Match\xp_task_on_item.png"   ;เมนู Task สีเขียว no use
Global $aTaskPos = [351, 471, 848, 865]
Global $sOpdPrice1 =   "\Match\opd_price1.png"   ;ข้อความ ค่าบริการผู้ป่วยนอก พื้นสีฟ้า will test
Global $sOpdPrice2 =   "\Match\opd_price2.png"  ;ข้อความ ค่าบริการผู้ป่วยนอก ไม่มีพื้น will test
Global $aOpdPricePos = [159, 713, 1083, 988]
Global $sEditPrice =   "\Match\xp_edit_price.png"   ; ข้อความ Inv Setting
Global $aEditPricePos = [154, 638, 1515, 941]
;Global $sFinalSave =   "\Match\xp_final_save1.png"  ;ไม่ใช้

Global $sFinalSave2 =   "\Match\xp_final_save2.png"  ;ภาพติ๊กถูกในวงกลมสีเขียวตอนบันทึกครั้งสุดท้าย
Global $aFinalS2Pos = [496, 197, 1081, 534]
Global $sLastOK =   "\Match\xp_lastOK.png"
Global $aLastOkPos = [728, 341, 1272, 701]
;=== end picture config =================================================================



;=== start element class and name config ================================================
Global $sOperationBoxTitle = "HOSxPDentalOperationEntryForm" ;THOSxPDentalOperationEntryForm
Global $sBtnSaveOperation = "[CLASS:TcxButton; INSTANCE:10]"

;=== end element class and name config ================================================



Global $ArraySit[8]  ;สิทธิที่ต้องแก้ค่าบริการ 50 บาทเป็น 0 บาท ส่งกลับบ้าน 103
$ArraySit[0] = "40"  ;เบิกได้
$ArraySit[1] = "10"  ;ชำระเงิน
$ArraySit[2] = "14"  ; ชำระเงินต่างชาติ กลุ่ม 1
$ArraySit[3] = "15"  ; ชำระเงินต่างชาติ กลุ่ม 2
$ArraySit[4] = "16"  ;ชำระเงินต่างชาติ กลุ่ม 3
$ArraySit[5] = "49"  ;ไม่ได้นำหลักฐานมาด้วย
$ArraySit[6] = "54"  ;สิทธิ์ว่าง
$ArraySit[7] = "90"  ;ชำระเงินต่างชาติ

Global $ArraySitKrg[6]  ;สิทธิที่ต้องส่งต่อห้องการเงิน 021
$ArraySitKrg[0] = "01"  ; ขรก. (องค์กรปกครองส่วนท้องถิ่น)
$ArraySitKrg[1] = "55"  ; ขรก.กรมบัญชีกลาง
$ArraySitKrg[2] = "34"  ; ขรก.กรุงเทพมหานคร
$ArraySitKrg[3] = "35"  ; ขรก.กกต.
$ArraySitKrg[4] = "25o"  ;นักเรียน นอกเขต  o=op anywhere ต่างจังหวัด
$ArraySitKrg[5] = "47o"  ;เด็ก 0-12 ปี นอกเขต  o=op anywhere ต่างจังหวัด




;=============================== start utility function ==============================================================================================
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

Func ctrlV() ;use with ClipPut()
   Send("{CTRLDOWN}")  ; Press and hold the Ctrl key
   Send("v")  ; Simulate the V keypress (paste)
   Send("{CTRLUP}")   ; Release the Ctrl key
   Sleep(100)
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

Func TestToClick($picPath = @ScriptDir&"\Match\dad_name.png", $x1 = 0, $y1 = 370, $x2 = 51, $y2 = 400, $tol = 0.75)
 Local Const $kScreen = 1 ;constant for screen size 125% = 1.25  for screen size 100% = 1
 Local Const $kPosition = 0  ;constant for screen size 125% position. Set $kPosition = 0 if screen size = 100%
 Local $Match = _MatchPicture("", $picPath, $x1*$kScreen, $y1*$kScreen, $x2*$kScreen, $y2*$kScreen, $tol)
 $realX = 0
 $realY = 0
 If $kPosition > 0 Then
	$realX = $Match[4] - ($Match[4]/$kPosition)
    $realY = $Match[5] - ($Match[5]/$kPosition)
 Else
	$realX = $Match[4]
    $realY = $Match[5]
 EndIf
 ;_ArrayDisplay($Match1)
 If $Match[0] > 0 Then
	MouseMove($realX,$realY,10)
	;MouseClick($MOUSE_CLICK_LEFT, $realX, $realY, 1, 2)
 EndIf

EndFunc

 Func FindToClick($picPath = @ScriptDir&"\Match\dad_name.png", $x1 = 0, $y1 = 370, $x2 = 51, $y2 = 400, $tol = 0.75)
 Local Const $kScreen = 1 ;constant for screen size 125% = 1.25  for screen size 100% = 1
 Local Const $kPosition = 0  ;constant for screen size 125% position. Set $kPosition = 0 if screen size = 100%
 Local $Match = _MatchPicture("", $picPath, $x1*$kScreen, $y1*$kScreen, $x2*$kScreen, $y2*$kScreen, $tol)
 $bClick = False
 $realX = 0
 $realY = 0
 If $kPosition > 0 Then
	$realX = $Match[4] - ($Match[4]/$kPosition)
    $realY = $Match[5] - ($Match[5]/$kPosition)
 Else
	$realX = $Match[4]
    $realY = $Match[5]
 EndIf
 ;_ArrayDisplay($Match1)
 If $Match[0] > 0 Then
	;MouseMove($realX,$realY,10)
	MouseClick($MOUSE_CLICK_LEFT, $realX, $realY, 1, 2)
	$bClick = True
	;Sleep(1000)
 EndIf
 Return $bClick
 EndFunc

 Func FindToClickRt($picPath = @ScriptDir&"\Match\dad_name.png", $x1 = 0, $y1 = 370, $x2 = 51, $y2 = 400, $tol = 0.75)
 Local Const $kScreen = 1 ;constant for screen size 125% = 1.25  for screen size 100% = 1
 Local Const $kPosition = 0  ;constant for screen size 125% position. Set $kPosition = 0 if screen size = 100%
 Local $Match = _MatchPicture("", $picPath, $x1*$kScreen, $y1*$kScreen, $x2*$kScreen, $y2*$kScreen, $tol)
 $realX = 0
 $realY = 0
 If $kPosition > 0 Then
	$realX = $Match[4] - ($Match[4]/$kPosition)
    $realY = $Match[5] - ($Match[5]/$kPosition)
 Else
	$realX = $Match[4]
    $realY = $Match[5]
 EndIf
 ;_ArrayDisplay($Match1)
 If $Match[0] > 0 Then
	;MouseMove($realX,$realY,10)
	MouseClick( $MOUSE_CLICK_RIGHT , $realX, $realY, 1, 2)
	;Sleep(1000)
 EndIf

EndFunc

Func FindToCon($picPath = @ScriptDir&"\Match\dad_name.png", $x1 = 0, $y1 = 370, $x2 = 51, $y2 = 400, $tol = 0.75)
 Local $kScreen = 1 ;constant for screen size 125% = 1.25  for screen size 100% = 1
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

;=============== send request nodejs localhost func ======================================
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

Func QueryPostgres1($sSQL)
    ;Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")
    Local $sUrl = "http://localhost:3074/query"
    Local $sData = '{"sql":"' & StringReplace($sSQL, '"', '\"') & '"}'

    $gHTTPNode.Open("POST", $sUrl, False)
    $gHTTPNode.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
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
	;ConsoleWrite("qpg1 : "&$sResponse&@CRLF)
    Return $sResponse
EndFunc

Func QueryPostgres2($sSQL)
    ;Local $oHTTP = ObjCreate("WinHttp.WinHttpRequest.5.1")
    Local $sUrl = "http://localhost:3074/query"
    Local $sData = '{"sql":"' & StringReplace($sSQL, '"', '\"') & '"}'
    ; ตั้ง timeout: resolve=5s, connect=5s, send=5s, receive=10s
    $gHTTPNode.SetTimeouts(5000, 5000, 5000, 10000)
    Local $sResponse = ""
    ; ลองส่ง request ถ้า COM error จะไปเข้าที่ _ComErrHandler
    $gHTTPNode.Open("POST", $sUrl, False)
    $gHTTPNode.SetRequestHeader("Content-Type", "application/json; charset=utf-8")
    $gHTTPNode.Send($sData)

    If Not @error Then
        $sResponse = $gHTTPNode.ResponseText
    EndIf

    If $sResponse = "" Then
        $sResponse = '{"error":"error time out"}'
		ConsoleWrite("qpg2 : " & $sResponse & @CRLF)
    EndIf

    Return $sResponse
EndFunc

Func GetSitFromDb1($iHnVal)
	Local $sSQL = "SELECT ovst.pttype, " & _
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
              "LIMIT 1;"
	Local $result = QueryPostgres1($sSQL)
	Local $pttype = StringRegExpReplace($result, '.*"pttype":"(.*?)".*', '\1')
	Local $lst    = StringRegExpReplace($result, '.*"lst":"(.*?)".*', '\1')
	Local $bw   = StringRegExpReplace($result, '.*"bw":("?)(\d+)\1.*', '\2')

	If $result = "[]" Or $pttype = "" Or $pttype = $result Then $pttype = "888"
	If ($pttype = "25" Or $pttype = "47") And _
	   ($lst = "d" Or $lst = "n" Or $lst = "o") Then
		$pttype &= $lst
	EndIf

	Local $ArrayDb[2]
	$ArrayDb[0] = $pttype
	$ArrayDb[1] = $bw

	Return $ArrayDb
EndFunc

Func GetSitFromDb2($iHnVal, $oFile, $iRow)
    Local $sSQL = "SELECT ovst.vn, COALESCE(ovst.oqueue, 0) AS oqueue, ovst.pttype, " & _
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

    Local $result = QueryPostgres2($sSQL)
    ; --- ตรวจสอบ response error จาก database / API ---
    If StringInStr($result, '"error"') Then
		SendTeleGram("Error Server Connect R= "&$iRow)
		FileWrite($oFile, "Error Server Connect R= "&$iRow&", ")
		Sleep(200)
        MsgBox(16, "Error", "API ไม่สามารถเชื่อมต่อฐานข้อมูลได้:" & @CRLF & $result)
        exit(0)
    EndIf

	 If $result = "[]" Then
        Local $ArrayDb[3]
        $ArrayDb[0] = "888"
        $ArrayDb[1] = "0"
        $ArrayDb[2] = "0"
        Return $ArrayDb
    EndIf

    Local $pttype = StringRegExpReplace($result, '.*"pttype":"(.*?)".*', '\1')
    Local $lst    = StringRegExpReplace($result, '.*"lst":"(.*?)".*', '\1')
    Local $bw     = StringRegExpReplace($result, '.*"bw":("?)(\d+)\1.*', '\2')
	 Local $oqueue  = StringRegExpReplace($result, '.*"oqueue":("?)(\d+)\1.*', '\2')
    ; --- ตั้งค่า default ถ้า response ว่าง ---
    If $result = "[]" Or $pttype = "" Or $pttype = $result Then $pttype = "888"
    ; --- ต่อท้าย lst ถ้าเป็น 25/47 และ lst มีค่า d/n/o ---
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
;=============== send request nodejs localhost func ======================================

Func _ComErrHandler($oError)
    ; เวลามี COM error (เช่น timeout) ให้เก็บ error message ไว้
    ConsoleWrite("COM Error: " & $oError.description & @CRLF)
    ; คืนค่าเฉย ๆ เพื่อไม่ให้ AutoIt หยุดทำงาน
EndFunc
;=============================== end utility function =====================================================================================

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

Func HideServer()
	 Local $hEcel = WinWait($sServer_Title, "", 2)
	 ;Local $hEcel = WinGetHandle("[CLASS:XLMAIN]")
	if $hEcel Then WinSetState($hEcel, "", @SW_MINIMIZE)
EndFunc

Func SwapXP()
		;_ArrayDisplay($aList)
	Local $aList = WinList($sHosXp_Title) ;Winlist all hos os
	Local $aList_Length = $aList[0][0]

	If $aList_Length > 0 Then
				For $i = 1 To $aList_Length
					$sTitle = $aList[$i][0]
					$hHandle = $aList[$i][1]
						If $sTitle <> "" And BitAND(WinGetState($hHandle), 2) Then
							;MsgBox($MB_SYSTEMMODAL, "", "Title: " & $aList[$i][0] & @CRLF & "Handle: " & $aList[$i][1])
							WinActivate($sTitle)
							Sleep(1000)
						    ;WinMove($sTitle, "", $aXpPos[0], $aXpPos[1], $aXpPos[2], $aXpPos[3])
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
		    MsgBox($MB_SYSTEMMODAL, "Warning!","ไม่พบโปรแกรม Hos XP",10)
			MsgStopServer()
		    Exit
    EndIf

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
Func KillFinance1()
    Local $bFoundFinance = False
    Local $aList = WinList()

    For $i = 1 To $aList[0][0]
        Local $hWnd = $aList[$i][1]
        If $hWnd <> 0 Then
            Local $sClass = _WinAPI_GetClassName($hWnd)
            If $sClass = "#32770" And StringInStr($aList[$i][0], "HOSxPMedicationOrderQueueGeneratorForm") Then
                ConsoleWrite("found popup: " & $hWnd & " -> WinClose/WinKill" & @CRLF)
                ; ปิด popup
                WinClose($hWnd)
                Sleep(500)
                If WinExists($hWnd) Then
                    WinKill($hWnd)
                    ConsoleWrite("WinKill popup" & @CRLF)
                EndIf
                $bFoundFinance = True
                ExitLoop ; เจอแล้วออก loop ทันที
            EndIf
        EndIf
    Next

    If $bFoundFinance Then
        Return True
    Else
        Return False
    EndIf
EndFunc





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
		If Findtocon(@ScriptDir&$sLoadPtSuccess2, $aLoadPtPos[0], $aLoadPtPos[1], $aLoadPtPos[2], $aLoadPtPos[3], 0.80)  Then ExitLoop ;diagram รูปฟันด้านล่าง
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
		if Findtocon(@ScriptDir&$sFoundVisit, $aFvPos[0], $aFvPos[1], $aFvPos[2], $aFvPos[3], 0.80)   Then ExitLoop   ;รูปคนตำแหน่งกลางค่อนมาทางซ้าย
	 WEnd
	 Sleep(1000)
	 Return True
EndFunc

Func PtExit($hWndXp, $hStartTime, $oLogStudent, $iRow)
	Local $hPtNote
	While True
		Sleep(1000)
		$hPtNote = WinGetHandle("PatientNoteViewDisplayForm")
		Sleep(100)
		if $hPtNote Then WinKill($hPtNote)
		Sleep(1000)
		If Findtocon(@ScriptDir&$sLoadPtSuccess2, $aLoadPtPos[0], $aLoadPtPos[1], $aLoadPtPos[2], $aLoadPtPos[3], 0.80) Then ExitLoop ;diagram รูปฟันด้านล่าง
	WEnd
    Sleep($iSleepAfterLoad)
	Local $clk
	While True
		$clk = ControlClick($hWndXp, "left", "TcxButton8") ;close pt btn
		Sleep(1000)
		 ExitMaxTime($hStartTime,$oLogStudent,$iRow)
		 If $clk = 1 Then ExitLoop
	WEnd
	;FindToClick(@ScriptDir&$sClosePt,982, 54,1121, 105,  0.75)  ;close pt btn
	;Sleep(1000)
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
		if Findtocon(@ScriptDir&$sFoundVisit, $aFvPos[0], $aFvPos[1], $aFvPos[2], $aFvPos[3], 0.80)  Then ExitLoop  ;รูปคนตำแหน่งกลางค่อนมาทางซ้าย
	 WEnd
	 Sleep(1000)
	 Return True
EndFunc
;=============================== end after load exit function ======================================================================

;=============================== start record dental function ======================================================================
Func SendDental($aArray, $hStartTime, $oLogStudent, $iRow)
	Local $hWnd   ;= WinWait("HOSxPDentalCareEntryForm", "", 60)
	While 1
		Sleep(500)
		$hWnd = WinGetHandle("HOSxPDentalCareEntryForm")
		if $hWnd Then
			Sleep(1000)   ;important to set sleep time before control class appear
			ExitLoop
		EndIf
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	WEnd

	Local $pteeth =  "TcxCustomInnerTextEdit22"
	Local $pcaries =  "TcxCustomInnerTextEdit21"
	Local $pfilling = "TcxCustomInnerTextEdit20"
	Local $pextract = "TcxCustomInnerTextEdit19"

	Local $dteeth = "TcxCustomInnerTextEdit18"
    Local $dcaries = "TcxCustomInnerTextEdit17"
    Local $dfilling = "TcxCustomInnerTextEdit16"
    Local $dextract = "TcxCustomInnerTextEdit15"

	Local $need_sealant = "TcxCustomInnerTextEdit13"  ;TcxCustomInnerTextEdit13
    Local $need_pfilling = "TcxCustomInnerTextEdit11" ;TcxCustomInnerTextEdit11
    Local $need_dfilling = "TcxCustomInnerTextEdit9" ;TcxCustomInnerTextEdit9
	Local $need_pextract = "TcxCustomInnerTextEdit7"
    Local $need_dextract = "TcxCustomInnerTextEdit5"


    If Number($aArray[0]) > 0 Then CtrlSendDt($hWnd,$pteeth,$aArray[0])
	If Number($aArray[1]) > 0 Then	CtrlSendDt($hWnd,$pcaries,$aArray[1])
	If Number($aArray[2]) > 0 Then CtrlSendDt($hWnd,$pfilling,$aArray[2])
	If Number($aArray[3]) > 0 Then CtrlSendDt($hWnd,$pextract,$aArray[3])

	If Number($aArray[4]) > 0 Then	CtrlSendDt($hWnd,$dteeth,$aArray[4])
	If Number($aArray[5]) > 0 Then	CtrlSendDt($hWnd,$dcaries,$aArray[5])
	If Number($aArray[6]) > 0 Then CtrlSendDt($hWnd,$dfilling,$aArray[6])
	If Number($aArray[7]) > 0 Then CtrlSendDt($hWnd,$dextract,$aArray[7])

	If Number($aArray[8]) > 0 Then CtrlSendDt($hWnd,$need_sealant,$aArray[8])
	If Number($aArray[9]) > 0 Then CtrlSendDt($hWnd,$need_pfilling,$aArray[9])
	If Number($aArray[10]) > 0 Then CtrlSendDt($hWnd,$need_dfilling,$aArray[10])
	If Number($aArray[11]) > 0 Then CtrlSendDt($hWnd,$need_dextract,$aArray[11])
	If Number($aArray[12]) > 0 Then CtrlSendDt($hWnd,$need_pextract,$aArray[12])

	SelectPtType($hWnd, $aArray[13])
	SelectPlace($hWnd)
	if $iSchoolType = 1 Then
	   SelectBoxEduRtb($hWnd) ;โรงเรียนรัฐบาล
	Else
	   SelectBoxEduEkc($hWnd)  ;โรงเรียนเอกชน
	EndIf
	SelectClass($hWnd,$aArray[13])
	;$dSaveBtnClass = "TcxButton3"
	;$dDontSave = "TcxButton4"
	Local $clk2
	While 1
		$clk2 = ControlClick($hWnd, "left","TcxButton3")
		Sleep(700)
		If $clk2 = 1 Then ExitLoop
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	WEnd
	WinClose($hWnd,"")
	Return True
EndFunc

Func SelectBoxEduRtb($hWnd)
	;Local $sTitle = "HOSxPDentalCareEntryForm"
	;Local $hWnd = WinWait($sTitle, "", 10)
	Local $edu  = "TcxCustomComboBoxInnerEdit3"
	ControlClick($hWnd, "left", $edu)
	Sleep(100)
	Local $iClick = 2
	For $i = 1 To $iClick
		ControlSend($hWnd, "",$edu , "{DOWN}")
		Sleep(100)
	Next
EndFunc

Func SelectBoxEduEkc($hWnd)
	;Local $sTitle = "HOSxPDentalCareEntryForm"
	;Local $hWnd = WinWait($sTitle, "", 10)
	Local $edu  = "TcxCustomComboBoxInnerEdit3"
	ControlClick($hWnd, "left", $edu)
	Sleep(100)
	Local $iClick = 5
	For $i = 1 To $iClick
		ControlSend($hWnd, "",$edu , "{DOWN}")
		Sleep(100)
	Next
EndFunc

Func SelectPtType($hWnd,$sShoolClass)
	;Local $sTitle = "HOSxPDentalCareEntryForm"
	;Local $hWnd = WinWait($sTitle, "", 10)
	Local $class_  = "TcxCustomComboBoxInnerEdit6"  ;TcxCustomComboBoxInnerEdit6
	ControlClick($hWnd, "left", $class_)
	Sleep(100)
   Local $iClick = 2
   Switch $sShoolClass
		Case 'อ.1'
           $iClick = 2
		Case 'อ.2'
			$iClick = 2
		Case 'อ.3'
            $iClick = 2
		Case 'ป.1'
            $iClick = 3
		Case 'ป.2'
            $iClick = 3
		Case 'ป.3'
            $iClick = 3
		Case 'ป.4'
			$iClick = 3
        Case 'ป.5'
            $iClick = 3
        Case 'ป.6'
            $iClick = 3
	    Case Else
			$iClick = 2
   EndSwitch
   For $i = 1 To $iClick
		ControlSend($hWnd, "",$class_ , "{DOWN}")
		Sleep(100)
	Next
EndFunc

Func SelectPlace($hWnd)
	;Local $sTitle = "HOSxPDentalCareEntryForm"
	;Local $hWnd = WinWait($sTitle, "", 10)
	Local $place  = "TcxCustomComboBoxInnerEdit5"  ;TcxCustomComboBoxInnerEdit5
	ControlClick($hWnd, "left", $place)
	Sleep(100)
	For $i = 1 To 2
		ControlSend($hWnd, "",$place , "{DOWN}")
		Sleep(100)
	Next
EndFunc

Func SelectClass($hWnd,$sShoolClass)
	;Local $sTitle = "HOSxPDentalCareEntryForm"
	;Local $hWnd = WinWait($sTitle, "", 10)
	Local $class_  = "TcxCustomComboBoxInnerEdit2"
	ControlClick($hWnd, "left", $class_)
	Sleep(100)
  Local $iClick = 1
   Switch $sShoolClass
		Case 'อ.1'
           $iClick = 1
		Case 'อ.2'
			$iClick = 1
		Case 'อ.3'
            $iClick = 1
		Case 'ป.1'
            $iClick = 2
		Case 'ป.2'
            $iClick = 2
		Case 'ป.3'
            $iClick = 2
		Case 'ป.4'
			$iClick = 2
        Case 'ป.5'
            $iClick = 2
        Case 'ป.6'
            $iClick = 2
	    Case Else
			$iClick = 1
   EndSwitch
   For $i = 1 To $iClick
		ControlSend($hWnd, "",$class_ , "{DOWN}")
		Sleep(100)
	Next
EndFunc
;===============================  end record dental function ======================================================================

;=============================== start record Hx function ========================================================================
Func ChiefComp($schclass, $hStartTime, $oLogStudent, $hWndXp, $iRow)
	Local $ccclick =  "TcxCustomInnerTextEdit25"
	Local $ccadd = "TcxButton52"  ;TcxButton50
	Local $cctextadd = $schclass&" "&$sSchool
	Local $cctext = ""
			;MouseClick("left",417, 426,1,1)
    ClipPut($cctextadd)
	Sleep(100)
	ControlClick($hWndXp, "left", $ccclick)
	Sleep(500)
	_Cpp_ctrlV()
	Sleep(500)
	$cctext = ControlGetText($hWndXp, "", "TcxDBTextEdit4")
	Local $a = 1
	Local $ccstatus = False
	While $a < 5
		if StringLen($cctext) > 10 Then
			$ccstatus = True
			ExitLoop
		EndIf
		Sleep(500)
	    $a += 1
	WEnd

    If $ccstatus Then
		 ControlClick($hWndXp, "left", $ccadd)
	Else
		Send("{ENTER}") ;for close not found CC text popup
		Sleep(700)
		ControlClick($hWndXp, "left", $ccclick)
		ClipPut($cctextadd)
		Sleep(100)
		_Cpp_ctrlA()
	    Sleep(700)
        _Cpp_ctrlV()
	    Sleep(700)
		Send("{ENTER}")
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
	 if  FindToClick(@ScriptDir&$sRec_allergy_iden, $aRecAllergyPos[0], $aRecAllergyPos[1], $aRecAllergyPos[2], $aRecAllergyPos[3], 0.75) Then
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

Func AddItem($hStartTime, $oLogStudent, $iRow)
	Local $hWndItem  ;= WinWait("HOSxPDentalOperationEntryForm", "", 60)
	$hWndItem = WinGetHandle($sOperationBoxTitle)
;~ 	While 1
;~ 		Sleep(200)
;~ 		$hWndItem = WinGetHandle("HOSxPDentalOperationEntryForm")
;~ 		Sleep(200)
;~ 		if $hWndItem Then
;~ 			 Sleep(1800)
;~ 			ExitLoop
;~ 		EndIf
;~ 		ExitMaxTime($hStartTime, $oLogStudent, $iRow)
;~ 	WEnd
	 ClipPut("priex")
	 Sleep(300)
	_Cpp_ctrlV()
	While True
		Sleep(900)
		if Findtocon(@ScriptDir&$sPratomItem, $aItemTxtPos[0],  $aItemTxtPos[1],  $aItemTxtPos[2],  $aItemTxtPos[3], 0.75)  Then
			ExitLoop
		Else
			ControlClick($hWndItem, "left", "TcxCustomInnerTextEdit7")
			Sleep(400)
			ClipPut("priex")
            Sleep(700)
		    _Cpp_ctrlA()
			Sleep(400)
			Send("{DELETE}")
	        Sleep(700)
            _Cpp_ctrlV()
			Sleep(1000)
		EndIf
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	WEnd
	;ControlSend($hWndItem, "","" , "{ENTER}")
	Send("{ENTER}")
	Sleep(2000)
	ClickSaveOperation($hWndItem)
	Sleep(500)
	Return True
EndFunc


Func AddItemFluoride($hStartTime, $oLogStudent, $iRow)
	Local $hWndItem  ;= WinWait("HOSxPDentalOperationEntryForm", "", 60)
	$hWndItem = WinGetHandle($sOperationBoxTitle)

	 ClipPut("fludex")
	 Sleep(300)
	_Cpp_ctrlV()
	While True
		Sleep(1000)
		if Findtocon(@ScriptDir&$sFluorideItem, $aItemTxtPos[0],  $aItemTxtPos[1],  $aItemTxtPos[2],  $aItemTxtPos[3], 0.75)  Then
			ExitLoop
		Else
			ControlClick($hWndItem, "left", "TcxCustomInnerTextEdit7")
			 Sleep(400)
			ClipPut("fludex")
            Sleep(700)
		    _Cpp_ctrlA()
			Sleep(400)
			Send("{DELETE}")
	        Sleep(700)
            _Cpp_ctrlV()
			Sleep(1000)
		EndIf
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	WEnd
	;ControlSend($hWndItem, "","" , "{ENTER}")
	Send("{ENTER}")
	Sleep(2000)
	ClickSaveOperation($hWndItem)
	Sleep(500)
	;FindToClick(@ScriptDir&$sSaveItem, $aSaveItemBtnPos[0], $aSaveItemBtnPos[1], $aSaveItemBtnPos[2], $aSaveItemBtnPos[3], 0.75)
	;WinClose($hWndItem, "")
	Return True
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




;=============================== start sit function =============================================================================
Func CheckSit($sText)
    Local $chk  = False
	For $i = 0 To UBound($ArraySit) - 1
         If $ArraySit[$i] = $sText Then
			 $chk = True
			 ExitLoop
		 EndIf
    Next
	;If $chk Then MsgBox($MB_SYSTEMMODAL, "", "The text in Edit1 is: " & $sText)
	Return $chk
EndFunc

Func CheckSitKrg($sText)
    Local $chk  = False
	For $i = 0 To UBound($ArraySitKrg) - 1
         If $ArraySitKrg[$i] = $sText Then
			 $chk = True
			 ExitLoop
		 EndIf
    Next
	;If $chk Then MsgBox($MB_SYSTEMMODAL, "", "The text in Edit1 is: " & $sText)
	Return $chk
EndFunc

Func SetZeroPrice($hWndXp, $hStartTime, $oLogStudent, $oLogStudentProgress, $iRow)
	While True
		if FindToCon(@ScriptDir&$sOpdPrice1, $aOpdPricePos[0], $aOpdPricePos[1], $aOpdPricePos[2], $aOpdPricePos[3], 0.75) Then
			FindToClickRt(@ScriptDir&$sOpdPrice1, $aOpdPricePos[0], $aOpdPricePos[1], $aOpdPricePos[2], $aOpdPricePos[3], 0.75)
			Sleep(1200)
		Else
			FindToClickRt(@ScriptDir&$sOpdPrice2, $aOpdPricePos[0], $aOpdPricePos[1], $aOpdPricePos[2], $aOpdPricePos[3], 0.75)
			Sleep(1200)
		EndIf
        Sleep(500)
		If FindToCon(@ScriptDir&$sEditPrice, $aEditPricePos[0], $aEditPricePos[1], $aEditPricePos[2], $aEditPricePos[3], 0.75) Then ExitLoop ;Text Inv Setting
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	 WEnd
	 Sleep(1000)
	 ControlSend($hWndXp, "","" , "q")
	 Sleep(1000)
	 Local  $hWndEdit
	 Local $a = 1
	 While $a < 10
		 Sleep(1000)
		 $hWndEdit = WinGetHandle("HOSxPMedicationOrderItemPriceEditForm")
		 If $hWndEdit Then
			 Sleep(1000)
			 ExitLoop
		 Else
			 ;Sleep(1000)
			 ControlSend($hWndXp, "","" , "q")
			 Sleep(2000)
		 EndIf
		 $a += 1
	 WEnd
	ControlSend($hWndEdit, "","TcxCustomInnerTextEdit2","{DELETE}")
	 Sleep(1000)
	 ControlSend($hWndEdit, "","TcxCustomInnerTextEdit2", "0")
	 Sleep(1200)
		;ControlSend($hWndEdit, "","TcxCustomInnerTextEdit2","{ENTER}")
	 Send("{ENTER}")
	 WinClose($hWndEdit, "")
	 Sleep(400)
     FileWrite($oLogStudentProgress, "Set 0 Price Success R= "&$iRow&@CRLF)
	 Local $b = 1
	 While $b < 5
		Sleep(300)
		$hWndEdit = WinGetHandle("HOSxPMedicationOrderItemPriceEditForm")
		if($hWndEdit) Then
			Sleep(1000)
			ControlClick($hWndEdit, "left", "TcxButton1")
			Sleep(1200)
			WinClose($hWndEdit, "")
			Sleep(400)
		EndIf
		$b += 1
    WEnd
	Return True
EndFunc
;===============================  end sit function =============================================================================


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
	WinActivate($hWnd)
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
		if FindToClick(@ScriptDir&$sFinalSave2, $aFinalS2Pos[0], $aFinalS2Pos[1], $aFinalS2Pos[2], $aFinalS2Pos[3], 0.75)  Then  ExitLoop
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

        FindToClick(@ScriptDir&$sLastOK, $aLastOkPos[0], $aLastOkPos[1], $aLastOkPos[2], $aLastOkPos[3], 0.75)
		Sleep(200)
		if Findtocon(@ScriptDir&$sFoundVisit, $aFvPos[0], $aFvPos[1], $aFvPos[2], $aFvPos[3], 0.75)  Then ExitLoop  ;รูปคนที่ตำแหน่งค่อนมาตรงกลาง
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
		if Findtocon(@ScriptDir&$sFinalSave2, $aFinalS2Pos[0], $aFinalS2Pos[1], $aFinalS2Pos[2], $aFinalS2Pos[3], 0.75)  Then ExitLoop
		ExitMaxTime($hStartTime,$oLogStudent,$iRow)
	WEnd

    Sleep(300)
    Select021($hStartTime, $oLogStudent, $iRow)
	;ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	FindToClick(@ScriptDir&$sFinalSave2, $aFinalS2Pos[0], $aFinalS2Pos[1], $aFinalS2Pos[2], $aFinalS2Pos[3], 0.75)
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

        FindToClick(@ScriptDir&$sLastOK, $aLastOkPos[0], $aLastOkPos[1], $aLastOkPos[2], $aLastOkPos[3], 0.75)
		Sleep(200)
		if Findtocon(@ScriptDir&$sFoundVisit, $aFvPos[0], $aFvPos[1], $aFvPos[2], $aFvPos[3], 0.75)  Then ExitLoop  ;รูปคนที่ตำแหน่งค่อนมาตรงกลาง
		ExitMaxTime($hStartTime,$oLogStudent,$iRow)
	WEnd
	Return True
EndFunc
;===============================  end save function ============================================================================================================



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



Func ClickAddOperation($hWndXp, $hStartTime, $oLogStudent, $iRow)
    Local $sFormTitle = $sOperationBoxTitle
    Local $sButton = "[CLASS:TcxButton; INSTANCE:31]"

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




;==================================== start bot loop ===========================================================================================================
Func BotLoop()
    $oExcel = _Excel_Open()
	Local $bReadOnly = True
	Local $bVisible = True
	$oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook, $bReadOnly, $bVisible)
	Local $aResult = _Excel_RangeRead($oWorkbook, $sSheet)

    If IsArray($aResult) Then
		$oLogStudent = FileOpen($logStudent, $FO_APPEND)
		$oLogStudentProgress = FileOpen($logStudenProgress, $FO_APPEND)
		$oLogPrice =  FileOpen($logPrice, $FO_APPEND) ;
		Sleep(1200)
		HideServer()
		HideExcel()
		HideSciTE()
		Sleep(1000)
		SwapXP()
		$hWndXp = WinWait($sHosXp_Title, "", 10)
		SetHnBox($hWndXp)
		Sleep(1000)
		AdlibRegister("ContXp1",10000)
		Local $hPtNote
	    Local $hDrgAllergy
	    Local $hAppoint
	    Local $hTodayAppoint
	    Local $hDentistLock
		Local $hFinLocked ;not use
	    Local $bTodayAppoint = False
	    Local $bDentistLock = False
		Local $bFinanceLock = False
	    Local $bSitKrg = False
		Local $sBw = "0"
		Local $sQueue = "0"
		Local $iCnt = 0
		Local $sSit = ""
   For $i = $iStartRow - 1 To $iEndRow - 1  ;UBound($aResult, 1) - 1
	  Sleep(500)
	  Local $hnClass  = "TcxCustomInnerTextEdit2"
	  Local $iHnVal = $aResult[$i][4]
	  FileWrite($oLogStudentProgress, "Start Get dbData R= "&$i+1&@CRLF)
	  Sleep(100)
	  Local $dbData = GetSitFromDb2($iHnVal, $oLogStudent, $i+1)
	  Sleep(100)
		  $sSit = $dbData[0]
		  $sBw = $dbData[1]
	      $sQueue = $dbData[2]
	  Sleep(100)
	  FileWrite($oLogStudentProgress, "End Get dbData R= "&$i+1&@CRLF)
	  ContXp2()

	  if Number($sQueue) < 1 Then
          FileWrite($oLogStudent, "Error Visit1 R= "&$i+1&", ")
		  SendTeleGram("Error Visit1 R= "&$i+1)
		  Sleep(200)
		  ContinueLoop
	  EndIf

	  Local $hQCtrl = ControlGetHandle($hWndXp, "", "[CLASS:TcxTextEdit; INSTANCE:1]")
	  ControlSetText($hWndXp, "", $hQCtrl, $sQueue)
	  Sleep(500)
	  ControlClick($hWndXp, "", $hQCtrl, "left")
	  Sleep(700)
	  ControlSend($hWndXp, "", $hQCtrl, "{ENTER}")
;~ 	  Send("{ENTER}")
;~ 	  ControlClick($hWndXp, "left", $hnClass)
;~ 	  Sleep(1000)
;~ 	  ContXp2()
;~ 	  CtrlSend($hWndXp,$hnClass,$iHnVal)
;~ 	  Sleep(1000)
;~ 	  ContXp2()
;~       Send("{ENTER}")
	  Local $hStartTime = TimerInit() ;เริ่มจับเวลาบันทึกข้อมูล
	  Sleep(500)
	  FileWrite($oLogStudentProgress, "Start Record R= "&$i+1&@CRLF)
	  ContXp2()
	  Local $a = 1
	  While $a < 10
		  Sleep(500)
		  $hDrgAllergy = WinGetHandle("[CLASS:THOSxPMedicationOrderDrugAllergyNoticeForm]")
		  Sleep(100)
		  if $hDrgAllergy Then WinKill($hDrgAllergy)
		  Sleep(300)

		  $hAppoint = WinGetHandle("HOSxPAppointmentInformationForm")
		  Sleep(100)
		  if $hAppoint Then WinKill($hAppoint)
		  Sleep(300)

		  $hTodayAppoint = WinGetHandle("HOSxPAppointmentVisitConfirmForm")
		  Sleep(100)
		  If $hTodayAppoint Then
			  $bTodayAppoint = True
		      WinKill($hTodayAppoint)
		  EndIf
		  Sleep(300)

		  $hDentistLock = WinGetHandle("Visit Number Locked")
		  Sleep(100)
		  If $hDentistLock Then
			  $bDentistLock = True
		      WinKill($hDentistLock)
		  EndIf
		  Sleep(300)

		 If Findtocon(@ScriptDir&$sFoundVisit, $aFvFirstPos[0], $aFvFirstPos[1], $aFvFirstPos[2], $aFvFirstPos[3], 0.8)  Then ;รูปคนที่อยู่ตำแหน่งซ้ายสุดของจอ
				Sleep(900)
				ExitLoop
		 EndIf
		  $a = $a+1
	  WEnd

	   If $a = 10 Then
		    FileWrite($oLogStudent, "Error Visit2 R= "&$i+1&", ")
			SendTeleGram("Error Visit2 R= "&$i+1)
			Sleep(200)
		    ContinueLoop  ;skip this data when no visit
	   EndIf
	  FileWrite($oLogStudentProgress, "Exit First PopUp Loop R= "&$i+1&@CRLF)
	  Sleep(50)
	  FileWrite($oLogStudentProgress, "Start Init Loading Loop R= "&$i+1&@CRLF)
	  ;== start hos xp loading loop ===================================
	   While True
		    Sleep(1000)
			$hPtNote = WinGetHandle("PatientNoteViewDisplayForm")
		    Sleep(100)
		    if $hPtNote Then WinKill($hPtNote)
			Sleep(100)
		    ;If FindToCon(@ScriptDir&$sFiLock, $aFiLockPos[0], $aFiLockPos[1], $aFiLockPos[2], $aFiLockPos[3], 0.75) Then
			if KillFinance2() Then
				$bFinanceLock = True
                FinanceLockExit($hWndXp, $hStartTime, $oLogStudent, $i+1)
				FileWrite($oLogStudent, "Error Finance R= "&$i+1&", ")
			    SendTeleGram("Error Finance R= "&$i+1)
				Sleep(200)
				ExitLoop
			ElseIf $bTodayAppoint Then
					if $bFinanceLock Then
						FinanceLockExit($hWndXp, $hStartTime, $oLogStudent, $i+1)
						FileWrite($oLogStudent, "Error Finance Today R= "&$i+1&", ")
						SendTeleGram("Error Finance Today R= "&$i+1)
						Sleep(200)
						ExitLoop
					Else
						PtExit($hWndXp, $hStartTime, $oLogStudent, $i+1)
						FileWrite($oLogStudent, "Error Today Appoint R= "&$i+1&", ")
						SendTeleGram("Error Today Appoint R= "&$i+1)
						Sleep(200)
						ExitLoop
					EndIf
			ElseIf Findtocon(@ScriptDir&$sLoadPtSuccess2, $aLoadPtPos[0], $aLoadPtPos[1], $aLoadPtPos[2], $aLoadPtPos[3], 0.80)  Then ;diagram รูปฟันด้านล่าง
				Sleep($iSleepAfterLoad)
				ExitLoop
			Else
				Sleep(500)
			EndIf
			ExitMaxTime($hStartTime,$oLogStudent,$i+1)
	   WEnd
	   ;== end hos xp loading loop ===================================
	   FileWrite($oLogStudentProgress, "End Init Loading Loop R= "&$i+1&@CRLF)

;~ 	   If $bTodayAppoint Then
;~ 		   $bTodayAppoint = False
;~ 		   ContinueLoop
;~ 	   EndIf

	   If $bFinanceLock Then
		   $bFinanceLock = False
		   ContinueLoop
	   EndIf

	   If $bDentistLock Then
            Sleep(100)
			FileWrite($oLogStudent, "Error Locked R= "&$i+1&", ")
			SendTeleGram("Error Locked R= "&$i+1)
		    Sleep(300)
			$bDentistLock = False
		    ContinueLoop
	   EndIf
	   ContXp2()
	   ;Local $ttest = TimerDiff($hStartTime)
	   ;MsgBox(0,"time1", $ttest, 2)
		FileWrite($oLogStudentProgress, "Start Click Open Dental Loop R= "&$i+1&@CRLF)
		Local $clk_dt
		While True
			$clk_dt = ControlClick($hWndXp, "left", "TcxButton9")  ; open dental care UI
			Sleep(100)
			if $clk_dt = 1 Then
				Sleep(500)
				ExitLoop
			EndIf
			Sleep(400)
			ExitMaxTime($hStartTime,$oLogStudent,$i+1)
		WEnd
		FileWrite($oLogStudentProgress, "End Click Open Dental Loop R= "&$i+1&@CRLF)
		ContXp2()

        Local $aDental[14]
		$aDental[0]  = $aResult[$i][6]  ;$pteeth
		$aDental[1]  = $aResult[$i][7] ;$pcaries
		$aDental[2]  = $aResult[$i][10] ;$pfilling
		$aDental[3]  = $aResult[$i][11] ;$pextract

		$aDental[4]  = $aResult[$i][8] ;$dteeth
		$aDental[5]  = $aResult[$i][9] ;$dcaries
		$aDental[6]  = $aResult[$i][13] ;$dfilling
		$aDental[7]  = $aResult[$i][14] ;$dextract

		$aDental[8]  = $aResult[$i][16]  ;$need_sealant
		$aDental[9]  = $aResult[$i][12] ;$need_pfilling
		$aDental[10]  = $aResult[$i][15] ;$need_dfilling
		$aDental[11]  = $aResult[$i][18] ;$need_dextract
		$aDental[12]  = $aResult[$i][17] ;$need_pextract
		$aDental[13] = $aResult[$i][1] ;class

		FileWrite($oLogStudentProgress, "Start SendDental Func R= "&$i+1&@CRLF)
		Sleep(100)
		SendDental($aDental, $hStartTime, $oLogStudent, $i+1)
		FileWrite($oLogStudentProgress, "End SendDental Func R= "&$i+1&@CRLF)
		ContXp2()
		;Send("{F1}") hot key F1 not work in sometimes

;~ 		FileWrite($oLogStudentProgress, "Start FindToClick F1 Menu R= "&$i+1&@CRLF)
;~ 		While 1
;~ 			Sleep(700)
;~ 			if FindToClick(@ScriptDir&$sF1, $aF1Pos[0], $aF1Pos[1], $aF1Pos[2], $aF1Pos[3], 0.80) Then ExitLoop  ;click F1 menu and go to add CC
;~ 			ExitMaxTime($hStartTime,$oLogStudent,$i+1)
;~ 		WEnd
;~         FileWrite($oLogStudentProgress, "End FindToClick F1 Menu R= "&$i+1&@CRLF)

		FileWrite($oLogStudentProgress, "Start Send F1 R= " & ($i + 1) & @CRLF)
		Click_F1($hWndXp, $hStartTime, $oLogStudent, $i + 1)
        FileWrite($oLogStudentProgress, "End Send F1 R= " & ($i + 1) & @CRLF)

        ContXp2()
		FileWrite($oLogStudentProgress, "Start FindToCon CC Pic R= "&$i+1&@CRLF)
		While 1
			Sleep(300)
			if FindToCon(@ScriptDir&$sCC, $aCcPos[0], $aCcPos[1], $aCcPos[2], $aCcPos[3], 0.80) Then ExitLoop  ;CC pic
			ExitMaxTime($hStartTime,$oLogStudent,$i+1)
		WEnd
        FileWrite($oLogStudentProgress, "End FindToCon CC Pic R= "&$i+1&@CRLF)
		Sleep(750)
		ContXp2()
		FileWrite($oLogStudentProgress, "Start Record CC And Allergy R= "&$i+1&@CRLF)
		ChiefComp($aDental[13] ,$hStartTime, $oLogStudent, $hWndXp, $i+1)
        Allergy()
		if Number($sBw) > 0 Then
			;CtrlSendDt($hWndXp, "TcxCustomInnerTextEdit59", $sBw)
			Local $hBwCtrl = ControlGetHandle($hWndXp, "", "[CLASS:TcxDBTextEdit; INSTANCE:18]")  ;TcxDBTextEdit17
			ControlSetText($hWndXp, "", $hBwCtrl, $sBw)
			Sleep(300)
		EndIf
		FileWrite($oLogStudentProgress, "End Record CC And Allergy R= "&$i+1&@CRLF)
		ContXp2()

		;Send("{F4}") hot key F4 not work in sometimes
		FileWrite($oLogStudentProgress, "Start FindToClick F4 Menu R= "&$i+1&@CRLF)
		Click_F4($hWndXp, $hStartTime, $oLogStudent, $i + 1)   ;send F4 shortcut key
		While 1
			Sleep(200)
			Local $hDialogBug = WinGetHandle("FindAllItemDialog")
			; ถ้าไม่พบหน้าต่าง ให้ออกจาก loop
			If Not $hDialogBug Then ExitLoop
			; ถ้าพบหน้าต่าง ให้ปิด
			WinKill($hDialogBug)
			Sleep(200)
			ExitMaxTime($hStartTime, $oLogStudent, $i + 1)
		WEnd

        ClickAddOperation($hWndXp, $hStartTime, $oLogStudent, $i+1)

        FileWrite($oLogStudentProgress, "End Click AddItem PlusBtn Success R= "&$i+1&@CRLF)

       ContXp2()
	   FileWrite($oLogStudentProgress, "Start AddItem Func R= "&$i+1&@CRLF)
       AddItem($hStartTime,$oLogStudent,$i+1)
	   FileWrite($oLogStudentProgress, "End AddItem Func Then Find Task Pic R= "&$i+1&@CRLF)

	    While True
			Sleep(500)
			Local $hOperationForm = WinGetHandle($sOperationBoxTitle)
			If $hOperationForm = 0 Then ExitLoop
			If Not _WinAPI_IsWindowVisible($hOperationForm) Then ExitLoop
			Sleep(700)
			ClickSaveOperation($hOperationForm)
			;FindToClick(@ScriptDir & $sSaveItem, $aSaveItemBtnPos[0], $aSaveItemBtnPos[1], $aSaveItemBtnPos[2], $aSaveItemBtnPos[3], 0.75)
			ExitMaxTime($hStartTime, $oLogStudent, $i + 1)
	    WEnd

        FileWrite($oLogStudentProgress, "AddItem Saved Then Close Add Item Box R= "&$i+1&@CRLF)
		ContXp2()

		If $iAddFluoride = 1 Then
			FileWrite($oLogStudentProgress, "Start Click AddItemFluoride  PlusBtn R= "&$i+1&@CRLF)
			ClickAddOperation($hWndXp, $hStartTime, $oLogStudent, $i+1)
			Sleep(1000)

			FileWrite($oLogStudentProgress, "End Click AddItemFluoride PlusBtn R= "&$i+1&@CRLF)
			Sleep(50)
			FileWrite($oLogStudentProgress, "Start AddItemFluoride Func R= "&$i+1&@CRLF)
			AddItemFluoride($hStartTime, $oLogStudent, $i+1)
			While True
				Sleep(500)
				Local $hOperationForm = WinGetHandle($sOperationBoxTitle)
				If $hOperationForm = 0 Then ExitLoop
				If Not _WinAPI_IsWindowVisible($hOperationForm) Then ExitLoop
				Sleep(700)
				ClickSaveOperation($hOperationForm)
				;FindToClick(@ScriptDir & $sSaveItem, $aSaveItemBtnPos[0], $aSaveItemBtnPos[1], $aSaveItemBtnPos[2], $aSaveItemBtnPos[3], 0.75)
				ExitMaxTime($hStartTime, $oLogStudent, $i + 1)
			WEnd
			FileWrite($oLogStudentProgress, "AddItemFluoride Saved Then Close Add Item Box R= "&$i+1&@CRLF)
		EndIf
        ;MsgBox(0,"show", $sSit, 2)
        If CheckSit($sSit) Then
			FileWrite($oLogStudentProgress, "Start SetZeroPrice Func R= "&$i+1&@CRLF)
			SetZeroPrice($hWndXp,$hStartTime,$oLogStudent, $oLogStudentProgress, $i+1)
			FileWrite($oLogPrice, "Set Zero Price R= "&$i+1&", ")
			Sleep(300)
			FileWrite($oLogStudentProgress, "End SetZeroPrice Func R= "&$i+1&@CRLF)
	    ElseIf  CheckSitKrg($sSit) Then
			$bSitKrg = True
			FileWrite($oLogPrice, "Send KRG R= "&$i+1&", ")
			Sleep(300)
		EndIf
	    ;if $sSit = "47o" Or $sSit = "25o"  Then SendTeleGram("OP HN= "&$iHnVal)
		If StringInStr("47o,25o", $sSit) Then SendTeleGram("OP HN= "&$iHnVal)
		FileWrite($oLogStudentProgress, "End CheckSit Condition R= "&$i+1&@CRLF)
        ExitMaxTime($hStartTime,$oLogStudent,$i+1)
        ;Sleep(500)
        ContXp2()
		FileWrite($oLogStudentProgress, "Start FinalSave1 Func R= "&$i+1&@CRLF)
		FinalSave1($hStartTime, $oLogStudent, $hWndXp, $i+1)
        Sleep(500)
		FileWrite($oLogStudentProgress, "End FinalSave1 Func R= "&$i+1&@CRLF)
		ContXp2()

		If $bSitKrg Then
			FileWrite($oLogStudentProgress, "Start FinalSave2Krg Func R= "&$i+1&@CRLF)
			FinalSave2Krg($hStartTime,$oLogStudent,$i+1)
			FileWrite($oLogStudentProgress, "End FinalSave2Krg Func R= "&$i+1&@CRLF)
		Else
			FileWrite($oLogStudentProgress, "Start FinalSave2 Func R= "&$i+1&@CRLF)
			FinalSave2($hStartTime,$oLogStudent,$i+1)
			FileWrite($oLogStudentProgress, "End FinalSave2 Func R= "&$i+1&@CRLF)
		EndIf

		$sSit = ""
		$sBw = "0"
		$sQueue = "0"
		$bTodayAppoint = False
		$bFinanceLock = False
		$bDentistLock = False
		$bSitKrg = False
		$iCnt += 1
		If ($iCnt = 5) Then
			SendTeleGram("Count= 5 R= "&$i+1)
			$iCnt = 0
		EndIf
		Sleep(200)
		FileWrite($oLogStudentProgress, "End Record R= "&$i+1&@CRLF)
		Sleep(300)
  Next
	  AdlibUnRegister("ContXp1")
	  FileClose($oLogStudent)
	  FileClose($oLogStudentProgress)
	  FileClose($oLogPrice)
	  ShowExcel()
	  _Excel_Close($oExcel)
	  ShowSciTE()
	  SendTeleGram("Success")
	  MsgBox(0,"Success","Success")
  Else
	  MsgBox(0,"Error","ไม่พบข้อมูลไฟล์ excel")
  EndIf
EndFunc
;==================================== end bot loop ==========================================================================

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
        Case $botButton
			$sWorkbook = GUICtrlRead($Input1)
			$sSheet = GUICtrlRead($Input2)
			$sSchool = GUICtrlRead($Input3)
			$iSchoolType = Number(GUICtrlRead($Input4))
            $iStartRow = Number(GUICtrlRead($Input5))
            $iEndRow = Number(GUICtrlRead($Input6))
			$iSleepAfterLoad = Number(GUICtrlRead($Input7))
			$iAddFluoride = Number(GUICtrlRead($Input8))
            ; Reset and show the alert label if errors exist
            GUICtrlSetState($AlertLabel, $GUI_HIDE)

			 if  StringLen($sWorkbook) < 9 Then
				 GUICtrlSetData($AlertLabel, "โปรดบันทึก Excel file path ให้ถูกต้อง!")
                 GUICtrlSetState($AlertLabel, $GUI_SHOW)
			 ElseIf StringLen($sSheet) < 1 Then
                GUICtrlSetData($AlertLabel, "โปรดบันทึก Sheet name ให้ถูกต้อง!")
                GUICtrlSetState($AlertLabel, $GUI_SHOW)
			 ElseIf StringLen($sSchool) < 3 Then
                GUICtrlSetData($AlertLabel, "โปรดบันทึก CC ให้ถูกต้อง!")
                GUICtrlSetState($AlertLabel, $GUI_SHOW)
			 ElseIf  $iSchoolType < 1 Or $iSchoolType > 2 Or Not IsInt($iSchoolType) Then
				GUICtrlSetData($AlertLabel, "โปรดบันทึก School Type ให้ถูกต้อง!")
                GUICtrlSetState($AlertLabel, $GUI_SHOW)
			ElseIf  $iAddFluoride < 1 Or $iAddFluoride > 2 Or Not IsInt($iAddFluoride) Then
				GUICtrlSetData($AlertLabel, "โปรดบันทึก Fluoride Type ให้ถูกต้อง!")
                GUICtrlSetState($AlertLabel, $GUI_SHOW)
			 ElseIf $iStartRow < 2 Or $iEndRow < $iStartRow Or Not IsInt($iStartRow) Or Not IsInt($iEndRow) Then
                GUICtrlSetData($AlertLabel, "โปรดเลือกช่วงแถวข้อมูลให้ถูกต้อง!")
                GUICtrlSetState($AlertLabel, $GUI_SHOW)
			 ElseIf $iSleepAfterLoad < 1500 Or $iSleepAfterLoad > 10000 Or Not IsInt($iSleepAfterLoad) Then
			    GUICtrlSetData($AlertLabel, "โปรดบันทึกค่า Delay ระว่าง 1500 ถึง 10000")
                GUICtrlSetState($AlertLabel, $GUI_SHOW)
             Else
                FileDelete($sConfigFile) ; Ensure clean write
                FileWrite($sConfigFile, $sWorkbook &@CRLF& $sSheet &@CRLF& $sSchool &@CRLF& $iSchoolType &@CRLF& $iStartRow &@CRLF& $iEndRow &@CRLF& $iSleepAfterLoad &@CRLF&$iAddFluoride)
                ; Properly exit the GUI loop and call TestLoop()
                GUIDelete($Form1)
                ExitLoop
            EndIf
	EndSwitch
WEnd

_OpenCV_Startup()  ;
_CppDllOpen()
$gTelegramToken = GetEnv("TELEGRAM_TOKEN")
$gTelegramChatID = GetEnv("TELEGRAM_CHAT")
StartServer()
Sleep(1000)
BotLoop()
_OpenCV_Shutdown()
_CppDllClose()
MsgStopServer()
Exit(0)