#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Match\back_gNW_icon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <Excel.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#Include <WinAPI.au3>
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
#include <Inet.au3>
#include <CppKeySend.au3>

Opt("MouseCoordMode", 1)
Opt("WinTitleMatchMode", 2)
HotKeySet("{ESC}", "Hkey")

;=============================== start GUI config ========================================================================
Global $sConfigFile = @ScriptDir &"\config\elderly_config.txt"
Global $iDefaultValue1 = "D:\schooldatasheet\sum2567\Dental_LR_2567.xlsx"
Global $iDefaultValue2 = "data"
Global $iDefaultValue3 =  "ออกหน่วยตรวจฟันผู้สูงอายุ"
Global $iDefaultValue4 = 1
Global $iDefaultValue5 = 2
Global $iDefaultValue6 = 27
Global $iDefaultValue7 = 2000

If FileExists($sConfigFile) Then
    Local $aConfig = FileReadToArray($sConfigFile)
	;_ArrayDisplay($aConfig)
    If IsArray($aConfig) Then
        If UBound($aConfig) >= 7 Then
            $iDefaultValue1 = $aConfig[0]
            $iDefaultValue2 = $aConfig[1]
			$iDefaultValue3 = $aConfig[2]
			$iDefaultValue4 = Number($aConfig[3])
			$iDefaultValue5 = Number($aConfig[4])
			$iDefaultValue6 = Number($aConfig[5])
			$iDefaultValue7 = Number($aConfig[6])
        EndIf
    EndIf
EndIf

#Region ### START Koda GUI section ### Form=
Global $Form1 = GUICreate("Elderly Bot", 634, 502, 192, 124)
Global $Input1 = GUICtrlCreateInput($iDefaultValue1, 24, 32, 577, 21)
Global $Input2 = GUICtrlCreateInput($iDefaultValue2, 24, 88, 289, 21)
Global $Input3 = GUICtrlCreateInput($iDefaultValue3, 24, 144, 577, 21)
$Label1 = GUICtrlCreateLabel("Excel file path", 24, 8, 70, 17)
$Label2 = GUICtrlCreateLabel("Sheet name", 24, 64, 61, 17)
$Label3 = GUICtrlCreateLabel("CC", 24, 120, 18, 17)
Global $Input4 = GUICtrlCreateInput($iDefaultValue4, 24, 216, 113, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
Global $Input5 = GUICtrlCreateInput($iDefaultValue5, 24, 280, 113, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
Global $Input6 = GUICtrlCreateInput($iDefaultValue6, 184, 280, 113, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
Global $Input7 = GUICtrlCreateInput($iDefaultValue7, 24, 344, 113, 21, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER))
$Label4 = GUICtrlCreateLabel("School Type 1 = รัฐบาล 2 = เอกชน", 24, 192, 168, 17)
$Label5 = GUICtrlCreateLabel("Start Row", 56, 256, 51, 17)
$Label6 = GUICtrlCreateLabel("End Row", 216, 256, 48, 17)
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
;=============================== end GUI config ========================================================================

Global $iPos1 = -8,  $iPos2 = -8,  $iSize1 = 1552, $iSize2 = 840
Global $sHosXp_Title = "BMS-HOSxP XE 4.0"
Global $sGoogle_Title = 'Google Chrome'
Global $logStudent = @ScriptDir &"\BotLog\log_elderly.txt"
Global $logStudenProgress = @ScriptDir &"\BotLog\log_elderly_progress.txt"
Global $logPrice = @ScriptDir &"\BotLog\log_elderly_price.txt"

Global $hWndXp
Global $oExcel
Global $oLogStudent
Global $oLogStudentProgress
Global $oLogPrice
Global $oWorkbook

Global $iMaxRecTime = 4*60*1000 ; 4 min
;Global $sDtMenu = "\Match\xp_dental_menu.png"
Global $sFoundVisit = "\Match\xp_hn_found_visit.png"
Global $sFiLock =  "\Match\xp_finance_lock.png"
;Global $sClosePt = "\Match\xp_close_pt.png"
Global $sLoadPtSuccess2 = "\Match\xp_load_pt_success2.png"
;Global $sDentalCare =  "\Match\xp_dental_care.png"
Global $sPratomItem =   "\Match\xp_item_found_elder.png"
Global $sOpdPrice1 =   "\Match\opd_price1.png"
Global $sOpdPrice2 =   "\Match\opd_price2.png"
Global $sEditPrice =   "\Match\xp_edit_price.png"
Global $sTask =   "\Match\xp_task_on_item.png"
;Global $sFinalSave =   "\Match\xp_final_save1.png"
Global $sFinalSave2 =   "\Match\xp_final_save2.png"
Global $sAddItem =   "\Match\xp_add_item.png"
Global $sSaveItem =   "\Match\xp_item_save.png"
Global $sF1 =   "\Match\xp_f1_click.png"
Global $sF4 =   "\Match\xp_f4_click.png"
Global $sCC =   "\Match\xp_cc.png"
;Global $sAddCC =   "\Match\xp_add_cc.png"
Global $sLastOK =   "\Match\xp_lastOK.png"
Global $sRec_allergy_iden =   "\Match\rec_allergy_iden.png"

Global $ArraySit[12]
$ArraySit[0] = "40 เบิกได้"
$ArraySit[1] = "10 ชำระเงิน"
$ArraySit[2] = "14 ชำระเงินต่างชาติ กลุ่ม 1"
$ArraySit[3] = "15 ชำระเงินต่างชาติ กลุ่ม 2"
$ArraySit[4] = "16 ชำระเงินต่างชาติ กลุ่ม 3"
$ArraySit[5] = "42 ประกันสังคม ในเขต"
$ArraySit[6] = "60 ประกันสังคม นอกเขต"
$ArraySit[7] = "49 ไม่ได้นำหลักฐานมาด้วย"
$ArraySit[8] = "54 สิทธิ์ว่าง"
$ArraySit[9] = "25 นักเรียน นอกเขต"
$ArraySit[10] = "47 เด็ก 0-12 ปี นอกเขต"
$ArraySit[11] = "38 กลุ่มคืนสิทธิ นอกเขต"

Global $ArraySitKrg[4]
$ArraySitKrg[0] = "01 ขรก. (องค์กรปกครองส่วนท้องถิ่น)"
$ArraySitKrg[1] = "55 ขรก.กรมบัญชีกลาง"
$ArraySitKrg[2] = "34 ขรก.กรุงเทพมหานคร"
$ArraySitKrg[3] = "35 ขรก.กกต."

;=============================== start utility function ========================================================================
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
 Local Const $kScreen = 1.25 ;constant for screen size 125% = 1.25  for screen size 100% = 1
 Local Const $kPosition = 5  ;constant for screen size 125% position. Set $kPosition = 0 if screen size = 100%
 Local $Match = _MatchPicture("", $picPath, $x1*$kScreen, $y1*$kScreen, $x2*$kScreen, $y2*$kScreen, $tol)
 Local $realX = 0,  $realY = 0
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
 Local Const $kScreen = 1.25 ;constant for screen size 125% = 1.25  for screen size 100% = 1
 Local Const $kPosition = 5  ;constant for screen size 125% position. Set $kPosition = 0 if screen size = 100%
 Local $Match = _MatchPicture("", $picPath, $x1*$kScreen, $y1*$kScreen, $x2*$kScreen, $y2*$kScreen, $tol)
 Local $bClick = False
 Local $realX = 0,  $realY = 0
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
 Local Const $kScreen = 1.25 ;constant for screen size 125% = 1.25  for screen size 100% = 1
 Local Const $kPosition = 5  ;constant for screen size 125% position. Set $kPosition = 0 if screen size = 100%
 Local $Match = _MatchPicture("", $picPath, $x1*$kScreen, $y1*$kScreen, $x2*$kScreen, $y2*$kScreen, $tol)
 Local $realX = 0,  $realY = 0

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
 Local $kScreen = 1.25 ;constant for screen size 125% = 1.25  for screen size 100% = 1
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
	if $hWndXp Then
	   Local $hErrPopUp = WinGetHandle("[CLASS:madExceptWndClass]")  ;madExceptWndClass
	   if $hErrPopUp Then WinKill($hErrPopUp) ;ConsoleWrite("kill error occurred"&@CRLF)
	Else
		 SendTeleGram("Not Found Hos XP")
		 MsgBox($MB_SYSTEMMODAL, "Error", "ไม่พบโปรแกรม Hos XP", 5)
		 Exit(0)
	EndIf
EndFunc

Func ContXp2()
	Local $hErrPopUp
	 While 1
		    Sleep(100)
		    $hErrPopUp = WinGetHandle("[CLASS:madExceptWndClass]")
			Sleep(100)
			if $hErrPopUp Then
				WinKill($hErrPopUp)   ;ConsoleWrite("kill error occurred"&@CRLF)
	        Else
				ExitLoop
			EndIf
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

Func SendTeleGram($sMessage)
	Local $sToken = ""
	Local $sChatID = "-"
	Local $sURL = "https://api.telegram.org/bot" & $sToken & "/sendMessage?chat_id=" & $sChatID & "&text=" & $sMessage
	; Send the request
	Local $sResponse = InetGet($sURL, 1)
	;ConsoleWrite(BinaryToString($sResponse) & @CRLF)
EndFunc
;=============================== end utility function ========================================================================

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
					$sTitle = $aList[$i][0]
					$hHandle = $aList[$i][1]
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
   ;$hWndXp = WinWait($sHosXp_Title, "", 10)
	Local $aPos = ControlGetPos( $hWndXp, "", "TcxGroupBox1" )
	Local $sSideBarHn = ControlGetHandle ($hWndXp, "", "TcxGroupBox1" )
	;MsgBox(0,"",$aPos[1])
	if $aPos[0] = 0 And  $aPos[1]  = 125 And $aPos[2] = 400 And  $aPos[3] = 676 Then
		Sleep(100)
	Else
		WinMove($sSideBarHn, "", 0, 125, 400, 676)  ;400 is width size
	EndIf
	Sleep(1000)
EndFunc

Func Success()
	SendTeleGram("Success")
EndFunc
;=============================== end preload function ======================================================================

;=============================== start after load exit function =====================================================================
Func FinanceLockExit($hWndXp, $hStartTime, $oLogStudent, $iRow)
	 Sleep(200)
	 Send("{ENTER}")
    Local $hPtNote
	While True
		Sleep(1000)
		$hPtNote = WinGetHandle("PatientNoteViewDisplayForm")
		Sleep(200)
		if $hPtNote Then WinKill($hPtNote)
		Sleep(1000)
		If Findtocon(@ScriptDir&$sLoadPtSuccess2, 698, 691,778, 749,  0.75)  Then ExitLoop ;diagram รูปฟันด้านล่าง
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
		if Findtocon(@ScriptDir&$sFoundVisit, 389, 131,470, 199,  0.75)   Then ExitLoop
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
		If Findtocon(@ScriptDir&$sLoadPtSuccess2, 698, 691,778, 749,  0.80)  Then ExitLoop ;diagram รูปฟันด้านล่าง
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
		Sleep(500)
		$hFeeSch = WinGetHandle("OvstFeeScheduleEntryForm")
		Sleep(100)
		 if $hFeeSch Then WinKill($hFeeSch)
		Sleep(500)
		$hSpsh = WinGetHandle("HOSxPNHSOConfirmPrivilegeForm")
		Sleep(100)
		if $hSpsh Then WinKill($hSpsh)
		Sleep(400)
		if Findtocon(@ScriptDir&$sFoundVisit, 389, 131,470, 199,  0.75)   Then ExitLoop  ;รูปคนตำแหน่งกลางค่อนมาทางซ้าย
	 WEnd
	 Sleep(1000)
	 Return True
EndFunc
;=============================== end after load exit function ======================================================================

;=============================== start record dental function ======================================================================
Func SendDental($aArray, $hStartTime, $oLogStudent, $iRow)
	;Local $sTitle = "HOSxPDentalCareEntryForm"
	Local $hWnd = WinWait("HOSxPDentalCareEntryForm", "", $iMaxRecTime)
	While True
		    Sleep(1000)
			if $hWnd  Then  ExitLoop
			ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	WEnd
    ;WinActivate($hWnd)
	Sleep(1000)
	;_ArrayDisplay($aArray, "Row Data")
	Local $pteeth =  "TcxCustomInnerTextEdit22"
	Local $pcaries =  "TcxCustomInnerTextEdit21"
	Local $pfilling = "TcxCustomInnerTextEdit20"
	Local $pextract = "TcxCustomInnerTextEdit19"

	;Local $dteeth = "TcxCustomInnerTextEdit18"
    ;Local $dcaries = "TcxCustomInnerTextEdit17"
    ;Local $dfilling = "TcxCustomInnerTextEdit16"
    Local $dextract = "TcxCustomInnerTextEdit15"

	;Local $need_sealant = "TcxCustomInnerTextEdit14"
    Local $need_pfilling = "TcxCustomInnerTextEdit13"
    ;Local $need_dfilling = "TcxCustomInnerTextEdit12"
    ;Local $need_dextract = "TcxCustomInnerTextEdit11"
    Local $need_pextract = "TcxCustomInnerTextEdit6"

	Local $permanent_perma = "TcxCustomInnerTextEdit10"
	Local $permanent_pros = "TcxCustomInnerTextEdit9"
	Local $prosthesis_pros = "TcxCustomInnerTextEdit7"
	Local $need_scaling = "TcxDBCheckBox1"

    If Number($aArray[0]) > 0 Then CtrlSendDt($hWnd,$pteeth,$aArray[0])
	If Number($aArray[1]) > 0 Then	CtrlSendDt($hWnd,$pcaries,$aArray[1])
	If Number($aArray[2]) > 0 Then CtrlSendDt($hWnd,$pfilling,$aArray[2])
	If Number($aArray[3]) > 0 Then CtrlSendDt($hWnd,$pextract,$aArray[3])

	;If Number($aArray[4]) > 0 Then	CtrlSendDt($hWnd,$dteeth,$aArray[4])
	;If Number($aArray[5]) > 0 Then	CtrlSendDt($hWnd,$dcaries,$aArray[5])
	;If Number($aArray[6]) > 0 Then CtrlSendDt($hWnd,$dfilling,$aArray[6])
	CtrlSendDt($hWnd, $dextract, 20)

	;If Number($aArray[8]) > 0 Then CtrlSendDt($hWnd,$need_sealant,$aArray[8])
	If Number($aArray[4]) > 0 Then CtrlSendDt($hWnd,$need_pfilling,$aArray[4])
	;If Number($aArray[10]) > 0 Then CtrlSendDt($hWnd,$need_dfilling,$aArray[10])
	;If Number($aArray[11]) > 0 Then CtrlSendDt($hWnd,$need_dextract,$aArray[11])
	If Number($aArray[5]) > 0 Then CtrlSendDt($hWnd,$need_pextract,$aArray[5])

	If Number($aArray[6]) > 0 Then CtrlSendDt($hWnd,$permanent_perma,$aArray[6])
	If Number($aArray[7]) > 0 Then CtrlSendDt($hWnd,$permanent_pros,$aArray[7])
	If Number($aArray[8]) > 0 Then CtrlSendDt($hWnd,$prosthesis_pros,$aArray[8])
	If Number($aArray[10]) > 0 Then
		ControlClick($hWnd, "left", $need_scaling)
		Sleep(500)
		Send("{SPACE}") ; กด Spacebar เพื่อติ๊ก/ยกเลิก
		;ControlSend($hWnd, "", $need_scaling, "{SPACE}")
    EndIf
    SelectNeedPros($hWnd, $aArray[9])
	SelectPtType($hWnd)
	SelectPlace($hWnd)
	Local $clk
	While 1
		$clk = ControlClick($hWnd, "left","TcxButton3")
		Sleep(700)
		If $clk = 1 Then ExitLoop
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	WEnd
	WinClose($hWnd,"")
	Return True
EndFunc

Func SelectPtType($hWnd)
	;Local $sTitle = "HOSxPDentalCareEntryForm"
	;Local $hWnd = WinWait($sTitle, "", 10)
	Local $class_  = "TcxCustomComboBoxInnerEdit12"
	ControlClick($hWnd, "left", $class_)
	Sleep(200)
    Local $iClick = 4

   For $i = 1 To $iClick
		ControlSend($hWnd, "",$class_ , "{DOWN}")
		Sleep(200)
	Next
EndFunc

Func SelectPlace($hWnd)
	;Local $sTitle = "HOSxPDentalCareEntryForm"
	;Local $hWnd = WinWait($sTitle, "", 10)
	Local $place  = "TcxCustomComboBoxInnerEdit11"
	ControlClick($hWnd, "left", $place)
	Sleep(200)
	For $i = 1 To 2
		ControlSend($hWnd, "",$place , "{DOWN}")
		Sleep(200)
	Next
EndFunc

Func SelectNeedPros($hWnd, $sNeedProsthesis)
	;Local $sTitle = "HOSxPDentalCareEntryForm"
	;Local $hWnd = WinWait($sTitle, "", 10)
	Local $class_  = "TcxCustomComboBoxInnerEdit10"
	ControlClick($hWnd, "left", $class_)
	Sleep(200)
    Local $iClick = 1
   Switch Number($sNeedProsthesis)
		Case 1
           $iClick = 1
		Case 2
			$iClick = 2
		Case 3
            $iClick = 3
		Case 4
            $iClick = 4
	    Case Else
			$iClick = 4
   EndSwitch
   For $i = 1 To $iClick
		ControlSend($hWnd, "",$class_ , "{DOWN}")
		Sleep(200)
	Next
EndFunc
;===============================  end record dental function ======================================================================

;=============================== start record Hx function ========================================================================
Func ChiefComp($cup, $date, $hStartTime, $oLogStudent, $hWndXp, $iRow)
	Local $ccclick =  "TcxCustomInnerTextEdit26"
	Local $ccadd = "TcxButton48"
	Local $cctextadd = $sSchool &" "& $cup &" "& $date
	Local $cctext = ""
		;MouseClick("left",417, 426,1,1)
    ClipPut($cctextadd)
	Sleep(100)
	ControlClick($hWndXp, "left", $ccclick)
	Sleep(500)
	;ctrlV()
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
	;Local $clk
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
	Sleep(800)
	ExitMaxTime($hStartTime,$oLogStudent,$iRow)
	Return True
EndFunc

Func Allergy()
  Sleep(200)
  Local $null_allergy = False
  Local $a = 1
  While $a < 5
	 Sleep(50)
	 if  FindToClick(@ScriptDir&$sRec_allergy_iden, 313, 230,547, 310,  0.75) Then
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
	Local $hWndItem = WinWait("HOSxPDentalOperationEntryForm", "", $iMaxRecTime)
	While True
		If $hWndItem Then
			ExitLoop
		Else
			Sleep(1000)
		EndIf
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	WEnd
	Sleep(600)
	;Local $search = "TcxCustomComboBoxInnerEdit5"
	ControlClick($hWndItem, "left", "TcxCustomInnerTextEdit7")
     ClipPut("elder")
	 Sleep(600)
	_Cpp_ctrlV()
	While True
		;ControlClick($hWndItem, "left", $search)
		;CtrlSend($hWndItem, "TcxCustomInnerTextEdit7", "priex")  ;not success with infinite loop
		Sleep(900)
		if Findtocon(@ScriptDir&$sPratomItem, 241, 252,719, 461,  0.75)  Then
			ExitLoop
		Else
			ControlClick($hWndItem, "left", "TcxCustomInnerTextEdit7")
            Sleep(500)
			ClipPut("elder")
		    Sleep(100)
		    _Cpp_ctrlA()
			Sleep(400)
			Send("{DELETE}")
	        Sleep(700)
            _Cpp_ctrlV()
		EndIf
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	WEnd
	;ControlSend($hWndItem, "","" , "{ENTER}")
	Send("{ENTER}")
	Sleep(2000)
	FindToClick(@ScriptDir&$sSaveItem, 1061, 702,1332, 807,  0.75)
	WinClose($hWndItem, "")
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

Func SetZeroPrice($hWndXp, $hStartTime, $oLogStudent, $iRow)
	 While True
		if FindToCon(@ScriptDir&$sOpdPrice1, 354, 572, 950, 771, 0.75) Then
			FindToClickRt(@ScriptDir&$sOpdPrice1,354, 572, 950, 771, 0.75)
			Sleep(1200)
		Else
			FindToClickRt(@ScriptDir&$sOpdPrice2, 354, 572, 950, 771, 0.75)
			Sleep(1200)
		EndIf
        Sleep(500)
		If FindToCon(@ScriptDir&$sEditPrice,194, 448, 1504, 656, 0.75) Then ExitLoop
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	 WEnd
	 ControlSend($hWndXp, "","" , "q")
	 Local  $hWndEdit
	 ;Local $edit = "TcxCustomInnerTextEdit2"
	  While True
		  Sleep(500)
		  $hWndEdit = WinWait("HOSxPMedicationOrderItemPriceEditForm", "", 10)
	      if $hWndEdit Then
			  CtrlSendDt($hWndEdit, "TcxCustomInnerTextEdit2", 0) ;set price = 0 baht
			  Sleep(1200)
			  ControlClick($hWndEdit, "left", "TcxButton1")
			  Sleep(1200)
			  WinClose($hWndEdit, "")
			  ExitLoop
		  Else
			  ControlSend($hWndXp, "","" , "q")
			  Sleep(1000)
		  EndIf
		  ExitMaxTime($hStartTime, $oLogStudent, $iRow)
		  ;WinClose($hWndEdit,"")
	  WEnd
	  Return True
EndFunc
;===============================  end sit function =============================================================================

;=============================== start save function ============================================================================
Func FinalSave1($hStartTime, $oLogStudent, $hWndXp, $iRow)
    Local $clk
	While True
		 Sleep(300)
		 $clk = ControlClick($hWndXp, "left", "TcxButton7")
		 ExitMaxTime($hStartTime,$oLogStudent,$iRow)
		 If $clk = 1 Then ExitLoop
	WEnd
	Return True
EndFunc

Func Select021()
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
		if FindToClick(@ScriptDir&$sFinalSave2, 420, 275, 761, 499,  0.75)   Then	ExitLoop
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

        FindToClick(@ScriptDir&$sLastOK, 311,18,1068,600,  0.75)
		Sleep(100)
		if Findtocon(@ScriptDir&$sFoundVisit, 389, 131, 470, 199, 0.75)  Then ExitLoop  ;รูปคนที่ตำแหน่งค่อนมาตรงกลาง
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	WEnd
	Return True
EndFunc

Func FinalSave2Krg($hStartTime,$oLogStudent,$iRow)
	 Local $hFeeSch
	 Local $hSpsh
	 Local $hNoAuthen

	While True
		Sleep(200)
		if Findtocon(@ScriptDir&$sFinalSave2, 420, 275, 761, 499,  0.75)   Then ExitLoop
		ExitMaxTime($hStartTime,$oLogStudent,$iRow)
	WEnd

    Sleep(200)
    Select021()
	ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	FindToClick(@ScriptDir&$sFinalSave2, 420, 275, 761, 499, 0.75)
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

        FindToClick(@ScriptDir&$sLastOK, 311, 18, 1068, 600, 0.75)
		Sleep(100)
		if Findtocon(@ScriptDir&$sFoundVisit, 389, 131, 470, 199, 0.75)  Then ExitLoop  ;รูปคนที่ตำแหน่งค่อนมาตรงกลาง
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	WEnd
	Return True
EndFunc
;===============================  end save function ============================================================================

;==================================== start bot loop ==========================================================================
Func BotLoop()
    $oExcel = _Excel_Open()
	Local $bReadOnly = True
	Local $bVisible = True
	$oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook, $bReadOnly, $bVisible)
	Local $aResult = _Excel_RangeRead($oWorkbook, $sSheet)

    If IsArray($aResult) Then
		$oLogStudent = FileOpen($logStudent, $FO_APPEND)
		$oLogStudentProgress = FileOpen($logStudenProgress, $FO_APPEND)
		$oLogPrice =  FileOpen($logPrice, $FO_APPEND)
		Sleep(1200)
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
		Local $hFinLocked
	    Local $bTodayAppoint = False
	    Local $bDentistLock = False
		Local $bFinanceLock = False
	    Local $bSitKrg = False
		Local $iCnt = 0
		Local $sSit = ""
    For $i = $iStartRow - 1 To $iEndRow - 1  ;UBound($aResult, 1) - 1
	  Sleep(800)
	  Local $hnClass  = "TcxCustomInnerTextEdit2"
	  Local $iHnVal = $aResult[$i][2]
	  ContXp2()
	  ControlClick($hWndXp, "left", $hnClass)
	  Sleep(1000)
	  ContXp2()
	  CtrlSend($hWndXp,$hnClass,$iHnVal)
	  Sleep(1000)
	  ContXp2()
      Send("{ENTER}")

	  Local $hStartTime = TimerInit() ;เริ่มจับเวลาบันทึกข้อมูล
	  Sleep(1000)
	  FileWrite($oLogStudentProgress, "Start Record R= "&$i+1&@CRLF)
	  ContXp2()
	  Local $a = 1
	  While $a < 10
		  Sleep(500)
		  $hDrgAllergy = WinGetHandle("HOSxPMedicationOrderDrugAllergyNoticeForm")
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

		 If Findtocon(@ScriptDir&$sFoundVisit, 0,81,75,169,  0.8)  Then ExitLoop ;รูปคนที่อยู่ตำแหน่งซ้ายสุดของจอ
		  $a += 1
	  WEnd

	   If $a = 10 Then
			;SwapFacebook("Error Visit R= "&$i+1)
		    FileWrite($oLogStudent, "Error Visit R= "&$i+1&", ")
			SendTeleGram("Error Visit R= "&$i+1)
			Sleep(200)
			;SwapXP()
		    ContinueLoop  ;skip this data when no visit
	   EndIf
	   FileWrite($oLogStudentProgress, "Exit First PopUp Loop R= "&$i+1&@CRLF)
	   Sleep(50)
	   FileWrite($oLogStudentProgress, "Start Init Loading Loop R= "&$i+1&@CRLF)
	   While True
		    Sleep(1000)
			$hPtNote = WinGetHandle("PatientNoteViewDisplayForm")
		    Sleep(100)
		    if $hPtNote Then WinKill($hPtNote)
			Sleep(100)
			#$hFinLocked = WinGetHandle("[CLASS:#32770]")
			Sleep(100)
		    If FindToCon(@ScriptDir&$sFiLock, 589, 354, 955, 470, 0.75) Then
			#If $hFinLocked Then
				$bFinanceLock = True
                FinanceLockExit($hWndXp, $hStartTime, $oLogStudent, $i+1)
				FileWrite($oLogStudent, "Error Finance R= "&$i+1&", ")
			    SendTeleGram("Error Finance R= "&$i+1)
				Sleep(200)
				ExitLoop
			ElseIf $bTodayAppoint Then
					if $bFinanceLock Then
						FinanceLockExit($hWndXp, $hStartTime, $oLogStudent, $i+1)
						FileWrite($oLogStudent, "Error Finance Today Appoint R= "&$i+1&", ")
						SendTeleGram("Error Finance Today Appoint R= "&$i+1)
						Sleep(200)
						ExitLoop
					Else
						PtExit($hWndXp, $hStartTime, $oLogStudent, $i+1)
						FileWrite($oLogStudent, "Error Today Appoint R= "&$i+1&", ")
						SendTeleGram("Error Today Appoint R= "&$i+1)
						Sleep(200)
						ExitLoop
					EndIf
			ElseIf Findtocon(@ScriptDir&$sLoadPtSuccess2, 698, 691,778, 749, 0.80)  Then ;diagram รูปฟันด้านล่าง
				Sleep($iSleepAfterLoad)
				ExitLoop
			Else
				Sleep(500)
			EndIf
			ExitMaxTime($hStartTime, $oLogStudent, $i+1)
	   WEnd
       FileWrite($oLogStudentProgress, "End Init Loading Loop R= "&$i+1&@CRLF)

	   If  $bTodayAppoint Then ContinueLoop
	   If  $bFinanceLock Then  ContinueLoop

	   If  $bDentistLock Then
            ;SwapFacebook("Error Locked R= "&$i+1)
			FileWrite($oLogStudent, "Error Locked R= "&$i+1&", ")
			SendTeleGram("Error Locked R= "&$i+1)
			;SwapXP()
		    Sleep(300)
		    ContinueLoop  ;skip this data when other dentist lock
	   EndIf
	   ContXp2()
	   ;Local $ttest = TimerDiff($hStartTime)
	   ;MsgBox(0,"time1", $ttest, 2)
	   FileWrite($oLogStudentProgress, "Start Do Loop Get Sit Text R= "&$i+1&@CRLF)
	   Do
	      $sSit = ControlGetText($hWndXp, "", "TcxCustomInnerTextEdit71")
		  Sleep(500)
	   Until  $sSit <> ""
	   FileWrite($oLogStudentProgress, "End Do Loop Get Sit Text R= "&$i+1&@CRLF)
        ContXp2()
		FileWrite($oLogStudentProgress, "Start Click Open Dental Loop R= "&$i+1&@CRLF)
		Local $clk_dt
		While True
			$clk_dt = ControlClick($hWndXp, "left", "TcxButton9")  ; open dental care UI
			if $clk_dt = 1 Then ExitLoop
			Sleep(1000)
		WEnd
		FileWrite($oLogStudentProgress, "End Click Open Dental Loop R= "&$i+1&@CRLF)
        ContXp2()

        Local $aDental[14]
		$aDental[0]  = $aResult[$i][5]  ;$pteeth
		$aDental[1]  = $aResult[$i][6] ;$pcaries
		$aDental[2]  = $aResult[$i][7] ;$pfilling
		$aDental[3]  = $aResult[$i][8] ;$pextract

		$aDental[4]  = $aResult[$i][9] ;$need_pfilling
		$aDental[5]  = $aResult[$i][10] ;$need_pextract

		$aDental[6]  = $aResult[$i][11] ;$permanent_perma
		$aDental[7]  = $aResult[$i][12] ;$permanent_pros
		$aDental[8]  = $aResult[$i][13]  ;$prosthesis_pros

		$aDental[9]  = $aResult[$i][14] ;$need_prosthesis
		$aDental[10]  = $aResult[$i][15] ;$need_scaling

        FileWrite($oLogStudentProgress, "Start SendDental Func R= "&$i+1&@CRLF)
	    SendDental($aDental, $hStartTime, $oLogStudent, $i+1)
		FileWrite($oLogStudentProgress, "End SendDental Func R= "&$i+1&@CRLF)
		ContXp2()
		;Sleep(500)
		;Send("{F1}") hot key F1 not work in sometimes
		FileWrite($oLogStudentProgress, "Start FindToClick F1 Menu R= "&$i+1&@CRLF)
		While 1
			Sleep(700)
			if FindToClick(@ScriptDir&$sF1, 6, 492,128, 557, 0.80) Then ExitLoop  ;go to add CC
			ExitMaxTime($hStartTime, $oLogStudent, $i+1)
		WEnd
        FileWrite($oLogStudentProgress, "End FindToClick F1 Menu R= "&$i+1&@CRLF)
        ContXp2()
		FileWrite($oLogStudentProgress, "Start FindToCon CC Pic R= "&$i+1&@CRLF)
		While 1
			Sleep(300)
			if FindToCon(@ScriptDir&$sCC, 253, 365,436, 485,  0.80) Then ExitLoop  ;CC pic
			ExitMaxTime($hStartTime,$oLogStudent,$i+1)
		WEnd
        FileWrite($oLogStudentProgress, "End FindToCon CC Pic R= "&$i+1&@CRLF)
		Sleep(750)
		ContXp2()
		FileWrite($oLogStudentProgress, "Start Record CC And Allergy R= "&$i+1&@CRLF)
		ChiefComp($aResult[$i][0], $aResult[$i][16], $hStartTime, $oLogStudent, $hWndXp, $i+1)
        Allergy()
		FileWrite($oLogStudentProgress, "End Record CC And Allergy R= "&$i+1&@CRLF)
		ContXp2()
		;Send("{F4}") hot key F4 not work in sometimes
		FileWrite($oLogStudentProgress, "Start FindToClick F4 Menu R= "&$i+1&@CRLF)
		FindToClick(@ScriptDir&$sF4, 8, 526,131, 596,  0.75) ;go to add item
		While 1
			Sleep(1000)
			if FindToClick(@ScriptDir&$sAddItem, 170, 118, 546, 241, 0.80) Then ExitLoop ;add item pic  TcxButton28
			ExitMaxTime($hStartTime, $oLogStudent, $i+1)
		WEnd
       FileWrite($oLogStudentProgress, "End Click AddItem PlusBtn Success R= "&$i+1&@CRLF)
       ContXp2()
	   FileWrite($oLogStudentProgress, "Start AddItem Func R= "&$i+1&@CRLF)
       AddItem($hStartTime, $oLogStudent, $i+1)
	   FileWrite($oLogStudentProgress, "End AddItem Func Then Find Task Pic R= "&$i+1&@CRLF)
		While True
			    Sleep(500)
              	if Findtocon(@ScriptDir&$sTask, 325, 366,629, 525,  0.80)  Then
					ExitLoop
			    Else
					Sleep(700)
					FindToClick(@ScriptDir&$sSaveItem, 1061, 702, 1332, 807, 0.75)
			    EndIf
				ExitMaxTime($hStartTime,$oLogStudent,$i+1)
		WEnd
        FileWrite($oLogStudentProgress, "AddItem Saved Then Close Add Item Box R= "&$i+1&@CRLF)
		ContXp2()
        If CheckSit($sSit) Then
			FileWrite($oLogStudentProgress, "Start SetZeroPrice Func R= "&$i+1&@CRLF)
			SetZeroPrice($hWndXp, $hStartTime, $oLogStudent, $i+1)
			FileWrite($oLogPrice, "Set Zero Price R= "&$i+1&", ")
			Sleep(300)
	    ElseIf  CheckSitKrg($sSit) Then
			$bSitKrg = True
			FileWrite($oLogPrice, "Send KRG R= "&$i+1&", ")
			Sleep(300)
		EndIf
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
			FinalSave2Krg($hStartTime, $oLogStudent, $i+1)
			FileWrite($oLogStudentProgress, "End FinalSave2Krg Func R= "&$i+1&@CRLF)
		Else
			FileWrite($oLogStudentProgress, "Start FinalSave2 Func R= "&$i+1&@CRLF)
			FinalSave2($hStartTime, $oLogStudent, $i+1)
			FileWrite($oLogStudentProgress, "End FinalSave2 Func R= "&$i+1&@CRLF)
		EndIf

		$sSit = ""
		$bTodayAppoint = False
		$bFinanceLock = False
		$bDentistLock = False
		$bSitKrg = False
		$iCnt += 1
		If ($iCnt = 5) Then
			SendTeleGram("Count= 5 R= "&$i+1)
			$iCnt = 0
		EndIf
		Sleep(500)
		FileWrite($oLogStudentProgress, "End Record R= "&$i+1&@CRLF)
		Sleep(500)
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

Func Test()
	    Local $hWndItem = WinWait("HOSxPDentalOperationEntryForm", "", 8)
	;WinActivate($hWndItem)
	Sleep(600)
	Local $search = "TcxCustomComboBoxInnerEdit5"

	While True
		ControlClick($hWndItem, "left", "TcxCustomInnerTextEdit7")
		Sleep(600)
		CtrlSend($hWndItem, "TcxCustomInnerTextEdit7", "priex") ;TcxCustomInnerTextEdit7
		;ClipPut("priex")
		;ctrlV()
		Sleep(900)
		if Findtocon(@ScriptDir&$sPratomItem, 241, 252,719, 461,  0.75)  Then ExitLoop

	WEnd
	;ControlSend($hWndItem, "","" , "{ENTER}")
	Send("{ENTER}")
	Sleep(1500)
	FindToClick(@ScriptDir&$sSaveItem, 1061, 702,1332, 807,  0.75)
	WinClose($hWndItem, "")
	Return True
EndFunc

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
			 ElseIf $iStartRow < 2 Or $iEndRow < $iStartRow Or Not IsInt($iStartRow) Or Not IsInt($iEndRow) Then
                GUICtrlSetData($AlertLabel, "โปรดเลือกช่วงแถวข้อมูลให้ถูกต้อง!")
                GUICtrlSetState($AlertLabel, $GUI_SHOW)
			 ElseIf $iSleepAfterLoad < 2000 Or $iSleepAfterLoad > 10000 Or Not IsInt($iSleepAfterLoad) Then
			    GUICtrlSetData($AlertLabel, "โปรดบันทึกค่า Delay ระว่าง 2000 ถึง 10000")
                GUICtrlSetState($AlertLabel, $GUI_SHOW)
             Else
                FileDelete($sConfigFile) ; Ensure clean write
                FileWrite($sConfigFile, $sWorkbook &@CRLF& $sSheet &@CRLF& $sSchool &@CRLF& $iSchoolType &@CRLF& $iStartRow &@CRLF& $iEndRow &@CRLF& $iSleepAfterLoad)
                ; Properly exit the GUI loop and call TestLoop()
                GUIDelete($Form1)
                ExitLoop
            EndIf
	EndSwitch
WEnd

_OpenCV_Startup()
_CppDllOpen()
Sleep(1000)
BotLoop()
;Test()

;TestToClick(@ScriptDir&$sLastOK, 311,18,1068,600,  0.75)
;TestToClick(@ScriptDir&$sFoundVisit, 389, 131,470, 199,  0.75)
;TestToClick(@ScriptDir&$sDtMenu, 10,617,146,689,  0.75)
;TestToClick(@ScriptDir&$sClosePt,982, 54,1121, 105,  0.75)
;TestToClick(@ScriptDir&$sDentalCare, 358, 182,459, 238,  0.75)
;TestToClick(@ScriptDir&$sFiLock, 589, 354,955, 470,  0.75)
;TestToClick(@ScriptDir&$sLoadPtSuccess2, 698, 691,778, 749,  0.75)
;TestToClick(@ScriptDir&$sPratomItem, 241, 252,719, 461,  0.75)
;TestToClick(@ScriptDir&$sOpdPrice1, 354, 572,950, 771,  0.75)
;TestToClick(@ScriptDir&$sOpdPrice2, 354, 572,950, 771,  0.75)
;TestToClick(@ScriptDir&$sFinalSave, 778, 50, 1095, 96,  0.75)
;TestToClick(@ScriptDir&$sFinalSave2, 420, 275, 761, 499,  0.75)
;TestToClick(@ScriptDir&$sTask, 325, 366,629, 525,  0.75)  ;
;TestToClick(@ScriptDir&$sAddCC, 714, 390,1121, 455,  0.75)
;TestToClick(@ScriptDir&$sRec_allergy_iden, 313, 230,547, 310,  0.75)
;TestToClick(@ScriptDir&$sF1, 6, 492,128, 557,  0.80)
;TestToClick(@ScriptDir&$sCC, 253, 365,436, 485,  0.80)
; TestToClick(@ScriptDir&$sEditPrice,810, 226, 1424, 670,  0.75)
;TestToClick(@ScriptDir&$sSaveItem, 1061, 702,1332, 807,  0.75)
;TestToClick(@ScriptDir&$sTest1234, 447, 400, 722, 540,  0.75)  ;447, 400, 722, 540
;FindToCon(@ScriptDir&$sEditPrice,810, 226, 1424, 670, 0.75)
;TestToClick(@ScriptDir&$sEditPrice,194, 448,1504, 656,  0.75) ;194, 448,1504, 656
_OpenCV_Shutdown()
_CppDllClose()
Exit(0)