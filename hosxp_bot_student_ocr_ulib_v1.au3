#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Match\dentist_Exp_icon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include "Lib\utility_global.au3"

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


;===================== start log path config ==================================================================
$logStudent =        $gLogDir & "\log_student.txt"
$logStudenProgress = $gLogDir & "\log_student_progress.txt"
$logPrice =          $gLogDir & "\log_student_price.txt"
;===================== end log path config ==================================================================


;=== start telegram and database request config ===============================================
$sReqPath = "query_student"
;=== end telegram  and database request config ===============================================


;=== start local picture config All Variable Use From Lib  ================================================
;Global $sOpdPrice1 =   $gMatchPicDir &  "\opd_price1.png"   ;ข้อความ ค่าบริการผู้ป่วยนอก พื้นสีฟ้า
;Global $sOpdPrice2 =   $gMatchPicDir & "\opd_price2.png"  ;ข้อความ ค่าบริการผู้ป่วยนอก ไม่มีพื้น
;Global $aOpdPricePos = [159, 713, 1083, 988]
;=== end local picture config =================================================================


;=== start local OCR text config ==============================================================
$sItemCheckTxt1 = "ตรวจสุขภาพช่องปาก โรงเรียนประถม"
$sItemCheckTxt2 = "เคลือบฟลูออไรด์วานิช (ในเด็ก)"
;$aItemTxtPos[4] = [417, 239, 1245, 658]  ;item txt position

$sGkeySearch1 = "priex"
$sGkeySearch2 = "fludex"
;=== end local OCR text config ==============================================================



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


;=============== send request telegram func ==========================================
;=============== send request telegram func ==========================================


;=============== send request nodejs ocr localhost func ======================================
;StartServer2 StopServer2 MsgStopServer2
;=============== end request nodejsocr localhost func ======================================


;=============================== start preload function ======================================================================
; HideExcel ShowExcel HideSciTE ShowSciTE SwapXP SetHnBox HideServerAPI2
;=============================== end preload function ======================================================================


;=============================== start after load exit function =====================================================================
; reuse 100% with utility_global.au   KillFinance2 FinanceLockExit PtExit
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

Func SelectPtType($hWnd,$sShoolClass)
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
	Local $place  = "TcxCustomComboBoxInnerEdit5"  ;TcxCustomComboBoxInnerEdit5
	ControlClick($hWnd, "left", $place)
	Sleep(100)
	For $i = 1 To 2
		ControlSend($hWnd, "",$place , "{DOWN}")
		Sleep(100)
	Next
EndFunc

Func SelectBoxEduRtb($hWnd)
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
	Local $edu  = "TcxCustomComboBoxInnerEdit3"
	ControlClick($hWnd, "left", $edu)
	Sleep(100)
	Local $iClick = 5
	For $i = 1 To $iClick
		ControlSend($hWnd, "",$edu , "{DOWN}")
		Sleep(100)
	Next
EndFunc


Func SelectClass($hWnd,$sShoolClass)
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
;ChiefComp Allergy AddItemOCR1 AddItemOCR2 ClickSaveOperation
;===============================  end record Hx function ========================================================================


;=============================== start sit function =============================================================================
; SetZeroPrice
Func CheckSit($sText)
    $sText = StringStripWS($sText, 3)

    For $i = 0 To UBound($ArraySit) - 1
        If StringStripWS($ArraySit[$i], 3) = $sText Then
            Return True
        EndIf
    Next

    Return False
EndFunc

Func CheckSitKrg($sText)
    ; ลบช่องว่างหน้า/หลัง
    $sText = StringStripWS($sText, 3)
    For $i = 0 To UBound($ArraySitKrg) - 1
        ; ข้าม index ที่ยังไม่ได้กำหนดค่า
        If $ArraySitKrg[$i] = "" Then ContinueLoop
        ; เทียบแบบตรงตัว
        If StringStripWS($ArraySitKrg[$i], 3) = $sText Then
            Return True
        EndIf
    Next

    Return False
EndFunc
;===============================  end sit function =============================================================================


;=============================== start save function ============================================================================
; reuse 100% with utility_global.au FinalSave1 Select021 FinalSave2 FinalSave2Krg
;===============================  end save function ===============================================================================


;=============================== start short cut F1 F4 function ========================================================================
; reuse 100% with utility_global.au Click_F1 Click_F4 ClickAddOperation
;=============================== end short cut F1 F4 function ========================================================================




;==================================== start bot loop ===========================================================================================================
Func BotLoop()
    $oExcel = _Excel_Open()
	Local $bReadOnly = True
	Local $bVisible = True
	$oWorkbook = _Excel_BookOpen($oExcel, $sWorkbook, $bReadOnly, $bVisible)
	Local $aResult = _Excel_RangeRead($oWorkbook, $sSheet)
    Local Const $bAllowRecWithTodayAppoint = True  ;ถ้านักเรียนมาตามข้อมูลนัดวันนี้ No Finance Locked and ต้องการบันทึก row ให้ set True

    If IsArray($aResult) Then
		$oLogStudent = FileOpen($logStudent, $FO_APPEND)
		$oLogStudentProgress = FileOpen($logStudenProgress, $FO_APPEND)
		$oLogPrice =  FileOpen($logPrice, $FO_APPEND) ;
		Sleep(1200)
		HideExcel()
		HideSciTE()
		HideServerAPI2()
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
	   $bTodayAppoint = False
		$bDentistLock = False
		$bFinanceLock = False
	    _OCRSetRow(0, 0, 0)
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

	  Local $hStartTime = TimerInit() ;เริ่มจับเวลาบันทึกข้อมูล
	  _OCRSetRow($hStartTime, $oLogStudent, $i+1)
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

		  $hDentistLock = WinGetHandle($sLockPtForm)  ;  "Visit Number Locked"
		  Sleep(100)
		  If $hDentistLock Then
				  $bDentistLock = True
				  WinKill($hDentistLock)
				  Sleep(300)
				  ExitLoop
		  EndIf
		  Sleep(300)

		 If Findtocon($sFoundVisit, $aFvFirstPos[0], $aFvFirstPos[1], $aFvFirstPos[2], $aFvFirstPos[3], 0.8)  Then ;รูปคนที่อยู่ตำแหน่งซ้ายสุดของจอ
				Sleep(900)
				ExitLoop
		 EndIf
		  $a = $a+1
	  WEnd

	   If $bDentistLock Then
            Sleep(100)
			FileWrite($oLogStudent, "Error Locked R= "&$i+1&", ")
			SendTeleGram("Error Locked R= "&$i+1)
		    Sleep(300)
		    ContinueLoop
	   EndIf

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
	  Local $bAllowAppointLogSend = True
	   While True
		    Sleep(1000)
			$hPtNote = WinGetHandle("PatientNoteViewDisplayForm")
		    Sleep(100)
		    if $hPtNote Then WinKill($hPtNote)
			Sleep(100)
		    ;If FindToCon(@ScriptDir&$sFiLock, $aFiLockPos[0], $aFiLockPos[1], $aFiLockPos[2], $aFiLockPos[3], 0.75) Then
			if KillFinance2() Then
					$bFinanceLock = True
					PtExit($hWndXp, $hStartTime, $oLogStudent, $i+1)
					FileWrite($oLogStudent, "Warning Finance Locked R= "&$i+1&", ")
					SendTeleGram("Warning Finance Locked R= "&$i+1)
					Sleep(200)
					ExitLoop
			ElseIf $bTodayAppoint And Not $bAllowRecWithTodayAppoint Then
					; ตรงนัดวันนี้และไม่อนุญาต: ออกจากหน้าผู้ป่วยเพื่อข้ามแถว
					PtExit($hWndXp, $hStartTime, $oLogStudent, $i+1)
					FileWrite($oLogStudent, "Warning Today Appoint No Rec R= " & ($i+1) & ", ")
					SendTeleGram("Warning Today Appoint No Rec R= " & ($i+1))
					Sleep(200)
					ExitLoop
			ElseIf $bTodayAppoint And $bAllowRecWithTodayAppoint And $bAllowAppointLogSend Then
					; ตรงนัดวันนี้และอนุญาต: แจ้งครั้งเดียว แล้ววนรอโหลดสำเร็จ
					FileWrite($oLogStudent, "Warning Today Appoint Allow Rec R= " & ($i+1) & ", ")
					SendTeleGram("Warning Today Appoint Allow Rec R= " & ($i+1))
					$bAllowAppointLogSend = False
					Sleep(200)
			ElseIf Findtocon($sLoadPtSuccess2, $aLoadPtPos[0], $aLoadPtPos[1], $aLoadPtPos[2], $aLoadPtPos[3], 0.80)  Then   ;refresh blue pic
				Sleep($iSleepAfterLoad)
				ExitLoop
			Else
				Sleep(500)
			EndIf
			ExitMaxTime($hStartTime,$oLogStudent,$i+1)
	   WEnd
	   ;== end hos xp loading loop ===================================
	   FileWrite($oLogStudentProgress, "End Init Loading Loop R= "&$i+1&@CRLF)

	   If $bFinanceLock Then  ContinueLoop

	   If $bTodayAppoint And Not $bAllowRecWithTodayAppoint Then ContinueLoop

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

		FileWrite($oLogStudentProgress, "Start Send F1 R= " & ($i + 1) & @CRLF)
		Click_F1($hWndXp, $hStartTime, $oLogStudent, $i + 1)
        FileWrite($oLogStudentProgress, "End Send F1 R= " & ($i + 1) & @CRLF)

        ContXp2()
		FileWrite($oLogStudentProgress, "Start FindToCon CC Pic R= "&$i+1&@CRLF)
		While 1
			Sleep(300)
			if FindToCon($sCC, $aCcPos[0], $aCcPos[1], $aCcPos[2], $aCcPos[3], 0.80) Then ExitLoop  ;CC pic
			ExitMaxTime($hStartTime,$oLogStudent,$i+1)
		WEnd
        FileWrite($oLogStudentProgress, "End FindToCon CC Pic R= "&$i+1&@CRLF)
		Sleep(750)
		ContXp2()
		FileWrite($oLogStudentProgress, "Start Record CC And Allergy R= "&$i+1&@CRLF)
		ChiefComp($aDental ,$hStartTime, $oLogStudent, $hWndXp, $i+1)  ; โยนทั้งก้อน array เพื่อใช้ $aDental[13] เชื่อมข้อความระดับชั้น
        Allergy()
		if Number($sBw) > 0 Then
			;CtrlSendDt($hWndXp, "TcxCustomInnerTextEdit59", $sBw)
			Local $hBwCtrl = ControlGetHandle($hWndXp, "", $gInputBW)
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
       AddItemOCR1($hStartTime,$oLogStudent,$i+1)
	   FileWrite($oLogStudentProgress, "End AddItem Func  R= "&$i+1&@CRLF)

	    While True
			Sleep(500)
			Local $hOperationForm = WinGetHandle($sOperationBoxTitle)
			If $hOperationForm = 0 Then ExitLoop
			If Not _WinAPI_IsWindowVisible($hOperationForm) Then ExitLoop
			Sleep(700)
			ClickSaveOperation($hOperationForm)
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
			AddItemOCR2($hStartTime, $oLogStudent, $i+1)
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
		EndIf

		If CheckSit($sSit) Then
			FileWrite($oLogStudentProgress, "Start SetZeroPrice Func R= "&$i+1&@CRLF)
			SetZeroPrice($hWndXp,$hStartTime,$oLogStudent, $oLogStudentProgress, $i+1)
			FileWrite($oLogPrice, "Set Zero Price R= "&$i+1&", ")
			Sleep(300)
			FileWrite($oLogStudentProgress, "End SetZeroPrice Func R= "&$i+1&@CRLF)
		EndIf

	    If  CheckSitKrg($sSit) Then
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
If Not StartServer2() Then Exit(1)
Sleep(1000)
BotLoop()
_OpenCV_Shutdown()
_CppDllClose()
MsgStopServer2()
Exit(0)