#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Match\telemed.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include "Lib\utility_global.au3"

Opt("MouseCoordMode", 1)
Opt("WinTitleMatchMode", 2)
HotKeySet("{ESC}", "Hkey")

;=============================== start GUI config ========================================================================
Global $sConfigFile = @ScriptDir &"\config\telemed_config.txt"
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
Global $Form1 = GUICreate("Telemed Bot", 634, 502, 192, 124)
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

$sSchool = $iDefaultValue3

Global $iSchoolType = $iDefaultValue4 ; รัฐบาล = 1 เอกชน = 2
Global $iStartRow = $iDefaultValue5
Global $iEndRow = $iDefaultValue6

$iSleepAfterLoad = $iDefaultValue7
;=============================== end GUI config ===================================================================


;===================== start log path config ==================================================================
$logStudent =        $gLogDir & "\log_telemed.txt"
$logStudenProgress = $gLogDir & "\log_telemed_progress.txt"
$logPrice =          $gLogDir & "\log_telemed_price.txt"
;===================== end log path config ==================================================================


;=== start nodejs database request config ===============================================
$sReqPath = "query_telemed"
;=== end nodejs database request config ===============================================


;=== start local picture config ==============================================================
Global Const $sTeleItem40Down = $gMatchPicDir &  "\xp_item_found_telemed40down.png"  ;ข้อความ telemed < 40 y
Global Const $sTeleItem40Up =   $gMatchPicDir & "\xp_item_found_telemed40up.png"  ;ข้อความ telemed >= 40 y
Global Const $aItemTxtPos = [417, 239, 1245, 658]  ;item txt position

; Global Const $sOpdPrice30_1 =  $gMatchPicDir & "\opd_price30.png"   ;ข้อความ ค่าบริการ 30 บาท ไม่มีพื้น
; Global Const $sOpdPrice30_2 =  $gMatchPicDir & "\opd_price30_2.png"  ;ข้อความ ค่าบริการ 30 บาท พื้นสีฟ้า
; Global Const $aOpdPricePos = [159, 713, 1083, 988]
;=== end local picture config =================================================================


Global $ArraySit30B[4] = []
$ArraySit30B[0] = "50"  ;12-59 นอกเขต
$ArraySit30B[1] = "50d"  ;12-59 นอกเขต  d=op anywhere ในจังหวัด
$ArraySit30B[2] = "56"  ;12-59 ในเขต
$ArraySit30B[3] = "50o"  ;12-59 นอกเขต  o=op anywhere ต่างจังหวัด

; Global $aCheckPttype[11] = ["17", "19", "21", "28", "30", "32", "37", "44", "50", "59", "99"]

Global $ArraySitKrg[15]  ;สิทธิที่ต้องส่งต่อห้องการเงิน 021
$ArraySitKrg[0] = "01"  ; ขรก. (องค์กรปกครองส่วนท้องถิ่น)
$ArraySitKrg[1] = "55"  ; ขรก.กรมบัญชีกลาง
$ArraySitKrg[2] = "34"  ; ขรก.กรุงเทพมหานคร
$ArraySitKrg[3] = "35"  ; ขรก.กกต.
$ArraySitKrg[4] = "17o"  ;ครอบครัวอสม นอกเขต  o=op anywhere ต่างจังหวัด
$ArraySitKrg[5] = "19o"  ;ครอบครัวผู้นำชุมชน นอกเขต  o=op anywhere ต่างจังหวัด
$ArraySitKrg[6] = "21o"  ;ครอบครัวทหารผ่านศึก นอกเขต  o=op anywhere ต่างจังหวัด
$ArraySitKrg[7] = "28o"  ;ทหารผ่านศึก นอกเขต  o=op anywhere ต่างจังหวัด
$ArraySitKrg[8] = "30o"  ;พระ นอกเขต  o=op anywhere ต่างจังหวัด
$ArraySitKrg[9] = "32o"  ;สูงอายุ นอกเขต  o=op anywhere ต่างจังหวัด
$ArraySitKrg[10] = "37o"  ;พิการ นอกเขต  o=op anywhere ต่างจังหวัด
$ArraySitKrg[11] = "44o"  ;อสม นอกเขต  o=op anywhere ต่างจังหวัด
$ArraySitKrg[12] = "50o"  ;12-59 นอกเขต  o=op anywhere ต่างจังหวัด
$ArraySitKrg[13] = "59o"  ;ผู้นำชุมชน นอกเขต  o=op anywhere ต่างจังหวัด
$ArraySitKrg[14] = "99o"  ;รายได้น้อย นอกเขต  o=op anywhere ต่างจังหวัด


;=============== send request telegram func ==========================================
;=============== send request telegram func ==========================================

;=============== send request nodejs localhost func ======================================
; StartServer1 StopServer1 MsgStopServer1
;=============== end request nodejs localhost func ======================================


;=============================== start preload function ======================================================================
; HideExcel ShowExcel HideSciTE ShowSciTE SwapXP SetHnBox HideServerAPI1
;=============================== end preload function ======================================================================


;=============================== start after load function ======================================================================
; reuse 100% with utility_global.au   KillFinance2 FinanceLockExit PtExit
;=============================== end after load function ======================================================================



;=============================== start record Hx function ========================================================================
;ChiefComp Allergy ClickSaveOperation
Func AddItemAge40Down($hStartTime, $oLogStudent, $iRow)
	Local $hWndItem
	$hWndItem = WinGetHandle($sOperationBoxTitle)

	Local Const $sKeySearch = "telemed1"
     ClipPut($sKeySearch)
	 Sleep(300)
	_Cpp_ctrlV()

	While True
		Sleep(900)
		if Findtocon($sTeleItem40Down, $aItemTxtPos[0],  $aItemTxtPos[1],  $aItemTxtPos[2],  $aItemTxtPos[3], 0.75)  Then
			ExitLoop
		Else
			Send("{ESC}") ;reset form and no active click
			Sleep(500)
			ControlClick($hWndItem, "left", "TcxCustomInnerTextEdit7")
			ClipPut($sKeySearch)
            Sleep(500)
            _Cpp_ctrlV()
			Sleep(1000)
		EndIf
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	WEnd
	;ControlSend($hWndItem, "","" , "{ENTER}")
	Send("{ENTER}")
	Sleep(1700)
	ClickSaveOperation($hWndItem)
	; รอให้โปรแกรมประมวลผลการบันทึก
	Sleep(500)
	Return True
EndFunc


Func AddItemAge40Up($hStartTime, $oLogStudent, $iRow)
	Local $hWndItem
	$hWndItem = WinGetHandle($sOperationBoxTitle)

	Local Const $sKeySearch = "telemed2"
     ClipPut($sKeySearch)
	 Sleep(300)
	_Cpp_ctrlV()

	While True
		Sleep(900)
		if Findtocon($sTeleItem40Up, $aItemTxtPos[0],  $aItemTxtPos[1],  $aItemTxtPos[2],  $aItemTxtPos[3], 0.75)  Then
			ExitLoop
		Else
			Send("{ESC}") ;reset form and no active click
			Sleep(500)
			ControlClick($hWndItem, "left", "TcxCustomInnerTextEdit7")
			ClipPut($sKeySearch)
            Sleep(500)
            _Cpp_ctrlV()
			Sleep(1000)
		EndIf
		ExitMaxTime($hStartTime, $oLogStudent, $iRow)
	WEnd
	;ControlSend($hWndItem, "","" , "{ENTER}")
	Send("{ENTER}")
	Sleep(1700)
	ClickSaveOperation($hWndItem)
	Sleep(500)
	Return True
EndFunc

;===============================  end record Hx function ========================================================================



;=============================== start sit function =============================================================================
Func CheckSit30B($sText)
    $sText = StringStripWS($sText, 3)

    For $i = 0 To UBound($ArraySit30B) - 1
        If StringStripWS($ArraySit30B[$i], 3) = $sText Then
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
;=============================== end save function ============================================================================


;=============================== short cut F1 F4 function ========================================================================
; reuse 100% with utility_global.au Click_F1 Click_F4 Ctrl_F2 ClickAddOperation
;=============================== end short cut F1 F4 function ====================================================================


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
		HideServerAPI1()
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
	  Sleep(500)
	  ; Local $hnClass  = "TcxCustomInnerTextEdit2"
	  Local $iHnVal = $aResult[$i][2]
	  $sQueue = $aResult[$i][1]
	  Local $iAgeVal = Number($aResult[$i][3])
	  $sSit = String($aResult[$i][5])

	  FileWrite($oLogStudentProgress, "Start Get dbData R= "&$i+1&@CRLF)
	  Sleep(100)
	  Local $dbData = GetSitFromDb2($iHnVal, $oLogStudent, $i+1)
	  Sleep(100)
		  ; $sSit = $dbData[0]
	  $sBw = $dbData[1]

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
	  Sleep(1000)
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

		  $hDentistLock = WinGetHandle($sLockPtForm)
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

	    If  $bDentistLock Then
            Sleep(100)
			FileWrite($oLogStudent, "Error Locked R= "&$i+1&", ")
			SendTeleGram("Error Locked R= "&$i+1)
		    Sleep(300)
		    ContinueLoop  ;skip this data when other dentist lock
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
	   While True
		    Sleep(1000)
			$hPtNote = WinGetHandle("PatientNoteViewDisplayForm")
		    Sleep(100)
		    if $hPtNote Then WinKill($hPtNote)
			Sleep(100)

			if KillFinance2() Then
					$bFinanceLock = True
					PtExit($hWndXp, $hStartTime, $oLogStudent, $i+1)
					FileWrite($oLogStudent, "Error Finance R= "&$i+1&", ")
					SendTeleGram("Error Finance R= "&$i+1)
					Sleep(200)
					ExitLoop
			ElseIf $bTodayAppoint Then
					PtExit($hWndXp, $hStartTime, $oLogStudent, $i+1)
					FileWrite($oLogStudent, "Warning Today Appoint R= "&$i+1&", ")
					SendTeleGram("Warning Today Appoint R= "&$i+1)
					Sleep(200)
					ExitLoop
			ElseIf Findtocon($sLoadPtSuccess2, $aLoadPtPos[0], $aLoadPtPos[1], $aLoadPtPos[2], $aLoadPtPos[3], 0.80)  Then   ;refresh blue pic
					Sleep($iSleepAfterLoad)
					ExitLoop
			Else
				Sleep(200)
			EndIf

			ExitMaxTime($hStartTime,$oLogStudent,$i+1)
	   WEnd
	;== end hos xp loading loop ===================================
	   FileWrite($oLogStudentProgress, "End Init Loading Loop R= "&$i+1&@CRLF)

	   If  $bFinanceLock Then  ContinueLoop
	   If  $bTodayAppoint Then  ContinueLoop

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
		Local $vDetail = ""    ;$aDetail[2]
		;$aDetail[0] = $sSchool ; cup
		;$aDetail[1] = $sSchool ; date exam
		ChiefComp($vDetail ,$hStartTime, $oLogStudent, $hWndXp, $i+1)
        Allergy()
		if Number($sBw) > 0 Then
			Local $hBwCtrl = ControlGetHandle($hWndXp, "", "[CLASS:TcxDBTextEdit; INSTANCE:18]")  ;TcxDBTextEdit18
			ControlSetText($hWndXp, "", $hBwCtrl, $sBw)
			Sleep(300)
		EndIf
		FileWrite($oLogStudentProgress, "End Record CC And Allergy R= "&$i+1&@CRLF)
		ContXp2()

		FileWrite($oLogStudentProgress, "Start Click HPI R= "&$i+1&@CRLF)
		Ctrl_F2($hWndXp, $hStartTime, $oLogStudent, $i + 1)
		FileWrite($oLogStudentProgress, "End Click HPI R= "&$i+1&@CRLF)

		FileWrite($oLogStudentProgress, "Start Add HPI Text R= "&$i+1&@CRLF)
		Local $hMemo = 0

		While Not $hMemo
			Sleep(1000)
			$hMemo = ControlGetHandle($hWndXp, "", "[CLASSNN:TcxCustomInnerMemo2]")
			ExitMaxTime($hStartTime,$oLogStudent,$i+1)
		WEnd

		While Not _WinAPI_IsWindowVisible($hMemo)
			Sleep(1000)
			ExitMaxTime($hStartTime,$oLogStudent,$i+1)
		WEnd

		ControlFocus($hWndXp, "", $hMemo)
		Sleep(500)
		ControlSetText($hWndXp, "", $hMemo, $sSchool)
		Sleep(500)
		ControlClick($hWndXp, "", $hMemo)
		Sleep(700)
		Send("{ENTER}")
        Sleep(500)
		FileWrite($oLogStudentProgress, "End Add HPI Text R= "&$i+1&@CRLF)

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
	   If $iAgeVal < 40 Then
		   AddItemAge40Down($hStartTime, $oLogStudent, $i+1)
	   Else
		   AddItemAge40Up($hStartTime, $oLogStudent, $i+1)
	   EndIf
	   FileWrite($oLogStudentProgress, "End AddItem Func Then check HOSxPDentalOperationEntryForm not invisible= "&$i+1&@CRLF)

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


        If CheckSit30B($sSit) Then
			FileWrite($oLogStudentProgress, "Start SetZeroPrice30B Func R= "&$i+1&@CRLF)
			SetZeroPrice30B($hWndXp,$hStartTime,$oLogStudent, $oLogStudentProgress, $i+1)
			FileWrite($oLogPrice, "Set Zero Price 30B R= "&$i+1&", ")
			Sleep(300)
			FileWrite($oLogStudentProgress, "End SetZeroPrice30B Func R= "&$i+1&@CRLF)
		EndIf

	    If  CheckSitKrg($sSit) Then
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
			 ElseIf $iSleepAfterLoad < 1500 Or $iSleepAfterLoad > 10000 Or Not IsInt($iSleepAfterLoad) Then
			    GUICtrlSetData($AlertLabel, "โปรดบันทึกค่า Delay ระว่าง 1500 ถึง 10000")
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

_OpenCV_Startup()  ;
_CppDllOpen()
$gTelegramToken = GetEnv("TELEGRAM_TOKEN")
$gTelegramChatID = GetEnv("TELEGRAM_CHAT")
StartServer1()
Sleep(1000)
BotLoop()
_OpenCV_Shutdown()
_CppDllClose()
MsgStopServer1()
Exit(0)


