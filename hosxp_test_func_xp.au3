#RequireAdmin
#include <Excel.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPI.au3>
#include <SendMessage.au3>
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
Global $sHosXp_Title = "BMS-HOSxP XE 4.0"
Global $iMaxTestTime = 10000

HotKeySet("{ESC}", "Hkey")

Func Hkey()
		Exit 0
EndFunc

Func SetHnBox($hWndXp)
	; Find the HN panel
	Local $hCtrl = ControlGetHandle($hWndXp, "", "[CLASS:TcxGroupBox; INSTANCE:1]")
	If @error Or $hCtrl = 0 Then Return False
	Local $aPos = ControlGetPos($hWndXp, "", $hCtrl)
	If @error Or Not IsArray($aPos) Then Return False
	; Move the panel only when necessary
	If $aPos[0] <> 0 Or $aPos[1] <> 125 Or $aPos[2] <> 400 Or $aPos[3] <> 676 Then
		If Not ControlMove($hWndXp, "", $hCtrl, 0, 125, 400, 676) Then Return False
		Sleep(500)
	EndIf
	Return True
EndFunc

Func ActHosXp()
	Local $hWndXp = WinGetHandle($sHosXp_Title)
	If @error Or $hWndXp = 0 Then
		MsgBox($MB_ICONERROR, "Error", "HOSxP window was not found.")
		Return 0
	EndIf
	; Activate once only if necessary
	If Not WinActive($hWndXp) Then
		WinActivate($hWndXp)
		If Not WinWaitActive($hWndXp, "", 5) Then
			MsgBox($MB_ICONERROR, "Error", "Unable to activate HOSxP.")
			Return 0
		EndIf
	EndIf
	Sleep(500)
;~ 	If Not SetHnBox($hWndXp) Then
;~ 		MsgBox($MB_ICONERROR, "Error", "Unable to position the HN panel.")
;~ 		Return 0
;~ 	EndIf
	Return $hWndXp
EndFunc

Func ExitMaxTime($hStartTime)
	; Stop the test without writing logs
	If TimerDiff($hStartTime) >= $iMaxTestTime Then
		MsgBox($MB_ICONERROR, "Error", "F1 click test timed out.")
		Exit(1)
	EndIf
EndFunc

Func Click_F1($hWndXp)
	If Not WinExists($hWndXp) Then
		MsgBox($MB_ICONERROR, "Error", "HOSxP window was not found.")
		Return False
	EndIf
	Local $hStartTime = TimerInit()
	; Activate once if HOSxP is not already active
	If Not WinActive($hWndXp) Then WinActivate($hWndXp)
	While Not WinActive($hWndXp)
		Sleep(200)
		ExitMaxTime($hStartTime)
	WEnd
	Sleep(300)
	If Not WinActive($hWndXp) Then
		MsgBox($MB_ICONERROR, "Error", "F1 cancelled: the active window changed.")
		Return False
	EndIf
	Send("{F1}")
	Sleep(2000)
	Return True
EndFunc

Func Click_F4($hWndXp)
	If Not WinExists($hWndXp) Then
		MsgBox($MB_ICONERROR, "Error", "HOSxP window was not found.")
		Return False
	EndIf
	Local $hStartTime = TimerInit()
	; Activate once if HOSxP is not already active
	If Not WinActive($hWndXp) Then WinActivate($hWndXp)
	While Not WinActive($hWndXp)
		Sleep(200)
		ExitMaxTime($hStartTime)
	WEnd
	Sleep(300)
	If Not WinActive($hWndXp) Then
		MsgBox($MB_ICONERROR, "Error", "F1 cancelled: the active window changed.")
		Return False
	EndIf
	Send("{F4}")
	Sleep(2000)
	Return True
EndFunc


Func ClickAddOperation($hWndXp)
	Local $sFormTitle = "HOSxPDentalOperationEntryForm"
	Local $sButton = "[CLASS:TcxButton; INSTANCE:31]"

	If Not WinExists($hWndXp) Then
		MsgBox($MB_ICONERROR, "Error", "HOSxP window was not found.")
		Return False
	EndIf
	; Prevent an existing form from producing a false success
	If WinExists($sFormTitle) Then
		MsgBox($MB_ICONWARNING, "Test", "Close the operation entry form before testing.")
		Return False
	EndIf
	Local $hButton = ControlGetHandle($hWndXp, "", $sButton)
	If @error Or $hButton = 0 Then
		MsgBox($MB_ICONERROR, "Error", "The Add Operation button was not found.")
		Return False
	EndIf
	; Verify the exact button text
	If ControlGetText($hWndXp, "", $hButton) <> "เพิ่มหัตถการ" Then
		MsgBox($MB_ICONERROR, "Error", "The button text does not match.")
		Return False
	EndIf
	If ControlClick($hWndXp, "", $hButton, "left", 1) <> 1 Then
		MsgBox($MB_ICONERROR, "Error", "ControlClick failed.")
		Return False
	EndIf
	; Confirm that the expected form actually becomes visible
	Local $hStartTime = TimerInit()
	While TimerDiff($hStartTime) < 10000
		Local $hForm = WinGetHandle($sFormTitle)
		If $hForm <> 0 Then
			If _WinAPI_IsWindowVisible($hForm) Then
				MsgBox($MB_ICONINFORMATION, "Success", "The operation entry form opened successfully.")
				Return True
			EndIf
		EndIf
		Sleep(200)
	WEnd
	MsgBox($MB_ICONERROR, "Error", "The operation entry form did not appear within 10 seconds.")
	Return False
EndFunc


;~ Local $hTestWnd = ActHosXp()
;~ If $hTestWnd <> 0 Then TestClickF1($hTestWnd)

;~ Local $hTestWnd = WinGetHandle($sHosXp_Title)
;~ If $hTestWnd <> 0 Then
;~ 	Click_F4($hTestWnd)
;~ Else
;~ 	MsgBox($MB_ICONERROR, "Error", "HOSxP window was not found.")
;~ EndIf



Local $hTestWnd = WinGetHandle($sHosXp_Title)
ClickAddOperation($hTestWnd)
Exit(0)