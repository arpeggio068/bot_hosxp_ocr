#RequireAdmin
#include-once
Global $hDLL
Func _CppDllOpen()
    $hDLL = DllOpen("DLLs\KeySimulator.dll")
    If $hDLL = -1 Then
        MsgBox(16, "Error", "Failed to load KeySimulator.dll")
        Exit
    EndIf
EndFunc

Func _CppDllClose()
      DllClose($hDLL)
EndFunc

Func  _Cpp_ctrlV()
     Local $aResult = DllCall($hDLL, "none", "SimulateCtrlV")
        If @error Then
            MsgBox(16, "DLL Call Error", "Error calling SimulateCtrlV")
        Else
            ;MsgBox(64, "Success", "SimulateCtrlV executed successfully!")
			ConsoleWrite("OK")
        EndIf
EndFunc

Func  _Cpp_ctrlA()
     Local $aResult = DllCall($hDLL, "none", "SimulateCtrlA")
        If @error Then
            MsgBox(16, "DLL Call Error", "Error calling SimulateCtrlA")
        Else
            ;MsgBox(64, "Success", "SimulateCtrlV executed successfully!")
			ConsoleWrite("OK")
        EndIf
EndFunc

Func  _Cpp_ctrlC()
     Local $aResult = DllCall($hDLL, "none", "SimulateCtrlC")
        If @error Then
            MsgBox(16, "DLL Call Error", "Error calling SimulateCtrlC")
        Else
            ;MsgBox(64, "Success", "SimulateCtrlV executed successfully!")
			ConsoleWrite("OK")
        EndIf
EndFunc

;~ _CppDllOpen()
;~ Sleep(1000)
;~ Local $hWnd = WinActivate("[CLASS:Notepad]")
;~ WinWaitActive($hWnd, "", 2)
;~ ClipPut("i am glad")
;~ Sleep(500)
;~ ControlClick($hWnd, "left",  "Edit1")
;~ Sleep(500)
;~ _Cpp_ctrlA()
;~ Sleep(500)
;~ _Cpp_ctrlV()
;~ _CppDllClose()
