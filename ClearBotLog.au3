#include <FileConstants.au3>
#include <MsgBoxConstants.au3>

; วางสคริปต์นี้ที่ root ของโปรเจกต์ และปิดบอทก่อนล้าง log
; ล้างเฉพาะไฟล์ .txt ใน BotLog ไม่รวมโฟลเดอร์ย่อย และไม่ลบไฟล์
Local $sLogDir = @ScriptDir & "\BotLog"
If Not FileExists($sLogDir) Then
    MsgBox($MB_ICONERROR, "Clear BotLog", "ไม่พบโฟลเดอร์: " & $sLogDir)
    Exit 1
EndIf

Local $hSearch = FileFindFirstFile($sLogDir & "\*.txt")
If $hSearch = -1 Then
    MsgBox($MB_ICONINFORMATION, "Clear BotLog", "ไม่พบไฟล์ .txt ใน BotLog")
    Exit 0
EndIf

Local $iCleared = 0, $iFailed = 0, $sFailed = ""
While True
    Local $sName = FileFindNextFile($hSearch)
    If @error Then ExitLoop
    Local $sPath = $sLogDir & "\" & $sName
    If StringInStr(FileGetAttrib($sPath), "D") Then ContinueLoop

    ; เปิดแบบ overwrite เพื่อล้างเนื้อหา โดยคงชื่อและไฟล์เดิมไว้
    Local $hFile = FileOpen($sPath, $FO_OVERWRITE)
    If $hFile = -1 Then
        $iFailed += 1
        $sFailed &= @CRLF & $sName
        ContinueLoop
    EndIf
    FileClose($hFile)
    $iCleared += 1
WEnd
FileClose($hSearch)

If $iFailed > 0 Then
    MsgBox($MB_ICONWARNING, "Clear BotLog", "ล้างแล้ว " & $iCleared & " ไฟล์" & @CRLF & "ล้างไม่สำเร็จ " & $iFailed & " ไฟล์:" & $sFailed)
    Exit 1
EndIf
MsgBox($MB_ICONINFORMATION, "Clear BotLog", "ล้างเนื้อหาแล้ว " & $iCleared & " ไฟล์ โดยไม่ลบไฟล์")
Exit 0
