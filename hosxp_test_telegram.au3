#include <MsgBoxConstants.au3>

Global $gEnvCache = 0       ; เก็บ cache (Dictionary ของ KEY=VALUE)
Global $gEnvFile = @ScriptDir & "\server\.env"   ; ใช้ path เดียวกับสคริปต์
Global $gEnvTimestamp = ""  ; เวลาล่าสุดของไฟล์
Global $gTelegramToken = ""
Global $gTelegramChatID = ""
Global $gHTTPTelegram = ObjCreate("WinHttp.WinHttpRequest.5.1")

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


; ====== ตัวอย่างการใช้งาน ======
;~ Local $oMap = LoadEnv($gEnvFile)
;~ For $key In $oMap.Keys
;~     ConsoleWrite($key & " = " & $oMap($key) & @CRLF)
;~ Next

$gTelegramToken = GetEnv("TELEGRAM_TOKEN")
$gTelegramChatID = GetEnv("TELEGRAM_CHAT")

SendTeleGram("test 999")


