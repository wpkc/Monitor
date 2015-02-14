Option Explicit
Const regKey = "kc-monitor"
Dim MsgDeleteSettings

MsgDeleteSettings = "Do you want to remove Monitor registry settings as well?" & vbNewLine & vbNewLine & _
                    "If you click No, your settings will not be deleted and" & vbNewLine & _
			"will be used again if you reinstall this script."

If MsgBox(MsgDeleteSettings, vbYesNo) = vbYes Then
	Dim WShell : Set WShell = CreateObject("WScript.Shell")
	WShell.RegDelete "HKCU\Software\MediaMonkey\Scripts\" & regKey & "\"
End If
