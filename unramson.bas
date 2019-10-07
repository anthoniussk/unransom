'*******************************************************************************************************
'
' Vypnutie spustania scriptov cez dvojklik - ochrana pred najcastejsim napadnutim cez prilohu emailu
'
' Autor: Anthonius, Anton Pitak, SOFTPAE.com
'
' Nutne spustat s administratorskymi pravami
'
' Arguments: Use -run to add Run command replacing Open command, use -quiet for logon script usage 
'
'*******************************************************************************************************

#Include "registry.bi"

Sub FixRegSettings(ByVal Key As String)
	
	Dim regpath As String 
	Dim keyname As String 
	Dim regpath_run As String 
	Dim strOldValue As String 
	
	keyname = ""
	
	regpath = Key + "\Shell\Open\Command"
	strOldValue = ReadRegistry(HKEY_CLASSES_ROOT, regpath, keyname)
	
	'add run command if required
	If Command(1) = "-run" Or  Command(2) = "-run" Then
		regpath_run = Key + "\Shell\Run\Command"
		WriteRegistry (HKEY_CLASSES_ROOT, regpath_run, keyname, ValXString, strOldValue)
	EndIf
	
	'backup old value
	If ReadRegistry(HKEY_CLASSES_ROOT, regpath, "OldValue") = "" Then
		WriteRegistry (HKEY_CLASSES_ROOT, regpath, "OldValue", ValXString, strOldValue)	
	EndIf
	
	'fix association, open script in Notepad
	WriteRegistry (HKEY_CLASSES_ROOT, regpath, keyname, ValXString, "%SystemRoot%\Notepad.exe ""%1"" %*")
	'WriteRegistry (HKEY_CLASSES_ROOT, regpath, keyname, ValXString, "")
	
End Sub

'Main sub

'%SystemRoot%\System32\WScript.exe "%1" %*

If Command(1) = "-h" Or  Command(1) = "-help" Then
	Print "UnRamson 1.0 - Disables standard behaviour for scripting languages in Windows, "
	Print "               fixes black hole left by M$. Disables script execution "
	Print "               by doubleclick for VBS, JS, VBE, WSF, JAR used by hackers."
	Print "Author:        Anthonius, Anton Pitak, SOFTPAE.com"
	Print ""
	Print "USAGE:"
	Print ""
	Print "unramson [-h|-help] [-run] [-quiet]"
	Print ""
	Print "-h|-help   show this help"
	Print "-run       add Run command to right menu that replaces Open command (doubleclick behaviour)"
	Print "-quiet     disables waiting for key press at the end, suitable for logon scripts"
	Print ""
	'Print "Press Any Key To Exit..."
	'Sleep
	
	GoTo myend
EndIf

FixRegSettings("VBSFile")
FixRegSettings("JSFile")
FixRegSettings("WSFFile")
FixRegSettings("VBEFile")
FixRegSettings("jarfile")

' use -quiet arg for logon script
If Command(1) = "-quiet" Or  Command(2) = "-quiet" Then
	Print "Done!"
Else
	Print "Done! Press Any Key..."
	Sleep
EndIf

myend:

End
