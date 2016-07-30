#include-once
#include <String.au3>
#include <File.au3>
#include <_Zip.au3>

; #INDEX# =======================================================================================================================
; Title .........: RemoteGmail
; Version........: 1.4
; AutoIt Version : 3.3.7.20++
; Language ......: English
; Author(s)......: Phoenix XL
; Librarie(s)....: _Zip, File and String
; Description ...: Functions for Executing Scripts or Single Functions from using Gmail
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
;	 RemoteGmail_Startup
;	 RemoteGmail
;	 EnableDebugging
;	_CheckGmail
;	 StringBetween
;	_GetMessageIds
;	 CallEx
;	 Zipit
;	 SendEmail2Gmail
;	 GetfromDirectLink
; ===============================================================================================================================

; #INTERNAL_USE_ONLY# ===========================================================================================================
;	_LogWrite
;	_ExecuteFromEmail
;	_ConvertTo
;	_INIgetKey
;	_AddToVar
;	_CompareID
;	_AssignID
;	_Exit
;	_INetSmtpMailCom
;	 MyErrFunc
; ===============================================================================================================================

; #VARIABLES# ===================================================================================================================
Global $__rg__sPassword = ''	;Account Password
Global $__rg__sUsername = ''	;Account Username
Global $__rg__sEmail = ''		;Complete Email ID

;Global Variables
Global $__rg__sIniFile = 'Message_id.ini', $__rg__sSectionName = "ExecutedMessage ID's"
Global $__rg__pMessage_Array[1] , $__rg__hDebug_File = FileOpen('Debug.log', 1 + 8)
Global $__rg__sOpenTag = 'PXL', $__rg__sCloseTag = '/PXL', $__rg__sDebug_ed = True
Global $__rg__iCount = 0, $__rg__fsSendRet = False

Global $__rg__oMyRet[2]
Global $__rg__oMyError = ObjEvent("AutoIt.Error", "MyErrFunc")

;Local Variable :P
Global $__rg__pMessageID = _GetMessageIds($__rg__sIniFile)
; ===============================================================================================================================

If IsArray($__rg__pMessageID) Then
	;If INI Existed and Returned
	;Add it to the Global Variable
	_AssignID($__rg__pMessage_Array, $__rg__pMessageID)
Else
	;Else set the first item to 0
	$__rg__pMessage_Array[0] = 0
EndIf
;Close File Handle Upon Exit
OnAutoItExitRegister('_exit');


; #FUNCTION# ====================================================================================================================
; Name ..........: RemoteGmail
; Description ...: The Main Function to check the Email and Execute the Function or Script
; Syntax ........: RemoteGmail()
; Parameters ....:
;
; Return values .: Success : Returns a positive integer
;				   Failure : Returns -1 and sets @error to
;					|1 - Error in getting from gmail and @extended is set to
;						|1 - Couldn't get new emails for some reason
;						|2 - No Username
;						|3 - No Password
;						|4 - Error Getting URL Source!
;						|5 - No response?
;						|6 - Unauthorized access, possibly a wrong username or password!
;
;					|2 - Error from the Function _ExecuteFromEmail, @extended is set to the error code of the same function.
;
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func RemoteGmail()
	;Check the Main for any New Mails
	Local $eReturn = _CheckGmail($__rg__sUsername, $__rg__sPassword, 'Phoenix XL')
	Local $Emails = @extended
	If @error Then Return SetError(1, @error, -1)
	For $x = 1 To $Emails
		If $eReturn[$x][0] = $__rg__sEmail Then
			;Check if the Email is from the Same Address
			If Not _CompareID($eReturn[$x][2]) Then
				If IsArray($__rg__pMessage_Array) Then
					$__rg__iCount = $__rg__pMessage_Array[0]
				Else
					$__rg__iCount += 1
				EndIf
				If IniWrite($__rg__sIniFile, $__rg__sSectionName, $__rg__iCount, $eReturn[$x][2]) Then _
				_AddToVar($__rg__pMessage_Array, $eReturn[$x][2])	;Add the New Message ID
				;Execute the Function or Script [after downloading]
				_ExecuteFromEmail($eReturn[$x][1])
				If @error Then Return SetError(2, @error, -1)
			EndIf
		EndIf
	Next

EndFunc   ;==>RemoteGmail

; #FUNCTION# ====================================================================================================================
; Name ..........: EnableDebugging
; Description ...: Enable Debugging
; Syntax ........: EnableDebugging([$Logging = True[, $Replying = True]])
; Parameters ....: 	$Logging         	- log to a local log file
;                 	$Replying           - send an Email for debugging
;
; Return values .: void
;
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func EnableDebugging($Replying = False, $Logging = True)
	If Not IsBool($Logging) And IsBool($Replying) Then Return SetError(1, 0, -1)
	$__rg__sDebug_ed = $Logging
	$__rg__fsSendRet = $Replying
EndFunc   ;==>EnableDebugging

; #FUNCTION# ====================================================================================================================
; Name ..........: _CheckGmail
; Description ...: Checks a Google Email for new emails.
; Syntax ........: _CheckGmail($UserName, $Pswd[, $UserAgentString = ""])
; Parameters ....: 	$UserName           - Username [not email]
;                 	$Pswd             	- Password of the gmail account
;                 	$UserAgentString   	- [optional] An unknown value. Default is "".
; Return values .: A 2d array with Email information as follows~
;                   1|Email
;                   2|Summary
;                   3|Message ID
;
; Author ........: dantay9
; Modified ......: THAT1ANONYMOUSDUDE, Phoenix XL
; Remarks .......:
; Related .......:
; Link ..........: http://www.autoitscript.com/forum/topic/...-checker/page__view__findpost_
; Example .......:
; ===============================================================================================================================
Func _CheckGmail($UserName, $Pswd, $UserAgentString = "") ;~ From the Forum and Modified by Phoenix XL
	If Not $UserName Then Return SetError(2, 0, 0)
	If Not $Pswd Then Return SetError(3, 0, 0)
	If $UserAgentString Then HttpSetUserAgent($UserAgentString)
	Local $source = InetRead("https://" & $UserName & ":" & $Pswd & "@gmail.google.com/gmail/feed/atom", 1)
	If @error Then
		;;ConsoleWrite("!>Error Getting URL Source!" & @CR & "     404>@Error =" & @error & @CR & "    404>@Extended =" & @extended & @CR)
		Return SetError(4, 0, 0)
	EndIf
	If $source Then
		$source = BinaryToString($source)
	Else
		Return SetError(5, 0, 0)
	EndIf
	If StringLeft(StringStripWS($source, 8), 46) == "<HTML><HEAD><TITLE>Unauthorized</TITLE></HEAD>" Then Return SetError(6, 0, 0)
	If Not Number(StringBetween($source, "<fullcount>", "</fullcount>")) Then Return SetError(0, 0, 0)
	Local $Email = _StringBetween($source, "<entry>", "</entry>")
	If @error Then Return SetError(1, 0, 0)
	Local $Count = UBound($Email)
	Local $Datum[$Count + 1][3]
	$Datum[0][0] = StringBetween($source, "<title>", "</title>")
	$Datum[0][1] = StringBetween($source, "<tagline>", "</tagline>")
	For $i = 0 To $Count - 1
		$Datum[$i + 1][0] = StringBetween($Email[$i], "<email>", "</email>")
		$Datum[$i + 1][1] = StringBetween($Email[$i], "<summary>", "</summary>")
		$Datum[$i + 1][2] = StringMid($Email[$i], StringInStr($Email[$i], 'message_id=', 2) + 11, 16)
	Next
	Return SetError(0, $Count, $Datum)
EndFunc   ;==>_CheckGmail

; #FUNCTION# ====================================================================================================================
; Name ..........: StringBetween
; Description ...: Returns the first string
; Syntax ........: StringBetween($Str, $S, $E)
; Parameters ....: 	$Str          		- the string
;                 	$S	            	- the first param to start the search
;                 	$E				  	- the last param to end the search
; Return values .: Returns the first string
;
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......: Similar to _StringBetween
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func StringBetween($Str, $S, $E)
	;Helper Function
	Local $B = _StringBetween($Str, $S, $E)
	If @error Then Return SetError(1, 0, 0)
	Return SetError(0, 0, $B[0])
EndFunc   ;==>StringBetween

; #FUNCTION# ====================================================================================================================
; Name ..........: _GetMessageIds
; Description ...: Get the message id from an INI file
; Syntax ........: _GetMessageIds($sINI)
; Parameters ....: 	$sINI         		- the Ini File
; Return values .: Returns an array, see IniReadSection
;
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _GetMessageIds($sINI)
	;Get the stored message ids from the INI
	If Not FileExists($sINI) Then Return 0
	Return IniReadSection($sINI, $__rg__sSectionName)
EndFunc   ;==>_GetMessageIds


; #FUNCTION# ====================================================================================================================
; Name ..........: RemoteGmail_Startup
; Description ...: Set the Username and the Password
; Syntax ........: RemoteGmail_Startup($Username, $Password)
; Parameters ....: $Username            - Username[ not email] of the Gmail Account.
;                  $Password            - Account Password.
; Return values .: Success - Returns 1
;				   Failure - Returns 0
;
; Author ........: Pheonix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func RemoteGmail_Startup($UserName, $Password)
	If $UserName = '' Then Return SetError(1, 0, 0)
	If $Password = '' Then Return SetError(2, 0, 0)
	$__rg__sUsername = $UserName
	$__rg__sPassword = $Password
	$__rg__sEmail = $__rg__sUsername & '@gmail.com'
	Return 1
EndFunc   ;==>RemoteGmail_Startup

; #FUNCTION# ====================================================================================================================
; Name ..........: CallEx
; Description ...: Similar to Call
; Syntax ........: CallEx ($sFuncName, $sParameter [, $sParam = '|' [, $sConcat = '+' ]])
; Parameters ....: 	$sFuncName        		- the name of the Function to call
;					$sParameter				- the string representation[raw] of parameters
;					$sParam					- the string for parameter separator
;					$sConcat				- the string for concatenation
; Return values .: Success - 1
;				   Failure - 0
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func CallEx($sFuncName, $sParameter, $sParam = '|', $sConcat = '+') ;
	Local $sData[1] = [0], $eData[1] = [0], $eSplit, $sSplit, $xType, $sJoin
	;Get the Parameters
	$sSplit = StringSplit($sParameter, $sParam, 1)
	For $i = 1 To $sSplit[0]
		;Get the Concatenation in the Parameters
		$eSplit = StringSplit($sSplit[$i], $sConcat, 1)
		If Not @error Then ;If concatenation is present
			For $n = 1 To $eSplit[0]
				$xType = StringMid($eSplit[$n], 1, 3)
				If Not StringLeft($xType, 1) == '(' Or Not StringRight($xType, 1) == ')' Then Return SetError(1, 0, -1)
				$eData[0] += 1
				ReDim $eData[$eData[0] + 1]
				;Get the Values from its Types
				$eData[$eData[0]] = _ConvertTo(StringMid($eSplit[$n], 3 + 1), $xType)
				If $n <> $eSplit[0] Then $eData[$eData[0]] &= '&'
				$sJoin &= $eData[$eData[0]]
			Next
			$sData[0] += 1
			ReDim $sData[$sData[0] + 1]
			$sData[$sData[0]] = $sJoin
		Else ;If concatenation is not present
			;Add new Parameters
			$xType = StringMid($sSplit[$i], 1, 3)
			If Not StringLeft($xType, 1) == '(' Or Not StringRight($xType, 1) == ')' Then Return SetError(1, 0, -1)
			$sData[0] += 1
			ReDim $sData[$sData[0] + 1]
			$sData[$sData[0]] = _ConvertTo(StringMid($sSplit[$i], 3 + 1), $xType)
		EndIf
	Next
	;Join the Parameters
	Local $sString = ''
	For $i = 1 To $sData[0]
		$sString &= $sData[$i]
		If $i <> $sData[0] Then $sString &= ','
	Next
	;Finally Execute it
	$sReturn = Execute($sFuncName & '(' & $sString & ')')
	_LogWrite(@HOUR & ':' & @MIN & ':' & @SEC & ' ' & @MDAY & '\' & @MON & '\' & @YEAR & _
			' Your Commands Executed : ' & $sFuncName & '(' & $sString & ')' & @TAB & _
			'Return Values : ' & $sReturn & @CRLF)
	Return 1
EndFunc   ;==>CallEx

; #FUNCTION# ====================================================================================================================
; Name ..........: GetfromDirectLink
; Description ...: Download file from direct link
; Syntax ........: GetfromDirectLink($nUrl[, $sLocal = Default])
;
; Parameters ....: $nUrl - A Direct URL to the Script [recommended DropBox]
;				   $slocal - The file to be saved
;
; Return values .: Success - Filename of the Downloaded file
;				   Failure - returns -1
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func GetfromDirectLink($nUrl, $sLocal = Default)
	If IsKeyword($sLocal) Or $sLocal = -1 Then $sLocal = @ScriptDir & "\Temp.au3"
	Local $hDownload = InetGet($nUrl, $sLocal, 3, 1)
	Do
		Sleep(250)
		If @error Then Return SetError(3, @error, -1)
	Until InetGetInfo($hDownload, 2) ; Check if the download is complete.
	InetClose($hDownload) ; Close the handle to release resources.
	If @error Then Return SetError(1, @error, -1)
	Return SetError(0, 0, $sLocal)
EndFunc   ;==>GetfromDirectLink

; #FUNCTION# ====================================================================================================================
; Name ..........: SendEmail2Gmail
; Description ...: Send Email to Gmail
; Syntax ........: SendEmail2Gmail($Body[,$AttachFiles=''[,$nZip=False]])
;
; Parameters ....: 	$Body 			- The Body of the Email
;					$AttachFiles	- The file(s) separated by semi-colon
;					$nZip			- Zip the File [Boolean value]
;					$nSubject		- The Subject of the Email
; Return values .: Failure - Returns -1
;				   Success - Returns 1
;
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func SendEmail2Gmail($Body, $AttachFiles = '', $nZip = False, $nSubject = 'Report')
	If Not ($AttachFiles = '') And $nZip = True Then $AttachFiles = Zipit($AttachFiles, $Body)
	_INetSmtpMailCom("smtp.gmail.com", $__rg__sUsername, $__rg__sEmail, $__rg__sEmail, $nSubject & " - Phoenix XL", _
			$Body, $AttachFiles, '', '', "Normal", $__rg__sUsername, $__rg__sPassword, 465, 1)
	If @error Then Return SetError(1, @error, -1)
	If $nZip Then FileDelete($AttachFiles)
	If @error Then Return SetError(2, @error, -1)
	Return 1
EndFunc   ;==>SendEmail2Gmail

; #FUNCTION# ====================================================================================================================
; Name ..........: Zipit
; Description ...: Zip Multiple Files
; Syntax ........: Zipit($sFileNames,ByRef $sText)
;
; Parameters ....: 	$sFileNames		- The file(s) separated by semi-colon
;					$sText			- Appends error value upon encountering
;
; Return values .: Failure - Returns -1
;				   Success - Returns the Filename of the ZipFile
;
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func Zipit($sFileNames, ByRef $sText)
	;If Blank then Return
	If $sFileNames = '' Then Return
	;Get the filenames if multiple
	$nSplit = StringSplit($sFileNames, ';', 1)
	;Create the Zip
	$iZip = _Zip_Create(@ScriptDir & '\Files.zip', 1)
	For $n = 1 To $nSplit[0]
		;Add Items
		_Zip_AddItem($iZip, $nSplit[$n])
		;Add an error message, if occured, to body of the Email
		If @error Then
			$sText &= @CRLF & ' | Error in Zipping "' & $nSplit[$n] & '"'
			Return -1
		EndIf
	Next
	Return $iZip
EndFunc   ;==>Zipit

; #FUNCTION# ====================================================================================================================
; Name ..........: _AssignID
; Description ...: Assigns value of Message-Id to a Global Var
; Syntax ........: _AssignID(ByRef $sVar, $__rg__pMessageID)
; Parameters ....: 	$sVar        		- the global var
;					$__rg__pMessageID			- the array of values
; Return values .: void
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AssignID(ByRef $sVar, $__rg__pMessageID)
	If Not IsArray($__rg__pMessageID) Then Return 0
	For $i = 1 To $__rg__pMessageID[0][0]
		;Assigns message ids to the database
		_AddToVar($sVar, $__rg__pMessageID[$i][1])
	Next
EndFunc   ;==>_AssignID

; #FUNCTION# ====================================================================================================================
; Name ..........: _CompareID
; Description ...: Match the id from the array of id(s)
; Syntax ........: _CompareID($sID)
; Parameters ....: $sID    		- the message id to compare
; Return values .: Match - 1
;				   NoMatch - 0
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _CompareID($sID)
	If Not IsArray($__rg__pMessageID) Then Return 0
	For $i = 1 To $__rg__pMessage_Array[0]
		;The Value matches an existing value
		If $__rg__pMessage_Array[$i] == $sID Then Return 1
	Next
	;not Matched
	Return 0
EndFunc   ;==>_CompareID

; #FUNCTION# ====================================================================================================================
; Name ..........: _AddToVar
; Description ...: Similar to _ArrayAdd
; Syntax ........: _AddToVar(ByRef $__rg__pMessage_Array, $__rg__pMessageID)
; Parameters ....: 	$__rg__pMessage_Array   		- the parent array
;					$__rg__pMessageID				- the value to add
; Return values .: Success - 1
;				   Failure - 0
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _AddToVar(ByRef $__rg__pMessage_Array, $__rg__pMessageID)
	If Not IsArray($__rg__pMessage_Array) Then Return 0
	$__rg__pMessage_Array[0] += 1
	ReDim $__rg__pMessage_Array[$__rg__pMessage_Array[0] + 1]
	$__rg__pMessage_Array[$__rg__pMessage_Array[0]] = $__rg__pMessageID
	Return 1
EndFunc   ;==>_AddToVar

; #FUNCTION# ====================================================================================================================
; Name ..........: _INIgetKey
; Description ...: Get IniKey by Ini Value
; Syntax ........: _INIgetKey($sValue)
; Parameters ....: 	$sValue     		- the value to check
; Return values .: Success - 1
;				   Failure - -1
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _INIgetKey($sValue)
	Local $x = IniReadSection($__rg__sIniFile, "$__rg__sSectionName")
	For $i = 1 To $x[0][0]
		If $x[$i][1] = $sValue Then Return $x[$i][0]
	Next
	Return -1
EndFunc   ;==>_INIgetKey


; #FUNCTION# ====================================================================================================================
; Name ..........: _ConvertTo
; Description ...: Get the value of an expression
; Syntax ........: _ConvertTo($sData, $sTo)
; Parameters ....: 	$sData       		- the data to convert
;					$sTo				- the conversion type enclosed in curved brackets
;						| (s) - String
;						| (n) - Number
;						| (f) - Floating Number
;						| (i) - Integer
;						| (b) - Binary
;						| (v) - Variable
;						| (h) - Hex
;						| (m) - macro
;						| (w) - Handle
;						| (p) - Pointer
; Return values .: Success - Returns the converted data
;				   Failure - Returns -1
;
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _ConvertTo($sData, $sTo)
	Local $nTo = StringBetween($sTo, '(', ')')
	Local $sReturn = ''
	Switch StringLeft($nTo, 1)
		Case 'n', 'f'
			$sReturn = Number($sData)
		Case 'i'
			$sReturn = Int($sData)
		Case 's'
			$sReturn = '"' & String($sData) & '"'
		Case 'b'
			$sReturn = Binary($sData)
		Case 'v'
			$sReturn = Eval($sData)
			If StringLeft(VarGetType(Eval($sData)), 1) = 's' Then $sReturn = '"' & $sReturn & '"'
		Case 'h'
			$sReturn = Hex(Binary($sData))
		Case 'm'
			If StringLeft($sData, 1) <> '@' Then $sData = '@' & $sData
			$sReturn = '"' & Execute($sData) & '"'
		Case 'w'
			$sReturn = HWnd($sData)
		Case 'p'
			$sReturn = Ptr($sData)
		Case Else
			$sReturn = SetError(1, $nTo, -1)
	EndSwitch
	Return SetError(0, 0, $sReturn)
EndFunc   ;==>_ConvertTo

; #FUNCTION# ====================================================================================================================
; Name ..........: _ExecuteFromEmail
; Description ...: Execute the Function or Download the Script and Execute
; Syntax ........: _ExecuteFromEmail($sData)
; Parameters ....: 	$sData - the expression which conveys the information
; Return values .: Success - Returns non-negative integer
;				   Failure - Returns -1
;
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _ExecuteFromEmail($sData)
	If Not StringInStr($sData, $__rg__sCloseTag, 2) Then Return SetError(1, 0, -1);End Tags Were Not Found, the Entire Email is Required
	;Get the Data between the Tags
	Local $sExecute = StringBetween($sData, $__rg__sOpenTag, $__rg__sCloseTag)
	;ConsoleWrite($sExecute & @CR)
	Local $pID
	;Check if it is a http address
	If StringLeft($sExecute, 4) = 'http' Then
		;Split the Filename and the Type
		Local $nType = StringSplit($sExecute, '^', 1)
		If Not ($nType[0] = 2) Then Return SetError(2, 0, -1)
		;Assign to Local Vars
		$sExecute = $nType[1]
		$nType = $nType[2]
		If $nType <> 'zip' And $nType <> 'au3' Then Return SetError(3, 0, -1)
		;Get the File from the Link
		Local $nFile = GetfromDirectLink($sExecute, @ScriptDir & '\Temp' & Random(1, 1000, 1) & '.' & $nType)
		If @error Or Not FileExists($nFile) Or $nFile = -1 Then Return SetError(4, @error, -1)
		If StringRight($nFile, 3) = 'zip' Then ;Check for Zip File
			Local $nArrays = _Zip_ListAll($nFile, 0)
			For $i = 1 To $nArrays[0]
				If StringRight($nArrays[$i], 3) = 'au3' Then
					;Unzip and then Execute the Autoit Scripts
					_Zip_Unzip($nFile, $nArrays[$i], @ScriptDir)
					$pID = Run(@AutoItExe & ' /AutoIt3ExecuteScript "' & @ScriptDir & '\' & $nArrays[$i] & '"')
					If @error Then Return SetError(5, @error, -1)
					;Write to Log and Send Email
					_LogWrite(@HOUR & ':' & @MIN & ':' & @SEC & ' ' & @MDAY & '\' & @MON & '\' & @YDAY & _
							' Script Executed : ' & @ScriptDir & '\' & $nArrays[$i] & @TAB & _
							'PID : ' & $pID & @CRLF)
				EndIf
			Next
			Return 1
		ElseIf StringRight($nFile, 3) = 'au3' Then ;Check for Autoit3 File
			;Execute using Autoit
			$pID = Run(@AutoItExe & ' /AutoIt3ExecuteScript "' & $nFile & '"')
			If @error Then Return SetError(6, @error, -1)
			;Write to Log
			_LogWrite(@HOUR & ':' & @MIN & ':' & @SEC & ' ' & @MDAY & '\' & @MON & '\' & @YDAY & _
					' Script Executed : ' & $nFile & @TAB & _
					'PID : ' & $pID & @CRLF)
			Return 1
		Else
			Return SetError(7, 0, -1)
		EndIf
	Else ;or a Single Function Execution
		;Get Function Name and the Parameters
		$sExecute = StringSplit($sExecute, ':', 1)
		;Split FuncName and Execute it
		If IsArray($sExecute) And UBound($sExecute) - 1 = 2 Then
			Return SetError(8, 0, CallEx($sExecute[1], $sExecute[2]))
		Else
			Return SetError(9, 0, -1)
		EndIf
	EndIf
EndFunc   ;==>_ExecuteFromEmail

; #FUNCTION# ====================================================================================================================
; Name ..........: _LogWrite
; Description ...: Write to log file
; Syntax ........: _LogWrite($sData)
; Parameters ....: 	$sData - the expression
; Return values .: Returns 0 [always]
;
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _LogWrite($sData)
	If $__rg__sDebug_ed Then FileWriteLine($__rg__hDebug_File, $sData)
	If $__rg__fsSendRet Then SendEmail2Gmail($sData)
	Return 0
EndFunc   ;==>_LogWrite

; #FUNCTION# ====================================================================================================================
; Name ..........: _Exit
; Description ...: Close the file handle upon exitting
; Syntax ........: _Exit()
; Parameters ....:
; Return values .:
;
; Author ........: Phoenix XL
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func _Exit()
	FileClose($__rg__hDebug_File)
EndFunc   ;==>_Exit


;by Jos
;Link - http://www.autoitscript.com/forum/topic/23860-smtp-mailer-that-supports-html-and-attachments/page__hl__%20smtp%20%20mailer%20%20attachments
Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Importance = "Normal", $s_Username = "", $s_Password = "", $IPPort = 25, $ssl = 0)
	Local $objEmail = ObjCreate("CDO.Message")
	$objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
	$objEmail.To = $s_ToAddress
	If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
	If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress
	$objEmail.Subject = $s_Subject
	If $s_AttachFiles <> "" Then
		Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
		For $x = 1 To $S_Files2Attach[0]
			$S_Files2Attach[$x] = _PathFull($S_Files2Attach[$x])
;~          ConsoleWrite('@@ Debug : $S_Files2Attach[$x] = ' & $S_Files2Attach[$x] & @LF & '>Error code: ' & @error & @LF) ;### Debug Console
			If FileExists($S_Files2Attach[$x]) Then
				ConsoleWrite('+> File attachment added: ' & $S_Files2Attach[$x] & @LF)
				$objEmail.AddAttachment($S_Files2Attach[$x])
			Else
				;ConsoleWrite('!> File not found to attach: ' & $S_Files2Attach[$x] & @LF)
				;SetError(1)
				;Return 0
				$as_Body &= @CR & ' | Error - The Attachment "' & $S_Files2Attach[$x] & '" was not found'
			EndIf
		Next
	EndIf
	If StringInStr($as_Body, "<") And StringInStr($as_Body, ">") Then
		$objEmail.HTMLBody = $as_Body
	Else
		$objEmail.Textbody = $as_Body & @CRLF
	EndIf
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
	If Number($IPPort) = 0 Then $IPPort = 25
	$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort
	;Authenticated SMTP
	If $s_Username <> "" Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
	EndIf
	If $ssl Then
		$objEmail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
	EndIf
	;Update settings
	$objEmail.Configuration.Fields.Update
	; Set email Importance
	Switch $s_Importance
		Case "High"
			$objEmail.Fields.Item("urn:schemas:mailheader:Importance") = "High"
		Case "Normal"
			$objEmail.Fields.Item("urn:schemas:mailheader:Importance") = "Normal"
		Case "Low"
			$objEmail.Fields.Item("urn:schemas:mailheader:Importance") = "Low"
	EndSwitch
	$objEmail.Fields.Update
	; Sent the Message
	$objEmail.Send
	If @error Then
		SetError(2)
		Return $__rg__oMyRet[1]
	EndIf
	$objEmail = ""
EndFunc   ;==>_INetSmtpMailCom

;by Jos
;Link - http://www.autoitscript.com/forum/topic/23860-smtp-mailer-that-supports-html-and-attachments/page__hl__%20smtp%20%20mailer%20%20attachments
; Com Error Handler
Func MyErrFunc()
	$HexNumber = Hex($__rg__oMyError.number, 8)
	$__rg__oMyRet[0] = $HexNumber
	$__rg__oMyRet[1] = StringStripWS($__rg__oMyError.description, 3)
	ConsoleWrite("### COM Error !  Number: " & $HexNumber & "   ScriptLine: " & $__rg__oMyError.scriptline & "   Description:" & $__rg__oMyRet[1] & @LF)
	SetError(1); something to check for when this function returns
	Return
EndFunc   ;==>MyErrFunc
