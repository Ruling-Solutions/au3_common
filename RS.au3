#include-once

Opt('MustDeclareVars', 1)

#include <APIConstants.au3>
#include <Array.au3>
#include <Crypt.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <Math.au3>
#include <ProcessConstants.au3>
#include <String.au3>
#include <WinAPIProc.au3>
#include <WinAPISys.au3>
#include <WinAPIShellEx.au3>
#include 'RS_Environ.au3'
#include 'RS_Encoding.au3'

Local Const $RS_FirefoxWindow = 'Mozilla Firefox'
Local $ASCIIChars

Local $Numbers = '0123456789,.'

; Creates ASCII characters string for escaping routines
For $i = 160 To 255
	$ASCIIChars &= Chr($i)
Next

; <== ARRAY PROCEDURES ============================================================================
; <=== RS_arrayToFile =============================================================================
; RS_arrayToFile(String, String)
; ; Write a array to a file.
; ;
; ; @param  String          Filename.
; ; @param  String          Data array.
; ; @return Boolean         True if successfull, 0 otherwise.
Func RS_arrayToFile($pFilePath, $pText)
	Local $hFileOpen = FileOpen($pFilePath, 2)
	If $hFileOpen = -1 Then
		MsgBox(16, 'Error', 'Error opening file.')
		Return False
	EndIf

	If Not FileWrite($hFileOpen, $pText) Then
		MsgBox(16, 'Error', 'Error writing file.')
		Return False
	EndIf

	FileClose($hFileOpen)
	Return True
EndFunc
; ============================================================================= RS_arrayToFile ===>
; ============================================================================ ARRAY PROCEDURES ==>

; <== BOOLEAN PROCEDURES ==========================================================================
; <=== RS_boolInteger =============================================================================
; RS_boolInteger(Boolean)
; ; Returns a integer value from a boolean.
; ;
; ; @param  Boolean         Boolean expression.
; ; @return Integer         1 if Boolean is True, 0 otherwise.
Func RS_boolInteger($pValue)
	Return ($pValue) ? 1 : 0
EndFunc
; ============================================================================= RS_boolInteger ===>
; <=== RS_intBoolean ==============================================================================
; RS_intBoolean(Integer)
; ; Returns a boolean value from a integer.
; ;
; ; @param  Integer         Integer expression.
; ; @return Boolean         True if Integer is 1, False otherwise.
Func RS_intBoolean($pValue)
	Return ($pValue = 1) ? True : False
EndFunc
; ============================================================================== RS_intBoolean ===>
; <=== RS_strBoolean ==============================================================================
; RS_strBoolean(String)
; ; Returns a boolean value from a literal string.
; ;
; ; @param  String          String expression.
; ; @return Boolean         True if String is 'true', False otherwise.
Func RS_strBoolean($pValue)
	Return (StringLower($pValue) = 'true') ? True : False
EndFunc
; ============================================================================== RS_strBoolean ===>
; ========================================================================== BOOLEAN PROCEDURES ==>

; <== COLOR PROCEDURES ============================================================================
; <== RS_hexToRGBA() ==============================================================================
; RS_hexToRGBA(String)
; ; Converts color hex values to RGBA color.
; ; @param  String          String with ARGB hex value.
; ; @return Integer()       Array with RGBA channels.
; ; @access Public
Func RS_hexToRGBA($pHexColor = '')
	Local $intChannels[4]

	; Remove hex symbols and invalid characters
	$pHexColor = StringReplace($pHexColor, '#', '')
	$pHexColor = RS_TrimLeft($pHexColor, '0x', True)
	$pHexColor = StringRegExpReplace($pHexColor, '([^0-9A-Fa-f@-])', '')

	If StringLen($pHexColor) = 0 Then
		$intChannels[0] = 0
		$intChannels[1] = 0
		$intChannels[2] = 0
		$intChannels[3] = ''
	Else
		Switch StringLen($pHexColor)
			Case 1
				$intChannels[0] = Dec($pHexColor & $pHexColor)
				$intChannels[1] = $intChannels[0]
				$intChannels[2] = $intChannels[0]
				$intChannels[3] = ''
			Case 2
				$intChannels[0] = Dec($pHexColor)
				$intChannels[1] = $intChannels[0]
				$intChannels[2] = $intChannels[0]
				$intChannels[3] = ''
			Case 3
				$intChannels[0] = Dec(StringLeft($pHexColor, 1) & StringLeft($pHexColor, 1))
				$intChannels[1] = Dec(StringMid($pHexColor, 2, 1) & StringMid($pHexColor, 2, 1))
				$intChannels[2] = Dec(StringRight($pHexColor, 1) & StringRight($pHexColor, 1))
				$intChannels[3] = ''
			Case 4
				$intChannels[3] = Dec(StringLeft($pHexColor, 1) & StringLeft($pHexColor, 1))
				$intChannels[0] = Dec(StringMid($pHexColor, 2, 1) & StringMid($pHexColor, 2, 1))
				$intChannels[1] = Dec(StringMid($pHexColor, 3, 1) & StringMid($pHexColor, 3, 1))
				$intChannels[2] = Dec(StringRight($pHexColor, 1) & StringRight($pHexColor, 1))
			Case 6
				$intChannels[0] = Dec(StringLeft($pHexColor, 2))
				$intChannels[1] = Dec(StringMid($pHexColor, 3, 2))
				$intChannels[2] = Dec(StringRight($pHexColor, 2))
				$intChannels[3] = ''
			Case 8
				$intChannels[3] = Dec(StringLeft($pHexColor, 2))
				$intChannels[0] = Dec(StringMid($pHexColor, 3, 2))
				$intChannels[1] = Dec(StringMid($pHexColor, 5, 2))
				$intChannels[2] = Dec(StringRight($pHexColor, 2))
			Case Else
				$intChannels[0] = 0
				$intChannels[1] = 0
				$intChannels[2] = 0
				$intChannels[3] = ''
		EndSwitch
	EndIf

	Return $intChannels
EndFunc
; ============================================================================== RS_hexToRGBA() ==>
; <== RS_RGBAToHex() ==============================================================================
; RS_RGBAToHex(Integer())
; ; Converts RGBA color to hex value.
; ; @param  Integer()       Array with RGBA channels.
; ; @return String          String with ARGB hex value.
; ; @access Public
Func RS_RGBAToHex($pChannels)
	If Not IsArray($pChannels) Then Return ''

	Return (StringLen($pChannels[3]) = 0) ? _
					Hex($pChannels[0], 2) & Hex($pChannels[1], 2) & Hex($pChannels[2], 2) : _
					Hex($pChannels[3], 2) & Hex($pChannels[0], 2) & Hex($pChannels[1], 2) & Hex($pChannels[2], 2)
EndFunc
; =============================================================================== RS_RGBAToHex ===>
; ============================================================================ COLOR PROCEDURES ==>

; <== ENCRYPTION PROCEDURES =======================================================================
; <=== RS_AESdecrypt ==============================================================================
; RS_AESdecrypt(String, String)
; ; Decrypt text using AES 256 and SHA 512 key.
; ;
; ; @param  String          Text to decrypt.
; ; @param  String          Password.
; ; @return String          Decrypted text.
Func RS_AESdecrypt($pText, $pPassword)
  If StringLen($pText) = 0 Or StringLen($pPassword) = 0 Then Return ''
  Local $strKey = _Crypt_DeriveKey($pPassword, $CALG_AES_256, $CALG_SHA_512)
  Local $strText = BinaryToString(_Crypt_DecryptData($pText, $strKey, $CALG_USERKEY))
  _Crypt_DestroyKey($strKey)
  Return $strText
EndFunc
; ============================================================================== RS_AESdecrypt ===>
; <=== RS_AESencrypt ==============================================================================
; RS_AESencrypt(String, String)
; ; Encrypt text using AES 256 and SHA 512 key.
; ;
; ; @param  String          Text to encrypt.
; ; @param  String          Password.
; ; @return String          Encrypted text.
Func RS_AESencrypt($pText, $pPassword)
  If StringLen($pText) = 0 Or StringLen($pPassword) = 0 Then Return ''
  Local $strKey = _Crypt_DeriveKey($pPassword, $CALG_AES_256, $CALG_SHA_512)
  Local $strText = BinaryToString(_Crypt_EncryptData($pText, $strKey, $CALG_USERKEY))
  _Crypt_DestroyKey($strKey)
  Return $strText
EndFunc
; ============================================================================== RS_AESencrypt ===>
; ======================================================================= ENCRYPTION PROCEDURES ==>

; <== NUMBER PROCEDURES ===========================================================================
; <=== RS_constrain ===============================================================================
; RS_constrain(Number, Number, Number)
; ; Returns a number constrained between two limits.
; ;
; ; @param  Number          Number.
; ; @param  Number          Lower limit.
; ; @param  Number          Upper limit.
; ; @return Number          Constrained number.
Func RS_constrain($pNumber, $pMin, $pMax)
	Return RS_max(RS_min($pNumber, $pMin), $pMax)
EndFunc
; =============================================================================== RS_constrain ===>
; <=== RS_decimalPoint ============================================================================
; RS_decimalPoint(Number, [Boolean])
; ; Swap decimal point between dot or comma.
; ;
; ; @param  Number          Number.
; ; @param  [Boolean]       True: Comma as decimal point. False: Dot. Default: False (Dot).
; ; @return Number/String   Modified value.
Func RS_decimalPoint($pNumber, $pComma = False)
	If $pComma Then
		Return StringReplace(StringReplace(StringReplace($pNumber, ',', Chr(1), 0, 2), '.', ',', 0, 2), Chr(1), '.', 0, 2)
	Else
		Return Number(StringReplace(StringReplace(StringReplace($pNumber, '.', Chr(1), 0, 2), ',', '.', 0, 2), Chr(1), ',', 0, 2))
	EndIf
EndFunc
; ============================================================================ RS_decimalPoint ===>
; <=== RS_digitGroup ==============================================================================
; RS_digitGroup(Integer, [String])
; ; Returns a number with digit grouping.
; ;
; ; @param  Integer         Number.
; ; @param  [String]        Delimiter. Default: Space.
; ; @return String          Number with digit grouping.
Func RS_digitGroup($pNumber, $pDelimiter = ' ')
	Local $sNumber

	$pNumber = StringStripWS($pNumber, 3)
	If RS_isNumber($pNumber) And StringLen($pNumber) > 3 Then
		$sNumber = StringTrimRight($pNumber, 3) & $pDelimiter & StringRight($pNumber, 3)

		Do
			If Not StringInStr(StringLeft($sNumber, 4), $pDelimiter) Then
				$sNumber = StringLeft($sNumber, StringInStr($sNumber, $pDelimiter) - 4) & $pDelimiter & StringRight($sNumber, StringLen($sNumber) - StringInStr($sNumber, $pDelimiter) + 4)
			EndIf
		Until StringInStr(StringLeft($sNumber, 4), $pDelimiter)
		Return $sNumber
	Else
		Return $pNumber
	EndIf
EndFunc
; ============================================================================== RS_digitGroup ===>
; <=== RS_isNumber ================================================================================
; RS_isNumber(String, Boolean)
; ; Checks if string has valid number digits.
; ;
; ; @param  String          Integer expression.
; ; @param  Boolean         Allows negative numbers. Default: False.
; ; @return Boolean         True if string contains only number digits, False otherwise.
Func RS_isNumber($pValue, $Negative = False)
	Local $sNumbers = ($Negative) ? $Numbers & '-' : $Numbers
	If StringLen($pValue) = 0 Then Return False

	Local $bNumber = True
	For $i = 1 To StringLen($pValue)
		If StringInStr($sNumbers, StringMid($pValue, $i, 1)) = 0 Then
			$bNumber = False
			ExitLoop
		EndIf
	Next

	Return $bNumber
EndFunc
; ================================================================================ RS_isNumber ===>
; <=== RS_max =====================================================================================
; RS_max(Number, Number, Number)
; ; Returns a number constrained to upper limit.
; ;
; ; @param  Number          Number.
; ; @param  Number          Upper limit.
; ; @return Number          Constrained number.
Func RS_max($pNumber, $pMax)
;~   Return (RS_isNumber($pNumber) And RS_isNumber($pMax)) ? (($pNumber > $pMax) ? $pMax : $pNumber) : ''
	Return ($pNumber > $pMax) ? $pMax : $pNumber
EndFunc
; ===================================================================================== RS_max ===>
; <=== RS_min =====================================================================================
; RS_min(Number, Number, Number)
; ; Returns a number constrained to lower limit.
; ;
; ; @param  Number          Number.
; ; @param  Number          Lower limit.
; ; @return Number          Constrained number.
Func RS_min($pNumber, $pMin)
;~   Return (RS_isNumber($pNumber) And RS_isNumber($pMin)) ? (($pNumber < $pMin) ? $pMin : $pNumber) : ''
	Return ($pNumber < $pMin) ? $pMin : $pNumber
EndFunc
; ===================================================================================== RS_min ===>
; <=== RS_noNumberPos =============================================================================
; RS_noNumberPos(String)
; ; Returns first non-numeric character position.
; ;
; ; @param  String          String with text.
; ; @return String          Found position.
Func RS_noNumberPos($pText)
	Local $bFound = False
	For $i = 1 To StringLen($pText)
		If Not StringInStr('0123456789,.', StringMid($pText, $i, 1), 2) Then
			$bFound = True
			ExitLoop
		EndIf
	Next
	Return ($bFound) ? $i : 0
EndFunc
; ============================================================================= RS_noNumberPos ===>
; <=== RS_within ==================================================================================
; RS_within(Number, Number, Number)
; ; Check if a number is inside two limits.
; ;
; ; @param  Number          Number.
; ; @param  Number          Lower limit.
; ; @param  Number          Upper limit.
; ; @return Boolean         True if number is inside limits.
Func RS_within($pNumber, $pMin, $pMax)
	If RS_isNumber($pNumber) And RS_isNumber($pMax) And RS_isNumber($pMax) Then
		$pNumber = Number($pNumber)
		$pMin = Number($pMin)
		$pMax = Number($pMax)
		Return (($pNumber = $pMin) ? True : (_Min($pNumber, $pMin) = $pNumber ? False : True)) ? _
		(($pNumber = $pMax) ? True : (_Max($pNumber, $pMax) = $pNumber ? False : True)) : _
		False
	Else
		Return SetError(1, 0, False)
	EndIf
EndFunc
; ================================================================================== RS_within ===>
; =========================================================================== NUMBER PROCEDURES ==>

; <== STRING PROCEDURES ===========================================================================
; <=== RS_absoluteURL =============================================================================
; RS_absoluteURL(String, String)
; ; Returns absolute URL from a relative path.
; ;
; ; @param  String          Base URL.
; ; @param  String          Relative URL.
; ; @return String          Absolute URL.
Func RS_absoluteURL($pBaseURL, $pRelativeURL)
	Local $iPos
	Local $sScheme

	; Modify base and relative URLs
	$pBaseURL = RS_trim(StringReplace($pBaseURL, '\', '/'), '/')
	$pRelativeURL = StringStripWS(StringReplace($pRelativeURL, '\', '/'), 3)

	If StringLen($pRelativeURL) = 0 Or $pBaseURL = $pRelativeURL Then Return $pBaseURL
	If StringLeft($pRelativeURL, 7) = 'http://' Or StringLeft($pRelativeURL, 8) = 'https://' Then Return $pRelativeURL

	; Remove scheme and page data from base URL
	$iPos = StringInStr($pBaseURL, '://', 2)
	If $iPos = 0 Then
		$sScheme = 'http://'
	Else
		$sScheme = StringLeft($pBaseURL, $iPos + 2)
		$pBaseURL = StringMid($pBaseURL, $iPos + 3)
	EndIf
	$iPos = StringInStr($pBaseURL, '/', 2, -1)
	If StringInStr($pBaseURL, '.', 2, -1) > $iPos Then $pBaseURL = StringLeft($pBaseURL, $iPos - 1)
	If StringLen($pBaseURL) = 0 Then Return ''

	If StringLeft($pRelativeURL, 2) = '//' Then
		; Relative path points to scheme
		$pBaseURL = ''
		$pRelativeURL = StringMid($pRelativeURL, 3)
	ElseIf StringLeft($pRelativeURL, 1) = '/' Then
		; Relative path points to root
		$iPos = StringInStr($pBaseURL, '/', 2)
		If $iPos > 0 Then $pBaseURL = StringLeft($pBaseURL, $iPos - 1)
		$pRelativeURL = StringMid($pRelativeURL, 2)
	Else
		; Update folder pointers
		While 1
			If StringLeft($pRelativeURL, 2) = './' Or StringLeft($pRelativeURL, 3) = '../' Then
				$iPos = StringInStr($pBaseURL, '/', 2, -1)
				If $iPos > 0 Then
					$pBaseURL = StringLeft($pBaseURL, $iPos - 1)
					$pRelativeURL = StringMid($pRelativeURL, StringInStr($pRelativeURL, '/', 2) + 1)
				Else
					Return ''
				EndIf
			Else
				ExitLoop
			EndIf
		WEnd
	EndIf

	If StringLen($pBaseURL) > 0 Then $pBaseURL &= '/'

	Return $sScheme & $pBaseURL & $pRelativeURL
EndFunc
; ============================================================================= RS_absoluteURL ===>
; <=== RS_addDir ==================================================================================
; RS_addDir(String, [String])
; ; Add directory to a filename. If dir was not defined, add Script dir.
; ;
; ; @param  String          Filename.
; ; @param  [String]        Dir to add. Default: Script dir.
; ; @return String          Filename and directory.
Func RS_addDir($pFile, $pDir = Default)
	If $pDir = Default Or StringLen($pDir) = 0 Then $pDir = RS_addSlash(@ScriptDir)
	If StringLen($pFile) > 0 And StringInStr($pFile, '\', 2) = 0 Then $pFile = $pDir & $pFile
	Return $pFile
EndFunc
; ================================================================================== RS_addDir ===>
; <=== RS_addSlash ================================================================================
; RS_addSlash(String)
; ; Add trailing slash removing leading and trailing spaces and double slashes.
; ;
; ; @param  String          Path to verify.
; ; @return String          Path with trailing slash.
Func RS_addSlash($pPath)
  $pPath = StringStripWS($pPath, 3)
  If StringLen($pPath) = 0 Then Return ''
  While StringRight($pPath, 1) = '\'
    $pPath = StringTrimRight($pPath, 1)
  WEnd
	Return $pPath & '\'
EndFunc
; ================================================================================ RS_addSlash ===>
; <=== RS_case ====================================================================================
; RS_case(Integer)
; ; Change string case.
; ;
; ; @param  String          String.
; ; @param  Integer         Case. -1: Unchanged, 0: lower, 1: UPPER, 2: Title case, 3: Proper Case,
; ;                         4: iNVERTED. Default: Unchanged
; ; @return String          String with new casing.
Func RS_case($pString, $pCase = Default)
	If StringLen($pString) = 0 Or $pCase = Default Or $pCase < 0 Or $pCase > 4 Then Return $pString

	Switch $pCase
		Case 1
			$pString = StringUpper($pString)
		Case 2
      $pString = StringUpper(StringLeft($pString, 1)) & StringLower(StringMid($pString, 2))
		Case 3
      $pString = _StringTitleCase($pString)
    Case 4
      $pString = StringLower(StringLeft($pString, 1)) & StringUpper(StringMid($pString, 2))
		Case Else
			$pString = StringLower($pString)
  EndSwitch

  Return $pString
EndFunc
; ==================================================================================== RS_case ===>
; <=== RS_cmdLine =================================================================================
; RS_cmdLine([Boolean])
; ; Return command line parameters in a single line avoiding debug return in scripting mode.
; ;
; ; @param  [Boolean]       True to use $CmdLineRaw, False for parameters iteration. Default: True.
; ; @return String          Command line parameter.
Func RS_cmdLine($pRawMode = True)
	Local $sCmdLine = ''

	If $pRawMode Then
		Local $sCmdLine = $CmdLineRaw
		If Not @Compiled Then
			Local $iPos = StringInStr($sCmdLine, @ScriptName)
			If $iPos > 0 Then
				$sCmdLine = StringStripWS(StringMid($sCmdLine, $iPos + StringLen(@ScriptName) + 1), 3)
			EndIf
		EndIf
	Else
		If $CmdLine[0] > 0 Then
			For $i = 1 To $CmdLine[0]
				$sCmdLine &= $CmdLine[$i] & ' '
			Next
			$sCmdLine = StringTrimRight($sCmdLine, 1)
		EndIf
	EndIf

	Return $sCmdLine
EndFunc
; ================================================================================= RS_cmdLine ===>
; <=== RS_cmdLines() ==============================================================================
; RS_cmdLines([Boolean], [String])
; ; Returns array with options or filenames instructions from command line.
; ;
; ; @param  [Boolean]   	  True to get options. False to get filenames. Default: False.
; ; @param  [String]   	    Single character to define options in command line. Default: '-'.
; ; @return String()   	    Array with info retrieved.
; ; @access Public
Func RS_cmdLines($pOptions = False, $pOptionChar = Default)
	If $pOptionChar = Default or StringLen($pOptionChar) = 0 Then $pOptionChar = '-'
	$pOptionChar = StringLeft($pOptionChar, 1)

	Local $strParams[1]
	For $i = 1 To $CmdLine[0]
		If $pOptions Then
			If StringLeft($CmdLine[$i], 1) = $pOptionChar Then _ArrayAdd($strParams, $CmdLine[$i])
		Else
			If StringLeft($CmdLine[$i], 1) <> $pOptionChar Then _ArrayAdd($strParams, $CmdLine[$i])
		EndIf
	Next
	$strParams[0] = UBound($strParams) - 1
	Return $strParams
EndFunc
; ============================================================================== RS_cmdLines() ===>
; <=== RS_cmdOption ===============================================================================
; RS_cmdOption(String, [Boolean])
; ; Returns value or existence of an option from command line instructions.
; ;
; ; @param  String   	      Command line option.
; ; @param  [Boolean]   	  True to get option existence, False to get value. Default: False.
; ; @return String          Value retrieved.
; ; @access Public
Func RS_cmdOption($pOption, $pExistence = Default)
	If $pExistence = Default Or StringLen($pExistence) = 0 Then $pExistence = False
	If StringLen($pOption) = 0 Then Return $pExistence ? False : ''

	Local $strOption = RS_cmdLines(True, StringLeft($pOption, 1))
	If $strOption[0] = 0 Then
		Return $pExistence ? False : ''
	Else
		Local $intLen = StringLen($pOption)
		For $i = 1 To $strOption[0]
			If StringLen($strOption[$i]) = $intLen Then
				If $strOption[$i] = $pOption Then Return $pExistence ? True : $strOption[$i]
			Else
				If StringLeft($strOption[$i], $intLen + 1) = $pOption & '=' Then
					Return $pExistence ? True : StringMid($strOption[$i], $intLen + 2)
				EndIf
			EndIf
		Next
		Return $pExistence ? False : ''
	EndIf
EndFunc
; =============================================================================== RS_cmdOption ===>
; <=== RS_countChars ==============================================================================
; RS_countChars(String, String, [Boolean])
; ; Count all characters ocurrences in a string.
; ;
; ; @param  String     	    String.
; ; @param  String   	      Characters to count. Default: Space.
; ; @param  [Boolean]   	  Count as individual characters or as a entire string. Default: False.
; ; @return Integer         Ocurrences.
Func RS_countChars($pText, $pCharacters, $pEntire = False)
	; Set default values
	If StringLen($pText) = 0 Then Return ''
	If StringLen($pCharacters) = 0 Or $pCharacters = Default Then $pCharacters = ' '
	If $pEntire = Default Then $pEntire = False

	If $pEntire Then
		Local $sTmp = StringReplace($pText, $pCharacters, $pCharacters)
		Return @extended
	Else
		Local $iCount = 0
		For $sCharacter In StringSplit($pCharacters, '', 2)
			$iCount += RS_countChars($pText, $sCharacter, True)
		Next
		Return $iCount
	EndIf
EndFunc
; =============================================================================== RS_countChars ===>
; <=== RS_countTokens ==============================================================================
; RS_countTokens(String, [String], [Boolean])
; ; Count tokens in a string separated by a delimiter.
; ;
; ; @param  String     	    String.
; ; @param  [String] 	      Characters to use as delimiters (case sensitive). Default: Space.
; ; @param  [Boolean]   	  Count as individual characters or as a entire string. Default: False.
; ; @return Integer         Tokens count.
Func RS_countTokens($pText, $pDelimiters = ' ', $pEntire = False)
	; Set default values
	If StringLen($pText) = 0 Then Return ''
	If StringLen($pDelimiters) = 0 Or $pDelimiters = Default Then $pDelimiters = ' '
	If $pEntire = Default Then $pEntire = False

	Return RS_countChars(RS_stripChars($pText, $pDelimiters, $pEntire), $pDelimiters, $pEntire) + 1
EndFunc
; ============================================================================= RS_countTokens ===>
; <=== RS_CRLF ====================================================================================
; RS_CRLF(String)
; ; Replace @CR and @LF with @CRLF.
; ;
; ; @param  String          Text with carriage returns.
; ; @return String          Text with only @CRLF.
Func RS_CRLF($pText)
	Return StringRegExpReplace($pText, '\n|\r', @CRLF)
EndFunc
; ==================================================================================== RS_CRLF ===>
; <=== RS_escape ==================================================================================
; RS_escape(String, [Boolean])
; ; Replaces or restores escape characters from string.
; ;
; ; @param  String     	    Text string to process.
; ; @param  [Boolean]  	    False to replace or True to restore escape characters. Default: False.
; ; @return String     	    Text string processed.
Func RS_escape($pText, $pToAscii = False)
	Local $iPos = 1
	Local $sChar

	If StringLen($pText) = 0 Then Return ''

	If $pToAscii Then
		$pText = StringReplace($pText, '\', '\\', 0, 2)
		$pText = StringReplace($pText, @CR, '\r', 0, 2)
		$pText = StringReplace($pText, @CRLF, '\n', 0, 2)
		$pText = StringReplace($pText, Chr(9), '\t', 0, 2)

		While StringLen($pText) > $iPos
			$sChar = StringMid($pText, $iPos, 1)
			If StringInStr($ASCIIChars, $sChar, 1) > 0 Then $pText = StringReplace($pText, $sChar, '\x' & Hex(Asc($sChar), 2), 0, 1)
			$iPos += 1
		WEnd
	Else
		$pText = StringReplace($pText, '\\', Chr(1), 0, 2)
		$pText = StringReplace($pText, '\r', @CR, 0, 1)
		$pText = StringReplace($pText, '\n', @CRLF, 0, 1)
		$pText = StringReplace($pText, '\t', Chr(9), 0, 1)

		While 1
			$iPos = StringInStr($pText, '\x', 1)
			If $iPos = 0 Then
				ExitLoop
			Else
				$sChar = StringMid($pText, $iPos + 2, 2)
				$pText = StringReplace($pText, '\x' & $sChar, Chr(Dec($sChar)), 0, 1)
			EndIf
		WEnd
		$pText = StringReplace($pText, Chr(1), '\')
	EndIf

	Return $pText
EndFunc
; ================================================================================== RS_escape ===>
; <=== RS_fromHex =================================================================================
; RS_fromHex(String)
; ; Returns a character string from a HEX sequence.
; ;
; ; @param  String          HEX ASCII values of characters.
; ; @return String          Characters string.
Func RS_fromHex($pHEX)
  If StringLen($pHEX) = 0 Then Return ''
  Local $strText
  For $i = 1 To StringLen($pHEX) Step 2
    $strText &= Chr(Dec(StringMid($pHEX, $i, 2)))
  Next
  Return $strText
EndFunc
; ================================================================================= RS_fromHex ===>
; <=== RS_padLeft =================================================================================
; RS_padLeft(String, Integer, String)
; ; Left pads a string with specified length and padding character. If string is longer than
; ; length, it will be truncated.
; ;
; ; @param  String          Original string.
; ; @param  Integer         Final string length.
; ; @param  String          Padding character. Default: ' '.
; ; @return String          New string.
Func RS_padLeft($pString, $pLength, $pChar = ' ')
	; Set default values
	If StringLen($pChar) = 0 Then $pChar = ' '

	; Apply left padding and return result.
	Return StringRight(_StringRepeat($pChar, $pLength) & $pString, $pLength)
EndFunc
; ================================================================================= RS_padLeft ===>
; <=== RS_padRight ================================================================================
; RS_padRight(String, Integer, String)
; ; Right pads a string with specified length and padding character. If string is longer than
; ; length, it will be truncated.
; ;
; ; @param  String          Original string.
; ; @param  Integer         Final string length.
; ; @param  String          Padding character. Default: ' '.
; ; @return String          New string.
Func RS_padRight($pString, $pLength, $pChar = ' ')
	; Set default values
	If StringLen($pChar) = 0 Then $pChar = ' '

	; Apply right padding and return result.
	Return StringLeft($pString & _StringRepeat($pChar, $pLength), $pLength)
EndFunc
; ================================================================================ RS_padRight ===>
; <=== RS_quote ===================================================================================
; RS_quote(String)
; ; Add quotes if filename has spaces, or removes if doesn't.
; ;
; ; @param  String			    Filename.
; ; @return String  		    Filename with or without quotes.
; ; @access Public
Func RS_quote($pFileName)
	If StringLen($pFileName) = 0 Then Return ''

	$pFileName = RS_trim(RS_trim($pFileName, '"'), "'")
  MsgBox(0, '', $pFileName)
	If StringInStr($pFileName, ' ') > 0 Then $pFileName = '"' & $pFileName & '"'
  MsgBox(0, '', $pFileName)
	Return $pFileName
EndFunc
; =================================================================================== RS_quote ===>
; <=== RS_removeExt ===============================================================================
; RS_removeExt(String)
; ; Remove extension from a filename.
; ;
; ; @param  String			Filename.
; ; @return String  		Filename without extension.
; ; @author guinness
; ; @access Public
Func RS_removeExt($pFileName)
	Return StringRegExpReplace($pFileName, '\.[^.\\/]*$', '')
EndFunc
; =============================================================================== RS_removeExt ===>
; <=== RS_split() =================================================================================
; RS_split(String, [String], [Integer], [Integer])
; ; Splits up a string into substrings depending on given delimiters avoiding empty entries.
; ; Specific number of returned entries can be defined.
; ;
; ; @param  String     	    Original string.
; ; @param  [String]   	    Character(s) to use as delimiters (case sensitive). Default: Space.
; ; @param  [Integer]       Flags for StringSplit. Default: 0 ($STR_CHRSPLIT).
; ; @param  [Integer]       Number of entries returned. Default: -1 (All).
; ; @return String()   	    String array.
Func RS_split($pText, $pDelimiters = ' ', $pFlags = 0, $pCount = -1)
	; Set default values
	If StringLen($pText) = 0 Then Return Null
	If StringLen($pDelimiters) = 0 Or $pDelimiters = Default Then $pDelimiters = ' '
	If $pFlags = Default Then $pFlags = 0
	If $pCount = Default Then $pCount = -1

	Local $bWhole = BitAND($pFlags, 1) ? True : False
	Local $iStart = BitAND($pFlags, 2) ? 0 : 1
	Local $iIndex = $iStart
	Local $sText
	Local $sToken

	; Get total entries count and adjust $pCount parameter
	Local $iEntriesCount = RS_countTokens($pText, $pDelimiters, $bWhole)
	If $pCount < 0 Or $pCount > $iEntriesCount Then $pCount = $iEntriesCount

	; Check entries count requested
	Local $sEntries
	If $pCount = 1 Then
		Dim $sEntries[$iStart + 1]
		If $iStart = 1 Then $sEntries[0] = 1
		$sEntries[$iStart] = RS_trim($pText, $pDelimiters, $bWhole)
	Else
		$sEntries = StringSplit(RS_stripChars($pText, $pDelimiters, $bWhole), $pDelimiters, $pFlags)
		If $pCount < $iEntriesCount Then
			$pCount += $iStart
			$sText = $pText
			For $i = $iStart To UBound($sEntries) - 1
				$iIndex += 1
				$sText = RS_trimLeft($sText, $pDelimiters, $bWhole)
				$sToken = StringLeft($sText, StringLen($sEntries[$i]))
				$sText = StringTrimLeft($sText, StringLen($sEntries[$i]))
				If $iIndex + 1 = $pCount Then ExitLoop
			Next
			$sEntries[$pCount - 1] = RS_trimLeft($sText, $pDelimiters, $bWhole)
			ReDim $sEntries[$pCount]
			If $iStart = 1 Then $sEntries[0] = $pCount - 1
		EndIf
	EndIf

	Return $sEntries
EndFunc
; ================================================================================= RS_split() ===>
; <=== RS_stripChars ==============================================================================
; RS_stripChars(String, [String], [Boolean])
; ; Removes all leading, trailing and double (or more) characters in a string.
; ;
; ; @param  String     	    Original string.
; ; @param  [String]   	    Characters to remove. Default: Space.
; ; @param  [Boolean]   	  Remove as individual characters or as a entire string. Default: False.
; ; @return String     	    New string.
Func RS_stripChars($pText, $pCharacters = ' ', $pEntire = False)
	; Set default values
	If StringLen($pText) = 0 Then Return ''
	If StringLen($pCharacters) = 0 Or $pCharacters = Default Then $pCharacters = ' '
	If $pEntire = Default Then $pEntire = False

	If $pEntire Then
		If $pCharacters = ' ' Then
			$pText = StringStripWS($pText, 7)
		Else
			$pText = RS_trim($pText, $pCharacters, $pEntire)
			Do
				$pText = StringReplace($pText, $pCharacters & $pCharacters, $pCharacters)
			Until @extended = 0
		EndIf
	Else
		For $sCharacter In StringSplit($pCharacters, '', 2)
			$pText = RS_stripChars($pText, $sCharacter, True)
		Next
	EndIf

	Return $pText
EndFunc
; ============================================================================== RS_stripChars ===>
; <=== RS_toHex ===================================================================================
; RS_toHex(String)
; ; Returns a HEX ASCII values of characters from a character string.
; ;
; ; @param  String          Characters string.
; ; @return String          HEX ASCII values of characters..
Func RS_toHex($pText)
  If StringLen($pText) = 0 Then Return ''
  Local $strHex
  For $strChar In StringSplit($pText, '', 2)
    $strHex &= Hex(Asc($strChar), 2)
  Next
  Return $strHex
EndFunc
; =================================================================================== RS_toHex ===>
; <=== RS_trim ====================================================================================
; RS_trim(String, [String], [Boolean])
; ; Removes defined leading and trailing characters at both ends.
; ;
; ; @param  String     	    Original text.
; ; @param  [String]   	    Characters to trim. Default: Space.
; ; @param  [Boolean]   	  Trim as individual characters or as a entire string. Default: False.
; ; @return String     	    Processed text.
Func RS_trim($pText, $pCharacters = ' ', $pEntire = True)
	Return RS_trimLeft(RS_trimRight($pText, $pCharacters, $pEntire), $pCharacters, $pEntire)
EndFunc
; ==================================================================================== RS_trim ===>
; <=== RS_trimLeft ================================================================================
; RS_trimLeft(String, [String], [Boolean])
; ; Removes defined leading characters.
; ;
; ; @param  String     	    Original text.
; ; @param  [String]   	    Characters to trim. Default: Space.
; ; @param  [Boolean]   	  Trim as individual characters or as a entire string. Default: True.
; ; @return String     	    Processed text.
Func RS_trimLeft($pText, $pCharacters = Default, $pEntire = Default)
	; Set default values
	If StringLen($pText) = 0 Then Return ''
	If StringLen($pCharacters) = 0 Then Return $pText
	If $pCharacters = Default Then $pCharacters = ' '
	If $pEntire = Default Then $pEntire = True

	If $pEntire Then
		Local $intLen = StringLen($pCharacters)
		While StringLeft($pText, $intLen) = $pCharacters
			$pText = StringTrimLeft($pText, $intLen)
		WEnd
	Else
		Local $boolExit
		Local $strChars = StringSplit($pCharacters, '', 2)
		Do
			$boolExit = True
			For $strChar In $strChars
				If StringInStr($pText, $strChar) Then $boolExit = False
			Next
			If $boolExit Then ExitLoop

			For $strChar In $strChars
				$pText = RS_trimLeft($pText, $strChar, True)
			Next
			For $i = UBound($strChars) - 1 To 0 Step -1
				$pText = RS_trimLeft($pText, $strChars[$i], True)
			Next
		Until $boolExit
	EndIf

	Return $pText
EndFunc
; ================================================================================ RS_trimLeft ===>
; <=== RS_trimRight ===============================================================================
; RS_trimRight(String, [String], [Boolean])
; ; Removes defined trailing characters.
; ;
; ; @param  String     	    Original text.
; ; @param  [String]   	    Characters to trim. Default: Space.
; ; @param  [Boolean]   	  Trim as individual characters or as a entire string. Default: True.
; ; @return String     	    Processed text.
Func RS_trimRight($pText, $pCharacters = Default, $pEntire = Default)
	; Set default values
	If StringLen($pText) = 0 Then Return ''
	If StringLen($pCharacters) = 0 Then Return $pText
	If $pCharacters = Default Then $pCharacters = ' '
	If $pEntire = Default Then $pEntire = True

	If $pEntire Then
		Local $iLen = StringLen($pCharacters)
		While StringRight($pText, $iLen) = $pCharacters
			$pText = StringTrimRight($pText, $iLen)
		WEnd
	Else
		Local $boolExit
		Local $strChars = StringSplit($pCharacters, '', 2)
		Do
			$boolExit = True
			For $strChar In $strChars
				If StringInStr($pText, $strChar) Then $boolExit = False
			Next
			If $boolExit Then ExitLoop

			For $strChar In $strChars
				$pText = RS_trimRight($pText, $strChar, True)
			Next
			For $i = UBound($strChars) - 1 To 0 Step -1
				$pText = RS_trimRight($pText, $strChars[$i], True)
			Next
		Until $boolExit
	EndIf

	Return $pText
EndFunc
; =============================================================================== RS_trimRight ===>
; =========================================================================== STRING PROCEDURES ==>

; <== IO PROCEDURES ===============================================================================
; <=== RS_dirMake =================================================================================
; RS_dirMake(String)
; ; Returns full path and create directory.
; ;
; ; @param  String          Directory.
; ; @return String          Full path directory created.
Func RS_dirMake($pFullPathDir)
	$pFullPathDir = RS_fileNameValid($pFullPathDir, 2, True)

	If StringLen($pFullPathDir) > 0 And Not FileExists($pFullPathDir) Then
		If DirCreate($pFullPathDir) = 0 Then Return SetError(1, @extended, '')
	EndIf

	Return $pFullPathDir
EndFunc
; ================================================================================= RS_dirMake ===>

; <=== RS_fileExists ==============================================================================
; RS_fileExists(String)
; ; Wrapper for FileExists. Removes leading and trailing quotes to avoid checking errors.
; ;
; ; @param  String          Filename with full path.
; ; @return Integer         1 if file exists, 0 otherwise.
Func RS_fileExists($pFullPath)
	Return FileExists(RS_trim(RS_trim($pFullPath, "'"), '"'))
EndFunc
; ============================================================================== RS_fileExists ===>
; <=== RS_fileNameInfo ============================================================================
; RS_fileNameInfo(String, Integer)
; ; Returns filename info: drive, directory, filename and/or extension and changing case.
; ;
; ; @param  String          Filename with full path.
; ; @param  Integer         Filename flags. 1: Name, 2: Extension, 4: Dir, 8: Drive,
; ;                         16: Change to uppercase, 32: Change to lowercase . Default: Name (1).
; ; @return String          Filename data: Drive, directory, filename and/or extension.
Func RS_fileNameInfo($pFullPath, $pFlags = Default)
	Local $iCase = 0
	Local $sDrive = ''
	Local $sDir = ''
	Local $sFileName = ''
	Local $sExtension = ''

  If $pFlags = Default Or StringLen($pFlags) = 0 Then $pFlags = 1
	If BitAND($pFlags, 16) Then $iCase = 1
	If BitAND($pFlags, 32) Then $iCase = 2

	_PathSplit($pFullPath, $sDrive, $sDir, $sFileName, $sExtension)
	Switch $iCase
		Case 1
			$sDrive = StringUpper($sDrive)
			$sDir = StringUpper($sDir)
			$sFileName = StringUpper($sFileName)
			$sExtension = StringUpper($sExtension)
		Case 2
			$sDrive = StringLower($sDrive)
			$sDir = StringLower($sDir)
			$sFileName = StringLower($sFileName)
			$sExtension = StringLower($sExtension)
	EndSwitch

	$pFullPath = ''
	If BitAND($pFlags, 8) Then $pFullPath = $sDrive
	If BitAND($pFlags, 4) Then $pFullPath &= $sDir
	If BitAND($pFlags, 1) Then $pFullPath &= $sFileName
	If BitAND($pFlags, 2) Then
		$pFullPath &= $sExtension
	ElseIf $pFlags = 2 Then
		$pFullPath = $sExtension
	EndIf

	Return $pFullPath
EndFunc
; ============================================================================ RS_fileNameInfo ===>
; <=== RS_fileNameValid ===========================================================================
; RS_fileNameValid(String, [Integer], [Boolean])
; ; Returns file or dir name removing invalid characters and adding backslash to directories.
; ;
; ; @param  String          File or directory name.
; ; @param  [Integer]       0: Filename, 1: URI with trailing backslash, 2: Full path. Default: 0.
; ; @param  [Boolean]       If True, expand environment variables. Default: False.
; ; @return String          File or directory name without invalid characters.
Func RS_fileNameValid($pName, $pType = 0, $pEnviron = False)
	If $pEnviron Then $pName = ENVIRON_replace($pName)
	$pName = StringStripWS(StringStripCR($pName), 3)

	Switch $pType
		Case 1
			$pName = StringRegExpReplace($pName, '\\|:|/|\*|\?|\"|\<|\>|\|', '') & '\'
		Case 2
			$pName = RS_trimRight(StringRegExpReplace($pName, '/|\*|\?|\"|\<|\>|\|', ''), '\') & '\'
		Case Else
			$pName = StringRegExpReplace($pName, '\\|:|/|\*|\?|\"|\<|\>|\|', '')
	EndSwitch
	If $pName = '\' Then $pName = ''

	Return $pName
EndFunc
; =========================================================================== RS_fileNameValid ===>
; <=== RS_fileToArray =============================================================================
; RS_fileToArray(String, [String], [Boolean])
; ; Returns an array with text lines from a file avoiding empty lines.
; ;
; ; @param  String          File name.
; ; @param  [String]        Characters to exclude. Default: None.
; ; @param  [Boolean]       True: No include count in first element. Default: False.
; ; @return String[]        Array with text lines.
Func RS_fileToArray($pFile, $pRemoveChars = '', $pNoCount = False)
	If StringLen($pFile) = 0 Or Not FileExists($pFile) Then Return SetError(1, 0, '')

	; Load raw data
	Local $hFile = FileOpen($pFile, $FO_READ)
	If $hFile = -1 Then Return SetError(1, 0, '')
	Local $sRaw = FileRead($hFile)
	FileClose($hFile)

	; Remove empty lines
	If StringLen($sRaw) = 0 Then
		Return SetError(1, 0, '')
	Else
		If StringLen($pRemoveChars) > 0 Then
			For $i = 1 To StringLen($pRemoveChars)
				$sRaw = StringReplace($sRaw, StringMid($pRemoveChars, $i, 1), '', 0, 2)
			Next
		EndIf
		$sRaw = RS_CRLF($sRaw)
		Do
			$sRaw = StringReplace($sRaw, @CRLF & @CRLF, @CRLF)
		Until @extended = 0
		$sRaw = StringStripWS($sRaw, 2)
		Return StringSplit($sRaw, @CRLF, $pNoCount ? 3 : 1)
	EndIf
EndFunc
; ============================================================================= RS_fileToArray ===>
; <=== RS_pathCombine =============================================================================
; RS_pathCombine(String, [String], [String], [String], [String], [String], [String], [String], [String], [String])
; ; Returns full path combining until 10 values.
; ;
; ; @param  String          First path string.
; ; @param  [String]        Second path string.
; ; @param  [String]        Third path string.
; ; @param  [String]        Fourth path string.
; ; @param  [String]        Fifth path string.
; ; @param  [String]        Sixth path string.
; ; @param  [String]        Seventh path string.
; ; @param  [String]        Eighth path string.
; ; @param  [String]        Nineth path string.
; ; @return String          Full path.
;~ Func _RS_path
Func RS_pathCombine($pPath0, $pPath1 = '', $pPath2 = '', $pPath3 = '', $pPath4 = '', $pPath5 = '', $pPath6 = '', $pPath7 = '', $pPath8 = '', $pPath9 = '')
	; Add non-empty strings adding a backslash at end.
	Local $sFullPath = ENVIRON_replace($pPath0)
	If StringLen($pPath1) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? ENVIRON_replace($pPath1) : RS_trimRight($sFullPath, '\') & '\' & ENVIRON_replace($pPath1)
	If StringLen($pPath2) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? ENVIRON_replace($pPath2) : RS_trimRight($sFullPath, '\') & '\' & ENVIRON_replace($pPath2)
	If StringLen($pPath3) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? ENVIRON_replace($pPath3) : RS_trimRight($sFullPath, '\') & '\' & ENVIRON_replace($pPath3)
	If StringLen($pPath4) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? ENVIRON_replace($pPath4) : RS_trimRight($sFullPath, '\') & '\' & ENVIRON_replace($pPath4)
	If StringLen($pPath5) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? ENVIRON_replace($pPath5) : RS_trimRight($sFullPath, '\') & '\' & ENVIRON_replace($pPath5)
	If StringLen($pPath6) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? ENVIRON_replace($pPath6) : RS_trimRight($sFullPath, '\') & '\' & ENVIRON_replace($pPath6)
	If StringLen($pPath7) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? ENVIRON_replace($pPath7) : RS_trimRight($sFullPath, '\') & '\' & ENVIRON_replace($pPath7)
	If StringLen($pPath8) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? ENVIRON_replace($pPath8) : RS_trimRight($sFullPath, '\') & '\' & ENVIRON_replace($pPath8)
	If StringLen($pPath9) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? ENVIRON_replace($pPath9) : RS_trimRight($sFullPath, '\') & '\' & ENVIRON_replace($pPath9)

	Return $sFullPath
EndFunc
; ============================================================================= RS_pathCombine ===>
; <=== RS_run =====================================================================================
; RS_run(String, [String], [String], [Integer], [Integer])
; ; Returns exit code from an external program.
; ;
; ; @param  String          Full path executable and parameters.
; ; @param  [String]        Working directory. Default = ''.
; ; @param  [String]        Tooltip message. Default = ''.
; ; @param  [Integer]       Show flag. Default: @SW_HIDE.
; ; @param  [Integer]       Option flags. Default: $STDOUT_CHILD.
; ; @return Integer         Exit code.
Func RS_run($pCommand, $pDir = '', $pTooltip = '', $pShowFlag = @SW_HIDE, $pOptFlags = $STDOUT_CHILD)
	Local $hProcess
	Local $aExitCode
	Local $bTooltip = Not (StringLen($pTooltip) = 0)
	Local $iPID = Run($pCommand, $pDir, $pShowFlag, $pOptFlags)
	If Not $iPID Then Return SetError(1)

	If _WinAPI_GetVersion() >= 6.0 Then
		$hProcess = _WinAPI_OpenProcess($PROCESS_QUERY_LIMITED_INFORMATION, 0, $iPID)
	Else
		$hProcess = _WinAPI_OpenProcess($PROCESS_QUERY_INFORMATION, 0, $iPID)
	EndIf
	If Not $hProcess Then Return SetError(1)

	While ProcessExists($iPID)
		Sleep(1)
		If $bTooltip Then ToolTip($pTooltip)
	WEnd
	ToolTip('')

	$aExitCode = DllCall('kernel32.dll', 'bool', 'GetExitCodeProcess', 'HANDLE', $hProcess, 'dword*', -1)
	_WinAPI_CloseHandle($hProcess)
	Return $aExitCode[2]
EndFunc
; ===================================================================================== RS_run ===>
; <=== RS_shell ===================================================================================
; RS_shell(String, [String], [String], [String], [Integer], [Integer])
; ; Returns StdOut from an external program.
; ;
; ; @param  String          Full path executable
; ; @param  [String]        Parameters. Default = ''.
; ; @param  [String]        Working directory. Default = ''.
; ; @param  [String]        Tooltip message. Default = ''.
; ; @param  [Integer]       Show flag. Default: @SW_HIDE.
; ; @param  [Integer]       Option flags. Default: $STDOUT_CHILD.
; ; @return String[]        Array with StdOut.
Func RS_shell($pExecutable, $pParameters = '', $pDir = '', $pTooltip = '', $pShowFlag = @SW_HIDE, $pOptFlags = $STDOUT_CHILD)
	; Add quotes to executable if needed and add parameters
	$pExecutable = RS_quote($pExecutable)
	If StringLen($pParameters) > 0 Then $pExecutable &= ' ' & $pParameters

	; Run external program
	Local $iPID = Run(@ComSpec & ' /C ""' & $pExecutable & '""', $pDir, $pShowFlag, $pOptFlags)

	; Wait until process has closed using returned PID.
	While 1
		Sleep(1)
		If StringLen($pTooltip) > 0 Then ToolTip($pTooltip)
		If ProcessExists($iPID) = 0 Then ExitLoop
	WEnd
	ToolTip('')

	; Read Stdout stream of returned PID.
	Return RS_split(StdoutRead($iPID), @CRLF, 3)
EndFunc
; =================================================================================== RS_shell ===>
; =============================================================================== IO PROCEDURES ==>

; <== GUI PROCEDURES ==============================================================================
; <=== RS_chkState ================================================================================
; RS_chkState(Integer)
; ; Checks checkbox state.
; ;
; ; @param  Integer         Control ID of check box.
; ; @return Boolean         True if check box is checked, False otherwise.
Func RS_chkState($pCheckBox)
	Return BitAND(GUICtrlRead($pCheckBox), $GUI_CHECKED) = $GUI_CHECKED
EndFunc
; ================================================================================ RS_chkState ===>
; <=== RS_dlgInput ================================================================================
; RS_dlgInput(String, String, [String], [Integer], [String])
; ; Show an InputBox and return result.
; ;
; ; @param  String          Input box initial value.
; ; @param  String          Input box message.
; ; @param  [String]        Input box title. Default: @ScriptName without extension.
; ; @param  [Integer]       Dialog position, 0: Top left, 1: Top center, 2: Top right,
; ;                         3: Middle left, 4: Middle center, 5: Middle right,
; ;                         6: Bottom left, 7: Bottom center, 8: Bottom right. Default: 4.
; ; @param  [String]        Password character. Default Empty to show original characters.
; ; @return String          String entered in Input box.
Func RS_dlgInput($pDefault, $pMsg, $pTitle = Default, $pPos = Default, $pPasswordChar = Default)
	If $pTitle = Default Or StringLen($pTitle) = 0 Then $pTitle = RS_removeExt(@ScriptName)
	If $pPos = Default Then $pPos = 4
	If $pPasswordChar = Default Then $pPasswordChar = ''

	; Set dialog position
	Local $intX
	Local $intY
	Switch $pPos
		Case 0
			$intX = 10
			$inty = 20
		Case 1
			$intX = @DesktopWidth / 2 - 160
			$inty = 20
		Case 2
			$intX = @DesktopWidth - 310
			$inty = 20
		Case 3
			$intX = 10
			$inty = @DesktopHeight / 2 - 90
		Case 5
			$intX = @DesktopWidth - 310
			$inty = @DesktopHeight / 2 - 90
		Case 6
			$intX = 10
			$inty = @DesktopHeight - 180
		Case 7
			$intX = @DesktopWidth / 2 - 160
			$inty = @DesktopHeight - 180
		Case 8
			$intX = @DesktopWidth - 310
			$inty = @DesktopHeight - 180
		Case Else
			$intX = @DesktopWidth / 2 - 160
			$inty = @DesktopHeight / 2 - 90
	EndSwitch

	$pDefault = InputBox($pTitle, $pMsg, $pDefault, '', 300, 130, $intX, $inty)
	If @error Then Return SetError(@error, @extended, Null)
	Return $pDefault
EndFunc
; ================================================================================ RS_dlgInput ===>
; <=== RS_eventState ==============================================================================
; RS_eventState(Integer, Integer)
; ; Checks if GUI event comes from given checkbox.
; ;
; ; @param  Integer         GUI event.
; ; @param  Integer         Control ID of check box.
; ; @return Boolean         True if GUI event comes from check box, False otherwise.
Func RS_eventState($pMsg, $pControl)
	Return $pMsg = $pControl And RS_chkState($pControl)
EndFunc
; ============================================================================== RS_eventState ===>
; <=== RS_fontExists ==============================================================================
; RS_fontExists(String)
; ; Checks if a font is installed.
; ;
; ; @param  String          GUI event.
; ; @return Boolean         True if font is installed, False otherwise.
Func RS_fontExists($pFontName)
  If StringLen($pFontName) = 0 Then Return False
	Return IsArray(_WinAPI_EnumFontFamilies(0, $pFontName))
EndFunc
; ============================================================================== RS_fontExists ===>
; <=== RS_controlType =============================================================================
; RS_controlType(Integer)
; ; Return control type.
; ;
; ; @param  Integer					Control handle.
; ; @return String        	Control type.
; ; @author	guiness					https://www.autoitscript.com/forum/topic/129129-how-to-obtain-the-type-of-gui-control/?do=findComment&comment=896780
Func RS_controlType($pControlHandle)
	Local Const $GWL_STYLE = -16
	Local $intLong
	Local $strClass

	If IsHWnd($pControlHandle) = 0 Then
		$pControlHandle = GUICtrlGetHandle($pControlHandle)
		If IsHWnd($pControlHandle) = 0 Then Return SetError(1, 0, "Unknown")
	EndIf

	$strClass = _WinAPI_GetClassName($pControlHandle)
	If @error Then Return "Unknown"

	$intLong = _WinAPI_GetWindowLong($pControlHandle, $GWL_STYLE)
	If @error Then Return SetError(2, 0, 0)

	Switch $strClass
		Case "Button"
			Select
				Case BitAND($intLong, $BS_GROUPBOX) = $BS_GROUPBOX
					Return "Group"
				Case BitAND($intLong, $BS_CHECKBOX) = $BS_CHECKBOX
					Return "Checkbox"
				Case BitAND($intLong, $BS_AUTOCHECKBOX) = $BS_AUTOCHECKBOX
					Return "Checkbox"
				Case BitAND($intLong, $BS_RADIOBUTTON) = $BS_RADIOBUTTON
					Return "Radio"
				Case BitAND($intLong, $BS_AUTORADIOBUTTON) = $BS_AUTORADIOBUTTON
					Return "Radio"
			EndSelect

		Case "Edit"
			Select
				Case BitAND($intLong, $ES_WANTRETURN) = $ES_WANTRETURN
					Return "Edit"
				Case Else
					Return "Input"
			EndSelect

		Case "Static"
			Select
				Case BitAND($intLong, $SS_BITMAP) = $SS_BITMAP
					Return "Pic"
				Case BitAND($intLong, $SS_ICON) = $SS_ICON
					Return "Icon"
				Case BitAND($intLong, $SS_LEFT) = $SS_LEFT
					If BitAND($intLong, $SS_NOTIFY) = $SS_NOTIFY Then Return "Label"
					Return "Graphic"
			EndSelect

		Case "ComboBox"
			Return "Combo"
		Case "ListBox"
			Return "ListBox"
		Case "msctls_progress32"
			Return "Progress"
		Case "msctls_trackbar32"
			Return "Slider"
		Case "SysDateTimePick32"
			Return "Date"
		Case "SysListView32"
			Return "ListView"
		Case "SysMonthCal32"
			Return "MonthCal"
		Case "SysTabControl32"
			Return "Tab"
		Case "SysTreeView32"
			Return "TreeView"
	EndSwitch

	Return $strClass
EndFunc   ;==>RS_controlType
; <=== RS_listViewSort ============================================================================
; RS_listViewSort(Integer, Integer, [Integer])
; ; Sorting ListView entries keeping icons.
; ;
; ; @param  Integer         Listview ID.
; ; @param  Integer         Array with entry IDs.
; ; @param  [Integer]       Column pivot.
; ; @return Boolean         Sort order. True: Descending, False: Ascending.
; ; @author R.Gilman (a.k.a rasim)
Func RS_listViewSort($pListView, $pEntryArray, $pEntryIndex = 0)
	Local $iColumnsCount
	Local $iCurPos
	Local $iDimension
	Local $iItemsCount
	Local $iImgSummand
	Local $aItemsTemp
	Local $aItemsText

	Static $bSortDirection

	$iColumnsCount = _GUICtrlListView_GetColumnCount($pListView)
	$iItemsCount = _GUICtrlListView_GetItemCount($pListView)
	$iDimension = $iColumnsCount * 2

	Local $aItemsTemp[1][$iDimension]
	For $i = 0 To $iItemsCount - 1
		$aItemsTemp[0][0] += 1
		ReDim $aItemsTemp[$aItemsTemp[0][0] + 1][$iDimension]

		$aItemsText = _GUICtrlListView_GetItemTextArray($pListView, $i)
		$iImgSummand = $aItemsText[0] - 1

		For $j = 1 To $aItemsText[0]
			$aItemsTemp[$aItemsTemp[0][0]][$j - 1] = $aItemsText[$j]
			$aItemsTemp[$aItemsTemp[0][0]][$j + $iImgSummand] = _GUICtrlListView_GetItemImage($pListView, $i, $j - 1)
		Next
	Next

	$iCurPos = $aItemsTemp[1][$pEntryIndex]
	_ArraySort($aItemsTemp, $bSortDirection, 1, 0, $pEntryIndex)
	$bSortDirection = Not $bSortDirection

	For $i = 1 To $aItemsTemp[0][0]
		For $j = 1 To $iColumnsCount
			_GUICtrlListView_SetItemText($pListView, $i - 1, $aItemsTemp[$i][$j - 1], $j - 1)
			_GUICtrlListView_SetItemImage($pListView, $i - 1, $aItemsTemp[$i][$j + $iImgSummand], $j - 1)
		Next
	Next
	Return $bSortDirection
EndFunc
; ============================================================================ RS_listViewSort ===>
; ============================================================================== GUI PROCEDURES ==>

; <== WINDOWS PROCEDURES ==========================================================================
; <=== RS_tabActivate =============================================================================
; RS_tabActivate(String, [String])
; ; Activates a browser tab with given text.
; ;
; ; @param  String          Tab title.
; ; @param  [String]        Browser window title. Default: $RECT_FirefoxWindow ('Mozilla Firefox').
; ; @return Integer         Window handle.
Func RS_tabActivate($pTargetTab, $pBrowser = Default, $pWaitTime = 50)
	Local $ihWnd
	Local $sFirstTab = ''
	Local $sActiveTab

	If $pBrowser = Default Then $pBrowser = $RS_FirefoxWindow
	$ihWnd = WinActivate($pBrowser)
	If $ihWnd = 0 Then Return SetError(1, 0, 0)

	While 1
		$sActiveTab = WinGetTitle($ihWnd) ; '[ACTIVE]')
		If StringLen($sFirstTab) = 0 Then
			$sFirstTab = $sActiveTab
		Else
			If $sFirstTab == $sActiveTab Then Return SetError(1, 0, 0)
		EndIf

		If StringInStr($sActiveTab, $pTargetTab) Then Return $ihWnd

		Send('^{PGDN}')
		Sleep($pWaitTime)
	WEnd
EndFunc
; ============================================================================= RS_tabActivate ===>
; ========================================================================== WINDOWS PROCEDURES ==>
