#include-once

#include <APIConstants.au3>
#include <Array.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <GuiListView.au3>
#include <Math.au3>
#include <ProcessConstants.au3>
#include <String.au3>
#include <WinAPIProc.au3>
#include <WinAPISys.au3>
#include <WinAPIShellEx.au3>

Opt('MustDeclareVars', 1)

Local Const $RS_FirefoxWindow = 'Mozilla Firefox'
Local $ASCIIChars
Local $URIChars = '-.0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz~'
Local $Numbers = '0123456789,.'

Local $Environs
Local $SpecialFolders
Local $UserEnvirons
Local $UserEnvironsCase
Local $UserEnvironsSymbol

; Creates ASCII characters string for escaping routines
For $i = 160 To 255
  $ASCIIChars &= Chr($i)
Next

; <== ARRAY PROCEDURES ============================================================================
; <=== RS_arrayToFile ==============================================================================
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
  Return StringReplace(RS_trimRight($pPath, '\') & '\', '\\', '\', 0, 2)
EndFunc
; ================================================================================ RS_addSlash ===>
; <=== RS_case ====================================================================================
; RS_case(Integer)
; ; Change string case.
; ;
; ; @param  String          String.
; ; @param  Integer         Case. 1: lower, 2: UPPER, 3: Title, 4: iNVERTED.
; ; @return String          String with new casing.
Func RS_case($pString, $pCase)
  If StringLen($pString) = 0 Or $pCase < 1 Or $pCase > 4 Then Return ''

  Switch $pCase
    Case 1
      $pString = StringLower($pString)
    Case 2
      $pString = StringUpper($pString)
    Case 3
      If StringLen($pString) > 1 Then
        $pString = StringUpper(StringLeft($pString, 1)) & StringLower(StringMid($pString, 2))
      EndIf
    Case 4
      If StringLen($pString) > 1 Then
        $pString = StringLower(StringLeft($pString, 1)) & StringUpper(StringMid($pString, 2))
      EndIf
  EndSwitch
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
; <=== RS_decodeURI ===============================================================================
; RS_decodeURI(String)
; ; Decode URI (URL and URN) with percent-encoding.
; ;
; ; @param  String          URI with percent-encoding.
; ; @return String          URI decoded.
Func RS_decodeURI($pURI)
  Local $iPos
  Local $sChar

  $pURI = StringStripWS(StringReplace(StringReplace($pURI, '+', ' '), '&amp;', '&'), 3)

  ; Decode hex values
  While 1
    $iPos = StringInStr($pURI, '%')
    If $iPos = 0 Then ExitLoop
    $sChar = StringMid($pURI, $iPos + 1, 2)
    $pURI = StringReplace($pURI, '%' & $sChar, Chr(Dec($sChar)))
  WEnd

  Return $pURI
EndFunc
; =============================================================================== RS_decodeURI ===>
; <=== RS_encodeURI ===============================================================================
; RS_encodeURI(String)
; ; Encode URI (URL and URN) with percent-encoding.
; ;
; ; @param  String          Sourece URI.
; ; @return String          URI with percent-encoding.
Func RS_encodeURI($pURI)
  Local $iPos = 1
  Local $sChar

  $pURI = StringStripWS(StringReplace($pURI, '%', Chr(1)), 3)

  ; Encode hex values
  While StringLen($pURI) >= $iPos
    $sChar = StringMid($pURI, $iPos, 1)
    If StringInStr($URIChars, $sChar, 1) = 0 And $sChar <> ' ' And $sChar <> Chr(1) Then
      $pURI = StringReplace($pURI, $sChar, '%' & Hex(Asc($sChar), 2), 1, 1)
      $iPos += 2
    EndIf
    $iPos += 1
  WEnd

  $pURI = StringReplace(StringReplace(StringReplace($pURI, Chr(1), '%25'), ' ', '+'), '&', '&amp;')

  Return $pURI
EndFunc
; =============================================================================== RS_encodeURI ===>
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
  If StringInStr($pFileName, ' ') > 0 Then $pFileName = '"' & $pFileName & '"'
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
; ==================================================================================== RS_dirMake ===>
; <=== RS_environ =================================================================================
; RS_environ(String)
; ; Returns a string with user/system environments and special folder variables expanded.
; ;
; ; @param  String          Raw string text.
; ; @return String          String with environments and special folder variables expanded.
Func RS_environ($pText)
  ; User environment variables
  If UBound($UserEnvirons) = 0 Then $UserEnvirons = RS_userEnvironCreate()
  For $i = 0 To UBound($UserEnvirons) - 1
    If StringInStr($pText, '%' & RS_trim($UserEnvirons[$i][0], '%') & '%', 2) > 0 Then
      $pText = StringReplace($pText, '%' & RS_trim($UserEnvirons[$i][0], '%') & '%', $UserEnvirons[$i][1], 0, 2)
    EndIf
  Next

  ; System environment variables
  If UBound($Environs) = 0 Then $Environs = RS_environLoad()
  If @error Then Return SetError(@error, @extended, '')
  For $i = 0 To UBound($Environs) - 1
    If StringInStr($pText, '%' & RS_trim($Environs[$i][0], '%') & '%', 2) > 0 Then
      $pText = StringReplace($pText, '%' & RS_trim($Environs[$i][0], '%') & '%', $Environs[$i][1], 0, 2)
    EndIf
  Next

  ; Special folders variables
  If UBound($SpecialFolders) = 0 Then $SpecialFolders = RS_specialFolderLoad()
  For $i = 0 To UBound($SpecialFolders) - 1
    If StringInStr($pText, '%' & RS_trim($SpecialFolders[$i][0], '%') & '%', 2) > 0 Then
      $pText = StringReplace($pText, '%' & RS_trim($SpecialFolders[$i][0], '%') & '%', RS_specialFolderPath($SpecialFolders[$i][1]), 0, 2)
    EndIf
  Next

  ; Fix double backslash if directories are removed
  $pText = StringReplace($pText, '\\', '\', 0, 2)
  Return $pText
EndFunc
; ================================================================================= RS_environ ===>
; <=== RS_environLoad =============================================================================
; RS_environLoad([Integer], [Boolean])
; ; Returns string array with system environment variables. Case and symbols can be set.
; ;
; ; @param  [Integer]       Case. 0: Unchanged, 1: lower, 2: UPPERC, 3: Title. Default: Unchanged.
; ; @param  [Boolean]       Symbols. True: Add '%'. False: Removes symbols. Default: False.
; ; @return String[]        String array with system environment variables.
Func RS_environLoad($pCase = 0, $pSymbol = False)
  Local $sArray
  Local $sLines
  Local $sEntry
  Local $hID = Run(@ComSpec & ' /c set', @SystemDir, @SW_HIDE, $STDOUT_CHILD)

  ProcessWaitClose($hID)
  $sLines = StdoutRead($hID)
  If @error Then Return SetError(@error, @extended, Null)

  ; Create array
  If StringRight($sLines, 2) = @CRLF Then $sLines = StringTrimRight($sLines, 2)
  $sLines = StringReplace($sLines, '=', '|')
  $sArray = _ArrayFromString($sLines)

  $pCase = $pCase = 1 ? 1 : ($pCase = 2 ? 2 : ($pCase = 3 ? 3 : 0))
  $pSymbol = $pSymbol ? True : False
  If $pCase > 0 Or $pSymbol Then
    For $i = 0 To UBound($sArray) - 1
      $sArray[$i][0] = RS_case($sArray[$i][0], $pCase)
      If $pSymbol Then $sArray[$i][0] = '%' & $sArray[$i][0] & '%'
    Next
  EndIf

  Return $sArray
EndFunc
; ============================================================================= RS_environLoad ===>
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
Func RS_fileNameInfo($pFullPath, $pFlags = 1)
  Local $iCase = 0
  Local $sDrive = ''
  Local $sDir = ''
  Local $sFileName = ''
  Local $sExtension = ''

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
  If $pEnviron Then $pName = RS_environ($pName)
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
  Local $sFullPath = RS_environ($pPath0)
  If StringLen($pPath1) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? RS_environ($pPath1) : RS_trimRight($sFullPath, '\') & '\' & RS_environ($pPath1)
  If StringLen($pPath2) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? RS_environ($pPath2) : RS_trimRight($sFullPath, '\') & '\' & RS_environ($pPath2)
  If StringLen($pPath3) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? RS_environ($pPath3) : RS_trimRight($sFullPath, '\') & '\' & RS_environ($pPath3)
  If StringLen($pPath4) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? RS_environ($pPath4) : RS_trimRight($sFullPath, '\') & '\' & RS_environ($pPath4)
  If StringLen($pPath5) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? RS_environ($pPath5) : RS_trimRight($sFullPath, '\') & '\' & RS_environ($pPath5)
  If StringLen($pPath6) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? RS_environ($pPath6) : RS_trimRight($sFullPath, '\') & '\' & RS_environ($pPath6)
  If StringLen($pPath7) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? RS_environ($pPath7) : RS_trimRight($sFullPath, '\') & '\' & RS_environ($pPath7)
  If StringLen($pPath8) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? RS_environ($pPath8) : RS_trimRight($sFullPath, '\') & '\' & RS_environ($pPath8)
  If StringLen($pPath9) > 0 Then $sFullPath = (StringLen($sFullPath) = 0) ? RS_environ($pPath9) : RS_trimRight($sFullPath, '\') & '\' & RS_environ($pPath9)

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
    ; If ProcessExists($iPID) = 0 Then ExitLoop
  WEnd
  ToolTip('')

  $aExitCode = DllCall('kernel32.dll', 'bool', 'GetExitCodeProcess', 'HANDLE', $hProcess, 'dword*', -1)
  _WinAPI_CloseHandle($hProcess)
  Return $aExitCode[2]
EndFunc

Func RS_scanF()
  Eval
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
  Local $iPID = Run(@ComSpec & ' /C "' & $pExecutable & '"', $pDir, $pShowFlag, $pOptFlags)

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
; <=== RS_specialFolderLoad =======================================================================
; RS_specialFolderLoad([Integer], [Boolean], [Boolean])
; ; Returns string array with special folders path. Case and symbols can be set.
; ; Includes variables not declared in APIShellExConstants.au3.
; ;
; ; @param  [Integer]       Case. 0: Unchanged, 1: lower, 2: UPPERC, 3: Title. Default: Unchanged.
; ; @param  [Boolean]       Symbols. True: Add '%'. False: Removes symbols. Default: False.
; ; @param  [Boolean]       If True, force XP special folders. Default: False.
; ; @return String[]        String array with special folders full path.
Func RS_specialFolderLoad($pCase = 0, $pSymbol = False, $pWindowsXP = False)
  ; Check Windows version
  If _WinAPI_GetVersion() < '6.0' Or $pWindowsXP Then
    ; Windows XP or lower version
    Local $sArray[55][2] = [['AdminTools', $CSIDL_ADMINTOOLS], _
      ['CDBurning', $CSIDL_CDBURN_AREA], _
      ['CommonAdminTools', $CSIDL_COMMON_ADMINTOOLS], _
      ['CommonAppData', $CSIDL_COMMON_APPDATA], _
      ['CommonFavorites', $CSIDL_COMMON_FAVORITES], _
      ['CommonPrograms', $CSIDL_COMMON_PROGRAMS], _
      ['CommonStartMenu', $CSIDL_COMMON_STARTMENU], _
      ['CommonStartup', $CSIDL_COMMON_ALTSTARTUP], _
      ['CommonStartup', $CSIDL_COMMON_STARTUP], _
      ['CommonTemplates', $CSIDL_COMMON_TEMPLATES], _
      ['ComputersNearMe', $CSIDL_COMPUTERSNEARME], _
      ['Connections', $CSIDL_CONNECTIONS], _
      ['Controls', $CSIDL_CONTROLS], _
      ['Cookies', $CSIDL_COOKIES], _
      ['Desktop', $CSIDL_DESKTOP], _
      ['DesktopDirectory', $CSIDL_DESKTOPDIRECTORY], _
      ['Documents', $CSIDL_PERSONAL], _
      ['Downloads', $CSIDL_WINDOWS], _ ; Windows XP default download folder: %windir%\Downloads
      ['Drives', $CSIDL_DRIVES], _
      ['Favorites', $CSIDL_FAVORITES], _
      ['Fonts', $CSIDL_FONTS], _
      ['History', $CSIDL_HISTORY], _
      ['Internet', 0x0001], _
      ['InternetCache', $CSIDL_INTERNET_CACHE], _
      ['LocalAppData', $CSIDL_LOCAL_APPDATA], _
      ['Music', $CSIDL_MYMUSIC], _
      ['NetHood', $CSIDL_NETHOOD], _
      ['Network', 0x0012], _
      ['Pictures', $CSIDL_MYPICTURES], _
      ['Printers', $CSIDL_PRINTERS], _
      ['PrintHood ', $CSIDL_PRINTHOOD], _
      ['Profile', $CSIDL_PROFILE], _
      ['ProgramFiles', $CSIDL_PROGRAM_FILES], _
      ['ProgramFilesCommon', $CSIDL_PROGRAM_FILES_COMMON], _
      ['ProgramFilesCommonX86', $CSIDL_PROGRAM_FILES_COMMONX86], _
      ['ProgramFilesX86', $CSIDL_PROGRAM_FILESX86], _
      ['Programs', $CSIDL_PROGRAMS], _
      ['PublicDesktop', $CSIDL_COMMON_DESKTOPDIRECTORY], _
      ['PublicDocuments', $CSIDL_COMMON_DOCUMENTS], _
      ['PublicMusic', $CSIDL_COMMON_MUSIC], _
      ['PublicPictures', $CSIDL_COMMON_PICTURES], _
      ['PublicVideos', $CSIDL_COMMON_VIDEO], _
      ['Recent', $CSIDL_RECENT], _
      ['RecycleBinFolder', $CSIDL_BITBUCKET], _
      ['Resources', 0x0038], _
      ['ResourcesLocalized', 0x0039], _
      ['RoamingAppData', $CSIDL_APPDATA], _
      ['SendTo', $CSIDL_SENDTO], _
      ['StartMenu', $CSIDL_STARTMENU], _
      ['Startup', $CSIDL_ALTSTARTUP], _
      ['Startup', $CSIDL_STARTUP], _
      ['System', $CSIDL_SYSTEM], _
      ['SystemX86', $CSIDL_SYSTEMX86], _
      ['Templates', $CSIDL_TEMPLATES], _
      ['Videos', $CSIDL_MYVIDEO], _
      ['Windows', $CSIDL_WINDOWS]]
  Else
    ; Windows Vista or higher
    Local $sArray[124][2] = [['AccountPictures', '{008ca0b1-55b4-4c56-b8a8-4de4b299d3be}'], _
      ['AddNewPrograms', $FOLDERID_AddNewPrograms], _
      ['AdminTools', $FOLDERID_AdminTools], _
      ['AppDataDesktop', '{B2C5E279-7ADD-439F-B28C-C41FE1BBF672}'], _
      ['AppDataDocuments', '{7BE16610-1F7F-44AC-BFF0-83E15F2FFCA1}'], _
      ['AppDataFavorites', '{7CFBEFBC-DE1F-45AA-B843-A542AC536CC9}'], _
      ['AppDataProgramData', '{559D40A3-A036-40FA-AF61-84CB430A4D34}'], _
      ['ApplicationShortcuts', '{A3918781-E5F2-4890-B3D9-A7E54332328C}'], _
      ['AppsFolder', '{1e87508d-89c2-42f0-8a7e-645a0f50ca58}'], _
      ['AppUpdates', $FOLDERID_AppUpdates], _
      ['CameraRoll', '{AB5FB87B-7CE2-4F83-915D-550846C9537B}'], _
      ['CDBurning', $FOLDERID_CDBurning], _
      ['ChangeRemovePrograms', $FOLDERID_ChangeRemovePrograms], _
      ['CommonAdminTools', $FOLDERID_CommonAdminTools], _
      ['CommonOEMLinks', $FOLDERID_CommonOEMLinks], _
      ['CommonPrograms', $FOLDERID_CommonPrograms], _
      ['CommonStartMenu', $FOLDERID_CommonStartMenu], _
      ['CommonStartup', $FOLDERID_CommonStartup], _
      ['CommonTemplates', $FOLDERID_CommonTemplates], _
      ['ComputerFolder', $FOLDERID_ComputerFolder], _
      ['ConflictFolder', $FOLDERID_ConflictFolder], _
      ['ConnectionsFolder', $FOLDERID_ConnectionsFolder], _
      ['Contacts', $FOLDERID_Contacts], _
      ['ControlPanelFolder', $FOLDERID_ControlPanelFolder], _
      ['Cookies', $FOLDERID_Cookies], _
      ['Desktop', $FOLDERID_Desktop], _
      ['DeviceMetadataStore', $FOLDERID_DeviceMetadataStore], _
      ['Documents', '{FDD39AD0-238F-46AF-ADB4-6C85480369C7}'], _
      ['DocumentsLibrary', $FOLDERID_DocumentsLibrary], _
      ['Downloads', $FOLDERID_Downloads], _
      ['Favorites', $FOLDERID_Favorites], _
      ['Fonts', $FOLDERID_Fonts], _
      ['Games', $FOLDERID_Games], _
      ['GameTasks', $FOLDERID_GameTasks], _
      ['History', $FOLDERID_History], _
      ['HomeGroup', $FOLDERID_HomeGroup], _
      ['HomeGroupCurrentUser', '{9B74B6A3-0DFD-4f11-9E78-5F7800F2E772}'], _
      ['ImplicitAppShortcuts', $FOLDERID_ImplicitAppShortcuts], _
      ['InternetCache', $FOLDERID_InternetCache], _
      ['InternetFolder', $FOLDERID_InternetFolder], _
      ['Libraries', $FOLDERID_Libraries], _
      ['Links', $FOLDERID_Links], _
      ['LocalAppData', $FOLDERID_LocalAppData], _
      ['LocalAppDataLow', $FOLDERID_LocalAppDataLow], _
      ['LocalizedResourcesDir', $FOLDERID_LocalizedResourcesDir], _
      ['Music', $FOLDERID_Music], _
      ['MusicLibrary', $FOLDERID_MusicLibrary], _
      ['NetHood', $FOLDERID_NetHood], _
      ['NetworkFolder', $FOLDERID_NetworkFolder], _
      ['Objects3D', '{31C0DD25-9439-4F12-BF41-7FF4EDA38722}'], _
      ['OneDrive', '{A52BBA46-E9E1-435f-B3D9-28DAA648C0F6}'], _
      ['OneDriveCameraRoll', '{767E6811-49CB-4273-87C2-20F355E1085B}'], _
      ['OneDriveDocuments', '{24D89E24-2F19-4534-9DDE-6A6671FBB8FE}'], _
      ['OneDrivePictures', '{339719B5-8C47-4894-94C2-D8F77ADD44A6}'], _
      ['OriginalImages', $FOLDERID_OriginalImages], _
      ['PhotoAlbums', $FOLDERID_PhotoAlbums], _
      ['PicturesLibrary', $FOLDERID_PicturesLibrary], _
      ['Pictures', $FOLDERID_Pictures], _
      ['Playlists', $FOLDERID_Playlists], _
      ['PrintersFolder', $FOLDERID_PrintersFolder], _
      ['PrintHood', $FOLDERID_PrintHood], _
      ['Profile', $FOLDERID_Profile], _
      ['ProgramData', $FOLDERID_ProgramData], _
      ['ProgramFiles', $FOLDERID_ProgramFiles], _
      ['ProgramFilesX64', $FOLDERID_ProgramFilesX64], _
      ['ProgramFilesX86', $FOLDERID_ProgramFilesX86], _
      ['ProgramFilesCommon', $FOLDERID_ProgramFilesCommon], _
      ['ProgramFilesCommonX64', $FOLDERID_ProgramFilesCommonX64], _
      ['ProgramFilesCommonX86', $FOLDERID_ProgramFilesCommonX86], _
      ['Programs', $FOLDERID_Programs], _
      ['Public', $FOLDERID_Public], _
      ['PublicDesktop', $FOLDERID_PublicDesktop], _
      ['PublicDocuments', $FOLDERID_PublicDocuments], _
      ['PublicDownloads', $FOLDERID_PublicDownloads], _
      ['PublicGameTasks', $FOLDERID_PublicGameTasks], _
      ['PublicLibraries', $FOLDERID_PublicLibraries], _
      ['PublicMusic', $FOLDERID_PublicMusic], _
      ['PublicPictures', $FOLDERID_PublicPictures], _
      ['PublicRingtones', $FOLDERID_PublicRingtones], _
      ['PublicUserTiles', '{0482af6c-08f1-4c34-8c90-e17ec98b1e17}'], _
      ['PublicVideos', $FOLDERID_PublicVideos], _
      ['QuickLaunch', $FOLDERID_QuickLaunch], _
      ['Recent', $FOLDERID_Recent], _
      ['RecordedTVLibrary', $FOLDERID_RecordedTVLibrary], _
      ['RecycleBinFolder', $FOLDERID_RecycleBinFolder], _
      ['ResourceDir', $FOLDERID_ResourceDir], _
      ['Ringtones', $FOLDERID_Ringtones], _
      ['RoamingAppData', $FOLDERID_RoamingAppData], _
      ['RoamedTileImages', '{AAA8D5A5-F1D6-4259-BAA8-78E7EF60835E}'], _
      ['RoamingTiles', '{00BCFC5A-ED94-4e48-96A1-3F6217F21990}'], _
      ['SampleMusic', $FOLDERID_SampleMusic], _
      ['SamplePictures', $FOLDERID_SamplePictures], _
      ['SamplePlaylists', $FOLDERID_SamplePlaylists], _
      ['SampleVideos', $FOLDERID_SampleVideos], _
      ['SavedGames', $FOLDERID_SavedGames], _
      ['SavedPictures', '{3B193882-D3AD-4eab-965A-69829D1FB59F}'], _
      ['SavedPicturesLibrary', '{E25B5812-BE88-4bd9-94B0-29233477B6C3}'], _
      ['SavedSearches', $FOLDERID_SavedSearches], _
      ['Screenshots', '{b7bede81-df94-4682-a7d8-57a52620b86f}'], _
      ['SearchHistory', '{0D4C3DB6-03A3-462F-A0E6-08924C41B5D4}'], _
      ['SearchHome', $FOLDERID_SearchHome], _
      ['SEARCH_CSC', $FOLDERID_SEARCH_CSC], _
      ['SEARCH_MAPI', $FOLDERID_SEARCH_MAPI], _
      ['SearchTemplates', '{7E636BFE-DFA9-4D5E-B456-D7B39851D8A9}'], _
      ['SendTo', $FOLDERID_SendTo], _
      ['SidebarDefaultParts', $FOLDERID_SidebarDefaultParts], _
      ['SidebarParts', $FOLDERID_SidebarParts], _
      ['StartMenu', $FOLDERID_StartMenu], _
      ['Startup', $FOLDERID_Startup], _
      ['SyncManagerFolder', $FOLDERID_SyncManagerFolder], _
      ['SyncResultsFolder', $FOLDERID_SyncResultsFolder], _
      ['SyncSetupFolder', $FOLDERID_SyncSetupFolder], _
      ['System', $FOLDERID_System], _
      ['SystemX86', $FOLDERID_SystemX86], _
      ['Templates', $FOLDERID_Templates], _
      ['UserPinned', $FOLDERID_UserPinned], _
      ['UserProfiles', $FOLDERID_UserProfiles], _
      ['UserProgramFiles', $FOLDERID_UserProgramFiles], _
      ['UserProgramFilesCommon', $FOLDERID_UserProgramFilesCommon], _
      ['UsersFiles', $FOLDERID_UsersFiles], _
      ['UsersLibraries', $FOLDERID_UsersLibraries], _
      ['Videos', $FOLDERID_Videos], _
      ['VideosLibrary', $FOLDERID_VideosLibrary], _
      ['Windows', $FOLDERID_Windows]]
  EndIf

  $pCase = $pCase = 1 ? 1 : ($pCase = 2 ? 2 : ($pCase = 3 ? 3 : 0))
  $pSymbol = $pSymbol ? True : False
  If $pCase > 0 Or $pSymbol Then
    For $i = 0 To UBound($sArray) - 1
      $sArray[$i][0] = RS_case($sArray[$i][0], $pCase)
      If $pSymbol Then $sArray[$i][0] = '%' & $sArray[$i][0] & '%'
    Next
  EndIf

  Return $sArray
EndFunc
; ======================================================================= RS_specialFolderLoad ===>
; <=== RS_specialFolderPath =======================================================================
; RS_specialFolderPath(String)
; ; Return special folders path according to OS version.
; ;
; ; @param  String			    String with Special Folder Member codename.
; ; @return String  		    Special folders full path.
Func RS_specialFolderPath($pFolder)
  Local $iIndex
  Local $sFolderPath

  $pFolder = StringLower(RS_trim($pFolder, '%'))
  If StringLen($pFolder) = 0 Then Return
;~ 	If $pFolder = 'script' Then Return @ScriptDir

  ; Get special folders array and search for folder name given
  If UBound($SpecialFolders) = 0 Then $SpecialFolders = RS_specialFolderLoad()
  $iIndex = _ArraySearch($SpecialFolders, $pFolder)
  If @error Then Return SetError(@error, @extended, '')

  ; Get special or known folder path according to Windows version
  If _WinAPI_GetVersion() < '6.0' Then
    ; XP or lower version of folder finder
    $sFolderPath = _WinAPI_ShellGetSpecialFolderPath($SpecialFolders[$iIndex][1])
    ; Windows XP default download folder: %Windows%\Downloads
    If $pFolder = 'downloads' Then $sFolderPath = StringLeft($sFolderPath, 3) & 'Downloads'
  Else
    ; Newer Windows version of folder finder
    $sFolderPath = _WinAPI_ShellGetKnownFolderPath($SpecialFolders[$iIndex][1])
  EndIf
  If @error Then Return SetError(_WinAPI_GetLastErrorMessage())

  Return $sFolderPath
EndFunc
; ======================================================================= RS_specialFolderPath ===>
; <=== RS_userEnvironAdd ==========================================================================
; RS_userEnvironAdd(String, String)
; ; Add user environment variable to array for use in enviroment replacements.
; ;
; ; @param  String          User enviroment variable name.
; ; @param  String          User enviroment value.
; ; @return NONE
Func RS_userEnvironAdd($pVariable, $pValue)
  If UBound($UserEnvirons) = 0 Then $UserEnvirons = RS_userEnvironCreate()
  If StringLen($pVariable) = 0 Then Return

  Switch $UserEnvironsCase
    Case 1
      $pVariable = StringLower($pVariable)
    Case 2
      $pVariable = StringUpper($pVariable)
    Case 3
      $pVariable = StringUpper(StringLeft($pVariable, 1)) & StringLower(StringMid($pVariable, 2))
  EndSwitch

  $pVariable = $UserEnvironsSymbol ? '%' & RS_trim($pVariable, '%') & '%': RS_trim($pVariable, '%')

  Local $iIndex = _ArraySearch($UserEnvirons, $pVariable)
  If $iIndex = -1 Then
    If @error = 6 Then
      $iIndex = UBound($UserEnvirons)
      ReDim $UserEnvirons[$iIndex + 1][2]
    Else
      Return SetError(@error, @extended, '')
    EndIf
  EndIf

  $UserEnvirons[$iIndex][0] = $pVariable
  $UserEnvirons[$iIndex][1] = $pValue
EndFunc
; ========================================================================== RS_userEnvironAdd ===>
; <=== RS_userEnvironCase =========================================================================
; RS_userEnvironCase(Integer)
; ; Change user environment variables casing.
; ;
; ; @param  Integer         Case. 1: lowercase, 2: UPPERCASE, 3: Title.
; ; @return NONE
Func RS_userEnvironCase($pCase)
  If UBound($UserEnvirons) = 0 Then $UserEnvirons = RS_userEnvironCreate()
  $UserEnvironsCase = ($pCase = 1) ? 1 : 2
  For $i = 0 To UBound($UserEnvirons) - 1
    Switch $UserEnvironsCase
      Case 1
        $UserEnvirons[$i][0] = StringLower($UserEnvirons[$i][0])
      Case 2
        $UserEnvirons[$i][0] = StringUpper($UserEnvirons[$i][0])
      Case 3
        $UserEnvirons[$i][0] = StringUpper(StringLeft($UserEnvirons[$i][0], 1)) & StringLower(StringMid($UserEnvirons[$i][0], 2))
    EndSwitch
  Next
EndFunc
; ========================================================================= RS_userEnvironCase ===>
; <=== RS_userEnvironCreate =======================================================================
; RS_userEnvironCreate([Integer], [Boolean])
; ; Returns string array with default user environment variables. Case and symbol use can be set.
; ;
; ; @param  [Integer]       Case. 0: Unchanged, 1: lower, 2: UPPERC, 3: Title. Default: Unchanged.
; ; @param  [Boolean]       Symbols. True: Add '%'. False: Removes symbols. Default: False.
; ; @return String[]        String array with user environment variables.
Func RS_userEnvironCreate($pCase = 2, $pSymbol = False)
  Local $sArray[2][2]
  $sArray[0][0] = 'APP_NAME'
  $sArray[0][1] = RS_removeExt(@ScriptName)
  $sArray[1][0] = 'APP_DIR'
  $sArray[1][1] = @ScriptDir

  $UserEnvironsCase = $pCase = 1 ? 1 : ($pCase = 2 ? 2 : ($pCase = 3 ? 3 : 0))
  $UserEnvironsSymbol = $pSymbol ? True : False
  If $UserEnvironsCase > 0 Or $UserEnvironsSymbol Then
    For $i = 0 To 1
      $sArray[$i][0] = RS_case($sArray[$i][0], $UserEnvironsCase)
      If $UserEnvironsSymbol Then $sArray[$i][0] = '%' & $sArray[$i][0] & '%'
    Next
  EndIf

  Return $sArray
EndFunc
; ======================================================================= RS_userEnvironCreate ===>
; <=== RS_userEnvironDel ==========================================================================
; RS_userEnvironDel(String)
; ; Remove user environment variable in array
; ;
; ; @param  String          User enviroment variable name.
; ; @return NONE
Func RS_userEnvironDel($pVariable)
  If UBound($UserEnvirons) = 0 Then $UserEnvirons = RS_userEnvironCreate()

  Switch $UserEnvironsCase
    Case 1
      $pVariable = StringLower($pVariable)
    Case 2
      $pVariable = StringUpper($pVariable)
    Case 3
      $pVariable = StringUpper(StringLeft($pVariable, 1)) & StringLower(StringMid($pVariable, 2))
  EndSwitch

  $pVariable = $UserEnvironsSymbol ? '%' & RS_trim($pVariable, '%') & '%': RS_trim($pVariable, '%')

  Local $iIndex = _ArraySearch($UserEnvirons, $pVariable)
  If @error Then
    Return SetError(@error, @extended, '')
  Else
    _ArrayDelete($UserEnvirons, $iIndex)
  EndIf
EndFunc
; ========================================================================== RS_userEnvironDel ===>
; <=== RS_userEnvironSymbol =======================================================================
; RS_userEnvironSymbol(Boolean)
; ; Change user environment variables symbol.
; ;
; ; @param  Boolean         Symbols. True: Add '%' at ends. False: Removes symbols.
; ; @return NONE
Func RS_userEnvironSymbol($pSymbol)
  If UBound($UserEnvirons) = 0 Then $UserEnvirons = RS_userEnvironCreate()
  $UserEnvironsSymbol = ($pSymbol = True) ? True : False
  For $i = 0 To UBound($UserEnvirons) - 1
    $UserEnvirons[$i][0] = RS_trim($UserEnvirons[$i][0], '%')
    If $UserEnvironsSymbol Then
      $UserEnvirons[$i][0] = '%' & $UserEnvirons[$i][0] & '%'
    EndIf
  Next
EndFunc
; ======================================================================= RS_userEnvironSymbol ===>
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
