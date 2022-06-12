#include-once
#include <Array.au3>

; Creates ASCII characters string
Local $sAscii
For $i = 160 To 255
  $sAscii &= Chr($i)
Next

; <== INI MANAGEMENT ==============================================================================
; <=== INI_append =================================================================================
; INI_append(String, String)
; ; Append a line of text to end of INI file.
; ;
; ; @param  String     	    INI file.
; ; @param  String     	    Text line.
; ; @return NONE
Func INI_append($pINIFile, $pLine)
	Local $iFile = FileOpen($pINIFile, 1) ; $FO_APPEND
	If @error Then Return SetError(5006, 0, '')
	FileWriteLine($iFile, $pLine)
	FileClose($iFile)
EndFunc   ;==>INI_lineAppend

; <=== INI_keys ===================================================================================
; INI_keys(String, String, [String], [Integer], [Boolean])
; ; Returns all keys from section in INI file.
; ;
; ; @param  String     	    INI file.
; ; @param  String     	    Section.
; ; @param  [String]   	    Text within key name.
; ; @param  [Integer]       Text included position, 1: Leading, 2: Trailing, 4: Within. Default: 1.
; ; @param  [Boolean]   	  Case sensitve comparation.
; ; @return String[]   	    Array with keys found.
Func INI_keys($pINIFile, $pSection, $pInclude = '', $pIncludePos = 1, $pCaseSensitive = True)
  Local $bIncluded
  Local $iIncludedLen = 0
	Local $iLinesCount
  Local $iPos
	Local $sLines
  Local $sLine
	Local $sKeys[0]

  If StringLen($pINIFile) = 0 Or StringLen($pSection) = 0 Or Not FileExists($pINIFile) Then Return SetError(5005, 0, '')

  ; Load lines from INI file into array
  $iIncludedLen = StringLen($pInclude)
  $sLines = FileReadToArray($pINIFile)
  If @error Then Return SetError(100, 0, '')
  $iLinesCount = @extended

  ; Search for sections
  $pSection = '[' & $pSection & ']'
  For $j = 0 To $iLinesCount - 1
    If $pSection == $sLines[$j] Then
      ; Section found, add keys to array
      For $i = $j + 1 To $iLinesCount - 1
        $sLine = StringStripWS($sLines[$i], 1)
        If StringLeft($sLine, 1) <> ';' Then
          If StringLeft($sLine, 1) = '[' Then
            ; New section, return array
            Return $sKeys
          Else
            ; Check and add keys
            $iPos = StringInStr($sLine, '=')
            If $iPos > 0 Then
              $sLine = StringStripWS(StringLeft($sLine, $iPos  - 1), 3)
              If $iIncludedLen > 0 Then
                If StringLen($sLine) > $iIncludedLen Then
                  If $pCaseSensitive Then
                    ; Search text in key name (case sensitive).
                    If BitAND($pIncludePos, 1) Then
                      $bIncluded = StringLeft($sLine, $iIncludedLen) == $pInclude
                    EndIf
                    If BitAND($pIncludePos, 2) Then
                      $bIncluded = $bIncluded Or StringRight($sLine, $iIncludedLen) == $pInclude
                    EndIf
                    If BitAND($pIncludePos, 4) Then
                      $bIncluded = $bIncluded Or StringInStr($sLine, $pInclude, 1, 1, 2, StringLen($sLine) - 2) > 0
                    EndIf
                  Else
                    ; Search text in key name (case insensitive).
                    If BitAND($pIncludePos, 1) Then
                      $bIncluded = StringLeft($sLine, $iIncludedLen) = $pInclude
                    EndIf
                    If BitAND($pIncludePos, 2) Then
                      $bIncluded = $bIncluded Or StringRight($sLine, $iIncludedLen) = $pInclude
                    EndIf
                    If BitAND($pIncludePos, 4) Then
                      $bIncluded = $bIncluded Or StringInStr($sLine, $pInclude, 2, 1, 2, StringLen($sLine) - 2) > 0
                    EndIf
                  EndIf
                EndIf
              Else
                $bIncluded = True
              EndIf

              If $bIncluded Then _ArrayAdd($sKeys, $sLine)
            EndIf
          EndIf
        EndIf
      Next
    EndIf
  Next
  Return $sKeys
EndFunc   ;==>INI_keys

; <=== INI_keyExists ==============================================================================
; INI_keyExists(String, String, String, [String], [Integer], [Boolean])
; ; Check if a key exists in a section from an INI file.
; ;
; ; @param  String     	    INI file.
; ; @param  String    	    Section.
; ; @param  String    	    Key.
; ; @param  [String]   	    Text within section name.
; ; @param  [Integer]       Text included position, 1: Leading, 2: Trailing, 4: Within. Default: 1.
; ; @param  [Boolean]   	  Case sensitve comparation.
; ; @return Boolean         True if Key was found, False otherwise.
Func INI_keyExists($pINIFile, $pSection, $pKey, $pInclude = '', $pIncludePos = 1, $pCaseSensitive = True)
  If StringLen($pINIFile) = 0 Or StringLen($pSection) = 0 Or StringLen($pKey) = 0 Or Not FileExists($pINIFile) Then Return SetError(5005, 0, False)

  Local $aKeys = INI_keys($pINIFile, $pSection, $pInclude, $pIncludePos, $pCaseSensitive)
  If @error Then Return SetError(101, 0, '')

  For $sKey In $aKeys
    If $sKey = $pKey Then Return True
  Next
  Return False
EndFunc

; <=== INI_objectName =============================================================================
; INI_objectName(String, String, String, String, String)
; ; Returns window or control title & class. If value is empty, use REGEXPTITLE.
; ;
; ; @param  String     	    INI file.
; ; @param  String    	    Section.
; ; @param  String    	    Key.
; ; @param  String   	      Default value.
; ; @param  String   	      Class name.
; ; @return String     	    Window or control advanced title.
Func INI_objectName($pINIFile, $pSection, $pKey, $pValue, $pClass = '')
  If StringLen($pKey) > 0 Then
    Local $sValue = INI_valueLoad($pINIFile, $pSection, $pKey, $pValue)
    If @error Then Return SetError(102, 0, '')
    If StringLen($sValue) = 0 Then $sValue = $pValue

    If StringLen($pClass) > 0 Then
      $pClass = StringReplace(StringReplace($pClass, ']', ']]', 0, 2), '[', '[[', 0, 2)
      If StringLen($sValue) = 0 Then
        Return '[REGEXPTITLE:^(?![\s\S]); CLASS:' & $pClass & ']'
      Else
        Return '[TITLE:' & StringReplace(StringReplace($sValue, ']', ']]', 0, 2), '[', '[[', 0, 2) & '; CLASS:' & $pClass & ']'
      EndIf
    Else
      Return $sValue
    EndIf
  Else
    Return ''
  EndIf
EndFunc

; <=== INI_pointLoad ==============================================================================
; INI_pointLoad(String, String, String, [String], [Boolean])
; ; Returns an array with a point value from INI file.
; ;
; ; @param  String     	    INI file.
; ; @param  String    	    Section.
; ; @param  String    	    Key.
; ; @param  [String]   	    Default value. Default: Empty string.
; ; @param  [String]   	    Separator. Space is always included. Default: Comma.
; ; @return Array     	    Point coordinates.
Func INI_pointLoad($pINIFile, $pSection, $pKey, $pValue = '', $pSeparator = Default)
  If $pSeparator = Default Or StringLen($pSeparator) = 0 Then $pSeparator = ','
  $pSeparator = StringStripWS($pSeparator, 8)

  Local $sEntries = INI_valueLoad($pINIFile, $pSection, $pKey, $pValue, False)
  If @error Then Return SetError(103, 0, '')
  $sEntries = StringReplace($sEntries, $pSeparator & $pSeparator, $pSeparator)
  Local $aEntries = StringSplit($sEntries, $pSeparator & ' ', 2)

  ; Remove empty entries
  For $i = UBound($aEntries) - 1 To 1 Step -1
    If StringLen($aEntries[$i]) = 0 Then _ArrayDelete($aEntries, $i)
  Next
  Return $aEntries
EndFunc

; <=== INI_sections ===============================================================================
; INI_sections(String, [String], [Integer], [Boolean])
; ; Returns all section from an INI file.
; ;
; ; @param  String     	    INI file.
; ; @param  [String]   	    Text within section name.
; ; @param  [Integer]       Text included position, 1: Leading, 2: Trailing, 4: Within. Default: 1.
; ; @param  [Boolean]   	  Case sensitve comparation.
; ; @return String[]   	    Array with sections found.
Func INI_sections($pINIFile, $pInclude = '', $pIncludePos = 1, $pCaseSensitive = True)
	Local $bIncluded
  Local $iIncludedLen = 0
  Local $sLines
	Local $sSections[0]

  If StringLen($pINIFile) = 0 Or Not FileExists($pINIFile) Then Return SetError(5004)

  ; Init variables and load lines from INI file into array
  $iIncludedLen = StringLen($pInclude)
  $sLines = FileReadToArray($pINIFile)
  If @error Then Return SetError(5002, 0, '')

  ; Search for sections
  For $sLine In $sLines
    If StringLeft($sLine, 1) = '[' Then
      ; Section found, add it to array
      $bIncluded = False
      $sLine = StringMid($sLine, 2, StringInstr($sLine, ']') - 2)
      If $iIncludedLen > 0 Then
        If StringLen($sLine) > $iIncludedLen Then
          If $pCaseSensitive Then
            ; Search text in section name (case sensitive).
            If BitAND($pIncludePos, 1) = 1 Then
              $bIncluded = StringLeft($sLine, $iIncludedLen) == $pInclude
            EndIf
            If BitAND($pIncludePos, 2) = 2 Then
              $bIncluded = $bIncluded Or StringRight($sLine, $iIncludedLen) == $pInclude
            EndIf
            If BitAND($pIncludePos, 4) = 4 Then
              $bIncluded = $bIncluded Or StringInStr($sLine, $pInclude, 1, 1, 2, StringLen($sLine) - 2) > 0
            EndIf
          Else
            ; Search text in section name (case insensitive).
            If BitAND($pIncludePos, 1) = 1 Then
              $bIncluded = StringLeft($sLine, $iIncludedLen) = $pInclude
            EndIf
            If BitAND($pIncludePos, 2) = 2 Then
              $bIncluded = $bIncluded Or StringRight($sLine, $iIncludedLen) = $pInclude
            EndIf
            If BitAND($pIncludePos, 4) = 4 Then
              $bIncluded = $bIncluded Or StringInStr($sLine, $pInclude, 2, 1, 2, StringLen($sLine) - 2) > 0
            EndIf
          EndIf
        EndIf
      Else
        $bIncluded = True
      EndIf

      If $bIncluded Then _ArrayAdd($sSections, $sLine)
    EndIf
  Next

	Return $sSections
EndFunc   ;==>INI_sections

; <=== INI_sectionExists ==========================================================================
; INI_sectionExists(String, String, [String], [Integer], [Boolean])
; ; Check if a section exists in an INI file.
; ;
; ; @param  String     	    INI file.
; ; @param  String    	    Section.
; ; @param  [String]   	    Text within section name.
; ; @param  [Integer]       Text included position, 1: Leading, 2: Trailing, 4: Within. Default: 1.
; ; @param  [Boolean]   	  Case sensitve comparation.
; ; @return Boolean         True if Key was found, False otherwise.
Func INI_sectionExists($pINIFile, $pSection, $pInclude = '', $pIncludePos = 1, $pCaseSensitive = True)
  If StringLen($pINIFile) = 0 Or StringLen($pSection) = 0 Or Not FileExists($pINIFile) Then Return SetError(5004, 0, False)

  Local $aSections = INI_sections($pINIFile, $pInclude, $pIncludePos, $pCaseSensitive)
  For $sSection In $aSections
    If $sSection = $pSection Then Return True
  Next
  Return False
EndFunc

; <=== INI_valueDelete ============================================================================
; INI_valueDelete(String, String, String)
; ; Removes a value from an INI file.
; ;
; ; @param  String     	    INI file.
; ; @param  String    	    Section.
; ; @param  String    	    Key.
; ; @return NONE
Func INI_valueDelete($pINIFile, $pSection, $pKey)
  _INI_value($pINIFile, $pSection, $pKey, '', 'deleting')
  If @error Then Return SetError(@error, @extended, '')
EndFunc   ;==>INI_valueDelete

; <=== INI_valueLoad ==============================================================================
; INI_valueLoad(String, String, String, [String], [Boolean])
; ; Reads a value from an INI file.
; ;
; ; @param  String     	    INI file.
; ; @param  String    	    Section.
; ; @param  String    	    Key.
; ; @param  [String]   	    Default value. Default: Empty string.
; ; @param  [Boolean]       Replace escape sequences. Default: True.
; ; @return String     	    Key value.
Func INI_valueLoad($pINIFile, $pSection, $pKey, $pValue = Default, $pReplaceEscapes = Default)
  Local $iPos
  Local $sChar

  If $pValue = Default Then $pValue = ''
  $pValue = _INI_value($pINIFile, $pSection, $pKey, $pValue, 'loading')
  If @error Then Return SetError(@error, @extended, '')

  ; Replace escape sequences
  If $pReplaceEscapes = Default Then $pReplaceEscapes = True
  If $pReplaceEscapes Then $pValue = _INI_escapeText($pValue)

  ; Return value
	Return $pValue
EndFunc   ;==>INI_valueLoad

; <=== INI_valueReplace ===========================================================================
; INI_valueReplace(String, String, String, String, [Boolean])
; ; Modify a value in INI file if is diferent from stored value.
; ;
; ; @param  String     	    INI file.
; ; @param  String    	    Section.
; ; @param  String    	    Key.
; ; @param  String   		Value.
; ; @param  [Boolean]       Restore escape sequences. Default: True.
; ; @return String    	    Key value.
Func INI_valueReplace($pINIFile, $pSection, $pKey, $pValue, $pReplaceEscapes = Default)
  Local $sValue = INI_valueLoad($pINIFile, $pSection, $pKey, '', $pReplaceEscapes)
  If @error Then Return SetError(@error, @extended, '')

  If $pValue = $sValue Then
    Return $pValue
  Else
    $sValue = INI_valueWrite($pINIFile, $pSection, $pKey, $pValue, $pReplaceEscapes)
    If @error Then Return SetError(@error, @extended, '')
    Return $sValue
  EndIf
EndFunc

; <=== INI_valueWrite =============================================================================
; INI_valueWrite(String, String, String, String, [Boolean])
; ; Writes a value in an INI file.
; ;
; ; @param  String     	    INI file.
; ; @param  String    	    Section.
; ; @param  String    	    Key.
; ; @param  String   		Value.
; ; @param  [Boolean]       Restore escape sequences. Default: True.
; ; @return String    	    Key value.
Func INI_valueWrite($pINIFile, $pSection, $pKey, $pValue, $pReplaceEscapes = Default)
  ; Restore escape sequences
  If $pReplaceEscapes = Default Then $pReplaceEscapes = True
  If $pReplaceEscapes Then $pValue = _INI_escapeText($pValue, True)

  Local $sValue = _INI_value($pINIFile, $pSection, $pKey, $pValue, 'writing')
  If @error Then Return SetError(@error, @extended, '')
  Return $sValue
EndFunc   ;==>INI_valueWrite
; ============================================================================== INI MANAGEMENT ==>

; <== INTERNAL PROCEDURES =========================================================================
; <=== _INI_trim ==================================================================================
; _INI_trim(String, String, [String])
; ; Removes leading and trailing characters.
; ;
; ; @param  String     	    Original text.
; ; @param  String    	    Characters to remove.
; ; @param  [String] 	      Process characters as entire string. Default: False.
; ; @return String    	    Trimmed text.
Func _INI_trim($pText, $pCharacters, $pEntire = False)
  Return _INI_trimRight(_INI_trimLeft($pText, $pCharacters, $pEntire), $pCharacters, $pEntire)
EndFunc
; <=== _INI_trimLeft ==============================================================================
; _INI_trimLeft(String, String, [String])
; ; Removes leading characters.
; ;
; ; @param  String     	    Original text.
; ; @param  String    	    Characters to remove.
; ; @param  [String] 	      Process characters as entire string. Default: False.
; ; @return String    	    Trimmed text.
Func _INI_trimLeft($pText, $pCharacters, $pEntire = False)
  If StringLen($pText) = 0 Then Return ''
  If StringLen($pCharacters) = 0 Then Return $pText

  If $pEntire Then
    Local $intLen = StringLen($pCharacters)
    While StringLeft($pText, $intLen) = $pCharacters
      $pText = StringTrimLeft($pText, $intLen)
    Wend
  Else
    For $sChar In StringSplit($pCharacters, '', 2)
      $pText = _INI_trimLeft($pText, $sChar, True)
    Next
  EndIf
  Return $pText
EndFunc
; <=== _INI_trimRight =============================================================================
; _INI_trimRight(String, String, [String])
; ; Removes trailing characters.
; ;
; ; @param  String     	    Original text.
; ; @param  String    	    Characters to remove.
; ; @param  [String] 	      Process characters as entire string. Default: False.
; ; @return String    	    Trimmed text.
Func _INI_trimRight($pText, $pCharacters, $pEntire = False)
  If StringLen($pText) = 0 Then Return ''
  If StringLen($pCharacters) = 0 Then Return $pText

  If $pEntire Then
    Local $intLen = StringLen($pCharacters)
    While StringRight($pText, $intLen) = $pCharacters
      $pText = StringTrimRight($pText, $intLen)
    Wend
  Else
    For $sChar In StringSplit($pCharacters, '', 2)
      $pText = _INI_trimRight($pText, $sChar, True)
    Next
  EndIf

  Return $pText
EndFunc

; <=== _INI_value =================================================================================
; _INI_value(String, String, String, String, String)
; ; Internal procedure for value management (supports encoding: UTF-8, ANSI or others).
; ;
; ; @param  String     	    INI file.
; ; @param  String    	    Section.
; ; @param  String    	    Key.
; ; @param  String    	    Value.
; ; @param  String    	    Action: deleting, loading, writing.
; ; @return String    	    Value according to action.
Func _INI_value($pINIFile, $pSection, $pKey, $pValue, $pAction)
  Local $bSectionFound
  Local $iLinesCount = 0
  Local $iPos
	Local $sLines
  Local $sLine

  ; Check parameters
	If StringLen($pINIFile) = 0 Or StringLen($pSection) = 0 Or StringLen($pKey) = 0 Then Return ''
  $pSection = '[' & _INI_trim($pSection, '[]') & ']'
  $pKey = StringStripWS($pKey, 3)
  $pAction = StringLower($pAction)

	If FileExists($pINIFile) Then
		; Load lines from INI file into array
		$sLines = FileReadToArray($pINIFile)
		If @error Then Return SetError(5003, 0, '')
		$iLinesCount = @extended

		For $j = 0 To $iLinesCount - 1
      $sLine = StringStripWS($sLines[$j], 1)
			If $pSection == $sLine Then
        ; Section match (case sensitive), search for key
        $bSectionFound = True
        For $i = $j + 1 To $iLinesCount - 1
          $sLine = StringStripWS($sLines[$i], 1)
          $iPos = StringInStr($sLine, '=', 2)
          If StringLeft($sLine, 1) <> ';' And $iPos > 0 Then
            If $pKey == StringStripWS(StringLeft($sLine, $iPos - 1), 3) Then
              ; Key match (case sensitive), execute action
              Switch $pAction
                Case 'deleting'
                  _ArrayDelete($sLines, $i)
                  If @error Then Return SetError(5003, 1, '')
                  _INI_fileWrite($pINIFile, $sLines)
                  If @error Then Return SetError(5003, 1, '')
                  Return ''
                Case 'loading'
                  Return StringMid($sLine, $iPos + 1)
                Case 'writing'
                  If StringMid($sLine, $iPos + 1) <> $pValue Then
                    $sLines[$i] = $pKey & '=' & $pValue
                    _INI_fileWrite($pINIFile, $sLines)
                    If @error Then Return SetError(5002, 1, '')
                  EndIf
                  Return $pValue
              EndSwitch
            Endif
          ElseIf StringLeft($sLine, 1) = '[' Then
            ; New Section and no Key found, execute action
            Switch $pAction
              Case 'deleting'
                Return SetError(5003, 2, '')
              Case 'loading'
                Return $pValue
              Case 'writing'
                _ArrayInsert($sLines, $i, $pKey & '=' & $pValue)
                If @error Then Return SetError(5002, 2, '')
                _INI_fileWrite($pINIFile, $sLines)
                If @error Then Return SetError(5002, 2, '')
                Return $pValue
            EndSwitch
          EndIf
        Next

        ; No Key found, execute action
        Switch $pAction
          Case 'deleting'
            Return SetError(5003, 3, '')
          Case 'loading'
            Return $pValue
          Case 'writing'
            If Not $bSectionFound Then
              _ArrayAdd($sLines, $pSection)
              If @error Then Return SetError(5002, 3, '')
            EndIf
            _ArrayAdd($sLines, $pKey & '=' & $pValue)
            If @error Then Return SetError(5002, 3, '')
            _INI_fileWrite($pINIFile, $sLines)
            If @error Then Return SetError(5002, 3, '')
            Return $pValue
        EndSwitch
      EndIf
		Next

		; No Section found, execute action
    Switch $pAction
      Case 'deleting'
        Return SetError(5003, 4, '')
      Case 'loading'
        Return $pValue
      Case 'writing'
        If $iLinesCount > 0 Then
          _ArrayAdd($sLines, $pSection)
          If @error Then Return SetError(5002, 4, '')
          _ArrayAdd($sLines, $pKey & '=' & $pValue)
          If @error Then Return SetError(5002, 4, '')
        Else
          Dim $sLines[2]
          $sLines[0] = $pSection
          $sLines[1] = $pKey & '=' & $pValue
        EndIf
        _INI_fileWrite($pINIFile, $sLines)
        If @error Then Return SetError(5002, 4, '')
        Return $pValue
    EndSwitch
	Else
		; No INI file found, execute action
    Switch $pAction
      Case 'deleting'
        Return SetError(5003, 0, '')
      Case 'loading'
        Return $pValue
      Case 'writing'
        Dim $sLines[2]
        $sLines[0] = $pSection
        $sLines[1] = $pKey & '=' & $pValue
        _INI_fileWrite($pINIFile, $sLines)
        If @error Then Return SetError(5002, 0, '')
        Return $pValue
    EndSwitch
	EndIf
EndFunc   ;==>_INI_value

; <=== _INI_escapeText ============================================================================
; _INI_escapeText(String, [Boolean])
; ; Replaces or restores escape characters from string.
; ;
; ; @param  String     	    Text string to process.
; ; @param  [Boolean]  	    False to replace or True to restore escape characters. Default: False.
; ; @return String     	    Text string processed.
Func _INI_escapeText($pText, $pRestoreEscapes = False)
  Local $iPos
  Local $sChar

  If StringLen($pText) = 0 Then Return ''

  If $pRestoreEscapes Then
    $pText = StringReplace($pText, '\', '\\', 0, 2)
    $pText = StringReplace($pText, @CRLF, '\n', 0, 2)
    $pText = StringReplace($pText, @CR, '\r', 0, 2)
    $pText = StringReplace($pText, Chr(9), '\t', 0, 2)
    For $i = 1 To StringLen($pText)
      $sChar = StringMid($pText, $i, 1)
      If StringInStr($sAscii, $sChar, 1) > 0 Then $pText = StringReplace($pText, $sChar, '\x' & Hex(Asc($sChar)), 0, 1)
    Next
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

; <=== _INI_fileWrite =============================================================================
; _INI_fileWrite(String, ByRef String[])
; ; Internal procedure for writing array to file (preserve encoding).
; ;
; ; @param  String     	    INI file.
; ; @param  String[]  	    String array.
; ; @return NONE
Func _INI_fileWrite($pFileName, ByRef $pArray)
  If StringLen($pFileName) = 0 Then Return

  ; Check file encoding or create with $FO_UTF8 encoding.
  Local $iEncoding = (FileExists($pFileName)) ? FileGetEncoding($pFileName, 2) : 128

  ; Open file for overwrite with encoding.
  Local $hFile = FileOpen($pFileName, 2 + $iEncoding) ; $FO_OVERWRITE (2)
  If $hFile = -1 Then Return SetError(1)

  ; Write array to file.
  For $i = 0 To UBound($pArray) - 1
    If FileWriteLine($hFile, $pArray[$i]) = 0 Then
      FileClose($hFile)
      Return SetError(1)
    EndIf
  Next

  ; Close handle.
  FileClose($hFile)
EndFunc
; ========================================================================= INTERNAL PROCEDURES ==>
