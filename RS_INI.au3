#include-once
#include <Array.au3>

Local $boolReplacedEscaped = True

; Creates ASCII characters string
Local $strASCIIchars
For $i = 160 To 255
  $strASCIIchars &= Chr($i)
Next

; <== INI MANAGEMENT ==============================================================================
; <=== INI_append =================================================================================
; INI_append(String, String)
; ; Append a line of text to end of INI file.
; ;
; ; @param  String           INI file.
; ; @param  String           Text line.
; ; @return NONE
Func INI_append($pINIFile, $pLine)
	Local $intFile = FileOpen($pINIFile, 1) ; $FO_APPEND
	If @error Then Return SetError(5006, 0, '')
	FileWriteLine($intFile, $pLine)
	FileClose($intFile)
EndFunc   ;==>INI_append
; <=== INI_append =================================================================================
; INI_append(Boolean)
; ; Default behaviour of escape sequences replacement.
; ;
; ; @param  Boolean           True to replace escaped sequences. False to avoid replacement.
; ; @return NONE
Func INI_replacedEscaped($pReplacedEscaped)
  $pReplacedEscaped = $boolReplacedEscaped
EndFunc   ;==>INI_replacedEscaped
; <=== INI_keys ===================================================================================
; INI_keys(String, String, [String], [Integer], [Boolean])
; ; Returns all keys from section in INI file.
; ;
; ; @param  String          INI file.
; ; @param  String          Section.
; ; @param  [String]        Text within key name.
; ; @param  [Integer]       Text included position, 1: Leading, 2: Trailing, 4: Within. Default: 1.
; ; @param  [Boolean]       Case sensitve comparation.
; ; @return String[]        Array with keys found.
Func INI_keys($pINIFile, $pSection, $pInclude = '', $pIncludePos = 1, $pCaseSensitive = True)
	Local $boolIncluded
	Local $intIncludeLen = 0
	Local $intLinesCount
	Local $intPos
	Local $strLines
	Local $strLine
	Local $strKeys[0]

	If StringLen($pINIFile) = 0 Or StringLen($pSection) = 0 Or Not FileExists($pINIFile) Then Return SetError(5005, 0, '')

	; Load lines from INI file into array
	$intIncludeLen = StringLen($pInclude)
	$strLines = FileReadToArray($pINIFile)
	If @error Then Return SetError(100, 0, '')
	$intLinesCount = @extended

	; Search for sections
	$pSection = '[' & $pSection & ']'
	For $j = 0 To $intLinesCount - 1
		If $pSection == $strLines[$j] Then
			; Section found, add keys to array
			For $i = $j + 1 To $intLinesCount - 1
				$strLine = StringStripWS($strLines[$i], 1)
				If StringLeft($strLine, 1) <> ';' Then
					If StringLeft($strLine, 1) = '[' Then
						; New section, return array
						Return $strKeys
					Else
						; Check and add keys
						$intPos = StringInStr($strLine, '=')
						If $intPos > 0 Then
							$strLine = StringStripWS(StringLeft($strLine, $intPos	- 1), 3)
							If $intIncludeLen > 0 Then
								If StringLen($strLine) > $intIncludeLen Then
									If $pCaseSensitive Then
										; Search text in key name (case sensitive).
										If BitAND($pIncludePos, 1) Then
											$boolIncluded = StringLeft($strLine, $intIncludeLen) == $pInclude
										EndIf
										If BitAND($pIncludePos, 2) Then
											$boolIncluded = $boolIncluded Or StringRight($strLine, $intIncludeLen) == $pInclude
										EndIf
										If BitAND($pIncludePos, 4) Then
											$boolIncluded = $boolIncluded Or StringInStr($strLine, $pInclude, 1, 1, 2, StringLen($strLine) - 2) > 0
										EndIf
									Else
										; Search text in key name (case insensitive).
										If BitAND($pIncludePos, 1) Then
											$boolIncluded = StringLeft($strLine, $intIncludeLen) = $pInclude
										EndIf
										If BitAND($pIncludePos, 2) Then
											$boolIncluded = $boolIncluded Or StringRight($strLine, $intIncludeLen) = $pInclude
										EndIf
										If BitAND($pIncludePos, 4) Then
											$boolIncluded = $boolIncluded Or StringInStr($strLine, $pInclude, 2, 1, 2, StringLen($strLine) - 2) > 0
										EndIf
									EndIf
								EndIf
							Else
								$boolIncluded = True
							EndIf

							If $boolIncluded Then _ArrayAdd($strKeys, $strLine)
						EndIf
					EndIf
				EndIf
			Next
		EndIf
	Next
	Return $strKeys
EndFunc   ;==>INI_keys
; <=== INI_keyExists ==============================================================================
; INI_keyExists(String, String, String, [String], [Integer], [Boolean])
; ; Check if a key exists in a section from an INI file.
; ;
; ; @param  String          INI file.
; ; @param  String          Section.
; ; @param  String          Key.
; ; @param  [String]        Text within section name.
; ; @param  [Integer]       Text included position, 1: Leading, 2: Trailing, 4: Within. Default: 1.
; ; @param  [Boolean]       Case sensitve comparation.
; ; @return Boolean         True if Key was found, False otherwise.
Func INI_keyExists($pINIFile, $pSection, $pKey, $pInclude = '', $pIncludePos = 1, $pCaseSensitive = True)
	If StringLen($pINIFile) = 0 Or StringLen($pSection) = 0 Or StringLen($pKey) = 0 Or Not FileExists($pINIFile) Then Return SetError(5005, 0, False)

	Local $strKeys = INI_keys($pINIFile, $pSection, $pInclude, $pIncludePos, $pCaseSensitive)
	If @error Then Return SetError(101, 0, '')

	For $strKey In $strKeys
		If $strKey = $pKey Then Return True
	Next
	Return False
EndFunc
; <=== INI_objectName =============================================================================
; INI_objectName(String, String, String, String, String)
; ; Returns window or control title & class. If value is empty, use REGEXPTITLE.
; ;
; ; @param  String          INI file.
; ; @param  String          Section.
; ; @param  String          Key.
; ; @param  String          Default value.
; ; @param  String          Class name.
; ; @return String          Window or control advanced title.
Func INI_objectName($pINIFile, $pSection, $pKey, $pValue, $pClass = '')
	If StringLen($pKey) > 0 Then
		Local $strValue = INI_valueLoad($pINIFile, $pSection, $pKey, $pValue)
		If @error Then Return SetError(102, 0, '')
		If StringLen($strValue) = 0 Then $strValue = $pValue

		If StringLen($pClass) > 0 Then
			$pClass = StringReplace(StringReplace($pClass, ']', ']]', 0, 2), '[', '[[', 0, 2)
			If StringLen($strValue) = 0 Then
				Return '[REGEXPTITLE:^(?![\s\S]); CLASS:' & $pClass & ']'
			Else
				Return '[TITLE:' & StringReplace(StringReplace($strValue, ']', ']]', 0, 2), '[', '[[', 0, 2) & '; CLASS:' & $pClass & ']'
			EndIf
		Else
			Return $strValue
		EndIf
	Else
		Return ''
	EndIf
EndFunc
; <=== INI_pointLoad ==============================================================================
; INI_pointLoad(String, String, String, [String], [Boolean])
; ; Returns an array with a point value from INI file.
; ;
; ; @param  String          INI file.
; ; @param  String          Section.
; ; @param  String          Key.
; ; @param  [String]        Default value. Default: Empty string.
; ; @param  [String]        Separator. Space is always included. Default: Comma.
; ; @return Array           Point coordinates.
Func INI_pointLoad($pINIFile, $pSection, $pKey, $pValue = '', $pSeparator = Default)
	If $pSeparator = Default Or StringLen($pSeparator) = 0 Then $pSeparator = ','
	$pSeparator = StringStripWS($pSeparator, 8)

	Local $strLine = INI_valueLoad($pINIFile, $pSection, $pKey, $pValue, False)
	If @error Then Return SetError(103, 0, '')
	$strLine = StringReplace($strLine, $pSeparator & $pSeparator, $pSeparator)
	Local $strEntries = StringSplit($strLine, $pSeparator & ' ', 2)

	; Remove empty entries
	For $i = UBound($strEntries) - 1 To 1 Step -1
		If StringLen($strEntries[$i]) = 0 Then _ArrayDelete($strEntries, $i)
	Next
	Return $strEntries
EndFunc
; <=== INI_sections ===============================================================================
; INI_sections(String, [String], [Integer], [Boolean])
; ; Returns all section from an INI file.
; ;
; ; @param  String          INI file.
; ; @param  [String]        Text within section name.
; ; @param  [Integer]       Text included position, 1: Leading, 2: Trailing, 4: Within. Default: 1.
; ; @param  [Boolean]       Case sensitve comparation.
; ; @return String[]        Array with sections found.
Func INI_sections($pINIFile, $pInclude = '', $pIncludePos = 1, $pCaseSensitive = True)
	Local $boolIncluded
	Local $intIncludeLen = 0
	Local $strLines
	Local $strSections[0]

	If StringLen($pINIFile) = 0 Or Not FileExists($pINIFile) Then Return SetError(5004)

	; Init variables and load lines from INI file into array
	$intIncludeLen = StringLen($pInclude)
	$strLines = FileReadToArray($pINIFile)
	If @error Then Return SetError(5002, 0, '')

	; Search for sections
	For $strLine In $strLines
		If StringLeft($strLine, 1) = '[' Then
			; Section found, add it to array
			$boolIncluded = False
			$strLine = StringMid($strLine, 2, StringInstr($strLine, ']') - 2)
			If $intIncludeLen > 0 Then
				If StringLen($strLine) > $intIncludeLen Then
					If $pCaseSensitive Then
						; Search text in section name (case sensitive).
						If BitAND($pIncludePos, 1) = 1 Then
							$boolIncluded = StringLeft($strLine, $intIncludeLen) == $pInclude
						EndIf
						If BitAND($pIncludePos, 2) = 2 Then
							$boolIncluded = $boolIncluded Or StringRight($strLine, $intIncludeLen) == $pInclude
						EndIf
						If BitAND($pIncludePos, 4) = 4 Then
							$boolIncluded = $boolIncluded Or StringInStr($strLine, $pInclude, 1, 1, 2, StringLen($strLine) - 2) > 0
						EndIf
					Else
						; Search text in section name (case insensitive).
						If BitAND($pIncludePos, 1) = 1 Then
							$boolIncluded = StringLeft($strLine, $intIncludeLen) = $pInclude
						EndIf
						If BitAND($pIncludePos, 2) = 2 Then
							$boolIncluded = $boolIncluded Or StringRight($strLine, $intIncludeLen) = $pInclude
						EndIf
						If BitAND($pIncludePos, 4) = 4 Then
							$boolIncluded = $boolIncluded Or StringInStr($strLine, $pInclude, 2, 1, 2, StringLen($strLine) - 2) > 0
						EndIf
					EndIf
				EndIf
			Else
				$boolIncluded = True
			EndIf

			If $boolIncluded Then _ArrayAdd($strSections, $strLine)
		EndIf
	Next

	Return $strSections
EndFunc   ;==>INI_sections
; <=== INI_sectionExists ==========================================================================
; INI_sectionExists(String, String, [String], [Integer], [Boolean])
; ; Check if a section exists in an INI file.
; ;
; ; @param  String          INI file.
; ; @param  String          Section.
; ; @param  [String]        Text within section name.
; ; @param  [Integer]       Text included position, 1: Leading, 2: Trailing, 4: Within. Default: 1.
; ; @param  [Boolean]       Case sensitve comparation.
; ; @return Boolean         True if Key was found, False otherwise.
Func INI_sectionExists($pINIFile, $pSection, $pInclude = '', $pIncludePos = 1, $pCaseSensitive = True)
	If StringLen($pINIFile) = 0 Or StringLen($pSection) = 0 Or Not FileExists($pINIFile) Then Return SetError(5004, 0, False)

	Local $strSections = INI_sections($pINIFile, $pInclude, $pIncludePos, $pCaseSensitive)
	For $strSection In $strSections
		If $strSection = $pSection Then Return True
	Next
	Return False
EndFunc
; <=== INI_valueDelete ============================================================================
; INI_valueDelete(String, String, String)
; ; Removes a value from an INI file.
; ;
; ; @param  String          INI file.
; ; @param  String          Section.
; ; @param  String          Key.
; ; @return NONE
Func INI_valueDelete($pINIFile, $pSection, $pKey)
	_INI_value($pINIFile, $pSection, $pKey, '', 'deleting')
	If @error Then Return SetError(@error, @extended, '')
EndFunc   ;==>INI_valueDelete
; <=== INI_valueLoad ==============================================================================
; INI_valueLoad(String, String, String, [String], [Boolean])
; ; Reads a value from an INI file.
; ;
; ; @param  String          INI file.
; ; @param  String          Section.
; ; @param  String          Key.
; ; @param  [String]        Default value. Default: Empty string.
; ; @param  [Boolean]       Replace escape sequences. Default: True.
; ; @return String          Key value.
Func INI_valueLoad($pINIFile, $pSection, $pKey, $pValue = Default, $pReplaceEscapes = Default)
	If $pValue = Default Then $pValue = ''
	If $pReplaceEscapes = Default Then $pReplaceEscapes = $boolReplacedEscaped

	$pValue = _INI_value($pINIFile, $pSection, $pKey, $pValue, 'loading')
	If @error Then Return SetError(@error, @extended, '')

	; Replace escape sequences
	If $pReplaceEscapes Then $pValue = _INI_escapeSeq($pValue)

	; Return value
	Return $pValue
EndFunc   ;==>INI_valueLoad
; <=== INI_valueReplace ===========================================================================
; INI_valueReplace(String, String, String, String, [Boolean])
; ; Modify a value in INI file if is diferent from stored value.
; ;
; ; @param  String          INI file.
; ; @param  String          Section.
; ; @param  String          Key.
; ; @param  String          Value.
; ; @param  [Boolean]       Restore escape sequences. Default: True.
; ; @return String          Key value.
Func INI_valueReplace($pINIFile, $pSection, $pKey, $pValue, $pReplaceEscapes = Default)
	Local $strValue = INI_valueLoad($pINIFile, $pSection, $pKey, '', $pReplaceEscapes)
	If @error Then Return SetError(@error, @extended, '')

	If $pValue = $strValue Then
		Return $pValue
	Else
		$strValue = INI_valueWrite($pINIFile, $pSection, $pKey, $pValue, $pReplaceEscapes)
		If @error Then Return SetError(@error, @extended, '')
		Return $strValue
	EndIf
EndFunc
; <=== INI_valueWrite =============================================================================
; INI_valueWrite(String, String, String, String, [Boolean])
; ; Writes a value in an INI file.
; ;
; ; @param  String          INI file.
; ; @param  String          Section.
; ; @param  String          Key.
; ; @param  String          Value.
; ; @param  [Boolean]       Restore escape sequences. Default: True.
; ; @return String          Key value.
Func INI_valueWrite($pINIFile, $pSection, $pKey, $pValue, $pReplaceEscapes = Default)
	If $pReplaceEscapes = Default Then $pReplaceEscapes = $boolReplacedEscaped

  ; Restore escape sequences
	If $pReplaceEscapes Then $pValue = _INI_escapeSeq($pValue, False)

	Local $strValue = _INI_value($pINIFile, $pSection, $pKey, $pValue, 'writing')
	If @error Then Return SetError(@error, @extended, '')
	Return $strValue
EndFunc   ;==>INI_valueWrite
; ============================================================================== INI MANAGEMENT ==>
; <== INTERNAL PROCEDURES =========================================================================
; <=== _INI_escapeSeq =============================================================================
; _INI_escapeSeq(String, [Boolean])
; ; Replaces or restores escape sequences in a string.
; ;
; ; @param  String           Text string.
; ; @param  [Boolean]        Direction. True: Resolve escapes. False: Restore them. Default: True.
; ; @return String           Processed text string.
Func _INI_escapeSeq($pText, $pDirection = True)
	If StringLen($pText) = 0 Then Return ''

	Local $intPos
	Local $strChar

	If $pDirection Then
		; Resolve escape sequences
		$pText = StringReplace($pText, '\\', Chr(1), 0, 2)
		$pText = StringReplace($pText, '\r', @CR, 0, 1)
		$pText = StringReplace($pText, '\n', @CRLF, 0, 1)
		$pText = StringReplace($pText, '\t', Chr(9), 0, 1)

		Do
			$intPos = StringInStr($pText, '\x', 1)
			If $intPos > 0 Then
				$strChar = StringMid($pText, $intPos + 2, 2)
				$pText = StringReplace($pText, '\x' & $strChar, Chr(Dec($strChar)), 0, 1)
			EndIf
		Until $intPos = 0
		$pText = StringReplace($pText, Chr(1), '\')
	Else
		; Restore escape sequences
		$pText = StringReplace($pText, '\', '\\', 0, 2)
		$pText = StringReplace($pText, @CRLF, '\n', 0, 2)
		$pText = StringReplace($pText, @LF, '\n', 0, 2)
		$pText = StringReplace($pText, @CR, '\r', 0, 2)
		$pText = StringReplace($pText, Chr(9), '\t', 0, 2)
		For $i = 1 To StringLen($pText)
			$strChar = StringMid($pText, $i, 1)
			If StringInStr($strASCIIchars, $strChar, 1) > 0 Then $pText = StringReplace($pText, $strChar, '\x' & Hex(Asc($strChar)), 0, 1)
		Next
	EndIf

	Return $pText
EndFunc
; <=== _INI_fileWrite =============================================================================
; _INI_fileWrite(String, ByRef String[])
; ; Internal procedure for writing array to file (preserve encoding).
; ;
; ; @param  String          INI file.
; ; @param  String[]        String array.
; ; @return NONE
Func _INI_fileWrite($pFileName, ByRef $pArray)
	If StringLen($pFileName) = 0 Then Return

	; Check file encoding or create it with $FO_UTF8 encoding.
	Local $intEncoding = (FileExists($pFileName)) ? FileGetEncoding($pFileName, 2) : 128

	; Open file for overwrite with encoding.
	Local $hFile = FileOpen($pFileName, BitOR(2, 8, $intEncoding)) ; $FO_OVERWRITE (2), $FO_CREATEPATH (8)
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
; <=== _INI_trim ==================================================================================
; _INI_trim(String, String, [String])
; ; Removes leading and trailing characters.
; ;
; ; @param  String          Original text.
; ; @param  String          Characters to remove.
; ; @param  [String]        Process characters as entire string. Default: False.
; ; @return String          Trimmed text.
Func _INI_trim($pText, $pCharacters, $pEntire = False)
	Return _INI_trimRight(_INI_trimLeft($pText, $pCharacters, $pEntire), $pCharacters, $pEntire)
EndFunc
; <=== _INI_trimLeft ==============================================================================
; _INI_trimLeft(String, String, [String])
; ; Removes leading characters.
; ;
; ; @param  String          Original text.
; ; @param  String          Characters to remove.
; ; @param  [String]        Process characters as entire string. Default: False.
; ; @return String          Trimmed text.
Func _INI_trimLeft($pText, $pCharacters, $pEntire = False)
	If StringLen($pText) = 0 Then Return ''
	If StringLen($pCharacters) = 0 Then Return $pText

	If $pEntire Then
		Local $intLen = StringLen($pCharacters)
		While StringLeft($pText, $intLen) = $pCharacters
			$pText = StringTrimLeft($pText, $intLen)
		Wend
	Else
		For $strChar In StringSplit($pCharacters, '', 2)
			$pText = _INI_trimLeft($pText, $strChar, True)
		Next
	EndIf
	Return $pText
EndFunc
; <=== _INI_trimRight =============================================================================
; _INI_trimRight(String, String, [String])
; ; Removes trailing characters.
; ;
; ; @param  String          Original text.
; ; @param  String          Characters to remove.
; ; @param  [String]        Process characters as entire string. Default: False.
; ; @return String          Trimmed text.
Func _INI_trimRight($pText, $pCharacters, $pEntire = False)
	If StringLen($pText) = 0 Then Return ''
	If StringLen($pCharacters) = 0 Then Return $pText

	If $pEntire Then
		Local $intLen = StringLen($pCharacters)
		While StringRight($pText, $intLen) = $pCharacters
			$pText = StringTrimRight($pText, $intLen)
		Wend
	Else
		For $strChar In StringSplit($pCharacters, '', 2)
			$pText = _INI_trimRight($pText, $strChar, True)
		Next
	EndIf

	Return $pText
EndFunc
; <=== _INI_value =================================================================================
; _INI_value(String, String, String, String, String)
; ; Internal procedure for value management (supports encoding: UTF-8, ANSI or others).
; ;
; ; @param  String          INI file.
; ; @param  String          Section.
; ; @param  String          Key.
; ; @param  String          Value.
; ; @param  String          Action: deleting, loading, writing.
; ; @return String          Value according to action.
Func _INI_value($pINIFile, $pSection, $pKey, $pValue, $pAction)
	Local $boolSectionFound
	Local $intLinesCount = 0
	Local $intPos
	Local $strLines
	Local $strLine

	; Check parameters
	If StringLen($pINIFile) = 0 Or StringLen($pSection) = 0 Or StringLen($pKey) = 0 Then Return ''
	$pSection = '[' & _INI_trim($pSection, '[]') & ']'
	$pKey = StringStripWS($pKey, 3)
	$pAction = StringLower($pAction)

	If FileExists($pINIFile) Then
		; Load lines from INI file into array
		$strLines = FileReadToArray($pINIFile)
		If @error Then Return SetError(5003, 0, '')
		$intLinesCount = @extended

		For $j = 0 To $intLinesCount - 1
			$strLine = StringStripWS($strLines[$j], 1)
			If $pSection == $strLine Then
				; Section match (case sensitive), search for key
				$boolSectionFound = True
				For $i = $j + 1 To $intLinesCount - 1
					$strLine = StringStripWS($strLines[$i], 1)
					$intPos = StringInStr($strLine, '=', 2)
					If StringLeft($strLine, 1) <> ';' And $intPos > 0 Then
						If $pKey == StringStripWS(StringLeft($strLine, $intPos - 1), 3) Then
							; Key match (case sensitive), execute action
							Switch $pAction
								Case 'deleting'
									_ArrayDelete($strLines, $i)
									If @error Then Return SetError(5003, 1, '')
									_INI_fileWrite($pINIFile, $strLines)
									If @error Then Return SetError(5003, 1, '')
									Return ''
								Case 'loading'
									Return StringMid($strLine, $intPos + 1)
								Case 'writing'
									If StringMid($strLine, $intPos + 1) <> $pValue Then
										$strLines[$i] = $pKey & '=' & $pValue
										_INI_fileWrite($pINIFile, $strLines)
										If @error Then Return SetError(5002, 1, '')
									EndIf
									Return $pValue
							EndSwitch
						Endif
					ElseIf StringLeft($strLine, 1) = '[' Then
						; New Section and no Key found, execute action
						Switch $pAction
							Case 'deleting'
								Return SetError(5003, 2, '')
							Case 'loading'
								Return $pValue
							Case 'writing'
								_ArrayInsert($strLines, $i, $pKey & '=' & $pValue)
								If @error Then Return SetError(5002, 2, '')
								_INI_fileWrite($pINIFile, $strLines)
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
						If Not $boolSectionFound Then
							_ArrayAdd($strLines, $pSection)
							If @error Then Return SetError(5002, 3, '')
						EndIf
						_ArrayAdd($strLines, $pKey & '=' & $pValue)
						If @error Then Return SetError(5002, 3, '')
						_INI_fileWrite($pINIFile, $strLines)
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
				If $intLinesCount > 0 Then
					_ArrayAdd($strLines, $pSection)
					If @error Then Return SetError(5002, 4, '')
					_ArrayAdd($strLines, $pKey & '=' & $pValue)
					If @error Then Return SetError(5002, 4, '')
				Else
					Dim $strLines[2]
					$strLines[0] = $pSection
					$strLines[1] = $pKey & '=' & $pValue
				EndIf
				_INI_fileWrite($pINIFile, $strLines)
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
				Dim $strLines[2]
				$strLines[0] = $pSection
				$strLines[1] = $pKey & '=' & $pValue
				_INI_fileWrite($pINIFile, $strLines)
				If @error Then Return SetError(5002, 0, '')
				Return $pValue
		EndSwitch
	EndIf
EndFunc   ;==>_INI_value
; ========================================================================= INTERNAL PROCEDURES ==>
