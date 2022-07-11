#include-once

Opt('MustDeclareVars', 1)

Local Const $Entities[][2] = [[38, 'amp'], [39, 'apos'], [60, 'lt'], [62, 'gt'], _
	[160, 'nbsp'], [161, 'iexcl'], [162, 'cent'], [163, 'pound'], [164, 'curren'], [165, 'yen'], _
	[166, 'brvbar'], [167, 'sect'], [168, 'uml'], [169, 'copy'], [170, 'ordf'], [171, 'laquo'], _
	[172, 'not'], [173, 'shy'], [174, 'reg'], [175, 'macr'], [176, 'deg'], [177, 'plusmn'], _
	[180, 'acute'], [181, 'micro'], [182, 'para'], [183, 'middot'], [184, 'cedil'], [186, 'ordm'], _
	[187, 'raquo'], [191, 'iquest'], [192, 'Agrave'], [193, 'Aacute'], [194, 'Acirc'], [195, 'Atilde'], _
	[196, 'Auml'], [197, 'Aring'], [198, 'AElig'], [199, 'Ccedil'], [200, 'Egrave'], [201, 'Eacute'], _
	[202, 'Ecirc'], [203, 'Euml'], [204, 'Igrave'], [205, 'Iacute'], [206, 'Icirc'], [207, 'Iuml'], _
	[208, 'ETH'], [209, 'Ntilde'], [210, 'Ograve'], [211, 'Oacute'], [212, 'Ocirc'], [213, 'Otilde'], _
	[214, 'Ouml'], [215, 'times'], [216, 'Oslash'], [217, 'Ugrave'], [218, 'Uacute'], [219, 'Ucirc'], _
	[220, 'Uuml'], [221, 'Yacute'], [222, 'THORN'], [223, 'szlig'], [224, 'agrave'], [225, 'aacute'], _
	[226, 'acirc'], [227, 'atilde'], [228, 'auml'], [229, 'aring'], [230, 'aelig'], [231, 'ccedil'], _
	[232, 'egrave'], [233, 'eacute'], [234, 'ecirc'], [235, 'euml'], [236, 'igrave'], [237, 'iacute'], _
	[238, 'icirc'], [239, 'iuml'], [240, 'eth'], [241, 'ntilde'], [242, 'ograve'], [243, 'oacute'], _
	[244, 'ocirc'], [245, 'otilde'], [246, 'ouml'], [247, 'divide'], [248, 'oslash'], [249, 'ugrave'], _
	[250, 'uacute'], [251, 'ucirc'], [252, 'uuml'], [253, 'yacute'], [254, 'thorn'], [255, 'yuml'], _
	[338, 'OElig'], [339, 'oelig'], [352, 'Scaron'], [353, 'scaron'], [376, 'Yuml'], [402, 'fnof'], _
	[710, 'circ'], [732, 'tilde'], [913, 'Alpha'], [914, 'Beta'], [915, 'Gamma'], [916, 'Delta'], _
	[917, 'Epsilon'], [918, 'Zeta'], [919, 'Eta'], [920, 'Theta'], [921, 'Iota'], [922, 'Kappa'], _
	[923, 'Lambda'], [924, 'Mu'], [925, 'Nu'], [926, 'Xi'], [927, 'Omicron'], [928, 'Pi'], [929, 'Rho'], _
	[931, 'Sigma'], [932, 'Tau'], [933, 'Upsilon'], [934, 'Phi'], [935, 'Chi'], [936, 'Psi'], _
	[937, 'Omega'], [945, 'alpha'], [946, 'beta'], [947, 'gamma'], [948, 'delta'], [949, 'epsilon'], _
	[950, 'zeta'], [951, 'eta'], [952, 'theta'], [953, 'iota'], [954, 'kappa'], [955, 'lambda'], _
	[956, 'mu'], [957, 'nu'], [958, 'xi'], [959, 'omicron'], [960, 'pi'], [961, 'rho'], [962, 'sigmaf'], _
	[963, 'sigma'], [964, 'tau'], [965, 'upsilon'], [966, 'phi'], [967, 'chi'], [968, 'psi'], _
	[969, 'omega'], [977, 'thetasym'], [978, 'upsih'], [982, 'piv'], [8194, 'ensp'], [8195, 'emsp'], _
	[8201, 'thinsp'], [8204, 'zwnj'], [8205, 'zwj'], [8206, 'lrm'], [8207, 'rlm'], [8211, 'ndash'], _
	[8212, 'mdash'], [8216, 'lsquo'], [8217, 'rsquo'], [8218, 'sbquo'], [8220, 'ldquo'], _
	[8221, 'rdquo'], [8222, 'bdquo'], [8224, 'dagger'], [8225, 'Dagger'], [8226, 'bull'], _
	[8230, 'hellip'], [8240, 'permil'], [8242, 'prime'], [8243, 'Prime'], [8249, 'lsaquo'], _
	[8250, 'rsaquo'], [8254, 'oline'], [8260, 'frasl'], [8364, 'euro'], [8465, 'image'], _
	[8472, 'weierp'], [8476, 'real'], [8482, 'trade'], [8501, 'alefsym'], [8592, 'larr'], [8593, 'uarr'], _
	[8594, 'rarr'], [8595, 'darr'], [8596, 'harr'], [8629, 'crarr'], [8656, 'lArr'], [8657, 'uArr'], _
	[8658, 'rArr'], [8659, 'dArr'], [8660, 'hArr'], [8704, 'forall'], [8706, 'part'], [8707, 'exist'], _
	[8709, 'empty'], [8711, 'nabla'], [8712, 'isin'], [8713, 'notin'], [8715, 'ni'], [8719, 'prod'], _
	[8721, 'sum'], [8722, 'minus'], [8727, 'lowast'], [8730, 'radic'], [8733, 'prop'], [8734, 'infin'], _
	[8736, 'ang'], [8743, 'and'], [8744, 'or'], [8745, 'cap'], [8746, 'cup'], [8747, 'int'], _
	[8764, 'sim'], [8773, 'cong'], [8776, 'asymp'], [8800, 'ne'], [8801, 'equiv'], [8804, 'le'], _
	[8805, 'ge'], [8834, 'sub'], [8835, 'sup'], [8836, 'nsub'], [8838, 'sube'], [8839, 'supe'], _
	[8853, 'oplus'], [8855, 'otimes'], [8869, 'perp'], [8901, 'sdot'], [8968, 'lceil'], [8969, 'rceil'], _
	[8970, 'lfloor'], [8971, 'rfloor'], [9001, 'lang'], [9002, 'rang'], [9674, 'loz'], [9824, 'spades'], _
	[9827, 'clubs'], [9829, 'hearts'], [9830, 'diams']]
Local $URLChars = '-.0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ_abcdefghijklmnopqrstuvwxyz~'

; <=== DECODE_URL =================================================================================
; DECODE_URL(String)
; ; Decode percent-encoding URL.
; ;
; ; @param	String					Percent-encoded URL.
; ; @return	String					Decoded URL.
Func DECODE_URL($pURL)
	Local $intPos
	Local $strChar

	; Decode hex values
	$pURL = StringStripWS(StringReplace($pURL, '+', ' ', 0, 2), 3)
	$pURL = DECODE_HTML($pURL)
	
	Do
		$intPos = StringInStr($pURL, '%')
		If $intPos > 0 Then
			$strChar = StringMid($pURL, $intPos + 1, 2)
			If Dec($strChar) > 0 Then
				; If a valid code is found, replace it.
				$pURL = StringReplace($pURL, '%' & $strChar, Chr(Dec($strChar)))
			Else
				; Otherwise, replace percent with SOH
				$pURL = StringLeft($pURL, $intPos - 1) & Chr(1) & StringMid($pURL, $intPos + 1)
			EndIf
		EndIf
	Until $intPos = 0

	; Replace SOHs with percents
	$pURL = StringReplace($pURL, Chr(1), '%', 0, 2)
	Return $pURL
EndFunc
; ================================================================================= DECODE_URL ===>
; <=== ENCODE_URL =================================================================================
; ENCODE_URL(String)
; ; Encode URL with percent-encoding.
; ;
; ; @param	String					Plain URL.
; ; @return	String					Percent-encoded URL.
Func ENCODE_URL($pURL)
	Local $intPos = 1
	Local $strChar
	
	; Encode hex values
	$pURL = StringStripWS(StringReplace($pURL, '%', Chr(1), 0, 2), 3)
	While StringLen($pURL) >= $intPos
		$strChar = StringMid($pURL, $intPos, 1)
		If StringInStr($URLChars, $strChar, 1) = 0 And $strChar <> ' ' And $strChar <> Chr(1) Then
			$pURL = StringReplace($pURL, $strChar, '%' & Hex(Asc($strChar), 2), 1, 1)
			$intPos += 2
		EndIf
		$intPos += 1
	WEnd
	$pURL = StringReplace(StringReplace(StringReplace($pURL, Chr(1), '%25', 0, 2), ' ', '+', 0, 2), '&', '&amp;', 0, 2)

	Return $pURL
EndFunc
; ================================================================================= ENCODE_URL ===>
; <=== DECODE_HTML ================================================================================
; DECODE_HTML(String)
; ; Decode HTML entities in text.
; ;
; ; @param	String					Text with HTML entities.
; ; @return	String					Decoded text.
; ; @author	WinWiesel				https://www.autoitscript.com/forum/topic/203629-convert-html-encoding-strings/?do=findComment&comment=1462238
Func DECODE_HTML($pText)
	; Replace quotes and quote entities with SOH to avoid errors in StringRegExpReplace
	$pText = StringReplace($pText, '"', Chr(1), 0, 2)
	$pText = StringReplace($pText, '&quot;', Chr(1), 0, 2)
	
	; Replace entities
	For $i = 0 To UBound($Entities) - 1
		If StringInStr($pText, "&" & $Entities[$i][1], 1) Then
			$pText = StringReplace($pText, "&" & $Entities[$i][1] & ";", ChrW($Entities[$i][0]), 0, 1)
		EndIf
	Next
	$pText = Execute('"' & StringRegExpReplace($pText, '\\u([[:xdigit:]]{4})', '" & ChrW(0x$1) & "') & '"')
	$pText = Execute('"' & StringRegExpReplace($pText, '&#(\d{3});', '" & Chr($1) & "') & '"')
	
	; Replace SOH with quotes
	$pText = StringReplace($pText, Chr(1), '"', 0, 2)
	Return $pText
EndFunc
; ================================================================================ DECODE_HTML ===>
; <=== ENCODE_HTML ================================================================================
; ENCODE_HTML(String)
; ; Encode HTML entities in text.
; ;
; ; @param	String					Plain text.
; ; @return	String					Encoded text with HTML entities.
; ; @author	WinWiesel				https://www.autoitscript.com/forum/topic/203629-convert-html-encoding-strings/?do=findComment&comment=1462238
Func ENCODE_HTML($pText)
	; Replace entities
	For $i = 0 To UBound($Entities) - 1
		If StringInStr($pText, ChrW($Entities[$i][0]), 1) Then
			$pText = StringReplace($pText, ChrW($Entities[$i][0]), "&" & $Entities[$i][1] & ";", 0, 1)
		EndIf
	Next

	; Replace quotes with qoute entities
	$pText = StringReplace($pText, '"', '&quot;', 0, 1)

	Return $pText
EndFunc
; ================================================================================ ENCODE_HTML ===>
