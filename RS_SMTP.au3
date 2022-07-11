#include-once

Opt('MustDeclareVars', 1)

#include <File.au3>

; Global variables
Global $objError = ObjEvent('AutoIt.Error', 'SMTP_error')
Global $objErrorDescription

; Local constants and variables
Local $X_Mailer = 'RS SMTP Sender'

; <== MAIN PROCEDURES =============================================================================
; <=== SMTP_send ==================================================================================
; SMTP_send(String, String, String, String, String, String, String, String, String, String, String, String, Integer, Boolean, Boolean)
; ; Send an email using SMTP CDO (Collaboration Data Objects) with optional adding to Send folder.
; ; NOTE: Since CDO doesn't allow adding a message copy to Send folder, a workaround is to use BCC
; ;       address AND creating a filter in sender account with Sender field = $pFrom value, and an
; ;       extra custom field called 'X-Mailer' with value 'RS SMTP Sender'.  So, when the email is
; ;       received, it will be automatically moved to Send folder.
; ;
; ; @param  String          SMTP server name.
; ; @param  String          Sender name.
; ; @param  String          Sender address.
; ; @param  String          Recipient address.
; ; @param  [String]        Email subject. Default: Empty.
; ; @param  [String]        Email body, can be plain or HMTL text. Default: Empty.
; ; @param  [String]        Email file attachments separated by semicolon. Default: Empty.
; ; @param  [String]        CC (Carbon-Copy) addresses. Default: Empty.
; ; @param  [String]        BCC (Blind-Carbon-Copy) addresses. Default: Empty.
; ; @param  [String]        Email priority. Default: Normal.
; ; @param  [String]        Authentication username. Default: Empty.
; ; @param  [String]        Authentication pasword. Default: Empty.
; ; @param  [Integer]       Port. Default: 25.
; ; @param  [Boolean]       Use SSL. Default: False.
; ; @param  [Boolean]       USE TLS. Default: False.
; ; @return String          Empty if successfull, error message otherwise.
Func SMTP_send($pSMTPServer, $pName, $pFrom, $pTo, $pSubject = Default, $pBody = Default, $pAttachments = Default, _
  $pCC = Default, $pBCC = Default, $pPriority = Default, $pUser = Default, $pPassword = Default, _
  $pPort = Default, $pSSL = Default, $pTLS = Default)

  ; Check default values
  If StringLen($pSMTPServer) = 0 Then Return SetError(1, 1, 'SMTP server not defined.')
  If StringLen($pName) = 0 Then Return SetError(1, 2, 'Sender name not defined.')
  If StringLen($pFrom) = 0 Then Return SetError(1, 3, 'Sender address not defined.')
  If StringLen($pTo) = 0 Then Return SetError(1, 4, 'Recipient address not defined.')

  If $pCC = Default Or StringLen($pCC) = 0 Then $pCC = ''
  If $pBCC = Default Then $pBCC = ''
  If $pSubject = Default Or StringLen($pSubject) = 0 Then $pSubject = ''
  If $pBody = Default Or StringLen($pBody) = 0 Then $pBody = ''
  If $pAttachments = Default Or StringLen($pAttachments) = 0 Then $pAttachments = ''
  If $pPriority = Default Or StringLen($pPriority) = 0 Then $pPriority = 'Normal'
  If $pUser = Default Or StringLen($pUser) = 0 Then $pUser = ''
  If $pPassword = Default Or StringLen($pPassword) = 0 Then $pPassword = ''
  If $pPort = Default Or Number($pPort) = 0 Or StringLen($pPort) = 0 Then $pPort = 25
  If $pSSL = Default Or StringLen($pSSL) = 0 Then $pSSL = False
  If $pTLS = Default Or StringLen($pTLS) = 0 Then $pTLS = False

	; Workaround to add email messages to Send folder. See function description to make it work.
	$pBCC = _SMTP_trim(StringStripWS($pBCC, 3))
	$pBCC = (StringLen($pBCC) = 0 ? $pFrom : (StringInStr($pBCC, $pFrom) = 0 ? $pBCC & ';' & $pFrom : $pBCC))

	Local $intError = 0
	Local $strError_Desciption = ''

	; Set CDO object
	Local $objEmail = ObjCreate('CDO.Message')
	$objEmail.BodyPart.ContentTransferEncoding = '8bit'
	$objEmail.BodyPart.CharSet = 'windows-1250'
	$objEmail.Fields('urn:schemas:mailheader:x-mailer') = $X_Mailer

  ; Fix parameter definition errors
  $pPriority = StringUpper(StringLeft($pPriority, 1)) & StringLower(StringMid($pPriority, 2))
  If $pPriority <> 'High' And $pPriority <> 'Low' Then $pPriority = 'Normal'

  ; Set sender and recipients
	$objEmail.From = '"' & $pName & '" <' & $pFrom & '>'
	$objEmail.To = $pTo
  If StringLen($pCC) > 0 Then $objEmail.Cc = $pCC
	If StringLen($pBCC) > 0 Then $objEmail.Bcc = $pBCC

  ; Set message content
	$objEmail.Subject = $pSubject
	If StringInStr($pBody, '<') And StringInStr($pBody, '>') Then
		$objEmail.HTMLBody = $pBody
	Else
		$objEmail.Textbody = $pBody & @CRLF
	EndIf

  ; Set attachments
	If StringLen($pAttachments) > 0 Then
		Local $strFiles = StringSplit($pAttachments, ';')
		For $i = 1 To $strFiles[0]
			$strFiles[$i] = _PathFull($strFiles[$i])
      ConsoleWrite('>Debug: $strFiles[' & $i & '] = ' & $strFiles[$i] & '. Error code: ' & @error & @LF)
			If FileExists($strFiles[$i]) Then
				ConsoleWrite('>Attaching file: ' & $strFiles[$i] & @LF)
				$objEmail.AddAttachment($strFiles[$i])
			Else
				ConsoleWrite('!File to attach not found: ' & $strFiles[$i] & @LF)
			EndIf
		Next
	EndIf

  ; Set email priority
  $objEmail.Fields.Item ('urn:schemas:mailheader:Importance') = $pPriority

  ; Set SMTP server
	$objEmail.Configuration.Fields.Item ('http://schemas.microsoft.com/cdo/configuration/smtpserver') = $pSMTPServer
	$objEmail.Configuration.Fields.Item ('http://schemas.microsoft.com/cdo/configuration/sendusing') = 2 ; cdoSendUsingPort
	$objEmail.Configuration.Fields.Item ('http://schemas.microsoft.com/cdo/configuration/smtpserverport') = $pPort

	; Set SMTP authentication
	If StringLen($pUser) > 0 Then
		$objEmail.Configuration.Fields.Item ('http://schemas.microsoft.com/cdo/configuration/smtpauthenticate') = 1
		$objEmail.Configuration.Fields.Item ('http://schemas.microsoft.com/cdo/configuration/sendusername') = $pUser
		$objEmail.Configuration.Fields.Item ('http://schemas.microsoft.com/cdo/configuration/sendpassword') = $pPassword
	EndIf

	; Set security
	If $pSSL Then $objEmail.Configuration.Fields.Item ('http://schemas.microsoft.com/cdo/configuration/smtpusessl') = True
	If $pTLS Then $objEmail.Configuration.Fields.Item ('http://schemas.microsoft.com/cdo/configuration/sendtls') = True

	; Update settings
	$objEmail.Configuration.Fields.Update
	$objEmail.Fields.Update

	; Send message
	$objEmail.Send

  ; Check for errors
	If @error Then Return SetError(2, @error, '0x' & Hex($objError.Number, 8) & ': ' & $objErrorDescription)

  ; Release object
  $objEmail = ''
  Return ''
EndFunc
; ================================================================================== SMTP_send ===>
; ============================================================================= MAIN PROCEDURES ==>
; <== INTERNAL PROCEDURES =========================================================================
; <=== _SMTP_trim =================================================================================
; _SMTP_trim(String, [String])
; ; Removes defined leading and trailing characters at trailing end.
; ;
; ; @param  String     	    Original text.
; ; @param  [String]   	    Characters to trim. Default: Semicolon.
; ; @return String     	    Processed text.
Func _SMTP_trim($pText, $pCharacters = Default)
	If StringLen($pText) = 0 Then Return ''
	If $pCharacters = Default Or StringLen($pCharacters) = 0 Then $pCharacters = ';'

	Local $intLen = StringLen($pCharacters)
	While StringRight($pText, $intLen) = $pCharacters
		$pText = StringTrimRight($pText, $intLen)
	WEnd
	Return $pText
EndFunc   ;==>_SMTP_trim
; ========================================================================= INTERNAL PROCEDURES ==>
; <== ERROR PROCEDURES ============================================================================
; <=== SMTP_error =================================================================================
; SMTP_error()
; ; Write a array to a file.
; ;
; ; @param  String          Filename.
; ; @param  String          Data array.
; ; @return Boolean         True if successfull, 0 otherwise.
; COM error handler
Func SMTP_error()
	$objErrorDescription = StringStripWS($objError.Description, 3)
	ConsoleWrite('!COM Error Number: 0x' & Hex($objError.Number, 8) & '. ScriptLine: ' & $objError.ScriptLine & '. Description: ' & $objErrorDescription & @LF)
	SetError(1) ; Something to check for when this function returns
	Return
EndFunc
; ================================================================================= SMTP_error ===>
; ============================================================================ ERROR PROCEDURES ==>
