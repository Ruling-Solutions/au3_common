#include-once

Opt('MustDeclareVars', 1)

#include <APIConstants.au3>
#include <Array.au3>
#include <String.au3>
#include <WinAPISys.au3>
#include <WinAPIShellEx.au3>

Local $Environs
Local $SpecialFolders
Local $UserEnvirons
Local $UserEnvironsCase
Local $UserEnvironsSymbol

; <=== ENVIRON_replace ============================================================================
; ENVIRON_replace(String)
; ; Returns a string with user/system environments and special folder variables expanded.
; ;
; ; @param  String          Raw string text.
; ; @return String          String with environments and special folder variables expanded.
Func ENVIRON_replace($pText)
  Local $strVariable

	; User environment variables
	If UBound($UserEnvirons) = 0 Then $UserEnvirons = ENVIRON_userLoad()
	For $i = 0 To UBound($UserEnvirons) - 1
    $strVariable = _ENVIRON_symbol($UserEnvirons[$i][0])
		If StringInStr($pText, $strVariable, 2) > 0 Then
			$pText = StringReplace($pText, $strVariable, $UserEnvirons[$i][1])
		EndIf
	Next

	; System environment variables
	If UBound($Environs) = 0 Then $Environs = ENVIRON_varLoad()
	If @error Then Return SetError(@error, @extended, '')
	For $i = 0 To UBound($Environs) - 1
    $strVariable = _ENVIRON_symbol($Environs[$i][0])
		If StringInStr($pText, $strVariable, 2) > 0 Then
			$pText = StringReplace($pText, $strVariable, $Environs[$i][1])
		EndIf
	Next

	; Special folders variables
	If UBound($SpecialFolders) = 0 Then $SpecialFolders = ENVIRON_specialFolderLoad()
	For $i = 0 To UBound($SpecialFolders) - 1
    $strVariable = _ENVIRON_symbol($SpecialFolders[$i][0])
		If StringInStr($pText, $strVariable, 2) > 0 Then
			$pText = StringReplace($pText, $strVariable, ENVIRON_specialFolder($strVariable))
		EndIf
	Next

	; Fix double backslash if directories are removed
	$pText = StringReplace($pText, '\\', '\', 0, 2)
	Return $pText
EndFunc
; ============================================================================ ENVIRON_replace ===>
; <=== ENVIRON_varLoad ============================================================================
; ENVIRON_varLoad([Integer], [Boolean])
; ; Returns string array with system environment variables. Case and symbols can be set.
; ;
; ; @param  [Integer]       Case. 0: Unchanged, 1: lower, 2: UPPERC, 3: Title. Default: Unchanged.
; ; @param  [Boolean]       Symbols. True: Add '%'. False: Removes symbols. Default: False.
; ; @return String[]        String array with system environment variables.
Func ENVIRON_varLoad($pCase = 0, $pSymbol = False)
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
			$sArray[$i][0] = _ENVIRON_case($sArray[$i][0], $pCase)
			If $pSymbol Then $sArray[$i][0] = '%' & $sArray[$i][0] & '%'
		Next
	EndIf

	Return $sArray
EndFunc
; ============================================================================ ENVIRON_varLoad ===>
; <=== ENVIRON_specialFolderLoad ==================================================================
; ENVIRON_specialFolderLoad([Integer], [Boolean], [Boolean])
; ; Returns string array with special folders path. Case and symbols can be set.
; ; Includes variables not declared in APIShellExConstants.au3.
; ;
; ; @param  [Integer]       String case. -1: Unchanged, 0: lower, 1: UPPER, 2: Title case,
; ;                         3: Proper case, 4: iNVERTED. Default: Unchanged.
; ; @param  [Boolean]       Symbols. True: Add '%'. False: No symbols. Default: False.
; ; @param  [Boolean]       If True, force XP special folders. Default: False.
; ; @return String[]        String array with special folders full path.
Func ENVIRON_specialFolderLoad($pCase = -1, $pSymbol = False, $pWindowsXP = False)
	; Check Windows version
	If _WinAPI_GetVersion() < '6.0' Or $pWindowsXP Then
		; Windows XP or lower version
		Local $sArray[56][2] = [['AdminTools', $CSIDL_ADMINTOOLS], _
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

  If $pCase < -1 Or $pCase > 4 Then $pCase = -1
	$pSymbol = $pSymbol ? True : False
	If $pCase > -1 Or $pSymbol Then
		For $i = 0 To UBound($sArray) - 1
			$sArray[$i][0] = _ENVIRON_case($sArray[$i][0], $pCase)
			If $pSymbol Then $sArray[$i][0] = '%' & $sArray[$i][0] & '%'
		Next
	EndIf

	Return $sArray
EndFunc
; ================================================================== ENVIRON_specialFolderLoad ===>
; <=== ENVIRON_specialFolder ======================================================================
; ENVIRON_specialFolder(String, [Integer], [Boolean])
; ; Return special folders path according to OS version.
; ;
; ; @param  String			    String with Special Folder Member codename.
; ; @param  [Integer]       String case. -1: Unchanged, 0: lower, 1: UPPER, 2: Title case,
; ;                         3: Proper case, 4: iNVERTED. Default: Unchanged.
; ; @param  [Boolean]       If True, force XP special folders. Default: False.
; ; @return String  		    Special folders full path.
Func ENVIRON_specialFolder($pFolder, $pCase = Default, $pForceXP = Default)
  If $pCase = Default Or StringLen($pCase) = 0 Then $pCase = -1
  If $pForceXP = Default Or StringLen($pForceXP) = 0 Then $pForceXP = False

	Local $iIndex
	Local $sFolderPath
  $pFolder = _ENVIRON_symbol($pFolder, False)
	If StringLen($pFolder) = 0 Then Return

	; Get special folders array and search for folder name given
	If UBound($SpecialFolders) = 0 Then $SpecialFolders = ENVIRON_specialFolderLoad(Default, Default, $pForceXP)
	$iIndex = _ArraySearch($SpecialFolders, $pFolder)
	If @error Then Return SetError(@error, @extended, '')

	; Get special or known folder path according to Windows version:
  ; Find folder in XP or lower version
	If $pForceXP Or _WinAPI_GetVersion() < '6.0' Then
    ; Windows XP doesn't have a default download folder, return [WindowsDrive]\Downloads
		If $pFolder = 'downloads' Then
      $iIndex = _ArraySearch($SpecialFolders, 'Windows')
      $sFolderPath = StringLeft(_WinAPI_ShellGetSpecialFolderPath($SpecialFolders[$iIndex][1]), 3) & 'Downloads'
    Else
      $iIndex = _ArraySearch($SpecialFolders, $pFolder)
      $sFolderPath = _WinAPI_ShellGetSpecialFolderPath($SpecialFolders[$iIndex][1])
    EndIf
	Else
		; Find folder in newer Windows version
		$sFolderPath = _WinAPI_ShellGetKnownFolderPath($SpecialFolders[$iIndex][1])
	EndIf
	If @error Then Return SetError(_WinAPI_GetLastErrorMessage())

	Return _ENVIRON_case($sFolderPath, $pCase)
EndFunc
; ====================================================================== ENVIRON_specialFolder ===>
; <=== ENVIRON_userAdd ============================================================================
; ENVIRON_userAdd(String, String)
; ; Add user environment variable to array for use in enviroment replacements.
; ;
; ; @param  String          User enviroment variable name.
; ; @param  String          User enviroment value.
; ; @return NONE
Func ENVIRON_userAdd($pVariable, $pValue)
	If UBound($UserEnvirons) = 0 Then $UserEnvirons = ENVIRON_userLoad()
	If StringLen($pVariable) = 0 Then Return

	$pVariable = _ENVIRON_case($pVariable, $UserEnvironsCase)
	$pVariable = _ENVIRON_symbol($pVariable, $UserEnvironsSymbol)

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
; ============================================================================ ENVIRON_userAdd ===>
; <=== ENVIRON_userCase ===========================================================================
; ENVIRON_userCase(Integer)
; ; Change case of all user environment variables.
; ;
; ; @param  Integer         Case. -1: Unchanged, 0: lower, 1: UPPER, 2: Title case, 3: Proper Case,
; ;                         4: iNVERTED. Default: UPPERCASE.
; ; @return NONE
Func ENVIRON_userCase($pCase = Default)
	If UBound($UserEnvirons) = 0 Then $UserEnvirons = ENVIRON_userLoad()
  If $pCase = Default Or StringLen($pCase) = 0 Or $pCase < 0 Or $pCase > 4 Then $pCase = 1
  $UserEnvironsCase = $pCase
	For $i = 0 To UBound($UserEnvirons) - 1
    $UserEnvirons[$i][0] = _ENVIRON_case($UserEnvirons[$i][0], $UserEnvironsCase)
	Next
EndFunc
; =========================================================================== ENVIRON_userCase ===>
; <=== ENVIRON_userDel ============================================================================
; ENVIRON_userDel(String)
; ; Remove user environment variable in array
; ;
; ; @param  String          User enviroment variable name.
; ; @return NONE
Func ENVIRON_userDel($pVariable)
	If UBound($UserEnvirons) = 0 Then $UserEnvirons = ENVIRON_userLoad()

  $pVariable = _ENVIRON_case($pVariable, $UserEnvironsCase)
	$pVariable = _ENVIRON_symbol($pVariable, $UserEnvironsSymbol)

	Local $iIndex = _ArraySearch($UserEnvirons, $pVariable)
	If @error Then
		Return SetError(@error, @extended, '')
	Else
		_ArrayDelete($UserEnvirons, $iIndex)
	EndIf
EndFunc
; ============================================================================ ENVIRON_userDel ===>
; <=== ENVIRON_userLoad ===========================================================================
; ENVIRON_userLoad([Integer], [Boolean])
; ; Returns string array with default user environment variables. Case and symbol use can be set.
; ;
; ; @param  Integer         Case. -1: Unchanged, 0: lower, 1: UPPER, 2: Title case, 3: Proper Case,
; ;                         4: iNVERTED. Default: UPPERCASE.
; ; @param  [Boolean]       Symbols. True: Add '%'. False: Removes symbols. Default: False.
; ; @return String[]        String array with user environment variables.
Func ENVIRON_userLoad($pCase = Default, $pSymbol = False)
	Local $sArray[2][2]
	$sArray[0][0] = 'APP_NAME'
	$sArray[0][1] = StringRegExpReplace(@ScriptName, '\.[^.\\/]*$', '')
	$sArray[1][0] = 'APP_DIR'
	$sArray[1][1] = @ScriptDir

  If $pCase = Default Or StringLen($pCase) = 0 Or $pCase < 0 Or $pCase > 4 Then $pCase = 1
	$UserEnvironsCase = $pCase
	$UserEnvironsSymbol = $pSymbol ? True : False
  For $i = 0 To 1
    $sArray[$i][0] = _ENVIRON_case($sArray[$i][0], $UserEnvironsCase)
    $sArray[$i][0] = _ENVIRON_symbol($sArray[$i][0], $UserEnvironsSymbol)
  Next

	Return $sArray
EndFunc
; =========================================================================== ENVIRON_userLoad ===>
; <=== ENVIRON_userSymbol =========================================================================
; ENVIRON_userSymbol(Boolean)
; ; Change user environment variables symbol.
; ;
; ; @param  Boolean         Symbols. True: Add '%' at ends. False: Removes symbols.
; ; @return NONE
Func ENVIRON_userSymbol($pSymbol)
	If UBound($UserEnvirons) = 0 Then $UserEnvirons = ENVIRON_userLoad()
	$UserEnvironsSymbol = $pSymbol ? True : False
	For $i = 0 To UBound($UserEnvirons) - 1
    $UserEnvirons[$i][0] = _ENVIRON_symbol($UserEnvirons[$i][0], $UserEnvironsSymbol)
	Next
EndFunc
; ========================================================================= ENVIRON_userSymbol ===>

; <=== _ENVIRON_case ==============================================================================
; _ENVIRON_case(Integer)
; ; Change string case.
; ;
; ; @param  String          String.
; ; @param  Integer         Case. -1: Unchanged, 0: lower, 1: UPPER, 2: Title case, 3: Proper Case,
; ;                         4: iNVERTED. Default: Unchanged.
; ; @return String          String with new casing.
Func _ENVIRON_case($pString, $pCase = Default)
	If $pCase = Default Or StringLen($pString) = 0 Or $pCase < 0 Or $pCase > 4 Then Return $pString

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
; ============================================================================== _ENVIRON_case ===>
; <=== _ENVIRON_symbol ============================================================================
; _ENVIRON_symbol(String, [Boolean])
; ; Returns a variable with or without leading and trailing symbols.
; ;
; ; @param  String          Variable name.
; ; @param  Boolean         Use symbols. True: Add '%'. False: Remove them. Default: True.
; ; @return String          Variable name.
Func _ENVIRON_symbol($pText, $pSymbol = Default)
  If $pSymbol = Default Or StringLen($pSymbol) = 0 Then $pSymbol = True
  $pText = _ENVIRON_trim($pText, '%')
  Return $pSymbol ? '%' & $pText & '%' : $pText
EndFunc
; ============================================================================ _ENVIRON_symbol ===>
; <=== _ENVIRON_trim ==============================================================================
; _ENVIRON_trim(String, [String])
; ; Removes defined leading and trailing characters at both ends.
; ;
; ; @param  String     	    Original text.
; ; @param  [String]   	    Characters to trim. Default: Space.
; ; @return String     	    Processed text.
Func _ENVIRON_trim($pText, $pCharacters = ' ')
	Return _ENVIRON_trimLeft(_ENVIRON_trimRight($pText, $pCharacters), $pCharacters)
EndFunc
; ============================================================================== _ENVIRON_trim ===>
; <=== _ENVIRON_trimLeft ==========================================================================
; _ENVIRON_trimLeft(String, [String])
; ; Removes defined leading characters.
; ;
; ; @param  String     	    Original text.
; ; @param  [String]   	    Characters to trim. Default: Space.
; ; @return String     	    Processed text.
Func _ENVIRON_trimLeft($pText, $pCharacters = Default)
	; Set default values
	If StringLen($pText) = 0 Then Return ''
	If $pCharacters = Default Or StringLen($pCharacters) = 0 Then $pCharacters = ' '

  Local $intLen = StringLen($pCharacters)
  While StringLeft($pText, $intLen) = $pCharacters
    $pText = StringTrimLeft($pText, $intLen)
  WEnd
	Return $pText
EndFunc
; ========================================================================== _ENVIRON_trimLeft ===>
; <=== _ENVIRON_trimRight =========================================================================
; _ENVIRON_trimRight(String, [String])
; ; Removes defined trailing characters.
; ;
; ; @param  String     	    Original text.
; ; @param  [String]   	    Characters to trim. Default: Space.
; ; @return String     	    Processed text.
Func _ENVIRON_trimRight($pText, $pCharacters = Default)
	; Set default values
	If StringLen($pText) = 0 Then Return ''
  If $pCharacters = Default Or StringLen($pCharacters) = 0 Then $pCharacters = ' '

  Local $iLen = StringLen($pCharacters)
  While StringRight($pText, $iLen) = $pCharacters
    $pText = StringTrimRight($pText, $iLen)
  WEnd
	Return $pText
EndFunc
; ========================================================================= _ENVIRON_trimRight ===>
