Attribute VB_Name = "modCommon"
Option Explicit

'------------------------------------------------------ STARTS
' to set the full window Opacity
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'------------------------------------------------------ ENDS

'---------------------------------------------------------------------------------------
Private Const WS_EX_LAYERED  As Long = &H80000
Private Const GWL_EXSTYLE  As Long = (-20)
Private Const LWA_COLORKEY  As Long = &H1       'to trans'
Private Const LWA_ALPHA  As Long = &H2          'to semi trans'

Private Declare Function PathGetCharType Lib "shlwapi.dll" Alias "PathGetCharTypeW" (ByVal ch As Integer) As Long

'Private Const GCT_INVALID = &H0
Private Const GCT_LFNCHAR = &H1
'Private Const GCT_SEPARATOR = &H8
Private Const GCT_SHORTCHAR = &H2
'Private Const GCT_WILD = &H4
'---------------------------------------------------------------------------------------



'---------------------------------------------------------------------------------------
Private Declare Function SHCreateStreamOnFileW Lib "shlwapi" (ByVal pszFile As Long, ByVal grfMode As Long, ppStream As IUnknown) As Long
Private Declare Function OleLoadPicture Lib "oleaut32" (ByVal lpStream As IUnknown, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, lplpvObj As IPicture) As Long
'---------------------------------------------------------------------------------------


'---------------------------------------------------------------------------------------
' Module    : Module1
' Author    : beededea
' Date      : 17/08/2021
' Purpose   :
'---------------------------------------------------------------------------------------


' https://www.vbforums.com/showthread.php?473677-VB6-Sorting-algorithms-(sort-array-sorting-arrays)


' API that Allows the execution of command svia the shell
Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

' API that is used to lock the listbox whilst the listboxes are updated from the array
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


'------------------------------------------------------ STARTS
' Types defined for reading process information
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

' Constants defined for reading process information
Private Const TH32CS_SNAPPROCESS As Long = 2&

' Variables defined for reading process information
Private uProcess   As PROCESSENTRY32
Private hSnapshot As Long

' APIs decLAred for reading process information
Private Declare Function ProcessFirst Lib "kernel32.dll" Alias "Process32First" (ByVal hSnapshot As Long, ByRef uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32.dll" Alias "Process32Next" (ByVal hSnapshot As Long, ByRef uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
'------------------------------------------------------ ENDS



'------------------------------------------------------ STARTS
' Constants defined for setting a theme to the prefs
Public Const COLOR_BTNFACE As Long = 15

' APIs declared for setting a theme to the prefs
Public Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function IsThemeActive Lib "uxtheme" () As Boolean
'------------------------------------------------------ ENDS



'------------------------------------------------------ STARTS
'constants defined for querying the registry
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Public Const HKEY_CURRENT_USER = &H80000001
Private Const REG_SZ  As Long = 1                          ' Unicode nul terminated string

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long
' APIs, constants defined for querying the registry ENDS

'------------------------------------------------------ STARTS
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
'' enumerate variables for special folder values
'Public Enum eSpecialFolders
'  SpecialFolder_AppData = &H1A        'for the current Windows user, on any computer on the network [Windows 98 or later]
'  SpecialFolder_CommonAppData = &H23  'for all Windows users on this computer [Windows 2000 or later]
'  SpecialFolder_LocalAppData = &H1C   'for the current Windows user, on this computer only [Windows 2000 or later]
'  SpecialFolder_Documents = &H5       'the Documents folder for the current Windows user
'End Enum
'------------------------------------------------------ ENDS



'------------------------------------------------------ STARTS
'constants used to choose a font via the system dialog window
Public Const LOGPIXELSY As Integer = 90
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const GHND As Long = (GMEM_MOVEABLE Or GMEM_ZEROINIT)
Private Const LF_FACESIZE As Integer = 32
Private Const CF_EFFECTS  As Long = &H100&
Private Const CF_INITTOLOGFONTSTRUCT  As Long = &H40&
Private Const CF_SCREENFONTS As Long = &H1

'type declaration used to choose a font via the system dialog window
Public Type FormFontInfo
  Name As String
  Weight As Integer
  Height As Integer
  UnderLine As Boolean
  Italic As Boolean
  Color As Long
End Type

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName(LF_FACESIZE) As Byte
End Type

Private Type FONTSTRUC
  lStructSize As Long
  hwnd As Long
  hdc As Long
  lpLogFont As Long
  iPointSize As Long
  Flags As Long
  rgbColors As Long
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
  hInstance As Long
  lpszStyle As String
  nFontType As Integer
  MISSING_ALIGNMENT As Integer
  nSizeMin As Long
  nSizeMax As Long
End Type

Private Type ChooseColorStruct
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'APIs used to choose a font via the system dialog window
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" _
(pChoosefont As FONTSTRUC) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" _
  (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetDeviceCaps Lib "gdi32" _
  (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'------------------------------------------------------ ENDS



'------------------------------------------------------ STARTS
'API Functions to read/write information from INI File start
Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long _
    , ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any _
    , ByVal lpString As Any, ByVal lpFileName As String) As Long
'------------------------------------------------------ ENDS



'------------------------------------------------------ STARTS
' Enums defined for opening a common dialog box to select files without OCX dependencies
Private Enum FileOpenConstants
    'ShowOpen, ShowSave constants.
    cdlOFNAllowMultiselect = &H200&
    cdlOFNCreatePrompt = &H2000&
    cdlOFNExplorer = &H80000
    cdlOFNExtensionDifferent = &H400&
    cdlOFNFileMustExist = &H1000&
    cdlOFNHideReadOnly = &H4&
    cdlOFNLongNames = &H200000
    cdlOFNNoChangeDir = &H8&
    cdlOFNNoDereferenceLinks = &H100000
    cdlOFNNoLongNames = &H40000
    cdlOFNNoReadOnlyReturn = &H8000&
    cdlOFNNoValidate = &H100&
    cdlOFNOverwritePrompt = &H2&
    cdlOFNPathMustExist = &H800&
    cdlOFNReadOnly = &H1&
    cdlOFNShareAware = &H4000&
End Enum

' Types defined for opening a common dialog box to select files without OCX dependencies
Public Type OPENFILENAME
    lStructSize As Long    'The size of this struct (Use the Len function)
    hwndOwner As Long       'The hWnd of the owner window. The dialog will be modal to this window
    hInstance As Long            'The instance of the calling thread. You can use the App.hInstance here.
    lpstrFilter As String        'Use this to filter what files are showen in the dialog. Separate each filter with Chr$(0). The string also has to end with a Chr(0).
    lpstrCustomFilter As String  'The pattern the user has choosed is saved here if you pass a non empty string. I never use this one
    nMaxCustFilter As Long       'The maximum saved custom filters. Since I never use the lpstrCustomFilter I always pass 0 to this.
    nFilterIndex As Long         'What filter (of lpstrFilter) is showed when the user opens the dialog.
    lpstrFile As String          'The path and name of the file the user has chosed. This must be at least MAX_PATH (260) character long.
    nMaxFile As Long             'The length of lpstrFile + 1
    lpstrFileTitle As String     'The name of the file. Should be MAX_PATH character long
    nMaxFileTitle As Long        'The length of lpstrFileTitle + 1
    lpstrInitialDir As String    'The path to the initial path :) If you pass an empty string the initial path is the current path.
    lpstrTitle As String         'The caption of the dialog.
    Flags As FileOpenConstants                'Flags. See the values in MSDN Library (you can look at the flags property of the common dialog control)
    nFileOffset As Integer       'Points to the what character in lpstrFile where the actual filename begins (zero based)
    nFileExtension As Integer    'Same as nFileOffset except that it points to the file extention.
    lpstrDefExt As String        'Can contain the extention Windows should add to a file if the user doesn't provide one (used with the GetSaveFileName API function)
    lCustData As Long            'Only used if you provide a Hook procedure (Making a Hook procedure is pretty messy in VB.
    lpfnHook As Long             'Pointer to the hook procedure.
    lpTemplateName As String     'A string that contains a dialog template resource name. Only used with the hook procedure.
End Type

Private Type BROWSEINFO
    hwndOwner As Long
    pidlRoot As Long 'LPCITEMIDLIST
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long  'BFFCALLBACK
    lParam As Long
    iImage As Long
End Type

' vars defined for opening a common dialog box to select files without OCX dependencies
Public x_OpenFilename As OPENFILENAME

' APIs declared for opening a common dialog box to select files without OCX dependencies
Public Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (lpofn As OPENFILENAME) As Long
Private Declare Function SHBrowseForFolderA Lib "Shell32.dll" (bInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDListA Lib "Shell32.dll" (ByVal pidl As Long, ByVal szPath As String) As Long
Private Declare Function CoTaskMemFree Lib "ole32.dll" (lp As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' Constants for hiding/adding horizontal scrollbars to the listboxes
Public Const LB_SETHORIZONTALEXTENT As Long = &H194
Public Const SB_VERT As Long = 1

' APIs for hiding/adding horizontal scrollbars to the listboxes
Public Declare Function SendMessageByNum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' Constants for playing sounds
Public Const SND_ASYNC As Long = &H1         '  play asynchronously
Public Const SND_FILENAME  As Long = &H20000     '  name is a file name

' APIs for playing sounds
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
'------------------------------------------------------ ENDS



'------------------------------------------------------ STARTS
' Type defined for testing a time difference used to initiate one of the hand-coded timers
Public Type LASTINPUTINFO
    cbSize As Long
    dwTime As Long
End Type

' APIs defined for testing a time difference used to initiate one of the hand-coded timers
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function GetLastInputInfo Lib "user32" (plii As Any) As Long
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
Private Const TIME_ZONE_ID_DAYLIGHT As Integer = 2

' Types for determining the timezone
Private Type SYSTEMTIME
    wYear                   As Integer
    wMonth                  As Integer
    wDayOfWeek              As Integer
    wDay                    As Integer
    wHour                   As Integer
    wMinute                 As Integer
    wSecond                 As Integer
    wMilliseconds           As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    bias                    As Long
    StandardName(63)        As Byte
    StandardDate            As SYSTEMTIME
    StandardBias            As Long
    DaylightName(63)        As Byte
    DaylightDate            As SYSTEMTIME
    DaylightBias            As Long
End Type

' APIs for determining the timezone
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function GetMem4 Lib "msvbvm60" (ByRef Source As Any, ByRef Dest As Any) As Long ' Always ignore the returned value, it's useless.
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' APIs defined for creating the hand-coded timers
Public Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' For unicode file names
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
' For unicode file names
Private Declare Function WideCharToMultiByte Lib "kernel32.dll" ( _
                         ByVal CodePage As Long, _
                         ByVal dwFlags As Long, _
                         ByVal lpWideCharStr As Long, _
                         ByVal cchWideChar As Long, _
                         ByVal lpMultiByteStr As Long, _
                         ByVal cbMultiByte As Long, _
                         ByVal lpDefaultChar As Long, _
                         ByVal lpUsedDefaultChar As Long) As Long
                         
Private Declare Function MultiByteToWideChar Lib "kernel32.dll" ( _
                         ByVal CodePage As Long, _
                         ByVal dwFlags As Long, _
                         ByVal lpMultiByteStr As Long, _
                         ByVal cbMultiByte As Long, _
                         ByVal lpWideCharStr As Long, _
                         ByVal cchWideChar As Long) As Long

Private Const CP_UTF8 As Long = 65001
'------------------------------------------------------ ENDS



'------------------------------------------------------ STARTS
'fnIsFileAlreadyOpen
Private Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Private Declare Function lClose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
'------------------------------------------------------ ENDS


'------------------------------------------------------ STARTS
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer
'------------------------------------------------------ ENDS
 
 
'------------------------------------------------------ STARTS
' API functions to scale VB picture objects according to DPI settings
'Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (lpPictDesc As Any, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
'Private Declare Function CopyImage Lib "user32.dll" (ByVal handle As Long, ByVal uType As Long, ByVal cx As Long, ByVal cy As Long, ByVal flags As Long) As Long
'------------------------------------------------------ ENDS

    Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long

    Public Enum FolderEnum
        feCDBurnArea = 59 ' \Docs & Settings\User\Local Settings\Application Data\Microsoft\CD Burning
        feCommonAppData = 35 ' \Docs & Settings\All Users\Application Data
        feCommonAdminTools = 47 ' \Docs & Settings\All Users\Start Menu\Programs\Administrative Tools
        feCommonDesktop = 25 ' \Docs & Settings\All Users\Desktop
        feCommonDocs = 46 ' \Docs & Settings\All Users\Documents
        feCommonPics = 54 ' \Docs & Settings\All Users\Documents\Pictures
        feCommonMusic = 53 ' \Docs & Settings\All Users\Documents\Music
        feCommonStartMenu = 22 ' \Docs & Settings\All Users\Start Menu
        feCommonStartMenuPrograms = 23 ' \Docs & Settings\All Users\Start Menu\Programs
        feCommonTemplates = 45 ' \Docs & Settings\All Users\Templates
        feCommonVideos = 55 ' \Docs & Settings\All Users\Documents\My Videos
        feLocalAppData = 28 ' \Docs & Settings\User\Local Settings\Application Data
        feLocalCDBurning = 59 ' \Docs & Settings\User\Local Settings\Application Data\Microsoft\CD Burning
        feLocalHistory = 34 ' \Docs & Settings\User\Local Settings\History
        feLocalTempInternetFiles = 32 ' \Docs & Settings\User\Local Settings\Temporary Internet Files
        feProgramFiles = 38 ' \Program Files
        feProgramFilesCommon = 43 ' \Program Files\Common Files
        'feRecycleBin = 10 ' ???
        feUser = 40 ' \Docs & Settings\User
        feUserAdminTools = 48 ' \Docs & Settings\User\Start Menu\Programs\Administrative Tools
        feUserAppData = 26 ' \Docs & Settings\User\Application Data
        feUserCache = 32 ' \Docs & Settings\User\Local Settings\Temporary Internet Files
        feUserCookies = 33 ' \Docs & Settings\User\Cookies
        feUserDesktop = 16 ' \Docs & Settings\User\Desktop
        feUserDocs = 5 ' \Docs & Settings\User\My Documents
        feUserFavorites = 6 ' \Docs & Settings\User\Favorites
        feUserMusic = 13 ' \Docs & Settings\User\My Documents\My Music
        feUserNetHood = 19 ' \Docs & Settings\User\NetHood
        feUserPics = 39 ' \Docs & Settings\User\My Documents\My Pictures
        feUserPrintHood = 27 ' \Docs & Settings\User\PrintHood
        feUserRecent = 8 ' \Docs & Settings\User\Recent
        feUserSendTo = 9 ' \Docs & Settings\User\SendTo
        feUserStartMenu = 11 ' \Docs & Settings\User\Start Menu
        feUserStartMenuPrograms = 2 ' \Docs & Settings\User\Start Menu\Programs
        feUserStartup = 7 ' \Docs & Settings\User\Start Menu\Programs\Startup
        feUserTemplates = 21 ' \Docs & Settings\User\Templates
        feUserVideos = 14  ' \Docs & Settings\User\My Documents\My Videos
        feWindows = 36 ' \Windows
        feWindowFonts = 20 ' \Windows\Fonts
        feWindowsResources = 56 ' \Windows\Resources
        feWindowsSystem = 37 ' \Windows\System32
    End Enum

'------------------------------------------------------ STARTS
Public Declare Function DwmGetWindowAttribute Lib "dwmapi.dll" _
            (ByVal hwnd As Long, ByVal dwAttribute As Long, ByRef pvAttribute As Any, ByVal cbAttribute As Long) As Long
'------------------------------------------------------ ENDS



Public borderSizeLeft As Long
Public borderSizeRight As Long
Public borderSizeTop As Long
Public borderSizeBottom As Long

' These are the IDs of the hand-coded timers.
Private pollingTimerID As Long
Private iconiseTimerID As Long
Private emailTimerID As Long
Private houseKeepingTimerID As Long



' required for LaVolpes reading of 'foreign' image types
Public cImage As c32bppDIB
Public origWidth As Long
Public origHeight As Long

' General variables declared
Public FCWSettingsDir As String
Public FCWSettingsFile As String
' Public toolSettingsFile  As String
Public debugflg As Integer
Public classicThemeCapable As Boolean
'Public suppliedFont As String
'Public suppliedSize As Integer
'Public suppliedWeight As Integer
'Public suppliedStyle As Boolean
'Public suppliedColour As Long
'Public suppliedItalics As Boolean
'Public suppliedUnderline As Boolean
'Public fontSelected As Boolean
Public FCWSkinTheme As String
Public storeThemeColour As Long
Public inputDataChangedFlag As Boolean
Public outputDataChangedFlag As Boolean
Public idleTime As Long
Public currindex As Integer
Public toolTipFlag As Boolean
Public validImageArrayList As Collection
Public executableSuffixArrayList As Collection
Public invalidImageArrayList As Collection
Public attachmentFilePath As String
Public recordingFilePath As String
Public displayedAttachmentFilePath As String
Public dropboxErrCnt As Integer
Public attachmentViewTime As Date
Public recordingViewTime As Date
Public nowBeingModifiedFlag As Boolean
Public pollingFlag As Boolean
Public msgBoxShowing As String
Public oldOutputLineCount As Long
Public inputLineCount As Long
Public outputLineCount As Long
Public combinedArrayCount As Long

Public inputFileModificationTime As Date
Public oldInputFileModificationTime As Date
Public outputFileModificationTime As Date
Public oldOutputFileModificationTime As Date

Public WindowsVer As String
Public remoteNetworkDisabled As Boolean


    
Private fso As Object
Private mbDebugMode As Boolean

' Vars for the arrays to store and sort the user data
Private inputFileArray() As String
Public outputFileArray() As String
Public combinedFileArray() As String

' Config Vars for storing the data
Public FCWSharedInputFile As String
Public FCWSharedOutputFile As String
Public FCWExchangeFolder As String
Public FCWRefreshIntervalIndex As String
Public FCWRefreshIntervalSecs As String
Public FCWAlarmSound As String
Public FCWAlarmSoundIndex As String

Public FCWPrefixString As String
Public FCWLoadBottom As String
Public FCWMaxLineLengthIndex As String
Public FCWMaxLineLength As String
Public FCWEnableScrollbars As String
Public FCWEnableTooltips As String
Public FCWEnableBalloonTooltips As String

Public FCWIconiseDelay As String
Public FCWSendEmails      As String
Public FCWSendErrorEmails  As String
Public FCWAdviceInterval  As String
Public FCWAdviceIntervalSecs As String
Public FCWLastEmail As String
Public FCWLastHouseKeep As String


Public FCWEmojiSetIndex As String
Public FCWEmojiSetDesc As String

Public FCWMainFont  As String
Public FCWMainFontSize  As String
Public FCWMainFontItalics  As String
'    fntWeight
Public FCWMainFontColour  As String
Public FCWMainFontUnderline As String

Public FCWPrefsFont  As String
Public FCWPrefsFontSize As String
Public FCWPrefsFontItalics  As String
Public FCWPrefsFontColour  As String

Public FCWWindowLevel As String
Public FCWOpacity  As String
Public FCWEnableSounds  As String
Public FCWEnableAlarmSound  As String

Public FCWPlayVolume  As String
Public FCWMinimiseFormX  As String
Public FCWMinimiseFormY  As String
Public FCWMaximiseFormX  As String
Public FCWMaximiseFormY  As String
Public FCWFormWidth  As String

Public FCWLastSoundPlayed As String
Public FCWLastPingResponse As String
Public FCWLastAwakeString As String
Public FCWLastShutdown As String
Public FCWAllowShutdowns As String
Public FCWClockStyle  As String

Public FCWSmtpServer As String
Public FCWSmtpConfig As String
Public FCWSmtpConfigName As String


Public FCWRecipientEmail As String
Public FCWEmailSubject As String
Public FCWEmailMessage As String
Public FCWSmtpUsername As String
Public FCWSmtpPassword As String
Public FCWSmtpPort As String
Public FCWSmtpAuthenticate As String
Public FCWSmtpSecurity As String

Public FCWSingleListBox As String
Public FCWImageDisplay As String
Public FCWOptHandleData As String
Public FCWOptWindowWidth As String
Public FCWAutomaticHousekeeping As String
Public FCWStartup As String

Public FCWArchiveDays As String
Public FCWArchiveDaysIndex As String

Public FCWBackupOnStart As String
Public FCWAutomaticBackups As String
Public FCWAutomaticBackupInterval As String
Public FCWServiceProvider As String
Public FCWCheckServiceProcesses As String
Public FCWCaptureDevices As String
Public FCWCaptureDevicesIndex As String
Public FCWRecordingQuality As String
Public FCWLastSelectedTab As String
Public FCWIconiseOpacity As String
Public FCWIconiseDesktop As String

Public FCWArchiveFolder As String
Public FCWBackupFolder As String

Public FCWMsgBox13Enabled As String


Public backupTimerCount As Integer

Public CTRL_1 As Boolean

Public messageQueue As Collection

Public ioMethodADO As Boolean

Public SoundName As String

'Public screenTwipsPerPixelX As Long ' .07 DAEB 26/04/2021 common.bas changed to use pixels alone, removed all unnecessary twip conversion
'Public screenTwipsPerPixelY As Long ' .07 DAEB 26/04/2021 common.bas changed to use pixels alone, removed all unnecessary twip conversion

Public recordingIsPossible As Boolean
Public binaryFlag As Boolean

Public emailTString As String

Public Bas64 As Base64

Public Const LNG_PROXY_PORT As Long = 10025

Public msgBoxOut As Boolean
Public msgLogOut As Boolean

Public currentOpacity As Integer


' called during startup to start the polling timers either VB6 or in code
'---------------------------------------------------------------------------------------
' Procedure : startThePollingTimers
' Author    : beededea
' Date      : 01/03/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub startThePollingTimers()

    Dim pollingIntervalMillisecs As Long
    
    Const lngSecs As Long = 65 ' just used to avoid multiplying two integers
    Const lngThousand As Long = 1000
    
    On Error GoTo startThePollingTimers_Error

    If Val(FCWRefreshIntervalSecs) = 0 Then
        debugLog "%Err-I-ErrorNumber 21 - The polling timer is not active, the prefs are set to No Timed Refresh" & vbCrLf & "Increase value if you want it to poll for new data,"
        Exit Sub
    End If
    ' start the polling timer in code
    If fInIDE Then
        ' VB6 timers cannot exceed 65 seconds (65535 ms)
        If Val(FCWRefreshIntervalSecs) > 65 Then
'            lngSecs = 65
'            lngThousand = 1000
            ' when multiplying two integer values and assigning to a long in the IDE it causes a failure as the IDE is handling the two numbers as integers
            ' pollingIntervalMillisecs = 65 * 1000 '  < this fails
            pollingIntervalMillisecs = lngSecs * lngThousand ' works!
        Else
            pollingIntervalMillisecs = Val(FCWRefreshIntervalSecs) * 1000
        End If
        FireCallMain.pollingTimer.Interval = pollingIntervalMillisecs
        FireCallMain.pollingTimer.Enabled = True
        debugLog "STARTING startPollingTimer using VB6 timer, ID = " & pollingTimerID & " at interval of " & pollingIntervalMillisecs & "ms", False

    Else
        ' using a timer in code rather than a VB6 timer as VB6 timers cannot exceed 65 seconds (65535 ms)
        ' and if you want a longer timer you have to roll your own.
        ' in addition, unfortunately the manual code timer method does not work in the IDE
        'Call pollingTimer_CodeTimer
                
        ' stop any possible running timer first
        Call stopPollingTimer
        
        ' using a timer in code rather than a VB6 timer as VB6 timers cannot exceed 65 seconds (65535 ms)
        'pollingIntervalMillisecs = FireCallPrefs.cmbRefreshInterval.ItemData(Val(FCWRefreshIntervalIndex)) * 1000
        pollingIntervalMillisecs = Val(FCWRefreshIntervalSecs) * 1000
        
        ' prevent starting this timer when working in the IDE
        If Not fInIDE Then
            'startPollingTimer pollingIntervalMillisecs
            ' Don't start the timer If it's already running.
            If pollingTimerID = 0 Then
                ' this has a callback routine that it jumps to on each interval completion
                pollingTimerID = SetTimer(0, 1, pollingIntervalMillisecs, AddressOf pollingTimer_CodeTimer)
                debugLog "STARTING startPollingTimer using API timer, ID = " & pollingTimerID & " at interval of " & pollingIntervalMillisecs & "ms", False
            End If
        Else ' only needs to be stated once
            MsgBox "Please note: Timers in code will not run in the IDE, defaulting to VB6 timers <65secs."
        End If
        
    End If

    On Error GoTo 0
    Exit Sub

startThePollingTimers_Error:

    With err
         If .Number <> 0 Then
            MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure startThePollingTimers of Module modCommon"
            Resume Next
          End If
    End With
End Sub



' Callback routine called byAddressOf used by the polling timer. Note: This function only operates at runtime,
' ie. it doesn't work in the IDE because in the IDE everything works in the main thread. Callback functions operate
' in a separate thread and this achieves basic multi threading but may limit some functionality - but basic commands seem to operate correctly

Public Sub pollingTimer_CodeTimer()
    Call pollingTimer_TimerLogic
End Sub

' the polling timer logic itself that is used by both the VB6 standard timer in the IDE and the hand-crafted timer at runtime
' the logic is in a separate routine as it is called directly by both the VB6 timer and the callback timer

Public Sub pollingTimer_TimerLogic()

    Call debugLog("Polled at " & Now, False)


    If Val(FCWRefreshIntervalSecs) = 0 Then Exit Sub

    If nowBeingModifiedFlag = True Then Exit Sub ' this is a switch set during sendSomething that allows/disallows the polling timer logic to run

    pollingFlag = True ' flag to indicate that polling is underway
    ' light the lamp for 5 seconds
    FireCallMain.picTimerLampBright.Visible = True
    FireCallMain.picTimerLampDull.Visible = False

    FireCallMain.lampTimer.Enabled = True ' this timer turns the lamp off after another 5 secs
    
    If FCWCheckServiceProcesses = "1" Then
        If fCheckDropboxRunning = False Then
            Exit Sub ' no point in polling for data when dropbox is unavailable
        Else
            ' if the carrying network (dropbox) was found to be turned off at startup then the timers will be disabled
            ' this test will automatically restart the polling timers when DB is found to be running again
            If remoteNetworkDisabled = True Then
                remoteNetworkDisabled = False
                
                debugLog "Found carrying network is disabled - restarting the polling timers", False
                ' restart the polling timers
                Call startThePollingTimers
            End If
        End If
    End If
    
    Call checkAndReadInputFile
    Call checkAndReadOutputFile ' no need to read the output file periodically
    If FCWSingleListBox = "1" Then Call populateCombinedBox
    
    Call debugLog("Completed polling " & Now, False)
    pollingFlag = False

End Sub



' The timer that stops the polling timer
Public Sub stopPollingTimer()
    ' Don't stop the timer If it isn't running.
    If pollingTimerID Then
        KillTimer 0, pollingTimerID
        pollingTimerID = 0
    End If
End Sub


' this is the iconising timer initiate routine that both stops and starts the timer
' note this timer only runs at runtime and not in the IDE
Public Sub initiateIconiseTimerInCode()

    Dim iconiseIntervalMillisecs As Long
    
    Call stopIconiseTimer
    ' using a timer in code rather than a VB6 timer as VB6 timers cannot exceed 65 seconds (65535 ms)
    iconiseIntervalMillisecs = Val(FCWIconiseDelay) * 1000
    ' disable this timer when working in the IDE
    If Not fInIDE Then
        If Val(FCWIconiseDelay) > 0 Then
            'MsgBox "startIconiseTimer " & Val(FCWIconiseDelay)
            startIconiseTimer iconiseIntervalMillisecs
        End If
    End If
End Sub
' The actual iconising timer that starts the timer
Public Sub startIconiseTimer(ByVal Timeout As Long)
    
    Call stopIconiseTimer
    If iconiseTimerID = 0 Then
        iconiseTimerID = SetTimer(0, 2, Timeout, AddressOf iconiseTimer_TimerA)
        'MsgBox "STARTING startIconiseTimer " & iconiseTimerID & " at interval of " & timeout & "ms"
    End If
End Sub
' The timer that stops the iconising timer
Public Sub stopIconiseTimer()
    ' Don't stop the timer If it isn't running.
    If iconiseTimerID Then
        KillTimer 0, iconiseTimerID
        iconiseTimerID = 0
    End If
End Sub

' Callback routine called byAddressOf used by the iconising timer. Note: This function only operates at runtime,
' ie. it doesn't work in the IDE because in the IDE everything works in the main thread. Callback functions operate
' in a separate thread and this achieves basic multi threading but may limit some functionality - but basic commands seem to operate correctly
Public Sub iconiseTimer_TimerA()

    If Val(FCWIconiseDelay) = 0 Then
        Call stopIconiseTimer
        Exit Sub
    End If

    Call getIdleTime

    If idleTime > Val(FCWIconiseDelay) * 1000 Then
        'MsgBox "timer fired"
        If FCWIconiseDesktop = "True" Then
            FireCallMain.opacityFadeOutTimer.Enabled = True
            MinimiseForm.Visible = True
'        Else
'            ' just set the opacity of the window
'            Call setMainWindowOpacity
        End If
        Call stopIconiseTimer
        
    End If

End Sub

Public Sub getIdleTime()
    Dim lastInputVar As LASTINPUTINFO
    
    lastInputVar.cbSize = Len(lastInputVar)
    Call GetLastInputInfo(lastInputVar)
    
    idleTime = GetTickCount - lastInputVar.dwTime
End Sub
    

' The actual polling timer that starts the timer
'Public Sub startPollingTimer(ByVal Timeout As Long)
'    ' Don't start the timer If it's already running.
'    If pollingTimerID = 0 Then
'        pollingTimerID = SetTimer(0, 0, Timeout, AddressOf pollingTimer_TimerA)
'        'MsgBox "STARTING startPollingTimer " & pollingTimerID & " at interval of " & timeout & "ms"
'    End If
'End Sub
'
'' The timer that stops the polling timer
'Public Sub stopPollingTimer()
'    ' Don't stop the timer If it isn't running.
'    If pollingTimerID Then
'        KillTimer 0, pollingTimerID
'        pollingTimerID = 0
'    End If
'End Sub

' set the opacity of the main window, emulating functionality of the YWE version
Private Sub setMainWindowOpacity()
    
    'Dim Opacity As Integer
    
    FireCallMain.opacityToTimer.Enabled = True

    'Opacity = 255 * (Val(FCWOpacity) / 100)
    
    'Call setOpacity(Opacity)

End Sub
 

 ' VB6 opacity at form level achieved using APIs
 Public Sub setOpacity(ByVal Opacity As Long)
    currentOpacity = Opacity
    Call SetWindowLong(FireCallMain.hwnd, GWL_EXSTYLE, GetWindowLong(FireCallMain.hwnd, GWL_EXSTYLE) Or WS_EX_LAYERED)
    Call SetLayeredWindowAttributes(FireCallMain.hwnd, RGB(255, 0, 255), Opacity, LWA_ALPHA Or LWA_COLORKEY)
End Sub

' checks the existence of the output file, the local user's file, checks the linecount and reads the data into the array and thence to the listbox
Private Sub checkAndReadOutputFile()

    Dim timeDifferenceInSecs As Long ' max 86 years as a LONG in secs
    
    timeDifferenceInSecs = 0
    
    If Not fFExists(FCWSharedOutputFile) Then
        If dropboxErrCnt >= 2 Then
            MsgBox ("%Err-I-ErrorNumber 12 - FCW was unable to access the shared output file. " & vbCrLf & FCWSharedOutputFile & vbCrLf & " with " & dropboxErrCnt & " attempts")
            dropboxErrCnt = 0
            'Call btnConfig_Click
            Exit Sub
        Else
            dropboxErrCnt = dropboxErrCnt + 1
        End If
    End If
        
    ' now do a quick lineCount check on the input file just in case you are using the .js version
    ' simultaneously. It does not matter if this is just tacked onto reading the input file
    ' and does not run every timer run as it is an unlikely event being catered for.
        
    outputFileModificationTime = FileDateTime(FCWSharedOutputFile)
    
    ' there are a lot of checks and tests here due to a runtime bug that occasionally rears its head
    
    ' on occasion when the program is left running (one day) we sometimes encounter a problem where the
    ' modification lamp repeatedly relights and then shortly afterward on an attempt to repoll, results in an
    ' overflow error being displayed. This should not ever be associated with the timeDifferenceInSecs variable
    ' as it is a LONG.
    
    ' We have a temporary error handling to try to compensate for the problem until it is resolved.
    ' It is as if the datediff produces an output that is a forced integer result, 32768, about 9.5hrs in secs.
    ' it does not always occur.
    
    If IsNull(outputFileModificationTime) Then
        MsgBox "outputFileModificationTime " & outputFileModificationTime & " is null!"
        Exit Sub
    End If
    
    If outputFileModificationTime = 0 Then
        MsgBox "outputFileModificationTime " & outputFileModificationTime & " is zero!"
        Exit Sub
    End If
    
    If outputFileModificationTime = #12/31/1899# Then
        MsgBox "outputFileModificationTime " & outputFileModificationTime & " is 12/31/1899!"
        Exit Sub
    End If
    
    If oldOutputFileModificationTime = #12/31/1899# Then
        MsgBox "oldoutputFileModificationTime " & oldOutputFileModificationTime & " is 12/31/1899!"
        Exit Sub
    End If
    
    timeDifferenceInSecs = DateDiff("s", oldOutputFileModificationTime, outputFileModificationTime)

    If timeDifferenceInSecs = 0 Then
        Exit Sub  ' to minimise CPU usage just exit
    End If
    
    outputLineCount = fLineCount(FCWSharedOutputFile)
    'If oldOutputLineCount = outputLineCount Then Exit Sub ' to minimise CPU usage just exit
    
    Call readOutputFileWriteArrayWriteListbox(FCWSharedOutputFile)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : checkAndReadInputFile
' Author    : beededea
' Date      : 03/08/2021
' Purpose   : checks the existence of the input file, the remote user's file, checks the
'             time/date stamp and reads the data into the array and thence to the listbox
'---------------------------------------------------------------------------------------
'
Private Sub checkAndReadInputFile()
    
    On Error GoTo l_checkAndReadInputFile_Error ' catch any error
    
    Dim timeDifferenceInSecs  As Long: timeDifferenceInSecs = 0 ' max 86 years as a LONG in secs

    If Not fFExists(FCWSharedInputFile) Then
        If dropboxErrCnt >= 2 Then
            If FCWMsgBox13Enabled = "1" Then MsgBox ("%Err-I-ErrorNumber 13 - FCW was unable to access the shared input file. " & vbCrLf & FCWSharedInputFile & vbCrLf & " with " & dropboxErrCnt & " attempts")
            dropboxErrCnt = 0
            Exit Sub
         Else
            dropboxErrCnt = dropboxErrCnt + 1
        End If
    End If
        
    inputFileModificationTime = FileDateTime(FCWSharedInputFile)
    
    ' there are a lot of checks and tests here due to a runtime bug that occasionally rears its head
    
    ' on occasion when the program is left running (one day) we sometimes encounter a problem where the
    ' modification lamp repeatedly relights and then shortly afterward on an attempt to repoll, results in an
    ' overflow error being displayed. This should not ever be associated with the timeDifferenceInSecs variable
    ' as it is a LONG.
    
    ' We have a temporary error handling to try to compensate for the problem until it is resolved.
    ' It is as if the datediff produces an output that is a forced integer result, 32768, about 9.5hrs in secs.
    ' it does not always occur.
    
    If IsNull(inputFileModificationTime) Then
        MsgBox "inputFileModificationTime " & inputFileModificationTime & " is null!"
        Exit Sub
    End If
    
    If inputFileModificationTime = 0 Then
        MsgBox "inputFileModificationTime " & inputFileModificationTime & " is zero!"
        Exit Sub
    End If
    
    If inputFileModificationTime = #12/31/1899# Then
        MsgBox "inputFileModificationTime " & inputFileModificationTime & " is 12/31/1899!"
        Exit Sub
    End If
    
    If oldInputFileModificationTime = #12/31/1899# Then
        MsgBox "oldinputFileModificationTime " & oldInputFileModificationTime & " is 12/31/1899!"
        Exit Sub
    End If
    
    timeDifferenceInSecs = DateDiff("s", oldInputFileModificationTime, inputFileModificationTime)

    If timeDifferenceInSecs = 0 Then
        Exit Sub  ' to minimise CPU usage just exit
    End If
    GoTo l_getInputLineCount ' bypass the error reporting on a normal condition
       
    
l_checkAndReadInputFile_Error: ' error handling location
    On Error GoTo 0
    debugLog "%Err-I-ErrorNumber 13 - the overflow error occurred and was handled"
    debugLog "oldinputFileModificationTime " & oldInputFileModificationTime
    debugLog "inputFileModificationTime " & inputFileModificationTime
    timeDifferenceInSecs = 0 ' reset the var that seems to be causing the problem
    
    If fFExists(FCWSettingsFile) Then ' write the overflow error to a file so we can see the data as it was when the overflow occurred
        PutINISetting "Software\FireCallWin", "oldinputFileModificationTime", oldInputFileModificationTime, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "inputFileModificationTime", inputFileModificationTime, FCWSettingsFile
        PutINISetting "Software\FireCallWin", "timeDifferenceInSecs", timeDifferenceInSecs, FCWSettingsFile
    End If
    Unload FireCallMain
    'Call Form_Unload_Sub ' exit the program

l_getInputLineCount: ' continue location

    ' on rare occasions it has still got this far with no apparent textual changes to the input file
    ' I have to assume that the partner, Harry had edited his file resulting in a change to the linecount
    
    MinimiseForm.pulseTimer.Enabled = True
    inputDataChangedFlag = True
    FireCallMain.lbxInputTextArea.Refresh
    
    FireCallMain.picTextChangeBright.Visible = True
    FireCallMain.picTextChangeDull.Visible = False
    
    Call readInputFileAndWriteArray(FCWSharedInputFile)


End Sub

' if Dropbox is unavailable, the process not running, then say so.
Public Function fCheckDropboxRunning() As Boolean
    Dim alarmFile As String
    Dim answer As VbMsgBoxResult
    
    'initialise the dimensioned variables
    answer = vbNo
    
    alarmFile = App.Path & "\Resources\Sounds\" & FCWAlarmSound
        
    If Not fIsRunning("dropbox.exe", vbNull) Then
        If FCWEnableAlarmSound = "1" Then
            If fFExists(alarmFile) Then PlaySound alarmFile, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
        
        'Call dropboxSendMail
            
        If msgBoxShowing = False Then
            msgBoxShowing = True
            answer = MsgBox("%Err-I-ErrorNumber 14 - Sharing is not currently active. Outgoing messages will be saved but will not progress further.", vbExclamation + vbOK)
            If answer <> 0 Then
                msgBoxShowing = False
            End If
        End If
        
        fCheckDropboxRunning = False
    Else
        fCheckDropboxRunning = True
    End If


End Function



'---------------------------------------------------------------------------------------
' Procedure : fGetINISetting
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : Get the INI Setting from the File
'---------------------------------------------------------------------------------------
'
Public Function fGetINISetting(ByVal sHeading As String, ByVal sKey As String, ByRef sINIFileName As String) As String
   On Error GoTo fGetINISetting_Error
    Const cparmLen As Integer = 500  ' maximum no of characters allowed in the returned string
    Dim sReturn As String * cparmLen
    Dim sDefault As String * cparmLen
    Dim lLength As Long
    
    ' we no longer use GetPrivateProfileString for reading all the vars as it cannot read certain special chars in the values
    ' the sort that might be generated by such things as the encryption routine.

    lLength = GetPrivateProfileString(sHeading, sKey, sDefault, sReturn, cparmLen, sINIFileName)
    fGetINISetting = Mid(sReturn, 1, lLength)

   On Error GoTo 0
   Exit Function

fGetINISetting_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure fGetINISetting of Module Common"
End Function

'
'---------------------------------------------------------------------------------------
' Procedure : PutINISetting
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : Save INI Setting in the File
'---------------------------------------------------------------------------------------
'
Public Sub PutINISetting(ByVal sHeading As String, ByVal sKey As String, ByVal sSetting As String, ByRef sINIFileName As String)

   On Error GoTo PutINISetting_Error

    Dim aLength As Long
    
    aLength = WritePrivateProfileString(sHeading, sKey _
            , sSetting, sINIFileName)

   On Error GoTo 0
   Exit Sub

PutINISetting_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure PutINISetting of Module Common"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : fFExists
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : file existence tester
'---------------------------------------------------------------------------------------
'
Public Function fFExists(ByRef OrigFile As String) As Boolean
    On Error GoTo fFExists_Error
    'If debugflg = 1 Then Debug.Print "%fnFExists"

    'Dim fileSystemObj As Object
'    Set fileSystemObj = CreateObject("Scripting.FileSystemObject")
'    fFExists = fileSystemObj.FileExists(OrigFile)
    
    ' test to see if a file exists
    Const INVALID_HANDLE_VALUE = -1&
    fFExists = Not (GetFileAttributesW(StrPtr(OrigFile)) = INVALID_HANDLE_VALUE)

   On Error GoTo 0
   Exit Function

fFExists_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure fFExists of Module Common"
End Function


'---------------------------------------------------------------------------------------
' Procedure : fDirExists
' Author    : beededea
' Date      : 17/10/2019
' Purpose   : folder/dir existence tester
'---------------------------------------------------------------------------------------
'
Public Function fDirExists(ByRef OrigFile As String) As Boolean
   On Error GoTo fDirExists_Error
   'If debugflg = 1 Then DebugPrint "%fDirExists"
   
   fDirExists = (GetFileAttributes(OrigFile) And vbDirectory + vbVolume) = vbDirectory
   
'    Dim fileSystemObj As Object
'    Set fileSystemObj = CreateObject("Scripting.FileSystemObject")
'    fDirExists = fileSystemObj.FolderExists(OrigFile)

   On Error GoTo 0
   Exit Function

fDirExists_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure fDirExists of Module Common"
End Function

'---------------------------------------------------------------------------------------
' Procedure : fDialogFont
' Author    : beededea
' Date      : 21/08/2020
' Purpose   : display the default windows dialog box that allows the user to select a font
'---------------------------------------------------------------------------------------
'
Public Function fDialogFont(ByRef f As FormFontInfo) As Boolean
      
    ' variables declared
    Dim logFnt As LOGFONT
    Dim ftStruc As FONTSTRUC
    Dim lLogFontAddress As Long
    Dim lMemHandle As Long
    Dim hWndAccessApp As Long
    
     On Error GoTo fDialogFont_Error
    
    logFnt.lfWeight = f.Weight
    logFnt.lfItalic = f.Italic * -1
    logFnt.lfUnderline = f.UnderLine * -1
    logFnt.lfHeight = -fMulDiv(CLng(f.Height), GetDeviceCaps(GetDC(hWndAccessApp), LOGPIXELSY), 72)
    Call StringToByte(f.Name, logFnt.lfFaceName())
    ftStruc.rgbColors = f.Color
    ftStruc.lStructSize = Len(ftStruc)
    
    lMemHandle = GlobalAlloc(GHND, Len(logFnt))
    If lMemHandle = 0 Then
      fDialogFont = False
      Exit Function
    End If

    lLogFontAddress = GlobalLock(lMemHandle)
    If lLogFontAddress = 0 Then
      fDialogFont = False
      Exit Function
    End If
    
    CopyMemory ByVal lLogFontAddress, logFnt, Len(logFnt)
    ftStruc.lpLogFont = lLogFontAddress
    'ftStruc.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT
    ftStruc.Flags = CF_SCREENFONTS Or CF_INITTOLOGFONTSTRUCT
    If ChooseFont(ftStruc) = 1 Then
      CopyMemory logFnt, ByVal lLogFontAddress, Len(logFnt)
      f.Weight = logFnt.lfWeight
      f.Italic = CBool(logFnt.lfItalic)
      f.UnderLine = CBool(logFnt.lfUnderline)
      f.Name = fByteToString(logFnt.lfFaceName())
      f.Height = CLng(ftStruc.iPointSize / 10)
      f.Color = ftStruc.rgbColors
      fDialogFont = True
    Else
      fDialogFont = False
    End If

   On Error GoTo 0
   Exit Function

fDialogFont_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure fDialogFont of Module modCommon"
End Function


'---------------------------------------------------------------------------------------
' Procedure : fMulDiv
' Author    :
' Date      : 21/08/2020
' Purpose   :  fMulDiv function multiplies two 32-bit values and then divides the 64-bit result by a third 32-bit value.
'---------------------------------------------------------------------------------------
'
Private Function fMulDiv(ByVal In1 As Long, ByVal In2 As Long, ByVal In3 As Long) As Long
    
    ' variables declared
    Dim lngTemp As Long
   On Error GoTo fMulDiv_Error

  On Error GoTo fMulDiv_err
  If In3 <> 0 Then
    lngTemp = In1 * In2
    lngTemp = lngTemp / In3
  Else
    lngTemp = -1
  End If
fMulDiv_end:
  fMulDiv = lngTemp
  Exit Function
fMulDiv_err:
  lngTemp = -1
  Resume fMulDiv_err

   On Error GoTo 0
   Exit Function

fMulDiv_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure fMulDiv of Module modCommon"
End Function



'---------------------------------------------------------------------------------------
' Procedure : StringToByte
' Author    : beededea
' Date      : 21/08/2020
' Purpose   : convert a provided string to a byte array
'---------------------------------------------------------------------------------------
'
Private Sub StringToByte(ByVal InString As String, ByRef ByteArray() As Byte)
    
    ' variables declared
    Dim intLbound As Integer
    Dim intUbound As Integer
    Dim intLen As Integer
    Dim intX As Integer
    On Error GoTo StringToByte_Error

    intLbound = LBound(ByteArray)
    intUbound = UBound(ByteArray)
    intLen = Len(InString)
    If intLen > intUbound - intLbound Then intLen = intUbound - intLbound
    For intX = 1 To intLen
        ByteArray(intX - 1 + intLbound) = Asc(Mid(InString, intX, 1))
    Next

   On Error GoTo 0
   Exit Sub

StringToByte_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure StringToByte of Module modCommon"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fByteToString
' Author    :
' Date      : 21/08/2020
' Purpose   : convert a byte array provided to a string
'---------------------------------------------------------------------------------------
'
Private Function fByteToString(ByRef aBytes() As Byte) As String
      
    ' variables declared
    Dim dwBytePoint As Long
    Dim dwByteVal As Long
    Dim szOut As String
    On Error GoTo fByteToString_Error

    dwBytePoint = LBound(aBytes)
    While dwBytePoint <= UBound(aBytes)
      dwByteVal = aBytes(dwBytePoint)
      If dwByteVal = 0 Then
        fByteToString = szOut
        Exit Function
      Else
        szOut = szOut & Chr$(dwByteVal)
      End If
      dwBytePoint = dwBytePoint + 1
    Wend
    fByteToString = szOut

   On Error GoTo 0
   Exit Function

fByteToString_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure fByteToString of Module modCommon"
End Function


'---------------------------------------------------------------------------------------
' Procedure : LoadFileToTB
' Author    : beededea
' Date      : 26/08/2019
' Purpose   :     'PURPOSE: Loads file specified by FilePath into textcontrol
    '(e.g., Text Box, Rich Text Box) specified by TxtBox
'---------------------------------------------------------------------------------------
'
Public Sub LoadFileToTB(ByRef TxtBox As Object, ByVal FilePath As String, Optional ByVal Append As Boolean = False)
       
    'If Append = true, then loaded text is appended to existing
    ' contents else existing contents are overwritten
    
    'Returns: True if Successful, false otherwise
    
    Dim iFile As Integer
    Dim readLine As String
    
   On Error GoTo LoadFileToTB_Error
      'If debugflg = 1 Then Debug.Print "%" & "LoadFileToTB"
   
   
   'If debugflg = 1 Then Debug.Print "%" & LoadFileToTB

    If Dir$(FilePath) = vbNullString Then Exit Sub
    
    On Error GoTo ErrorHandler:
    readLine = TxtBox.Text
    
    iFile = FreeFile
    Open FilePath For Input As #iFile
    readLine = Input(LOF(iFile), #iFile)
    If Append Then
        TxtBox.Text = TxtBox.Text & readLine
    Else
        TxtBox.Text = readLine
    End If
    
    'LoadFileToTB = True
    
ErrorHandler:
    If iFile > 0 Then Close #iFile

   On Error GoTo 0
   Exit Sub

LoadFileToTB_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure LoadFileToTB of Form common"

End Sub





' Procedure : fGetstring
' Author    : beededea
' Date      : 05/07/2019
' Purpose   : get a string from the registry
'---------------------------------------------------------------------------------------
'
Public Function fGetstring(ByRef hKey As Long, ByRef strPath As String, ByRef strvalue As String) As String

    Dim keyhand As Long
    'Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim rvar As Integer
    'in .NET the variant type will need to be replaced by object? This code will go altogether as .NET has native functions to read the registry

    Dim lValueType As Variant

   ' On Error GoTo fGetstring_Error

    rvar = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strvalue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String$(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strvalue, 0&, 0&, ByVal strBuf, lDataBufSize)
        Dim ERROR_SUCCESS As Variant
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                fGetstring = Left$(strBuf, intZeroPos - 1)
            Else
                fGetstring = strBuf
            End If
        End If
    End If

   On Error GoTo 0
   Exit Function

fGetstring_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure fGetstring of Module Common"
End Function

'----------------------------------------
'Name: TestWinVer
'Description: Tests the multiplicity of Windows versions and returns some values
'----------------------------------------
Public Sub testWinVer(ByRef classicThemeCapable As Boolean)

    '=================================
    '2000 / XP / NT / 7 / 8 / 10
    '=================================
    ' On Error GoTo TestWinVer_Error

    ' variables declared
    
    Dim ProgramFilesDir As String
    'Dim WindowsVer As String
    Dim strString As String

    'initialise the dimensioned variables
    strString = vbNullString
    classicThemeCapable = False
    WindowsVer = vbNullString
    ProgramFilesDir = vbNullString
    
    ' other variable assignments
    strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName")
    WindowsVer = strString

    
    ' note that when running in compatibility mode the o/s will respond with "Windows XP"
    ' The IDE runs in compatibility mode so it may report the wrong working folder
    
    'MsgBox WindowsVer

    'Get the value of "ProgramFiles", or "ProgramFilesDir"
    
    Select Case WindowsVer
    Case "Microsoft Windows NT4"
        classicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Microsoft Windows 2000"
        classicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Microsoft Windows XP"
        classicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
    Case "Microsoft Windows 2003"
        classicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Microsoft Vista"

        classicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case "Microsoft 7"

        classicThemeCapable = True
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProgramFilesDir")
    Case Else

        classicThemeCapable = False
        strString = fGetstring(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion", "ProgramFilesDir")
    End Select

    'MsgBox strString
    

    ProgramFilesDir = strString
    If ProgramFilesDir = vbNullString Then ProgramFilesDir = "c:\program files (x86)" ' 64bit systems
    If Not fDirExists(ProgramFilesDir) Then
        ProgramFilesDir = "c:\program files" ' 32 bit systems
    End If
    
    'If debugflg = 1 Then DebugPrint "%" & "ProgramFilesDir = " & ProgramFilesDir
    


    '======================================================
    'END routine error handler
    '======================================================

   
    On Error GoTo 0: Exit Sub

TestWinVer_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure TestWinVer of Module Common"

End Sub

' select a font for the fnt form
Public Sub changeFont(ByRef frm As Form, ByRef fntNow As Boolean, ByRef fntFont As String, ByRef fntSize As Integer, ByRef fntWeight As Integer, ByRef fntStyle As Boolean, ByRef fntColour As Long, ByRef fntItalics As Boolean, ByRef fntUnderline As Boolean, ByRef fntFontResult As Boolean)

    'initialise the dimensioned variables
    
    fntWeight = 0
    fntStyle = False
    'fntColour = 0
    'fntBold = False
    'fntUnderline = False
    fntFontResult = False
    
    If debugflg = 1 Then Debug.Print "%mnuFont_Click"

    displayFontSelector fntFont, fntSize, fntWeight, fntStyle, fntColour, fntItalics, fntUnderline, fntFontResult
    If fntFontResult = False Then Exit Sub


    If fntWeight > 700 Then
        'fntBold = True
    Else
        'fntBold = False
    End If
    
    If fntFont <> vbNullString And fntNow = True Then
        Call changeFormFont(frm, fntFont, Val(fntSize), fntWeight, fntStyle, fntItalics, fntColour)
    End If
    
    If frm.Name = "FireCallPrefs" Then
        Call resetComboBoxHighlight
    Else
        'FireCallMain.cmbEmojiSelection.SelLength = 0
    End If
    
End Sub

Public Sub resetComboBoxHighlight()

        ' The comboboxes all autoselect when the font is changed, we need to reset this afterwards

    'FireCallPrefs.cmbRefreshInterval.SelLength = 0
    'FireCallPrefs.cmbAlarmSound.SelLength = 0
    'FireCallPrefs.cmbButtonPositions.SelLength = 0
    'FireCallPrefs.cmbMaxLineLength.SelLength = 0
    'FireCallPrefs.cmbAdviceInterval.SelLength = 0
    'FireCallPrefs.cmbSmtpAuthenticate.SelLength = 0
    'FireCallPrefs.cmbSmtpSecurity.SelLength = 0
    
    
    
    'FireCallPrefs.cmbEmojiSet.SelLength = 0
    'FireCallPrefs.cmbWindowLevel.SelLength = 0
    
'    FireCallPrefs.cmbTTFN.SelLength = 0
'    FireCallPrefs.cmbWell.SelLength = 0
'    FireCallPrefs.cmbNews.SelLength = 0
'    FireCallPrefs.cmbMorn.SelLength = 0
'    FireCallPrefs.cmbWot.SelLength = 0
'    FireCallPrefs.cmbWth.SelLength = 0
'    FireCallPrefs.cmbPrg.SelLength = 0
'    FireCallPrefs.cmbGdn.SelLength = 0
'    FireCallPrefs.cmbBusy.SelLength = 0
'    FireCallPrefs.cmbCod.SelLength = 0
'    FireCallPrefs.cmbOut.SelLength = 0
'
'    FireCallPrefs.cmbCaptureDevices.SelLength = 0
    
End Sub

' this populates the input box, ie. the remote user data list box, sets the scrollbars immediately after.

Public Sub populateInputBox()

    Dim lLength As Long
    
    If FCWOptHandleData = "0" Then
        ioMethodADO = False
        FireCallMain.picFsoLampBright.Visible = True
        FireCallMain.picFsoLampDull.Visible = False
        FireCallMain.picUtf8LampBright.Visible = False
        FireCallMain.picUtf8LampDull.Visible = True
        
    Else
        ioMethodADO = True
        FireCallMain.picFsoLampBright.Visible = False
        FireCallMain.picFsoLampDull.Visible = True
        FireCallMain.picUtf8LampBright.Visible = True
        FireCallMain.picUtf8LampDull.Visible = False
    End If

    ' read the defined input file and write the input array
    Call readInputFileAndWriteArray(FCWSharedInputFile)
    
    ' the scrollbar config code needs to be here after the reading of the output data
    If FCWEnableScrollbars = 0 Then
        'the next two line must be in this order
        Call SendMessageByNum(FireCallMain.lbxInputTextArea.hwnd, LB_SETHORIZONTALEXTENT, 0, 0&) ' hides the horizontal scrollbar
        Call ShowScrollBar(FireCallMain.lbxInputTextArea.hwnd, SB_VERT, False)  ' hides the vertical scrollbar
    Else
        Call ShowScrollBar(FireCallMain.lbxInputTextArea.hwnd, SB_VERT, True) ' shows the vertical scrollbar
        ' add the horizontal scroll bar to the upper listbox
        lLength = 2 * (FireCallMain.lbxInputTextArea.Width / Screen.TwipsPerPixelX)
        Call SendMessageByNum(FireCallMain.lbxInputTextArea.hwnd, LB_SETHORIZONTALEXTENT, lLength, 0&)
    End If
    
    ' set the position to the last entry in the listbox
    If Val(FCWLoadBottom) = 1 Then
        FireCallMain.lbxInputTextArea.ListIndex = FireCallMain.lbxInputTextArea.ListCount - 1
    Else
        FireCallMain.lbxInputTextArea.ListIndex = 0
    End If

End Sub

' this populates the output box, ie. the local user data list box, sets the scrollbars immediately after.

Public Sub populateOutputBox()

    Dim lLength As Long
    
    'now do the same for the output
    outputLineCount = fLineCount(FCWSharedOutputFile)

    If outputLineCount >= 32766 Then
        debugLog "%Err-I-ErrorNumber 15 - The output file is close to the maximum limit, please split and shorten the output file"
    End If
        
    If outputLineCount >= 32766 Then
        debugLog "%Err-I-ErrorNumber 16 - The output file is too long at 32,766 lines long, please split and shorten the output file. FCW will not process new messages."
        Exit Sub
    End If
    ' read the file chosen as the output file
    ' write an array the same length as your output file
    ' write the listbox using the array
    Call readOutputFileWriteArrayWriteListbox(FCWSharedOutputFile)
    
    ' the scrollbar config code must be here after the reading of the output data
    If FCWEnableScrollbars = 0 Then
        'the next two line must be in this order
        Call SendMessageByNum(FireCallMain.lbxOutputTextArea.hwnd, LB_SETHORIZONTALEXTENT, 0, 0&) ' hides the horizontal scrollbar
        Call ShowScrollBar(FireCallMain.lbxOutputTextArea.hwnd, SB_VERT, False)  ' hides the vertical scrollbar
    Else
        Call ShowScrollBar(FireCallMain.lbxOutputTextArea.hwnd, SB_VERT, True) ' shows the vertical scrollbar
        ' add the horizontal scroll bar to the upper listbox
        lLength = 2 * (FireCallMain.lbxOutputTextArea.Width / Screen.TwipsPerPixelX)
        Call SendMessageByNum(FireCallMain.lbxOutputTextArea.hwnd, LB_SETHORIZONTALEXTENT, lLength, 0&)
    End If
    
    'set to the latest item in the listbox
    If Val(FCWLoadBottom) = 1 Then
        FireCallMain.lbxOutputTextArea.ListIndex = FireCallMain.lbxOutputTextArea.ListCount - 1
    Else
        FireCallMain.lbxOutputTextArea.ListIndex = 0
    End If
    
End Sub
' populates a third array for sorting and thence to a listbox containing both the inputs and outputs
Public Sub populateCombinedBox()

    Dim lLength As Long

    'If inputDataChangedFlag = True Or outputDataChangedFlag = True Then
    Call readListBoxesAndWriteCombinedArray
    Call performQuickSort
    Call readCombinedArrayAndWriteListbox

    ' the scrollbar config code must be here after the reading of the output data
    If FCWEnableScrollbars = 0 Then
        'the next two line must be in this order
        Call SendMessageByNum(FireCallMain.lbxCombinedTextArea.hwnd, LB_SETHORIZONTALEXTENT, 0, 0&) ' hides the horizontal scrollbar
        Call ShowScrollBar(FireCallMain.lbxCombinedTextArea.hwnd, SB_VERT, False)  ' hides the vertical scrollbar
    Else
        Call ShowScrollBar(FireCallMain.lbxCombinedTextArea.hwnd, SB_VERT, True) ' shows the vertical scrollbar
        ' add the horizontal scroll bar to the upper listbox
        lLength = 2 * (FireCallMain.lbxCombinedTextArea.Width / Screen.TwipsPerPixelX)
        Call SendMessageByNum(FireCallMain.lbxCombinedTextArea.hwnd, LB_SETHORIZONTALEXTENT, lLength, 0&)
    End If
    
    'set to the latest item in the listbox
    If Val(FCWLoadBottom) = 1 Then
        FireCallMain.lbxCombinedTextArea.ListIndex = FireCallMain.lbxCombinedTextArea.ListCount - 1
    Else
        FireCallMain.lbxCombinedTextArea.ListIndex = 0
    End If

End Sub
' reads the two arrays when required to do so, prior to sorting
Public Sub readListBoxesAndWriteCombinedArray()
    Dim useloop As Integer
    Dim stringToWrite As String
    Dim combinedArrayCount As Long
    
    stringToWrite = vbNullString
    combinedArrayCount = outputLineCount + inputLineCount
    
    ReDim combinedFileArray(combinedArrayCount)
    
    If combinedArrayCount >= 32766 Then
        debugLog "%Err-I-ErrorNumber 17 - The combined chat box is close to the maximum limit of lines of text, please split and shorten the input/output files or select the two chatbox option"
    End If
        
    If combinedArrayCount >= 32766 Then
        debugLog "%Err-I-ErrorNumber 18 - The combined chat box is too long at 32,766 lines long, please split and shorten the input/output files or select the two chatbox option. FCW will not process new messages."
        Exit Sub
    End If
    ' this reads the input listbox (with fully procesed text) and populates the combined array
    For useloop = 1 To inputLineCount
        stringToWrite = FireCallMain.lbxInputTextArea.List(useloop - 1) 'inputFileArray(useloop)
        combinedFileArray(useloop) = stringToWrite
    Next useloop
                    
    ' this reads the output listbox and populates the combined array
    For useloop = 1 To outputLineCount
        stringToWrite = FireCallMain.lbxOutputTextArea.List(useloop - 1) ' outputFileArray(useloop)
        combinedFileArray(useloop + inputLineCount) = stringToWrite
    Next useloop
    

End Sub ' sorts the combined array
Public Sub performQuickSort()

    Call QuickSort(combinedFileArray)
    
End Sub

' reads the sorted third array and write to the third listbox containing both the inputs and outputs
Public Sub readCombinedArrayAndWriteListbox()
    'Exit Sub ' remove this when code is ready to run
    
    Dim combinedArrayCount As Long
    Dim useloop As Long
    Dim stringToWrite As String
    Dim lbxCount As Long
    
    stringToWrite = vbNullString
    combinedArrayCount = outputLineCount + inputLineCount
    lbxCount = 0
    
    ' locks the combined listbox whilst the listbox is updated from the array to prevent rippling/flickering
    LockWindowUpdate FireCallMain.lbxCombinedTextArea.hwnd

    ' this reads the output array and populates the listbox
    For useloop = 1 To combinedArrayCount
        stringToWrite = vbNullString
        If combinedFileArray(useloop) <> vbNullString Then
            If FireCallMain.lbxCombinedTextArea.List(lbxCount) = combinedFileArray(useloop) Then
'                 the listboxes are much slower to update so this comparison is essential for speed.
'                 we do not clear down the listbox and so a comparison of the list content to the array contents is performed
'                 and if they are the same we make no changes, this prevents the full listbox update
            Else
                stringToWrite = combinedFileArray(useloop)
                FireCallMain.lbxCombinedTextArea.List(lbxCount) = stringToWrite
            End If
            lbxCount = lbxCount + 1
        End If
    Next useloop
        
    ' litle fix to prevent potential duplication of content in the listboxes
    Dim listCounter As Long
    listCounter = FireCallMain.lbxCombinedTextArea.ListCount
    If listCounter > combinedArrayCount Then
        'Call debugLog("This message pops up to prevent a duplication occurring in the INPUT. Please report if this occurs.")
        For useloop = (listCounter + 1) To combinedArrayCount
            FireCallMain.lbxCombinedTextArea.List(useloop) = ""
        Next useloop
    End If
    
    ' at this point the listbox has been written so we now unlock it after the update
    LockWindowUpdate 0& '
    
End Sub

' counts the number of lines in a file
Public Function fLineCount(ByRef sFName As String) As Long


' timer code BEGINS - requires QueryPerformanceCounter API declaration to be enabled
'    Dim lngReturn As Long
'    Dim curFreq As Currency
'    Dim curStart As Currency
'    Dim curEnd As Currency
'    Dim sngTime As Single
'
'    lngReturn = QueryPerformanceFrequency(curFreq)
'    If lngReturn = 0 Then MsgBox "Your Hardware does not support a high-resolution timer"
'
'    If lngReturn <> 0 Then
'        lngReturn = QueryPerformanceCounter(curStart)
' timer code ENDS

    ' new code to count the lines using ADODB.Stream STARTS ' 3.4ms
    'If ioMethodADO = True Then
        '3.4ms
'        Dim Stm As ADODB.Stream
'        Dim Line As String
'
'        Set Stm = New ADODB.Stream
'        With Stm
'            .Open
'            .LoadFromFile sFName
'            .Type = 2
'            .Charset = "utf-8"
'            .LineSeparator = -1
'            Do Until .EOS
'                Line = .ReadText(-2)
'                fLineCount = fLineCount + 1
'            Loop
'            .Close
'        End With
'    Else
        ' old code to count the lines using FileSystemObject STARTS ' 10.99ms
        Dim dummyRead As String
        Const ForReading As Integer = 1
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim thisFile As Object
        Set thisFile = fso.OpenTextFile(sFName, ForReading, False, 0)
        Do While Not thisFile.AtEndOfStream
            dummyRead = thisFile.readLine
            fLineCount = fLineCount + 1
        Loop
        thisFile.Close

        If fFExists(sFName) Then

            '-- subtract blank line at the bottom of the file
            If Right$(dummyRead, 1) = vbLf Then fLineCount = fLineCount - 1

        End If
        ' old code to count the lines using FileSystemObject ENDS ' 10.99ms

 '   End If
    
' timer code BEGINS
'        lngReturn = QueryPerformanceCounter(curEnd)
'        sngTime = (curEnd - curStart) * 1000 / curFreq
'        Debug.Print "Execution Time " & sngTime & " mS"
'        MsgBox "Execution Time " & sngTime & " mS"
'    End If
' timer code ENDS
End Function



' read the remote user's file line by line, read it into an interim array and thence into a listbox.
' called by checkAndReadInputFile during polling & populateInputBox during startup
Public Sub readInputFileAndWriteArray(ByVal sFName As String)

    Dim lIndex As Long
    'Dim fileString As String
    Dim useloop As Long
    Dim lbxCount As Integer
    Dim startLoc As Long
    Dim endLoc As Long
    Dim stepNo As Integer
    Dim stringToWrite As String
    Dim emojiFilenamePos As Integer
    Dim emojiFilename As String
    Dim emojiFilePath As String
    Dim recordingString As String
    Dim recordingFilenamePos As Integer
    Dim recordingFilename As String
    Dim attachmentString As String
    Dim attachmentFilenamePos As Integer
    Dim attachmentFilename As String
    Dim suffix As String
    Dim suffixNoDot As String
    Dim emojiString As String
    Dim pingString As String
    Dim buzzerString As String
    Dim awakeString As String
    Dim buzzerStamp As String
    'Dim finalStamp As String
    Dim pingStamp As String
    Dim awakeStamp As String
    Dim unixTimeStamp As String
    Dim shutdownString As String
    Dim shutdownStamp As String
    Dim shutDateTime As String
    Dim folderDisplayed As Boolean
    Dim soundtoplay As String
    Dim attachmentTimeDiffInSecs As Long
    Dim recordingTimeDiffInSecs As Long
    Dim shutdownTimeDiffInSecs As Currency
    Dim thisFile As Object
    Dim answer As VbMsgBoxResult
    Dim answer2 As VbMsgBoxResult
    Dim inStm As ADODB.Stream
    Dim tmpRead As String
    Dim findStr As Integer
    Dim findStartPos As Integer
    Dim stringWithoutPrefix As String
        
    Dim nowInSecs As Long
    Dim lastShutdownDiff As Long
        
    Const ForReading As Integer = 1
    
    useloop = 0
    buzzerString = vbNullString
    buzzerStamp = vbNullString
    'finalStamp = vbNullString
    pingStamp = vbNullString
    emojiString = vbNullString
    emojiFilenamePos = 0
    emojiFilename = vbNullString
    emojiFilePath = vbNullString
    attachmentString = vbNullString
    attachmentFilenamePos = 0
    attachmentFilename = vbNullString
    attachmentFilePath = vbNullString
    suffix = vbNullString
    pingString = vbNullString
    awakeString = vbNullString
    awakeStamp = vbNullString
    unixTimeStamp = vbNullString
    shutdownString = vbNullString
    shutdownStamp = vbNullString
    shutDateTime = vbNullString
    folderDisplayed = False
    attachmentTimeDiffInSecs = 0
    shutdownTimeDiffInSecs = 0
    lIndex = 1
    answer = vbNo
    answer2 = vbNo
    findStartPos = 0
    stringWithoutPrefix = ""
    
    ' get the line count
    inputLineCount = fLineCount(FCWSharedInputFile)
    
    If inputLineCount >= 32766 Then
        debugLog "%Err-I-ErrorNumber 19 - The input file is close to the maximum limit, please split and shorten the input file"
    End If
        
    If inputLineCount >= 32766 Then
        debugLog "%Err-I-ErrorNumber 20 - The input file is too long at 32,766 lines long, please split and shorten the input file. FCW will not process new messages"
        Exit Sub
    End If
    
    ' resize the array to the new linecount
    ReDim inputFileArray(inputLineCount)
    ' timer code BEGINS - requires QueryPerformanceCounter API declaration to be enabled
'    Dim lngReturn As Long
'    Dim curFreq As Currency
'    Dim curStart As Currency
'    Dim curEnd As Currency
'    Dim sngTime As Single
'
'    lngReturn = QueryPerformanceFrequency(curFreq)
'    If lngReturn = 0 Then MsgBox "Your Hardware does not support a high-resolution timer"
'
'    lngReturn = QueryPerformanceCounter(curStart)
    ' timer code ENDS

    
    If ioMethodADO = False Then
        ' use of the FileSystemObject as it handles EOL with CrLf whereas INPUT LINE does not.
        ' when working with a linux version of the utility this is vital.
        Set fso = CreateObject("Scripting.FileSystemObject") '11.62ms in the IDE 13ms when compiled
        Set thisFile = fso.OpenTextFile(sFName, ForReading, False, 0)
    ''    ' read the file into an interim storage array for the input data
        Do While Not thisFile.AtEndOfStream
            inputFileArray(lIndex) = thisFile.readLine
            lIndex = lIndex + 1
            If lIndex > inputLineCount Then
                Exit Do
            End If
        Loop
        thisFile.Close
    Else
    
        ' code to replace the usage of the file system object i/o, supposedly the FSO object
        ' is slow, bloated and poor at UTF-8 support. Use of the ADO object requires enabling
        ' of Microsoft ActiveX Data Objects 2.8 Library in References.
        ' required Projects>References>Microsoft ActiveX Data Objects 2.8 Library.
        ' it can handle Charset = "utf-8" and LineSeparator = -1 properly.
    
        Set inStm = New ADODB.Stream ' 27.75ms in the IDE, slower than FSO by factor of 2 but 7-8ms when compiled! 1,400 lines
        With inStm
            .Open
            .LoadFromFile sFName
            .Type = 2
            .Charset = "utf-8"
            .LineSeparator = -1 ' adCRLF 'vbCrLf
                                '
                                'adCRLF   -1    Default. Carriage return line feed
                                'adLF     10    Line feed only
                                'adCR     13    Carriage return only
            ' read the file into an interim storage array for the input data
            .Position = 2

            Do While Not .EOS
                inputFileArray(lIndex) = .ReadText(-2)    ' -2 Reads the next line from the stream
                lIndex = lIndex + 1
                If lIndex > inputLineCount Then
                    Exit Do
                End If
            Loop
            .Close
        End With
    End If
    
' timer code BEGINS
'        lngReturn = QueryPerformanceCounter(curEnd)
'        sngTime = (curEnd - curStart) * 1000 / curFreq
'        Debug.Print "Execution Time " & sngTime & " mS"
'        MsgBox "Execution Time " & sngTime & " mS"
'
' timer code ENDS

    lbxCount = 0
        
    ' determine the start point at which we read from the array, beginning or end.
    If Val(FCWLoadBottom) = 1 Then
            startLoc = inputLineCount
            endLoc = 1
            stepNo = -1
    Else
            startLoc = 1
            endLoc = inputLineCount
            stepNo = 1
    End If
        
    ' locks the input listbox whilst the listbox is updated from the array to prevent rippling/flickering
    LockWindowUpdate FireCallMain.lbxInputTextArea.hwnd
    
    ' read the array and write the output to the listbox, replacing the known tags with the correct text
    ' also store timestamps and variables associated with each type of tag
    ' known tags = <><> <o><o> <p><p> <b><b> <t><t> <z><z>
    '
    For useloop = startLoc To endLoc Step stepNo
        stringToWrite = vbNullString
        If inputFileArray(useloop) <> vbNullString Then 'differs from the input file in that we turn the array upside down
            If FireCallMain.lbxInputTextArea.List(lbxCount) = inputFileArray(useloop) Then
                ' we do not clear down the listbox and so a comparison of the array contents to the file is performed
                ' and if the two lines are the same we make no changes, this avoids the flickering penalty of a full listbox update
            Else
                'find the line start point minus the timestamp and prefix:
                findStartPos = InStr(inputFileArray(useloop), ":    ")
                stringWithoutPrefix = Mid$(inputFileArray(useloop), findStartPos + 5, Len(inputFileArray(useloop)))
                
                ' only update the lines that have changed
                If InStr(stringWithoutPrefix, "<><>") = 1 Then ' we only allow these tags to be recognised at position 1
                    ' <><> attachment
                    attachmentString = inputFileArray(useloop)
                    stringToWrite = Replace$(inputFileArray(useloop), "<><>", "New File:")
                ElseIf InStr(stringWithoutPrefix, "<f><f>") = 1 Then
                    ' <f><f> attachment
                    attachmentString = inputFileArray(useloop)
                    stringToWrite = Replace$(inputFileArray(useloop), "<f><f>", "New Folder:")
                ElseIf InStr(stringWithoutPrefix, "<o><o>") = 1 Then
                    ' <o><o> emoji
                    emojiString = inputFileArray(useloop)
                    stringToWrite = Replace$(inputFileArray(useloop), "<o><o>", "New Emoji:")
                ElseIf InStr(stringWithoutPrefix, "<p><p>") = 1 Then
                    ' <p><p> ping
                    pingString = inputFileArray(useloop)
                    pingStamp = Left$(inputFileArray(useloop), 22)
                    stringToWrite = Replace$(inputFileArray(useloop), "<p><p>", "Ping Request.")
                ElseIf InStr(stringWithoutPrefix, "<b><b>") = 1 Then
                    ' <b><b> buzzer
                    buzzerString = inputFileArray(useloop)
                    buzzerStamp = Left$(inputFileArray(useloop), 22)
                    stringToWrite = Replace$(inputFileArray(useloop), "<b><b>", "Attention!")
                ElseIf InStr(stringWithoutPrefix, "<t><t>") = 1 Then
                    ' <t><t> Awake
                    awakeString = inputFileArray(useloop)
                    awakeStamp = Left$(inputFileArray(useloop), 22)
                    findStr = InStr(inputFileArray(useloop), "<t><t>")
                    unixTimeStamp = Mid$(inputFileArray(useloop), findStr + 6, Len(inputFileArray(useloop)))
                    stringToWrite = Replace$(inputFileArray(useloop), "<t><t>", "Awake at:")
                ElseIf InStr(stringWithoutPrefix, "<z><z>") = 1 Then
                    ' <z><z> shutdown
                    shutdownString = inputFileArray(useloop)
                    shutdownStamp = Left$(inputFileArray(useloop), 22)
                    findStr = InStr(inputFileArray(useloop), "<z><z>")
                    unixTimeStamp = Mid$(inputFileArray(useloop), findStr + 6, Len(inputFileArray(useloop)))
                    stringToWrite = Replace$(inputFileArray(useloop), "<z><z>", "Shutdown at:")
                ElseIf InStr(stringWithoutPrefix, "<r><r>") = 1 Then
                    ' <r><r> recording
                    recordingString = inputFileArray(useloop)
                    stringToWrite = Replace$(inputFileArray(useloop), "<r><r>", "New Recording:")
                Else
                    stringToWrite = inputFileArray(useloop)
                End If
                FireCallMain.lbxInputTextArea.List(lbxCount) = stringToWrite
            End If
            lbxCount = lbxCount + 1
        End If
        'finalStamp = Left$(inputFileArray(useloop), 22)
    Next useloop
    
    ' at this point the listbox has been written so we now unlock it after the update
    LockWindowUpdate 0& '

    ' litle fix to prevent potential duplication of content in the listboxes
    Dim listCounter As Long
    listCounter = FireCallMain.lbxInputTextArea.ListCount
    If listCounter > inputLineCount Then
        Call debugLog("This message pops up to prevent a duplication occurring in the INPUT. Please report if this occurs.")
        For useloop = (listCounter + 1) To inputLineCount
            FireCallMain.lbxInputTextArea.List(useloop) = ""
        Next useloop
    End If
    
    ' Post processing according to what we found in the file, generally we respond to the last occurrence of each event
    ' hold the event time and compare that with a stored value.
    
    'respond to a ping request and store the time of the last ping so that we do not respond multiple times
    If pingString <> vbNullString And FCWLastPingResponse <> pingStamp Then
        
        ' next line caused corruption during the refresh process
        'Call sendSomething("Ping response. Refresh interval: " & FireCallPrefs.cmbRefreshInterval.ItemData(Val(FCWRefreshIntervalIndex)) & "  OS:" & WindowsVer & "  Version:" & App.Major & "." & App.Minor & "." & App.Revision)
        
        'old method of passing a single command to a timer to run it asynchronously
        'FireCallMain.sendCommandTimer.Tag = "Ping response. Refresh interval: " & FireCallPrefs.cmbRefreshInterval.ItemData(Val(FCWRefreshIntervalIndex)) & "  OS:" & WindowsVer & "  Version:" & App.Major & "." & App.Minor & "." & App.Revision
        
        'new method of passing a command to a message queue for a timer to run each asynchronously
        messageQueue.Add "Ping response. Refresh interval: " & FireCallPrefs.cmbRefreshInterval.ItemData(Val(FCWRefreshIntervalIndex)) & "  OS:" & WindowsVer & "  Version:" & App.Major & "." & App.Minor & "." & App.Revision ' Add to the end of the collection
        
        FireCallMain.sendCommandTimer.Enabled = True ' this does a Call sendSomething(stringToSend)
                                                     ' but it does it ensuring this tranche of current polling is complete
        FCWLastPingResponse = pingStamp
        If fFExists(FCWSettingsFile) Then
            PutINISetting "Software\FireCallWin", "lastPingResponse", FCWLastPingResponse, FCWSettingsFile
        End If
    End If
    
    ' if an emoji code is sent then set the emoji image locally, this is always the last emoji received
    If emojiString <> vbNullString Then
        emojiFilenamePos = InStr(emojiString, "<o><o>") + 6
        emojiFilename = Mid$(emojiString, emojiFilenamePos, Len(emojiString)) & ".jpg"
        emojiFilePath = App.Path & "\Resources\Emojis\standard\base\" & emojiFilename
        If fFExists(emojiFilePath) Then
            FireCallMain.picEmoji.Picture = LoadPicture(emojiFilePath)
            'If mainBtnLidVisible = False Then Call clickOnPicEmoji
        End If
    End If
    
    ' if an attachment is sent then attempt to display it always shows the last
    If attachmentString <> vbNullString Then
        If InStr(attachmentString, "<><>") > 0 Then
            attachmentFilenamePos = InStr(attachmentString, "<><>") + 4 ' file
        Else
            attachmentFilenamePos = InStr(attachmentString, "<f><f>") + 6 ' folder
            folderDisplayed = True
        End If
        attachmentFilename = Mid$(attachmentString, attachmentFilenamePos, Len(attachmentString))
        ' next line removes a spurious character (vbCrLf?) that appeared after changing the method of reading to ADO
        'attachmentFilename = Mid$(attachmentFilename, 1, Len(attachmentFilename))
        attachmentFilePath = FCWExchangeFolder & "\" & attachmentFilename
        
        ' if the current image display, caused by a recent click on an image in the chat happened less than
        ' three minutes ago, then bypass the display of the last image found. This allows the user to retain
        ' his recently clicked-upon image even if a repoll for new data happens in that time.
        
        If attachmentViewTime <> "00:00:00" Then
            attachmentTimeDiffInSecs = DateDiff("s", attachmentViewTime, Now)
        End If
        
        If attachmentTimeDiffInSecs = 0 Or attachmentTimeDiffInSecs >= 180 Then
            If folderDisplayed = False Then
                ' we store the full file path as the attachmentFilePath variable will be overwritten by subsequent sutomatic clicks
                ' when setting the input listBox to the last position a click is always generated
                displayedAttachmentFilePath = attachmentFilePath
                suffix = fExtractSuffixWithDot(displayedAttachmentFilePath)
                suffixNoDot = fExtractSuffix(displayedAttachmentFilePath)

                If fInstrSuffix(validImageArrayList, LCase(suffix)) Then
                    Call displaySelectedImage(displayedAttachmentFilePath)
                ElseIf fInstrSuffix(invalidImageArrayList, LCase(suffix)) Then
                    Call displaySelectedImage(App.Path & "\resources\images\documentIcons\document-unknown" & ".png")
                Else
                    Call displaySelectedImage(App.Path & "\resources\images\documentIcons\document-" & suffixNoDot & ".png")
                End If
                If FCWEnableTooltips = "1" Then FireCallMain.picImagePrintOut.ToolTipText = attachmentFilename & " - double click to open the file using default app."
                
            Else ' folder
                Call displaySelectedImage(App.Path & "\resources\images\documentIcons\document-dir.png")
                If FCWEnableTooltips = "1" Then FireCallMain.picImagePrintOut.ToolTipText = attachmentFilename & " - double click to open the folder in Explorer."
            End If
        End If

    End If
    
    ' buzzer code received, stores the last buzz time so we only buzz once
    If buzzerString <> vbNullString And FCWLastSoundPlayed <> buzzerStamp Then
        
        If FCWPlayVolume = "1" Then
            soundtoplay = App.Path & "\Resources\Sounds\" & "buzzer.wav"
        Else
            soundtoplay = App.Path & "\Resources\Sounds\" & "buzzerQuiet.wav"
        End If
        
        If fFExists(soundtoplay) Then
             PlaySound soundtoplay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
        FCWLastSoundPlayed = buzzerStamp

        FireCallMain.picBuzzerDullLamp.Visible = False
        FireCallMain.picBuzzerBrightLamp.Visible = True
        
        If fFExists(FCWSettingsFile) Then
            PutINISetting "Software\FireCallWin", "lastSoundPlayed", FCWLastSoundPlayed, FCWSettingsFile
        End If
    End If
    
    ' we received an awake string, store the time so we respond only once
    If awakeString <> vbNullString And FCWLastAwakeString <> awakeStamp Then
        ' respond to an awake request by interpreting the unix epoch date and time string
        ' rather than using the timestamp at the beginning of the incoming string
        ' probably overkill as I could have just reformatted the date stamp at the beginning of the string
        ' but I wanted to have the code to read unix timestamps, and it is more elegant.
        '
        'FireCallMain.sendCommandTimer.Tag = "Awake response. Request time: " & fConvertEpochToTimeString(unixTimeStamp)
        
        messageQueue.Add "Awake response. Request time: " & fConvertEpochToTimeString(unixTimeStamp)

        FireCallMain.sendCommandTimer.Enabled = True ' this does a Call sendSomething(stringToSend)
                                                     ' but it does it ensuring the current polling is complete
        FCWLastAwakeString = awakeStamp
        
        If fFExists(FCWSettingsFile) Then
            PutINISetting "Software\FireCallWin", "lastAwakeString", FCWLastAwakeString, FCWSettingsFile
        End If
    End If
    

    ' remote shutdown code received, stores the last shutdown time so we only respond to the last shutdown request
    If shutdownString <> vbNullString And FCWLastShutdown <> shutdownStamp Then
    
        FCWLastShutdown = shutdownStamp
        If fFExists(FCWSettingsFile) Then
            PutINISetting "Software\FireCallWin", "lastShutdown", FCWLastShutdown, FCWSettingsFile
        End If
        
        ' if the shutdown time is old > 5 mins then ignore it (this assumes it arrived when FCW was either asleep or not running)
        shutdownTimeDiffInSecs = fSecondsFromDateString(shutdownStamp)
        nowInSecs = fSecondsFromDateString(Now)
        lastShutdownDiff = nowInSecs - shutdownTimeDiffInSecs
    
        If lastShutdownDiff <= 300 Then
            If FCWAllowShutdowns = "1" Then
 
                answer = MsgBox("The remote chat partner has requested a temporary FCW shutdown, whilst maintenance takes place. " & vbCrLf & _
                        "OK to shutdown?", vbExclamation + vbYesNo)
                        
                If answer = vbYes Then
                    messageQueue.Add "Positive shutdown response to shutdown request at " & shutdownStamp & " GO AHEAD."
                    FireCallMain.sendCommandTimer.Enabled = True ' this does a Call sendSomething(stringToSend)
                                                          ' but it does it ensuring the current polling is complete
                    FireCallMain.shutdownTimer.Enabled = True ' we call a timer to shut it down after 5 secs.
                Else
                    messageQueue.Add "Negative shutdown response to shutdown request at " & shutdownStamp
                    FireCallMain.sendCommandTimer.Enabled = True ' this does a Call sendSomething(stringToSend)
                End If
            Else
                messageQueue.Add "Negative shutdown response to shutdown request at " & shutdownStamp
                FireCallMain.sendCommandTimer.Enabled = True ' this does a Call sendSomething(stringToSend)
            End If
        End If

    End If
    
    
        ' if an attachment is sent then attempt to display it always shows the last
    If recordingString <> vbNullString Then
        If InStr(recordingString, "<r><r>") > 0 Then
            recordingFilenamePos = InStr(recordingString, "<r><r>") + 6 ' folder
        End If
        recordingFilename = Mid$(recordingString, recordingFilenamePos, 23)
        
        ' next line removes a spurious character (vbCrLf?) that appeared after changing the method of reading to ADO
        'recordingFilename = Mid$(recordingFilename, 1, Len(recordingFilename))
        recordingFilePath = FCWExchangeFolder & "\" & recordingFilename
        
        ' if the current image display, caused by a recent click on an image in the chat happened less than
        ' three minutes ago, then bypass the display of the last image found. This allows the user to retain
        ' his recently clicked-upon image even if a repoll for new data happens in that time.
        
        If recordingViewTime <> "00:00:00" Then
            recordingTimeDiffInSecs = DateDiff("s", recordingViewTime, Now)
        End If
        
        If recordingTimeDiffInSecs = 0 Or recordingTimeDiffInSecs >= 180 Then
                ' we store the full file path as the recordingFilePath variable will be overwritten by subsequent sutomatic clicks
                ' when setting the input listBox to the last position a click is always generated
                

                displayedAttachmentFilePath = recordingFilePath

                Call displaySelectedImage(App.Path & "\resources\images\documentIcons\document-rec" & ".png")
                If FCWEnableTooltips = "1" Then FireCallMain.picImagePrintOut.ToolTipText = recordingFilename & " - double click to play the recording."
        End If

    End If
    ' now store the file characteristics so that we can use them to compare on the next run
    oldInputFileModificationTime = inputFileModificationTime

End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : fInstrSuffix
' Author    : beededea
' Date      : 23/10/2024
' Purpose   : searches through a supplied collection for a matching string
'---------------------------------------------------------------------------------------
'
Public Function fInstrSuffix(arrayList As Collection, thisSuffix As String)
    Dim arrayMember As Variant
    On Error GoTo fInstrSuffix_Error

    fInstrSuffix = False
    For Each arrayMember In arrayList
        If LCase(thisSuffix) = arrayMember Then
            fInstrSuffix = True
            Exit For
        End If
    Next

   On Error GoTo 0
   Exit Function

fInstrSuffix_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fInstrSuffix of Module modCommon"
End Function


'---------------------------------------------------------------------------------------
' Procedure : displaySelectedImage
' Author    : beededea
' Date      : 10/07/2021
' Purpose   : This uses LaVolpe's code for the reading of 'foreign' image types
'---------------------------------------------------------------------------------------
'
Public Sub displaySelectedImage(ByVal fullFilePath As String)

    Dim suffix As String: suffix = vbNullString
    Dim suffixNoDot As String: suffixNoDot = vbNullString
    Dim imageCreated As Boolean: imageCreated = False
    Dim currFilePath As String: currFilePath = vbNullString
    Dim imgFilePath As String: imgFilePath = vbNullString
    Dim rectifiedFileName As Boolean: rectifiedFileName = False
            
    imageCreated = False

    ' On Error GoTo displaySelectedImage_Error

    ' default positions prior to any resizing/shifting
    Call putImageInPlace
    
    ' dispose of the image prior to use
    Set FireCallMain.picImagePrintOut.Picture = Nothing ' added because the two methods of drawing an image conflict leaving an image behind
    FireCallMain.picImagePrintOut.cls

    If fFExists(fullFilePath) <> 0 Then
    
        rectifiedFileName = ValidFileName(fGetFileNameFromPath(fullFilePath))
        If rectifiedFileName = False Then Exit Sub
            
        If FireCallMain.picImagePrintOut.Visible = False Then
            'FireCallMain.picImagePrintOut.Visible = True
            'FireCallMain.picPrintOutShadow.Visible = True
            If FCWImageDisplay = "1" Then
                imgFilePath = App.Path & "\Resources\images\lidBackgroundDullShadowed.jpg"
                If fFExists(imgFilePath) Then
                    FireCallMain.picLidBackground.Picture = LoadPictureEx(imgFilePath)
                End If
            End If
'            imgFilePath = App.Path & "\Resources\images\lidBackgroundDullShadowed.jpg"
'            FireCallMain.picLidBackground.Picture = LoadPicture(imgFilePath)
            
            FireCallMain.picEmojiKnobRight.Visible = False
            
        End If
                
        suffix = Trim$(fExtractSuffix(fullFilePath))
        
        If InStr("png,tif,tiff,cur,wmf,emf", LCase(suffix)) <> 0 Then
            ' using Lavolpe's later method as it allows for resizing of PNGs and all other types
            ' cPNGParser, cTGAParser, later formats foreign to VB6
            Set cImage = New c32bppDIB
            imageCreated = cImage.LoadPictureFile(fullFilePath, 256, 256, False, 32)
            Call resizeNonNative(FireCallMain.picImagePrintOut, origWidth, origHeight)
            Call renderPicBox(FireCallMain.picImagePrintOut, origWidth, origHeight)
            
        ElseIf InStr("ico", LCase(suffix)) <> 0 Then
            ' *.ico
            ' using Lavolpe's earlier StdPictureEx method as it allows for correct display of ICOs.
            ' The later method (using ICOParser.cls) has a bug when displaying several ICOs on the same form
            ' causing image corruption, so we use the older version which is more stable.
            
            ' because the earlier method draws the ico images from the top left of the
            ' pictureBox we have to manually set the picbox to size and position for each icon size
               
            FireCallMain.picImagePrintOut.Left = 250
            'FireCallMain.picPrintOutShadow.Left = 254
            
            FireCallMain.picImagePrintOut.Width = 1920 '1920 twips or 128 pixels
            FireCallMain.picImagePrintOut.Height = 1920 ' icons are always square
            
            Set FireCallMain.picImagePrintOut.Picture = StdPictureEx.LoadPicture(fullFilePath, lpsCustom, , 128, 128)
        End If
    
        If InStr("jpg,bmp,jpeg,gif", LCase(suffix)) <> 0 Then
            'for image types known to VB6 we do not use the native methods of displaying in a picbox as they cannot handle unicode characters
            ' instead we use a unicode compatible
            Dim lPic As Picture
            Set lPic = LoadPictureEx(fullFilePath)
            
            Call resizeNative(FireCallMain.picImagePrintOut, lPic)
        End If
    Else
        currFilePath = App.Path & "\resources\images\documentIcons\document-unknown" & ".png"
        If fFExists(currFilePath) Then
            Set cImage = New c32bppDIB
            imageCreated = cImage.LoadPictureFile(currFilePath, 144, 144, False, 32)
            Call renderPicBox(FireCallMain.picImagePrintOut, 144, 144)
        End If
    End If

    On Error GoTo 0
    Exit Sub

displaySelectedImage_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure displaySelectedImage of Module Module1"

End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : LoadPictureEx
' Author    : vangoghgaming
' Date      : 22/10/2024
' Purpose   : unicode version of loadPicture that handles unicode char.s in filenames
'---------------------------------------------------------------------------------------
'
Private Function LoadPictureEx(sFilename As String) As IPicture
    Dim IID_IUnknown(0 To 1) As Currency, ppStream As IUnknown
    
    On Error GoTo LoadPictureEx_Error

    IID_IUnknown(1) = 504403158265495.5712@
    If SHCreateStreamOnFileW(StrPtr(sFilename), 0, ppStream) = 0 Then OleLoadPicture ppStream, 0, 0, IID_IUnknown(0), LoadPictureEx

   On Error GoTo 0
   Exit Function

LoadPictureEx_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure LoadPictureEx of Module modCommon"
End Function


'
'---------------------------------------------------------------------------------------
' Procedure : ValidFileName
' Author    : eduardo
' Date      : 22/10/2024
' Purpose   : flags any non-standard character in a filename
'---------------------------------------------------------------------------------------
'
Public Function ValidFileName(nProposedFileName As String, Optional ForOldFileFormat_8Dot3 As Boolean = False) As Boolean
    Dim iChar As String: iChar = vbNullString
    Dim C  As Long: C = 0
    Dim iFlag As Long: iFlag = 0
    
    On Error GoTo ValidFileName_Error

    If ForOldFileFormat_8Dot3 Then
        iFlag = GCT_SHORTCHAR
    Else
        iFlag = GCT_LFNCHAR
    End If
    ValidFileName = False
    For C = 1 To Len(nProposedFileName)
        iChar = Mid$(nProposedFileName, C, 1)
        If (PathGetCharType(AscW(iChar)) And iFlag) = iFlag Then
            ValidFileName = True
        Else
            ValidFileName = False
        End If
    Next C

   On Error GoTo 0
   Exit Function

ValidFileName_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure ValidFileName of Module modCommon"
End Function
' credit jcis https://www.vbforums.com/member.php?40893-jcis
' resize image types known to VB6
'---------------------------------------------------------------------------------------
' Procedure : resizeNative
' Author    : beededea
' Date      : 23/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub resizeNative(ByRef pBox As PictureBox, ByRef pPic As Picture)
    Dim lWidth     As Single: lWidth = 0
    Dim lHeight    As Single: lHeight = 0
    Dim lnewWidth  As Single: lnewWidth = 0
    Dim lnewHeight As Single: lnewHeight = 0
 
    'Clear the Picture in the PictureBox
    On Error GoTo resizeNative_Error

    pBox.Picture = Nothing
    
    'Clear the Image  in the Picturebox
    pBox.cls
    
    'Get the size of the Image, but in the same Scale than the scale used by the PictureBox
    lWidth = pBox.ScaleX(pPic.Width, vbHimetric, pBox.ScaleMode)
    lHeight = pBox.ScaleY(pPic.Height, vbHimetric, pBox.ScaleMode)
    
    ' resize Width
    If lWidth > pBox.ScaleWidth Then
        lnewWidth = pBox.ScaleWidth
        lHeight = lHeight * (lnewWidth / lWidth) 'Resize height proportionally
    Else
        lnewWidth = lWidth                       'No changes required
    End If
    
    ' resize Height
    If lHeight > pBox.ScaleHeight Then
        lnewHeight = pBox.ScaleHeight
        lnewWidth = lnewWidth * (lnewHeight / lHeight)  'Resize width proportionally
    Else
        lnewHeight = lHeight                            'No changes required
    End If
    
    'add resized and centred to Picturebox
    pBox.PaintPicture pPic, (pBox.ScaleWidth - lnewWidth) / 2, _
                            (pBox.ScaleHeight - lnewHeight) / 2, _
                            lnewWidth, lnewHeight
                            
    'Update the Picture with the new image
    Set pBox.Picture = pBox.Image

   On Error GoTo 0
   Exit Sub

resizeNative_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure resizeNative of Module modCommon"
End Sub

' Calculate new dimensions of the picturebox
'---------------------------------------------------------------------------------------
' Procedure : resizeNonNative
' Author    : beededea
' Date      : 23/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub resizeNonNative(ByRef pBox As PictureBox, ByRef lWidth As Long, ByRef lHeight As Long)
    Dim lnewWidth As Single: lnewWidth = 0
    Dim lnewHeight As Single: lnewHeight = 0
    
    ' note that the size of the Image is already in the same Scale as the picBox

    ' resize Width
   On Error GoTo resizeNonNative_Error

    If lWidth > pBox.ScaleWidth Then
        lnewWidth = pBox.ScaleWidth
        lHeight = lHeight * (lnewWidth / lWidth) 'Resize height proportionally
    Else
        lnewWidth = lWidth                       'No changes required
    End If

    ' resize Height
    If lHeight > pBox.ScaleHeight Then
        lnewHeight = pBox.ScaleHeight
        lnewWidth = lnewWidth * (lnewHeight / lHeight)  'Resize width proportionally
    Else
        lnewHeight = lHeight                            'No changes required
    End If
    
    lWidth = lnewWidth  ' pass the new values back to the byRef vars
    lHeight = lnewHeight

   On Error GoTo 0
   Exit Sub

resizeNonNative_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure resizeNonNative of Module modCommon"
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : renderPicBox
' Author    : lavolpe
' Date      : 14/07/2019
' Purpose   :
    ' Generally, when rotating and/or resizing, it is easier to calculate the center of where you want the image rotated vs
    '   calculating the top/left coordinate of the resized, rotated image.  The last parameter of the Render call (CenterOnDestXY)
    '   will render around that center point if that paremeter is set.  So, what about when an image is not rotated? The Render
    '   function will still draw around that center point if the parameter is true. Or render, starting at the passed
    '   DestX,DestY coordinates if that parameter is false.
    ' The Render call only has one required parameter.  All others are optional and defaulted as follows
        ' srcX, srcY, destX, destY defaults are zero
        ' srcWidth, destWidth defaults are the image's width
        ' srcHeight, destHeight defaults are the image's height
        ' Opacity (Global Alpha) default is 100% opaque, pixel LigthAdjustmnet default is zero (no additional adjustment)
        ' GrayScale default is not grayscaled
        ' Rotation angle is at zero degrees
        ' Rendering image around a center point is false
    ' the cboAngle entries are at 15 degree intervals, so we simply multiply ListIndex by 15
'---------------------------------------------------------------------------------------
'
Private Sub renderPicBox(ByRef picBox As PictureBox, ByVal iconWidth As Integer, ByVal iconHeight As Integer)

    Dim newWidth As Long
    Dim newHeight As Long
'    Dim mirrorOffsetX As Long
'    Dim mirrorOffsetY As Long
    Dim positionX As Long
    Dim positionY As Long
    'Dim ShadowOffset As Long
    'Dim LightAdjustment As Single
    Dim imageCreated As Boolean
    
    imageCreated = False
        
    ' On Error GoTo renderPicBox_Error
    If debugflg = 1 Then Debug.Print "%" & "renderPicBox"

'    mirrorOffsetX = 1
'    mirrorOffsetY = 1

    newWidth = iconWidth: newHeight = iconHeight
    
    positionX = (picBox.ScaleWidth - newWidth) \ 2
    positionY = (picBox.ScaleHeight - newHeight) \ 2
    
'   See c32bppDIB.CreateDropShadow for more
'   Color, Opacity, Blur Effect,
'   and positionX,positionY Position are adjustable

    picBox.cls
'    If Not cShadow Is Nothing Then
'        picBox.CurrentX = 20
'        picBox.CurrentY = 5
'        picBox.CurrentX = 20
'        picBox.CurrentX = 20
'    End If
    
'    If Not cShadow Is Nothing Then
'        ' the 55 below is the shadow's opacity; hardcoded here but can be modified to your heart's delight
'        cShadow.Render picBox.hDC, positionX + newWidth \ 2 + ShadowOffset, positionY + newHeight \ 2 + ShadowOffset, newWidth * mirrorOffsetX, newHeight * mirrorOffsetY, , , , , _
'            55, , , , , LightAdjustment, 0, True
'    End If
    
    imageCreated = cImage.Render(picBox.hdc, positionX + newWidth \ 2, positionY + newHeight \ 2, newWidth * 1, newHeight * 1, , , , , _
        100, , , , -1, 0, 0, True)
    
    picBox.Refresh

   On Error GoTo 0
   Exit Sub

renderPicBox_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure renderPicBox of Form FireCallMain"

End Sub




' read the file chosen as the output file
' write an array the same length as your output file
' write the listbox using the array

' read the local user's file line by line, read it into an interim array and thence into a listbox.
' called by checkAndReadOutputFile during polling & populateOutputBox during startup

Public Sub readOutputFileWriteArrayWriteListbox(ByVal sFName As String)
    Dim lIndex As Long
    'Dim fileString As String
    Dim useloop As Integer
    Dim lbxCount As Integer
    Dim startLoc As Long
    Dim endLoc As Long
    Dim stepNo As Integer
    Dim stringToWrite As String
    Dim outfile As Object
    Dim outStm As ADODB.Stream
    Dim findStartPos As Integer
    Dim stringWithoutPrefix As String

    
    Const ForReading As Integer = 1
    
    lIndex = outputLineCount 'differs from the input file in that we turn the array upside down
                             ' as we need to write to the first record in the file
    
    ' resize the array to the new linecount
    ReDim outputFileArray(outputLineCount)
    

    
    If ioMethodADO = False Then
    '     we use the FSO method rather than VB6 input as it is more friendly to unix crlf EOLs
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set outfile = fso.OpenTextFile(sFName, ForReading, False, 0)
    '    ' this reads the output file and populates the output array
        Do While Not outfile.AtEndOfStream
            outputFileArray(lIndex) = outfile.readLine
            lIndex = lIndex - 1 'differs from the input file in that we turn the array upside down
            If lIndex <= 0 Then '
                Exit Do
            End If
        Loop
        outfile.Close
    Else
    
        Set outStm = New ADODB.Stream ' 47ms in the IDE, slower than FSO by factor of 2 but 9-11ms when compiled! 2,400 lines
        With outStm
            .Open
            .LoadFromFile sFName
            .Type = 2
            .Charset = "utf-8"
            .LineSeparator = -1
            .Position = 2
            
            ' this reads the output file and populates the output array
            Do While Not .EOS

                outputFileArray(lIndex) = .ReadText(-2)

                lIndex = lIndex - 1
                If lIndex <= 0 Then
                    Exit Do
                End If
            Loop
            .Close
        End With
    End If
    
 ' timer code BEGINS
'        lngReturn = QueryPerformanceCounter(curEnd)
'        sngTime = (curEnd - curStart) * 1000 / curFreq
'        Debug.Print "Execution Time " & sngTime & " mS"
'        MsgBox "Execution Time " & sngTime & " mS"
'
' timer code ENDS
    


    lbxCount = 0
    If Val(FCWLoadBottom) = 1 Then
            startLoc = outputLineCount
            endLoc = 0
            stepNo = -1
    Else
            startLoc = 0
            endLoc = outputLineCount
            stepNo = 1
    End If
    
    ' locks the input listbox whilst the listbox is updated from the array
    LockWindowUpdate FireCallMain.lbxOutputTextArea.hwnd
    
    
    
        
    ' this reads the output array and populates the listbox
    For useloop = startLoc To endLoc Step stepNo
        stringToWrite = vbNullString
        If outputFileArray(outputLineCount - useloop) <> vbNullString Then 'differs from the input file in that we turn the array upside down
            'If InStr(outputFileArray(outputLineCount - useloop), "my bloody") Then MsgBox "bloody"
            
            
            If FireCallMain.lbxOutputTextArea.List(lbxCount) = outputFileArray(outputLineCount - useloop) Then
                ' we do not clear down the array and so a comparison of the array contents to the file is performed
                ' and if they are the same we make no changes, this prevents the full listbox update
            Else
                ' only update the lines that have changed
                ' <><> attachment
                'find the line start point minus the timestamp and prefix:
                findStartPos = InStr(outputFileArray(outputLineCount - useloop), ":    ")
                stringWithoutPrefix = Mid$(outputFileArray(outputLineCount - useloop), findStartPos + 5, Len(outputFileArray(outputLineCount - useloop)))
                
                If InStr(stringWithoutPrefix, "<><>") = 1 Then
                   stringToWrite = Replace$(outputFileArray(outputLineCount - useloop), "<><>", "New File:")
                ElseIf InStr(stringWithoutPrefix, "<f><f>") = 1 Then
'                ' <f><f> emoji
                   stringToWrite = Replace$(outputFileArray(outputLineCount - useloop), "<f><f>", "New Folder:")
                ElseIf InStr(stringWithoutPrefix, "<o><o>") = 1 Then
'                ' <o><o> emoji
                   stringToWrite = Replace$(outputFileArray(outputLineCount - useloop), "<o><o>", "New Emoji:")
                ElseIf InStr(stringWithoutPrefix, "<p><p>") = 1 Then
'                ' <p><p> ping
                    stringToWrite = Replace$(outputFileArray(outputLineCount - useloop), "<p><p>", "Ping Request.")
                ElseIf InStr(stringWithoutPrefix, "<b><b>") = 1 Then
'                ' <b><b> Buzzer
                    stringToWrite = Replace$(outputFileArray(outputLineCount - useloop), "<b><b>", "Attention!")
                ElseIf InStr(stringWithoutPrefix, "<t><t>") = 1 Then
                    ' <t><t> Awake
                    stringToWrite = Replace$(outputFileArray(outputLineCount - useloop), "<t><t>", "Awake at:")
                ElseIf InStr(stringWithoutPrefix, "<z><z>") = 1 Then
                    ' <z><z> Shutdown
                    stringToWrite = Replace$(outputFileArray(outputLineCount - useloop), "<z><z>", "Shutdown at:")
                ElseIf InStr(stringWithoutPrefix, "<r><r>") = 1 Then
                    ' <z><z> Recording
                    stringToWrite = Replace$(outputFileArray(outputLineCount - useloop), "<r><r>", "New Recording:")
                Else
                    stringToWrite = outputFileArray(outputLineCount - useloop)
                End If
                FireCallMain.lbxOutputTextArea.List(lbxCount) = stringToWrite
                
            End If
            
            lbxCount = lbxCount + 1
        End If
    Next useloop
    
    LockWindowUpdate 0&
    
    ' litle fix to prevent potential duplication of content in the listboxes
    Dim listCounter As Long
    listCounter = FireCallMain.lbxOutputTextArea.ListCount
    If listCounter > outputLineCount Then
        'MsgBox "This message pops up to prevent a duplication occurring in the OUTPUT. Please report if this occurs."
        For useloop = (listCounter + 1) To outputLineCount
            FireCallMain.lbxOutputTextArea.List(useloop) = ""
        Next useloop
    End If
    
    oldOutputLineCount = outputLineCount
    
 End Sub

' click on the slim strip of paper that shows an emerging Emoji
Public Sub clickOnPicEmoji()
    Dim soundtoplay As String
    Dim nought As String
    'Dim fullPath As String
    
    If FireCallMain.picEmoji.Top < 2000 Then
        If FireCallMain.printerTimer.Enabled = True Then
            FireCallMain.printerTimer.Enabled = False

            If FCWEnableSounds = "1" Then PlaySound nought, ByVal 0&, SND_FILENAME Or SND_ASYNC
            FireCallMain.picEmoji.Top = 2000
        Else
        
            FireCallMain.brightTimer.Enabled = True
            
'            fullpath = App.Path & "\resources\images\" & "lidBackgroundBright.jpg"
'
'            if fFExists(fullpath) Then
'                FireCallMain.picLidBackground.Picture = LoadPicture(fullpath)
'            End If
            
            If toolTipFlag = True Then FireCallMain.picEmoji.ToolTipText = "Click on me to stop the printing"
            FireCallMain.printerTimer.Enabled = True
            
            If FCWPlayVolume = "1" Then
                soundtoplay = App.Path & "\Resources\Sounds\" & "computalk2.wav"
            Else
                soundtoplay = App.Path & "\Resources\Sounds\" & "computalk2Quiet.wav"
            End If
        
            If fFExists(soundtoplay) And FCWEnableSounds = "1" Then PlaySound soundtoplay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        End If
    Else
        If FireCallMain.shredderTimer.Enabled = False Then
            If toolTipFlag = True Then FireCallMain.picEmoji.ToolTipText = "Click on me to stop the shredding"
            FireCallMain.shredderTimer.Enabled = True
            
            If FCWPlayVolume = "1" Then
                soundtoplay = App.Path & "\Resources\Sounds\" & "shredder.wav"
            Else
                soundtoplay = App.Path & "\Resources\Sounds\" & "shredderQuiet.wav"
            End If
        
            If fFExists(soundtoplay) And FCWEnableSounds = "1" Then PlaySound soundtoplay, ByVal 0&, SND_FILENAME Or SND_ASYNC
        Else
            FireCallMain.shredderTimer.Enabled = False
            If toolTipFlag = True Then FireCallMain.picEmoji.ToolTipText = "Click on me to stop the shredding"

            If fFExists(nought) And FCWEnableSounds = "1" Then PlaySound nought, ByVal 0&, SND_FILENAME Or SND_ASYNC
            FireCallMain.picEmoji.Top = -1200
        End If
        
        Call FireCallMain.dropTimer_TimerSub
    End If

End Sub
'---------------------------------------------------------------------------------------
' Procedure : readSettingsFile
' Author    : beededea
' Date      : 12/05/2020
' Purpose   : read the application's setting file and assign values to public vars
'---------------------------------------------------------------------------------------
'
Public Sub readSettingsFile(ByVal location As String, ByVal FCWSettingsFile As String)
    
    
    ' On Error GoTo readSettingsFile_Error
    'If debugflg = 1 Then DebugPrint "%readFCWFCWSettingsFile"

    If fFExists(FCWSettingsFile) Then
        'General Tab Items
        FCWSharedInputFile = fGetINISetting(location, "sharedInputFile", FCWSettingsFile)
        FCWSharedOutputFile = fGetINISetting(location, "sharedOutputFile", FCWSettingsFile)
        FCWExchangeFolder = fGetINISetting(location, "exchangeFolder", FCWSettingsFile)
        FCWRefreshIntervalIndex = fGetINISetting(location, "refreshIntervalIndex", FCWSettingsFile)
        FCWRefreshIntervalSecs = fGetINISetting(location, "refreshIntervalSecs", FCWSettingsFile)
        FCWAlarmSound = fGetINISetting(location, "alarmSound", FCWSettingsFile)
        FCWAlarmSoundIndex = fGetINISetting(location, "alarmSoundIndex", FCWSettingsFile)
        
        'General Config Items
        FCWPrefixString = fGetINISetting(location, "prefixString", FCWSettingsFile)
        FCWLoadBottom = fGetINISetting(location, "loadBottom", FCWSettingsFile)
        FCWMaxLineLengthIndex = fGetINISetting(location, "maxLineLengthIndex", FCWSettingsFile)
        FCWMaxLineLength = fGetINISetting(location, "maxLineLength", FCWSettingsFile)
        FCWEnableScrollbars = fGetINISetting(location, "enableScrollbars", FCWSettingsFile)
        FCWEnableTooltips = fGetINISetting(location, "enableTooltips", FCWSettingsFile)
        FCWEnableBalloonTooltips = fGetINISetting(location, "enableBalloonTooltips", FCWSettingsFile)
        
        
        FCWIconiseDelay = fGetINISetting(location, "iconiseDelay", FCWSettingsFile)
        
        'save the values from the Emojis Config Items
        FCWEmojiSetIndex = fGetINISetting(location, "emojiSet", FCWSettingsFile)
        FCWEmojiSetDesc = fGetINISetting(location, "emojiSetDesc", FCWSettingsFile)
        
        'Email Config Items
        FCWSendEmails = fGetINISetting(location, "sendEmails", FCWSettingsFile)
        FCWSendErrorEmails = fGetINISetting(location, "sendErrorEmails", FCWSettingsFile)
        
        
        'FCWEmailAddress = fGetINISetting(location, "emailAddress", FCWSettingsFile)
        FCWAdviceInterval = fGetINISetting(location, "adviceInterval", FCWSettingsFile)
        FCWAdviceIntervalSecs = fGetINISetting(location, "adviceIntervalSecs", FCWSettingsFile)
        FCWLastEmail = fGetINISetting(location, "lastEmail", FCWSettingsFile)
        FCWLastHouseKeep = fGetINISetting(location, "lastHouseKeep", FCWSettingsFile)
        
        
        
        'Fonts Config Items
        FCWMainFont = fGetINISetting(location, "mainFont", FCWSettingsFile)
        FCWMainFontSize = fGetINISetting(location, "mainFontSize", FCWSettingsFile)
        FCWMainFontItalics = fGetINISetting(location, "mainFontItalics", FCWSettingsFile)
        FCWMainFontColour = fGetINISetting(location, "mainFontColour", FCWSettingsFile)
        
        
        FCWPrefsFont = fGetINISetting(location, "prefsFont", FCWSettingsFile)
        FCWPrefsFontSize = fGetINISetting(location, "prefsFontSize", FCWSettingsFile)
        FCWPrefsFontItalics = fGetINISetting(location, "prefsFontItalics", FCWSettingsFile)
        FCWPrefsFontColour = fGetINISetting(location, "prefsFontColour", FCWSettingsFile)
        
    
        'save the values from the Windows Config Items
        FCWWindowLevel = fGetINISetting(location, "WindowLevel", FCWSettingsFile)
        FCWOpacity = fGetINISetting(location, "Opacity", FCWSettingsFile)
        
        FCWMinimiseFormX = fGetINISetting(location, "minimiseFormX", FCWSettingsFile)
        FCWMinimiseFormY = fGetINISetting(location, "minimiseFormY", FCWSettingsFile)
        FCWMaximiseFormX = fGetINISetting(location, "maximiseFormX", FCWSettingsFile)
        FCWMaximiseFormY = fGetINISetting(location, "maximiseFormY", FCWSettingsFile)
        FCWFormWidth = fGetINISetting(location, "formWidth", FCWSettingsFile)
        FCWLastSoundPlayed = fGetINISetting(location, "lastSoundPlayed", FCWSettingsFile)
        FCWLastPingResponse = fGetINISetting(location, "lastPingResponse", FCWSettingsFile)
        FCWLastAwakeString = fGetINISetting(location, "lastAwakeString", FCWSettingsFile)
        FCWLastShutdown = fGetINISetting(location, "lastShutdown", FCWSettingsFile)
        FCWAllowShutdowns = fGetINISetting(location, "allowShutdowns", FCWSettingsFile)
        FCWClockStyle = fGetINISetting(location, "clockStyle", FCWSettingsFile)
        FCWEnableSounds = fGetINISetting(location, "enableSounds", FCWSettingsFile)
        FCWPlayVolume = fGetINISetting(location, "playVolume", FCWSettingsFile)
        
        FCWSmtpConfig = fGetINISetting(location, "smtpConfig", FCWSettingsFile)

        If FCWRecipientEmail = "" Then
        
            Call readSmtpConfigDetails("Software\FireCallWin", Val(FCWSmtpConfig))

'
'            FCWRecipientEmail = fGetINISetting(location, "recipientEmail", FCWSettingsFile)
'            FCWEmailSubject = fGetINISetting(location, "emailSubject", FCWSettingsFile)
'            FCWEmailMessage = fGetINISetting(location, "emailMessage", FCWSettingsFile)
        End If
        

        'Call readSmtpConfigDetails(location, Val(FCWSmtpConfig))
        
        FCWSingleListBox = fGetINISetting(location, "singleListBox", FCWSettingsFile)
        FCWImageDisplay = fGetINISetting(location, "imageDisplay", FCWSettingsFile)
        FCWOptHandleData = fGetINISetting(location, "optHandleData", FCWSettingsFile)
        FCWOptWindowWidth = fGetINISetting(location, "optWindowWidth", FCWSettingsFile)
        FCWAutomaticHousekeeping = fGetINISetting(location, "automaticHousekeeping", FCWSettingsFile)
        FCWStartup = fGetINISetting(location, "startup", FCWSettingsFile)

        FCWArchiveDays = fGetINISetting(location, "archiveDays", FCWSettingsFile)
        FCWArchiveDaysIndex = fGetINISetting(location, "archiveDaysIndex", FCWSettingsFile)
        
        FCWBackupOnStart = fGetINISetting(location, "backupOnStart", FCWSettingsFile)
        FCWAutomaticBackups = fGetINISetting(location, "automaticBackups", FCWSettingsFile)
        FCWAutomaticBackupInterval = fGetINISetting(location, "automaticBackupInterval", FCWSettingsFile)
        FCWServiceProvider = fGetINISetting(location, "serviceProvider", FCWSettingsFile)
        FCWCheckServiceProcesses = fGetINISetting(location, "checkServiceProcesses", FCWSettingsFile)
        
        FCWMsgBox13Enabled = fGetINISetting(location, "msgBox13Enabled", FCWSettingsFile)
        
        FCWCaptureDevices = fGetINISetting(location, "captureDevices", FCWSettingsFile)
        FCWCaptureDevicesIndex = fGetINISetting(location, "captureDevicesIndex", FCWSettingsFile)
        FCWRecordingQuality = fGetINISetting(location, "recordingQuality", FCWSettingsFile)
        FCWLastSelectedTab = fGetINISetting(location, "lastSelectedTab", FCWSettingsFile)
        FCWIconiseOpacity = fGetINISetting(location, "iconiseOpacity", FCWSettingsFile)
        FCWIconiseDesktop = fGetINISetting(location, "iconiseDesktop", FCWSettingsFile)
        
        FCWArchiveFolder = fGetINISetting(location, "archiveFolder", FCWSettingsFile)
        FCWBackupFolder = fGetINISetting(location, "backupFolder", FCWSettingsFile)
        
        
        FCWSkinTheme = fGetINISetting(location, "skinTheme", FCWSettingsFile)

    End If

   On Error GoTo 0
   Exit Sub

readSettingsFile_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure readSettingsFile of Module common2"

End Sub

Public Sub readSmtpConfigDetails(ByVal location As String, smtpConfigVal As Integer)
        Dim b64FCWSMTPPassword As String
        
        Set Bas64 = New Base64
        
        FCWSmtpConfigName = fGetINISetting(location, "smtpConfigName" & smtpConfigVal, FCWSettingsFile)
        FCWSmtpServer = fGetINISetting(location, "smtpServer" & smtpConfigVal, FCWSettingsFile)
        FCWSmtpUsername = fGetINISetting(location, "SMTPUsername" & smtpConfigVal, FCWSettingsFile)
        
        b64FCWSMTPPassword = fGetINISetting(location, "SMTPPassword" & smtpConfigVal, FCWSettingsFile)
    
        Bas64.Base64Buf = b64FCWSMTPPassword
        Call Bas64.Base64Decode

        FCWSmtpPassword = Bas64.sBuffer
        
        ' ûûÖÍÍÍ    settings.ini
        ' ûûÖÍÍÍ      GetPrivateProfileString
        ' ûûÖÍÍÍ    altGetPrivateProfileString
        
        ' we no longer use GetPrivateProfileString in fGetINISetting as it cannot read certain special chars
        ' generated by the encryption routine, tried base64 encoding it to no avail.
        
        'FCWSMTPPassword = altGetPrivateProfileString(location, "SMTPPassword", FCWSettingsFile)
        
        FCWSmtpPort = fGetINISetting(location, "smtpPort" & smtpConfigVal, FCWSettingsFile)
        FCWSmtpAuthenticate = fGetINISetting(location, "smtpAuthenticate" & smtpConfigVal, FCWSettingsFile)
        FCWSmtpSecurity = fGetINISetting(location, "smtpSecurity" & smtpConfigVal, FCWSettingsFile)
        
        FCWRecipientEmail = fGetINISetting(location, "recipientEmail" & smtpConfigVal, FCWSettingsFile)
        FCWEmailSubject = fGetINISetting(location, "emailSubject" & smtpConfigVal, FCWSettingsFile)
        FCWEmailMessage = fGetINISetting(location, "emailMessage" & smtpConfigVal, FCWSettingsFile)

        
End Sub

Public Sub validateSmtpInputs()

        If FCWSmtpServer = vbNullString Then FCWSmtpServer = ""
        If FCWSmtpUsername = vbNullString Then FCWSmtpUsername = ""
        If FCWSmtpPassword = vbNullString Then FCWSmtpPassword = ""
        If FCWSmtpPort = vbNullString Then FCWSmtpPort = "0"
        If FCWSmtpAuthenticate = vbNullString Then FCWSmtpAuthenticate = "0"
        If FCWSmtpSecurity = vbNullString Then FCWSmtpSecurity = "0"
End Sub




Private Function altGetPrivateProfileString(strSection As String, strKey As String, strFilePath As String) As String
    ' we no longer use GetPrivateProfileString in fGetINISetting as it cannot read certain special chars
    ' generated by the encryption routine, tried base64 encoding it to no avail.
        
    Const ForReading = 1
        
    Dim ReadIni As String
    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strLeftString, strLine

    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ReadIni = ""

    If objFSO.FileExists(strFilePath) Then
        Set objIniFile = objFSO.OpenTextFile(strFilePath, ForReading, False)
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim(objIniFile.readLine)

            ' Check if section is found in the current line
            If LCase(strLine) = "[" & LCase(strSection) & "]" Then
                strLine = Trim(objIniFile.readLine)

                ' Parse lines until the next section is reached
                Do While Left(strLine, 1) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr(1, strLine, "=", 1)
                    If intEqualPos > 0 Then
                        strLeftString = Trim(Left(strLine, intEqualPos - 1))
                        ' Check if item is found in the current line
                        If LCase(strLeftString) = LCase(strKey) Then
                            ReadIni = Trim(Mid(strLine, intEqualPos + 1))
                            ' In case the item exists but value is blank
                            If ReadIni = "" Then
                                ReadIni = " "
                            End If
                            altGetPrivateProfileString = ReadIni
                            ' Abort loop when item is found
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim(objIniFile.readLine)
                Loop
            Exit Do
            End If
        Loop
        objIniFile.Close
    End If
        
End Function

Public Sub altPutPrivateProfileString(strSection As String, strKey As String, stringToWrite As String, strFilePath As String)
    ' we no longer use WritePrivateProfileString in PutINISetting as it cannot write certain special chars
    ' generated by the encryption routine. This is a lot slower but it works!
        
    Const ForReading = 1
    Const ForWriting = 2
        
    Dim ReadIni As String
    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strLeftString, strLine
    Dim temporaryfile As String
    Dim outfile As Object
    Dim foundStr As Boolean
    
    
'    Dim objStream As ADODB.Stream
'    Set objStream = CreateObject("ADODB.Stream")
'    objStream.Charset = "utf-8"
'    objStream.Open
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    ReadIni = ""
    temporaryfile = strFilePath & "1"

    Set outfile = objFSO.OpenTextFile(temporaryfile, ForWriting, True)

    If objFSO.FileExists(strFilePath) Then
        Set objIniFile = objFSO.OpenTextFile(strFilePath, ForReading, False)
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim(objIniFile.readLine)
            foundStr = False

            ' Check if section is found in the current line
            If LCase(strLine) = "[" & LCase(strSection) & "]" Then
                outfile.WriteLine strLine
                strLine = Trim(objIniFile.readLine)
                
                ' Parse lines until the next section is reached
                Do While Left(strLine, 1) <> "["
                    ' Find position of equal sign in the line
                    intEqualPos = InStr(1, strLine, "=", 1)
                    If intEqualPos > 0 Then
                        strLeftString = Trim(Left(strLine, intEqualPos - 1))
                        ' Check if item is found in the current line
                        If LCase(strLeftString) = LCase(strKey) Then
                            ReadIni = Trim(Mid(strLine, intEqualPos + 1))
                            foundStr = True
                            
                            ' Abort loop when item is found
                            ' we do not write this one old value line to the output file
                            Exit Do
                        End If
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    If foundStr = False Then
                        ' write ALL the other lines to the temporary file
                        outfile.WriteLine strLine
                        
                        ' Continue with next line
                        strLine = Trim(objIniFile.readLine)

                    End If
                Loop
            End If
        Loop
        ' write the new line to the temporary file
        outfile.WriteLine strKey & "=" & stringToWrite
        objIniFile.Close
        'objStream.SaveToFile temporaryfile, 2
        outfile.Close
    End If

    If fFExists(temporaryfile) Then
        If fFExists(strFilePath) Then Kill strFilePath
        FileCopy temporaryfile, strFilePath
        If fFExists(temporaryfile) Then Kill temporaryfile
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : displayFontSelector
' Author    : beededea
' Date      : 29/02/2020
' Purpose   : select a font for the supplied form
'---------------------------------------------------------------------------------------
'
Private Sub displayFontSelector(ByRef currFont As String, ByRef currSize As Integer, ByRef currWeight As Integer, ByRef currStyle As Boolean, ByRef currColour As Long, ByRef currItalics As Boolean, ByRef currUnderline As Boolean, ByRef fontResult As Boolean)

       
    ' variables declared
    Dim thisFont As FormFontInfo
        
    'initialise the dimensioned variables
    'thisFont =
   
   ' On Error GoTo displayFontSelector_Error
   If debugflg = 1 Then Debug.Print "%displayFontSelector"

    With thisFont
      .Color = currColour
      .Height = currSize
      .Weight = currWeight
      '400     Font is normal.
      '700     Font is bold.
      .Italic = currItalics
      .UnderLine = currUnderline
      .Name = currFont
    End With
    
    fontResult = fDialogFont(thisFont)
    If fontResult = False Then Exit Sub
    
    ' some fonts have naming problems and the result is an empty font name field on the font selector
    If thisFont.Name = vbNullString Then thisFont.Name = "times new roman"
    If thisFont.Name = vbNullString Then Exit Sub
    
    With thisFont
        currFont = .Name
        currSize = .Height
        currWeight = .Weight
        currItalics = .Italic
        currUnderline = .UnderLine
        currColour = .Color
        'ctl = .Name & " - Size:" & .Height
    End With

   On Error GoTo 0
   Exit Sub

displayFontSelector_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure displayFontSelector of Form dockSettings"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : changeFormFont
' Author    : beededea
' Date      : 12/07/2019
' Purpose   : change the font throughout the whole form
'---------------------------------------------------------------------------------------
'
Public Sub changeFormFont(ByRef formName As Object, ByVal suppliedFont As String, ByVal suppliedSize As Integer, ByVal suppliedWeight As Integer, ByVal suppliedStyle As Boolean, ByVal suppliedItalics As Boolean, ByVal suppliedColour As Long)
        
    ' variables declared
    'Dim useloop As Integer
    Dim ctrl As Control
        
    'initialise the dimensioned variables
    'useloop = 0
    'Ctrl
    
    ' On Error GoTo changeFormFont_Error
    
    If debugflg = 1 Then Debug.Print "%" & "changeFormFont"
      
    ' a method of looping through all the controls and identifying the labels and text boxes
    For Each ctrl In formName.Controls
        If formName.Name = "FireCallPrefs" And ctrl.Name = "txtTextFont" Then

        Else
            If (TypeOf ctrl Is CommandButton) Or (TypeOf ctrl Is TextBox) Or (TypeOf ctrl Is FileListBox) Or (TypeOf ctrl Is Label) Or (TypeOf ctrl Is ComboBox) Or (TypeOf ctrl Is CheckBox) Or (TypeOf ctrl Is OptionButton) Or (TypeOf ctrl Is Frame) Or (TypeOf ctrl Is ListBox) Then
                If suppliedFont <> vbNullString Then ctrl.Font.Name = suppliedFont
                If suppliedSize > 0 Then ctrl.Font.Size = suppliedSize
                ctrl.Font.Italic = suppliedItalics
                
                Select Case True
                    Case (TypeOf ctrl Is CommandButton)
                        ' stupif fecking VB6 will not let you change the font of the forecolour on a button!
                        'Ctrl.ForeColor = suppliedColour
                    Case Else
                        ctrl.ForeColor = suppliedColour
                End Select
                
            End If
        End If
    Next
    
     
   On Error GoTo 0
   Exit Sub

changeFormFont_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure changeFormFont of Form dockSettings"
    
End Sub




'these callback functions need to be in a BAS module and not a form or the AddressOf does not work.

'---------------------------------------------------------------------------------------
' Procedure : BrowseCallbackProc
' Author    : beededea
' Date      : 20/08/2020
' Purpose   : create a folder window using APIs, the routine called by the callback in fBrowseFolder
'---------------------------------------------------------------------------------------
'
Private Function BrowseCallbackProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal lp As Long, ByVal InitDir As String) As Long
   Const BFFM_INITIALIZED As Long = 1
   Const BFFM_SETSELECTION As Long = &H466
   ' On Error GoTo BrowseCallbackProc_Error

   If (Msg = BFFM_INITIALIZED) And (InitDir <> vbNullString) Then
      Call SendMessage(hwnd, BFFM_SETSELECTION, 1, ByVal InitDir)
   End If
   BrowseCallbackProc = 0

   On Error GoTo 0
   Exit Function

BrowseCallbackProc_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure BrowseCallbackProc of Module Common"
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetAddress
' Author    :
' Date      : 20/08/2020
' Purpose   : stub routine to allow the callback above to be called
'---------------------------------------------------------------------------------------
'
Private Function GetAddress(ByVal Addr As Long) As Long
   ' On Error GoTo GetAddress_Error

   GetAddress = Addr

   On Error GoTo 0
   Exit Function

GetAddress_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure GetAddress of Module Common"
End Function

'---------------------------------------------------------------------------------------
' Procedure : fBrowseFolder
' Author    : beededea
' Date      : 20/08/2020
' Purpose   : create a folder window using APIs
'---------------------------------------------------------------------------------------
'
Public Function fBrowseFolder(ByVal hwndOwner As Long, ByVal DefFolder As String) As String
   Dim bInfo As BROWSEINFO
   Dim pidl As Long
   Dim newPath As String

   ' On Error GoTo fBrowseFolder_Error

   bInfo.hwndOwner = hwndOwner
   bInfo.lpfn = GetAddress(AddressOf BrowseCallbackProc)
   bInfo.lParam = StrPtr(DefFolder)
   pidl = SHBrowseForFolderA(bInfo)
   If (pidl) Then
      newPath = String(260, 0)
      If SHGetPathFromIDListA(pidl, newPath) Then
         newPath = Left$(newPath, InStr(1, newPath, Chr(0)) - 1)
         fBrowseFolder = newPath
      End If
      Call CoTaskMemFree(ByVal pidl)
   End If

   On Error GoTo 0
   Exit Function

fnBrowseFolder_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure fBrowseFolder of Module Common"
End Function



'---------------------------------------------------------------------------------------
' Procedure : addTargetfile
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : open a dialogbox to select a file as the target, normally a binary
'---------------------------------------------------------------------------------------
'
Public Sub addTargetFile(ByVal fieldValue As String, ByRef retFileName As String)
    Dim FilePath As String
    'Dim dllPath As String
    Dim dialogInitDir As String
    Dim retfileTitle As String
    Const x_MaxBuffer As Integer = 256
    
    If debugflg = 1 Then Debug.Print "%" & "addTargetfile"
    
    On Error Resume Next
    
    ' set the default folder to the existing reference
    If Not fieldValue = vbNullString Then
        If fFExists(fieldValue) Then
            ' extract the folder name from the string
            FilePath = fGetDirectory(fieldValue)
            ' set the default folder to the existing reference
            dialogInitDir = FilePath 'start dir, might be "C:\" or so also
        ElseIf fDirExists(fieldValue) Then ' this caters for the entry being just a folder name
            ' set the default folder to the existing reference
            dialogInitDir = fieldValue 'start dir, might be "C:\" or so also
        Else
            dialogInitDir = App.Path 'start dir, might be "C:\" or so also
        End If
    End If
    
  With x_OpenFilename
'    .hwndOwner = Me.hWnd
    .hInstance = App.hInstance
    .lpstrTitle = "Select a File Target"
    .lpstrInitialDir = dialogInitDir
    
    .lpstrFilter = "Text Files" & vbNullChar & "*.txt" & vbNullChar & "All Files" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    .nFilterIndex = 2
    
    .lpstrFile = String(x_MaxBuffer, 0)
    .nMaxFile = x_MaxBuffer - 1
    .lpstrFileTitle = .lpstrFile
    .nMaxFileTitle = x_MaxBuffer - 1
    .lStructSize = Len(x_OpenFilename)
  End With
  

  Call obtainOpenFileName(retFileName, retfileTitle) ' retfile will be buffered to 256 bytes

   On Error GoTo 0
   
   Exit Sub

addTargetfile_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure addTargetfile of Form FireCallMain"
 
End Sub


'---------------------------------------------------------------------------------------
' Procedure : obtainOpenFileName
' Author    : beededea
' Date      : 02/09/2019
' Purpose   : using GetOpenFileName API rturns file name and title, the filename will be buffered to 256 bytes
'---------------------------------------------------------------------------------------
'
Public Sub obtainOpenFileName(ByRef retFileName As String, ByRef retfileTitle As String)
   ' On Error GoTo obtainOpenFileName_Error
   If debugflg = 1 Then Debug.Print "%obtainOpenFileName"

  If GetOpenFileName(x_OpenFilename) <> 0 Then
    If x_OpenFilename.lpstrFile = "*.*" Then
        'txtTarget.Text = savLblTarget
    Else
        retfileTitle = x_OpenFilename.lpstrFileTitle
        retFileName = x_OpenFilename.lpstrFile
    End If
  Else
    'The CANCEL button was pressed
    'MsgBox "Cancel"
  End If

   On Error GoTo 0
   Exit Sub

obtainOpenFileName_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure obtainOpenFileName of Form FireCallMain"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : fGetDirectory
' Author    : beededea
' Date      : 11/07/2019
' Purpose   : get the folder or directory path as a string not including the last backslash
'---------------------------------------------------------------------------------------
'
Public Function fGetDirectory(ByRef Path As String) As String

   ' On Error GoTo fGetDirectory_Error
   'If debugflg = 1 Then DebugPrint "%" & "fnGetDirectory"

    If InStrRev(Path, "\") = 0 Then
        fGetDirectory = vbNullString
        Exit Function
    End If
    fGetDirectory = Left$(Path, InStrRev(Path, "\") - 1)

   On Error GoTo 0
   Exit Function

fnGetDirectory_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure fGetDirectory of Module Common"
End Function


'---------------------------------------------------------------------------------------
' Procedure : fExtractSuffix
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : extract the suffix from a filename
'---------------------------------------------------------------------------------------
'
Public Function fExtractSuffix(ByVal strPath As String) As String

    ' variables declared
    Dim AY() As String ' string array
    Dim Max As Integer
    
    'initialise the dimensioned variables
    Max = 0
    
    ' On Error GoTo fExtractSuffix_Error
    'If debugflg = 1 Then DebugPrint "%" & "fnExtractSuffix"
   
    If strPath = vbNullString Then
        fExtractSuffix = vbNullString
        Exit Function
    End If
        
    If InStr(strPath, ".") <> 0 Then
        AY = Split(strPath, ".")
        Max = UBound(AY)
        fExtractSuffix = AY(Max)
    Else
        fExtractSuffix = strPath
    End If

   On Error GoTo 0
   Exit Function

fnExtractSuffix_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure fExtractSuffix of Module Common"
End Function
'---------------------------------------------------------------------------------------
' Procedure : fExtractSuffixWithDot
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : extract the suffix from a filename
'---------------------------------------------------------------------------------------
'
Public Function fExtractSuffixWithDot(ByVal strPath As String) As String

    ' variables declared
    Dim AY() As String ' string array
    Dim Max As Integer:    Max = 0
    
    On Error GoTo fExtractSuffixWithDot_Error
    'If debugflg = 1 Then DebugPrint "%" & "fExtractSuffixWithDot"
   
    If strPath = vbNullString Then
        fExtractSuffixWithDot = vbNullString
        Exit Function
    End If
        
    If InStr(strPath, ".") <> 0 Then
        AY = Split(strPath, ".")
        Max = UBound(AY)
        fExtractSuffixWithDot = "." & AY(Max)
    Else
        fExtractSuffixWithDot = vbNullString
    End If

   On Error GoTo 0
   Exit Function

fExtractSuffixWithDot_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure fExtractSuffixWithDot of Module Common"
End Function

'---------------------------------------------------------------------------------------
' Procedure : fExtractFileNameNoSuffix
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : extract the filename without a suffix
'---------------------------------------------------------------------------------------
'
Public Function fExtractFileNameNoSuffix(ByVal strPath As String) As String

    ' variables declared
    Dim AY() As String ' string array
    Dim Min As Integer
    
    'initialise the dimensioned variables
    Min = 0
    
    ' On Error GoTo fExtractFileNameNoSuffix_Error
    'If debugflg = 1 Then DebugPrint "%" & "fnExtractFileNameNoSuffix"
   
    If strPath = vbNullString Then
        fExtractFileNameNoSuffix = vbNullString
        Exit Function
    End If
        
    If InStr(strPath, ".") <> 0 Then
        AY = Split(strPath, ".")
        Min = LBound(AY)
        fExtractFileNameNoSuffix = AY(Min)
    Else
        fExtractFileNameNoSuffix = strPath
    End If

   On Error GoTo 0
   Exit Function

fnExtractFileNameNoSuffix_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure fExtractFileNameNoSuffix of Module Common"
End Function


'---------------------------------------------------------------------------------------
' Procedure : fInIDE
' Author    :
' Date      : 09/02/2021
' Purpose   : checks whether the code is running in the VB6 IDE or not
'---------------------------------------------------------------------------------------
'
Public Function fInIDE() As Boolean

   ' On Error GoTo fInIDE_Error

    ' .30 DAEB 03/03/2021 frmMain.frm replaced the fInIDE function that used a variant to one without
    ' This will only be done if in the IDE
    Debug.Assert InDebugMode
    If mbDebugMode Then
        fInIDE = True
    End If

   On Error GoTo 0
   Exit Function

fnInIDE_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure fInIDE of Form dock"
End Function

'---------------------------------------------------------------------------------------
' Procedure : InDebugMode
' Author    :
' Date      : 02/03/2021
' Purpose   : using Debug.Assert sets the value of InDebugMode
'---------------------------------------------------------------------------------------
'
Private Function InDebugMode() As Boolean
   ' On Error GoTo InDebugMode_Error

    mbDebugMode = True
    InDebugMode = True

   On Error GoTo 0
   Exit Function

InDebugMode_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure InDebugMode of Form dock"
End Function

'
'---------------------------------------------------------------------------------------
' Procedure : fGetFileNameFromPath
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : A function to fGetFileNameFromPath
'---------------------------------------------------------------------------------------
'
Public Function fGetFileNameFromPath(ByRef strFullPath As String) As String
   ' On Error GoTo fGetFileNameFromPath_Error
   'If debugflg = 1 Then DebugPrint "%" & "fnGetFileNameFromPath"
   
   fGetFileNameFromPath = Right$(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))

   On Error GoTo 0
   Exit Function

fnGetFileNameFromPath_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure fGetFileNameFromPath of Module Common"
End Function



    
'---------------------------------------------------------------------------------------
' Procedure : setThemeShade
' Author    : beededea
' Date      : 26/09/2019
' Purpose   : set the theme shade, Windows classic dark/new lighter theme colours
'---------------------------------------------------------------------------------------
'
Public Sub setThemeShade(ByVal redC As Integer, ByVal greenC As Integer, ByVal blueC As Integer)
        
    ' variables declared
    'Dim a As Long
    Dim ctrl As Control
    'Dim useloop As Integer
    
    'initialise the dimensioned variables
    ' a = 0
     'Ctrl As Control
    ' useloop = 0
    
    ' RGB(redC, greenC, blueC) is the background colour used by the lighter themes
    
    FireCallPrefs.BackColor = RGB(redC, greenC, blueC)
    
    ' a method of looping through all the controls that require reversion of any background colouring
'    For Each Ctrl In FireCallMain.Controls
'        If (TypeOf Ctrl Is CommandButton) Or (TypeOf Ctrl Is CheckBox) Or (TypeOf Ctrl Is Label) Or (TypeOf Ctrl Is OptionButton) Or (TypeOf Ctrl Is Frame) Then
'          Ctrl.BackColor = RGB(redC, greenC, blueC)
'        End If
'    Next

    ' all buttons must be set to graphical
    For Each ctrl In FireCallPrefs.Controls
        If (TypeOf ctrl Is CommandButton) Or (TypeOf ctrl Is CheckBox) Or (TypeOf ctrl Is Label) Or (TypeOf ctrl Is OptionButton) Or (TypeOf ctrl Is Frame) Then
          ctrl.BackColor = RGB(redC, greenC, blueC)
        End If
    Next
    
    'FireCallPrefs.btnSave.BackColor = RGB(redC, greenC, blueC)
    
    If redC = 212 Then
        'classicTheme = True
        FireCallPrefs.mnuLight.Checked = False
        FireCallPrefs.mnuDark.Checked = True
    
        ' the general tab icon does not need alternative images as it is a square image on a background
        If fFExists(App.Path & "\config-icon.jpg") Then FireCallPrefs.picConfig.Picture = LoadPicture(App.Path & "\config-icon.jpg")
        If fFExists(App.Path & "\pennyred.jpg") Then FireCallPrefs.picEmail.Picture = LoadPicture(App.Path & "\pennyred.jpg")
        If fFExists(App.Path & "\emoji-icon.jpg") Then FireCallPrefs.picEmojis.Picture = LoadPicture(App.Path & "\emoji-icon.jpg")
        If fFExists(App.Path & "\font-icon.jpg") Then FireCallPrefs.picFonts.Picture = LoadPicture(App.Path & "\font-icon.jpg")
        If fFExists(App.Path & "\texts-icon.jpg") Then FireCallPrefs.picTexts.Picture = LoadPicture(App.Path & "\texts-icon.jpg")
        If fFExists(App.Path & "\sounds-icon.jpg") Then FireCallPrefs.picSounds.Picture = LoadPicture(App.Path & "\sounds-icon.jpg")
        If fFExists(App.Path & "\housekeepingIcon.jpg") Then FireCallPrefs.picHousekeeping.Picture = LoadPicture(App.Path & "\housekeepingIcon.jpg")
        If fFExists(App.Path & "\windowsScreenMagnify.jpg") Then FireCallPrefs.picWindow.Picture = LoadPicture(App.Path & "\windowsScreenMagnify.gif")
        
    Else
        'classicTheme = False
        FireCallPrefs.mnuLight.Checked = True
        FireCallPrefs.mnuDark.Checked = False
    
        If fFExists(App.Path & "\config-icon-light.jpg") Then FireCallPrefs.picConfig.Picture = LoadPicture(App.Path & "\config-icon-light.jpg")
        If fFExists(App.Path & "\pennyredlight.jpg") Then FireCallPrefs.picEmail.Picture = LoadPicture(App.Path & "\pennyredlight.jpg")
        If fFExists(App.Path & "\emoji-icon-light.jpg") Then FireCallPrefs.picEmojis.Picture = LoadPicture(App.Path & "\emoji-icon-light.jpg")
        If fFExists(App.Path & "\font-icon-light.jpg") Then FireCallPrefs.picFonts.Picture = LoadPicture(App.Path & "\font-icon-light.jpg")
        If fFExists(App.Path & "\texts-icon-light.jpg") Then FireCallPrefs.picTexts.Picture = LoadPicture(App.Path & "\texts-icon-light.jpg")
        If fFExists(App.Path & "\sounds-icon-light.jpg") Then FireCallPrefs.picSounds.Picture = LoadPicture(App.Path & "\sounds-icon-light.jpg")
        If fFExists(App.Path & "\housekeeping-icon-light.jpg") Then FireCallPrefs.picHousekeeping.Picture = LoadPicture(App.Path & "\housekeeping-icon-light.jpg")
        If fFExists(App.Path & "\windowsScreenMagnifyLight.jpg") Then FireCallPrefs.picWindow.Picture = LoadPicture(App.Path & "\windowsScreenMagnifyLight.jpg")
        
    End If
    
    FireCallPrefs.sliIconiseDelay.BackColor = RGB(redC, greenC, blueC)
    'FireCallPrefs.sliEmojiSize.BackColor = RGB(redC, greenC, blueC)
    'FireCallPrefs.sliBackgroundOpacity.BackColor = RGB(redC, greenC, blueC)
    FireCallPrefs.sliOpacity.BackColor = RGB(redC, greenC, blueC)
    FireCallPrefs.sliAutomaticBackupInterval.BackColor = RGB(redC, greenC, blueC)
    FireCallPrefs.sliRecordingQuality.BackColor = RGB(redC, greenC, blueC)
    
    ' these elements are normal elements that should have their styling reverted
    ' the loop above changes the background colour and we don't want that for all items
        
    ' buttons need to be set to graphical in the IDE to allow background colour change
    FireCallPrefs.btnTestEmail.BackColor = RGB(redC, greenC, blueC)

    
    PutINISetting "Software\FireCallWin", "skinTheme", FCWSkinTheme, FCWSettingsFile ' now saved to the toolsettingsfile
End Sub



'---------------------------------------------------------------------------------------
' Procedure : setThemeColour
' Author    : beededea
' Date      : 19/09/2019
' Purpose   : if the o/s is capable of supporting the classic theme it tests every 10 secs
'             to see if a theme has been switched
'
'---------------------------------------------------------------------------------------
'
Public Sub setThemeColour()
    ' variables declared
    Dim SysClr As Long
        
    'initialise the dimensioned variables
    SysClr = 0
    
   ' On Error GoTo setThemeColour_Error
   If debugflg = 1 Then Debug.Print "%setThemeColour"

    If IsThemeActive() = False Then
        'MsgBox "Windows Classic Theme detected"
        'set themed buttons to none
        Call setThemeShade(212, 208, 199)
        SysClr = GetSysColor(COLOR_BTNFACE)
        FCWSkinTheme = "dark"
        
        FireCallPrefs.mnuDark.Caption = "Dark Theme Enabled"
        FireCallPrefs.mnuLight.Caption = "Light Theme Enable"

    Else
        Call setModernThemeColours
        FireCallPrefs.mnuDark.Caption = "Dark Theme Enable"
        FireCallPrefs.mnuLight.Caption = "Light Theme Enabled"
    End If

    storeThemeColour = SysClr

   On Error GoTo 0
   Exit Sub

setThemeColour_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure setThemeColour of Form dockSettings"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : fIsRunning
' Author    : beededea
' Date      : 21/09/2019
' Purpose   : determines if a process is running or not
'---------------------------------------------------------------------------------------
'
Public Function fIsRunning(ByRef NameProcess As String, ByRef processID As Long) As Boolean

    Dim AppCount As Integer
    Dim RProcessFound As Long
    Dim SzExename As String
    'Dim ExitCode As Long
    Dim procId As Long
    Dim i As Integer
    'Dim WinDirEnv As String
    Dim binaryName As String
    'Dim folderName As String
    'Dim runningProcessFolder As String

   ' On Error GoTo fIsRunning_Error
   'If debugflg = 1 Then DebugPrint "%fnIsRunning"

    If NameProcess <> vbNullString Then
          AppCount = 0
          binaryName = fGetFileNameFromPath(NameProcess)
          'folderName = fGetDirectory(NameProcess) ' folder name of the binary in the stored process array
          uProcess.dwSize = Len(uProcess)
          hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
          RProcessFound = ProcessFirst(hSnapshot, uProcess)
          Do
            i = InStr(1, uProcess.szexeFile, Chr$(0))
            SzExename = LCase$(Left$(uProcess.szexeFile, i - 1))
            'WinDirEnv = Environ("Windir") + "\"
            'WinDirEnv = LCase$(WinDirEnv)

            If Right$(SzExename, Len(binaryName)) = LCase$(binaryName) Then

                    AppCount = AppCount + 1
                    procId = uProcess.th32ProcessID
                    'runningProcessFolder = fGetDirectory(GetExePathFromPID(procId))
'                    If LCase(runningProcessFolder) = LCase(folderName) Then
'                        fIsRunning = True
'                        processID = procId
'                    Else
'                        'MsgBox runningProcessFolder & " " & binaryName
'                        fIsRunning = False
'                    End If

                    fIsRunning = True
                    processID = procId
                    
                    Exit Function
            End If
            RProcessFound = ProcessNext(hSnapshot, uProcess)

          Loop While RProcessFound
          Call CloseHandle(hSnapshot)
    End If


   On Error GoTo 0
   Exit Function

fnIsRunning_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure fIsRunning of Module common"

End Function


' show the preferences for the application
'Public Sub btnConfig_Click()
'
'    If FireCallMain.Left + FireCallMain.Width + 200 + FireCallPrefs.Width > fScreenWidth Then
'        FireCallPrefs.Left = FireCallMain.Left - (FireCallPrefs.Width + 200)
'    Else
'        FireCallPrefs.Left = FireCallMain.Left + FireCallMain.Width + 200
'    End If
'
'    FireCallPrefs.Show
'End Sub

'
'---------------------------------------------------------------------------------------
' Procedure : Form_Unload_Sub
' Author    : beededea
' Date      : 18/08/2021
' Purpose   : a sub routine to close all open forms that we can call from other locations
'---------------------------------------------------------------------------------------
'
'Public Sub Form_Unload_Sub()
'    Dim frm As Form
'    On Error GoTo Form_Unload_Sub_Error
'
'    Call stopPollingTimer
'    Call stopIconiseTimer
'
'    End ' <- naughty!
'
'    On Error GoTo 0
'    Exit Sub
'
'Form_Unload_Sub_Error:
'
'    With Err
'         If .Number <> 0 Then
'            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Unload_Sub of Module Module1"
'            Resume Next
'          End If
'    End With
'
'End Sub


' Omit plngLeft & plngRight; they are used internally during recursion
Public Sub QuickSort(ByRef pvarArray As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim varMid As Variant
    Dim varSwap As Variant
    
    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    varMid = pvarArray((plngLeft + plngRight) \ 2)
    Do
        Do While pvarArray(lngFirst) < varMid And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While varMid < pvarArray(lngLast) And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            varSwap = pvarArray(lngFirst)
            pvarArray(lngFirst) = pvarArray(lngLast)
            pvarArray(lngLast) = varSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop Until lngFirst > lngLast
    If plngLeft < lngLast Then QuickSort pvarArray, plngLeft, lngLast
    If lngFirst < plngRight Then QuickSort pvarArray, lngFirst, plngRight
End Sub


' credit Matthew Gates
Function fIsFileAlreadyOpen(FileName As String) As Boolean
 Dim hFile As Long
 Dim lastErr As Long

 ' Initialize file handle and error variable.
 hFile = -1
 lastErr = 0

 ' Open for for read and exclusive sharing.
 hFile = lOpen(FileName, &H10)

 ' If we couldn't open the file, get the last error.
 If hFile = -1 Then
    lastErr = err.LastDllError
    Else
    ' Make sure we close the file on success.
    lClose (hFile)
 End If

 ' Check for sharing violation error.
 If (hFile = -1) And (lastErr = 32) Then
    fIsFileAlreadyOpen = True
    Else
    fIsFileAlreadyOpen = False
 End If
End Function


'---------------------------------------------------------------------------------------
' Procedure : writeSingleLineToEndOfOutputArray
' Author    : beededea
' Date      : 28/01/2022
' Purpose   : cslled by sendSomething, it only updates the outputfile array
'---------------------------------------------------------------------------------------
'
Public Sub writeSingleLineToEndOfOutputArray(ByVal thingToSend As String)
    Dim lineToSend As String
    Dim arbitrarilyCroppedLine As String
    Dim timestamp As String
    Dim lastTimeStamp As String
    Dim useloop As Long
    Dim lineLoop As Integer
    Dim lineLength As Integer
    Dim maxLineLength As Integer
    Dim thisLineCount As Single
    Dim realStartPoint As Integer
    Dim realEndPoint As Integer
    Dim nEndPoint As Integer

    On Error GoTo writeSingleLineToEndOfOutputArray_Error

    lineToSend = ""
    arbitrarilyCroppedLine = ""
    timestamp = ""
    lastTimeStamp = ""
    useloop = 0
    lineLoop = 0
    lineLength = 0
    maxLineLength = 0
    thisLineCount = 0
    realStartPoint = 0
    realEndPoint = 0
    nEndPoint = 0
    
    ' disble the polling timer - this is a switch that allows the polling timer logic to be disabled
    ' we are modifying the file so we want no-one else to do anything
    nowBeingModifiedFlag = True

    'remove tabs
    'thingToSend = replace$(thingToSend, vbTab, "")
    'thingToSend = replace$(thingToSend, "\t", "")
    
    maxLineLength = Val(FCWMaxLineLength)
        
    'count the number of potential lines in the string
    If Len(thingToSend) > maxLineLength Then
        lineLength = Len(thingToSend)
        thisLineCount = lineLength / maxLineLength
        thisLineCount = Int(thisLineCount) + 1
    Else
        thisLineCount = 1
    End If
        
    ' if it is a URL then do not chop the string into chunks
    If fMultiInstr(thingToSend, "ANY", "http", "https", "HTTP", "HTTPS", "www.", "WWW.") >= 0 Then
        lineToSend = thingToSend

        ' increment the global count of lines in the output
        outputLineCount = outputLineCount + 1

        ' redimension the array to store the additional new line, preserving the data contained therein
        ReDim Preserve outputFileArray(outputLineCount)

        ' get the last timestamp and remember it!
        lastTimeStamp = timestamp

        ' a comparison is made here and the timestamp incremented a little to prevent the later non-stable quicksort re-ordering two lines with the same timestamp
        Do
            timestamp = fGetDateInUniversalFormat ' add the timestamp format 2018-01-01 00:00:00.000 to the array
        Loop While timestamp = lastTimeStamp

        If timestamp = "" Then debugLog "%Err-I-ErrorNumber 24 - No valid timestamp generated." ' this occurred just once.

        ' write the new line to the array
        outputFileArray(outputLineCount) = timestamp & " " & FCWPrefixString & ":    " & lineToSend
    Else
        ' chop the string into sized chunks and add each line one at a time, ie. write it to the outputFileArray
        realEndPoint = 0
        realStartPoint = 1
        realEndPoint = maxLineLength
        For lineLoop = 1 To thisLineCount
            If thisLineCount > 1 Then
                arbitrarilyCroppedLine = Mid(thingToSend, realStartPoint, maxLineLength)
    
                If Len(arbitrarilyCroppedLine) >= maxLineLength Then ' all multiline pastes except for the last line normally.
                    ' if the clipped line ends on a space, ie. last char = "" then set that location as the new cutting point
                    If InStr(arbitrarilyCroppedLine, " ") = Len(arbitrarilyCroppedLine) Then
                        nEndPoint = Len(arbitrarilyCroppedLine)
                    ElseIf InStrRev(arbitrarilyCroppedLine, " ") <> 0 Then  ' if there is a space then...
                        nEndPoint = InStrRev(arbitrarilyCroppedLine, " ")   ' get the last occurrence of a space char
                    Else ' no space in the line use the whole line
                        nEndPoint = Len(arbitrarilyCroppedLine)
                    End If
    
                    realEndPoint = realStartPoint + nEndPoint
                    lineToSend = Mid$(thingToSend, realStartPoint, nEndPoint)
                    realStartPoint = realEndPoint ' reset the start point
                Else ' the last line less than the maximum line length
                    lineToSend = arbitrarilyCroppedLine
                End If
            Else
                lineToSend = thingToSend
            End If
    
            ' increment the global count of lines in the output
            outputLineCount = outputLineCount + 1
    
            ' redimension the array to store the the new line, preserving the data contained therein
            ReDim Preserve outputFileArray(outputLineCount)
    
            ' get the last timestamp and remember it!
            lastTimeStamp = timestamp
    
            ' a comparison is made here and the timestamp incremented a little to prevent the later non-stable quicksort re-ordering two lines with the same timestamp
            Do
                timestamp = fGetDateInUniversalFormat ' add the timestamp format 2018-01-01 00:00:00.000 to the array
            Loop While timestamp = lastTimeStamp
    
            If timestamp = "" Then debugLog "%Err-I-ErrorNumber 24 - No valid timestamp generated." ' this occurred just once.
    
            ' write the new line to the array
            outputFileArray(outputLineCount) = timestamp & " " & FCWPrefixString & ":    " & lineToSend
    
        Next lineLoop
    End If
    
    ' turn off the polling timer - this is a switch that allows the polling timer logic to be disabled
    ' we are modifying the file so we want no-one else to do anything
    nowBeingModifiedFlag = False

    On Error GoTo 0
    Exit Sub

writeSingleLineToEndOfOutputArray_Error:

    With err
         If .Number <> 0 Then
            debugLog "Error " & err.Number & " (" & err.Description & ") in procedure writeSingleLineToEndOfOutputArray of Module modCommon"
            Resume Next
          End If
    End With
End Sub

'---------------------------------------------------------------------------------------
' Procedure : insertLineIntoOutputArray
' Author    : beededea
' Date      : 28/01/2022
' Purpose   : called by insertSomething, it inserts text into the current position of the outputfile array
'             if it is a single line
'               increment the array size by one line
'               read the array from the end and write each value up one position
'               until the current insert position has been reached
'             if it is multiple lines then increment the array size by the number of lines
'
'---------------------------------------------------------------------------------------
'
Public Sub insertLineIntoOutputArray(ByVal thingToSend As String, ByVal thisLineNumber As Long)
    Dim lineToSend As String
    Dim arbitrarilyCroppedLine As String
    Dim timestamp As String
    Dim lastTimeStamp As String
    Dim useloop As Long
    Dim lineLoop As Integer
    Dim lineLength As Integer
    Dim maxLineLength As Integer
    Dim thisLineCount As Single
    Dim realStartPoint As Integer
    Dim realEndPoint As Integer
    Dim nEndPoint As Integer
    Dim lineToWrite As String
    Dim holdingOutputArray() As String
    Dim tempStorageArray() As String
    Dim currentArrayPosition As Integer
    Dim currPos As Integer
    Dim reorderloop As Integer
    
    On Error GoTo insertLineIntoOutputArray_Error

    lineToSend = ""
    arbitrarilyCroppedLine = ""
    timestamp = ""
    lastTimeStamp = ""
    useloop = 0
    lineLoop = 0
    lineLength = 0
    maxLineLength = 0
    thisLineCount = 0
    realStartPoint = 0
    realEndPoint = 0
    nEndPoint = 0
    lineToWrite = ""
    currentArrayPosition = 0
    currPos = 0
    reorderloop = 0
        
    ' disble the polling timer - this is a switch that allows the polling timer logic to be disabled
    ' we are modifying the file so we want no-one else to do anything
    nowBeingModifiedFlag = True

    'remove tabs
    'thingToSend = replace$(thingToSend, vbTab, "")
    'thingToSend = replace$(thingToSend, "\t", "")
    
    maxLineLength = Val(FCWMaxLineLength)
        
    'count the number of potential lines in the string
    If Len(thingToSend) > maxLineLength Then
        lineLength = Len(thingToSend)
        thisLineCount = lineLength / maxLineLength
        thisLineCount = Int(thisLineCount) + 1
    Else
        thisLineCount = 1
    End If
            
    ' resize the temporary holding array to the number of lines
    ReDim holdingOutputArray(thisLineCount)
            
    ' if it is a URL then do not chop the string into chunks
    If fMultiInstr(thingToSend, "ANY", "http", "https", "HTTP", "HTTPS", "www.", "WWW.") > 0 Then
        lineToSend = thingToSend

        currentArrayPosition = 0
                        
        ' get the last timestamp and remember it!
        lastTimeStamp = timestamp

        ' a comparison is made here and the timestamp incremented a little to prevent the later non-stable quicksort re-ordering two lines with the same timestamp
        Do
            timestamp = fGetDateInUniversalFormat ' add the timestamp format 2018-01-01 00:00:00.000 to the array
        Loop While timestamp = lastTimeStamp

        If timestamp = "" Then debugLog "%Err-I-ErrorNumber 24 - No valid timestamp generated." ' this occurred just once.

        holdingOutputArray(currentArrayPosition) = timestamp & " " & FCWPrefixString & ":    " & lineToSend

    Else
        ' chop the string into sized chunks and add each line one at a time, ie. write it to the outputFileArray
        realEndPoint = 0
        realStartPoint = 1
        realEndPoint = maxLineLength
        For lineLoop = 1 To thisLineCount
            If thisLineCount > 1 Then
                arbitrarilyCroppedLine = Mid(thingToSend, realStartPoint, maxLineLength)
    
                If Len(arbitrarilyCroppedLine) >= maxLineLength Then ' all multiline pastes except for the last line normally.
                    ' if the clipped line ends on a space, ie. last char = "" then set that location as the new cutting point
                    If InStr(arbitrarilyCroppedLine, " ") = Len(arbitrarilyCroppedLine) Then
                        nEndPoint = Len(arbitrarilyCroppedLine)
                    ElseIf InStrRev(arbitrarilyCroppedLine, " ") <> 0 Then  ' if there is a space then...
                        nEndPoint = InStrRev(arbitrarilyCroppedLine, " ")   ' get the last occurrence of a space char
                    Else ' no space in the line use the whole line
                        nEndPoint = Len(arbitrarilyCroppedLine)
                    End If
    
                    realEndPoint = realStartPoint + nEndPoint
                    lineToSend = Mid$(thingToSend, realStartPoint, nEndPoint)
                    realStartPoint = realEndPoint ' reset the start point
                Else ' the last line less than the maximum line length
                    lineToSend = arbitrarilyCroppedLine
                End If
            Else
                lineToSend = thingToSend
            End If
    
            ' get the last timestamp and remember it!
            lastTimeStamp = timestamp
    
            ' a comparison is made here and the timestamp incremented a little to prevent the later non-stable quicksort re-ordering two lines with the same timestamp
            Do
                timestamp = fGetDateInUniversalFormat ' add the timestamp format 2018-01-01 00:00:00.000 to the array
            Loop While timestamp = lastTimeStamp
    
            If timestamp = "" Then debugLog "%Err-I-ErrorNumber 24 - No valid timestamp generated." ' this occurred just once.
    
            ' write the new line to the holding array
            holdingOutputArray(lineLoop) = timestamp & " " & FCWPrefixString & ":    " & lineToSend
    
        Next lineLoop
    End If
        
    ReDim tempStorageArray(outputLineCount - thisLineNumber)
        
    currPos = 0
    ' extract the last x lines and populate temporary storage array
    For reorderloop = thisLineNumber To outputLineCount
        tempStorageArray(currPos) = outputFileArray(reorderloop) '
        outputFileArray(reorderloop) = ""
        currPos = currPos + 1
    Next reorderloop
    
    ' increment the global count of lines in the output
    outputLineCount = outputLineCount + thisLineCount
    
    ' redimension the main array to store the additional new lines, preserving the data contained therein
    ReDim Preserve outputFileArray(outputLineCount)
        
    currPos = 0
    'write the new data
    For reorderloop = thisLineNumber To thisLineNumber + thisLineCount
        outputFileArray(reorderloop) = holdingOutputArray(currPos) '
        currPos = currPos + 1
    Next reorderloop
    
    currPos = 0
    ' bring the 'pushed up' data back in
    For reorderloop = (thisLineNumber + thisLineCount) To outputLineCount
        outputFileArray(reorderloop) = tempStorageArray(currPos)
        currPos = currPos + 1
    Next reorderloop
    
    ' turn off the polling timer - this is a switch that allows the polling timer logic to be disabled
    ' we are modifying the file so we want no-one else to do anything
    nowBeingModifiedFlag = False

    On Error GoTo 0
    Exit Sub

insertLineIntoOutputArray_Error:

    With err
         If .Number <> 0 Then
            debugLog "Error " & err.Number & " (" & err.Description & ") in procedure insertLineIntoOutputArray of Module modCommon"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : sendSomething
' Author    : beededea
' Date      : 12/06/2021
' Purpose   : sends text to the output file, writes the output array and then writes the output file
'             it is only used for single lines of text
'---------------------------------------------------------------------------------------
'
Public Sub sendSomething(ByVal thingToSend As String)

    On Error GoTo sendSomething_Error
    
    Call writeSingleLineToEndOfOutputArray(thingToSend)
    
    nowBeingModifiedFlag = True
    Call writeOutputFile(FCWSharedOutputFile, outputLineCount)
    
    ' re-read the file chosen as the output file
    ' re populate the array the same length as your output file
    ' update the listbox using the array
    Call readOutputFileWriteArrayWriteListbox(FCWSharedOutputFile)
    
    CTRL_1 = False ' ensuring that the automatic click caused by the next few commands does not cause any URL to
                   ' automatically show in the browser (Ctrl+click)
                   
    ' this needs to be incremented properly on multi-line addition TBD
    If Val(FCWLoadBottom) = 1 Then
        FireCallMain.lbxOutputTextArea.ListIndex = FireCallMain.lbxOutputTextArea.ListCount - 1
    Else
        FireCallMain.lbxOutputTextArea.ListIndex = 0
    End If
    
    outputDataChangedFlag = True
    
    ' populate the combined listbox
    If FCWSingleListBox = "1" Then Call populateCombinedBox

    ' turn the polling timer back on - this is a switch that allows the polling timer logic to run
    nowBeingModifiedFlag = False
    
    On Error GoTo 0
    Exit Sub

sendSomething_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure sendSomething of Form FireCallMain"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : insertSomething
' Author    : beededea
' Date      : 12/06/2021
' Purpose   : sends text to the output file, writes the output array and then writes the output file
'             it is only used for single lines of text
'---------------------------------------------------------------------------------------
'
Public Sub insertSomething(ByVal thingToSend As String, ByVal thisLineNumber As Long)

    On Error GoTo insertSomething_Error
    
    Call insertLineIntoOutputArray(thingToSend, thisLineNumber)
    
    nowBeingModifiedFlag = True
    Call writeOutputFile(FCWSharedOutputFile, outputLineCount)
    
    ' re-read the file chosen as the output file
    ' re populate the array the same length as your output file
    ' update the listbox using the array
    Call readOutputFileWriteArrayWriteListbox(FCWSharedOutputFile)
    
    CTRL_1 = False ' ensuring that the automatic click caused by the next few commands does not cause any URL to
                   ' automatically show in the browser (Ctrl+click)
    
    ' this needs to be incremented properly on multi-line addition TBD
    If Val(FCWLoadBottom) = 1 Then
        FireCallMain.lbxOutputTextArea.ListIndex = FireCallMain.lbxOutputTextArea.ListCount - 1
    Else
        FireCallMain.lbxOutputTextArea.ListIndex = 0
    End If
    
    outputDataChangedFlag = True
    
    ' populate the combined listbox
    If FCWSingleListBox = "1" Then Call populateCombinedBox

    ' turn the polling timer back on - this is a switch that allows the polling timer logic to run
    nowBeingModifiedFlag = False
    
    On Error GoTo 0
    Exit Sub

insertSomething_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure insertSomething of Form FireCallMain"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : sendMultipleThings
' Author    : beededea
' Date      : 28/01/2022
' Purpose   : Same as sendSomething but it only writes the output array and does NOT write the file for each line sent.
'             At least not until after the last line of text has been processed.
'             This allows for less cpu, i/o and it avoids the file lock error that can occur on multiple file writes
'---------------------------------------------------------------------------------------
'
Public Sub sendMultipleThings()
    
    On Error GoTo sendMultipleThings_Error
    
    nowBeingModifiedFlag = True
    Call writeOutputFile(FCWSharedOutputFile, outputLineCount)

    ' re-read the file chosen as the output file
    ' re populate the array the same length as your output file
    ' update the listbox using the array
    Call readOutputFileWriteArrayWriteListbox(FCWSharedOutputFile)
    
    CTRL_1 = False ' ensuring that automatic click caused by the next few commands does not cause any URL to
                   ' automatically show in the browser (Ctrl+click)
    If Val(FCWLoadBottom) = 1 Then
        FireCallMain.lbxOutputTextArea.ListIndex = FireCallMain.lbxOutputTextArea.ListCount - 1
    Else
        FireCallMain.lbxOutputTextArea.ListIndex = 0
    End If
    
    outputDataChangedFlag = True
    
    ' populate the combined listbox
    If FCWSingleListBox = "1" Then Call populateCombinedBox

    ' turn the polling timer back on - this is a switch that allows the polling timer logic to run
    nowBeingModifiedFlag = False
    

    On Error GoTo 0
    Exit Sub

sendMultipleThings_Error:

    With err
         If .Number <> 0 Then
            debugLog "Error " & err.Number & " (" & err.Description & ") in procedure sendMultipleThings of Module modCommon"
            Resume Next
          End If
    End With

End Sub

'---------------------------------------------------------------------------------------
' Procedure : insertMultipleThings
' Author    : beededea
' Date      : 28/01/2022
' Purpose   : Same as sendSomething but it only writes the output array and does NOT write the file
'             This allows for less cpu, i/o and it avoids the file lock error that can occur on multiple file writes
'---------------------------------------------------------------------------------------
'
Public Sub insertMultipleThings()
    
    On Error GoTo insertMultipleThings_Error
    
    nowBeingModifiedFlag = True
    Call writeOutputFile(FCWSharedOutputFile, outputLineCount)

    ' re-read the file chosen as the output file
    ' re populate the array the same length as your output file
    ' update the listbox using the array
    Call readOutputFileWriteArrayWriteListbox(FCWSharedOutputFile)
    
    CTRL_1 = False ' ensuring that automatic click caused by the next few commands does not cause any URL to
                   ' automatically show in the browser (Ctrl+click)
    If Val(FCWLoadBottom) = 1 Then
        FireCallMain.lbxOutputTextArea.ListIndex = FireCallMain.lbxOutputTextArea.ListCount - 1
    Else
        FireCallMain.lbxOutputTextArea.ListIndex = 0
    End If
    
    outputDataChangedFlag = True
    
    ' populate the combined listbox
    If FCWSingleListBox = "1" Then Call populateCombinedBox

    ' turn the polling timer back on - this is a switch that allows the polling timer logic to run
    nowBeingModifiedFlag = False
    

    On Error GoTo 0
    Exit Sub

insertMultipleThings_Error:

    With err
         If .Number <> 0 Then
            debugLog "Error " & err.Number & " (" & err.Description & ") in procedure insertMultipleThings of Module modCommon"
            Resume Next
          End If
    End With

End Sub
'---------------------------------------------------------------------------------------
' Procedure : writeOutputFile
' Author    : beededea
' Date      : 28/01/2022
' Purpose   : called by sendSomething and ,
'             It writes the selected output file typically after the assocated array has been updated
'---------------------------------------------------------------------------------------
'
Private Sub writeOutputFile(ByVal fileToUpdate As String, ByVal thisLineCount As Long)

    Dim outfile As Object
    Dim wendCount As Long
    Dim useloop As Long
    Dim stringToWrite As String
    
    Dim sendStm As ADODB.Stream
    Dim BinaryStream As New ADODB.Stream
    
    Const ForWriting As Integer = 2
    Const adStateClosed As Long = 0 'Indicates that the object is closed.
    Const adStateOpen As Long = 1 'Indicates that the object is open.
    Const adStateConnecting As Long = 2 'Indicates that the object is connecting.
    Const adStateExecuting As Long = 4 'Indicates that the object is executing a command.
    Const adStateFetching As Long = 8 'Indicates that the rows of the object are being retrieved.
    
    ' now write the new line to the file, we have to rewrite the whole file for each line (!) as the newest data is
    ' always at the beginning and you cannot just write to the beginning of a Windows file in sequential mode as
    ' they are not record orientated.
    
    ' Alternative options:
    ' read/write in binary mode storing each line with a key of some sort using WriteUTF8FileEx <-preferred
    ' write each line to a temporary file and then append the full text file to the new one each time
    
    ' turn off the polling timer - this is a switch that allows the polling timer logic to be disabled
    ' we are modifying the file so we want no-one else to do anything
    On Error GoTo writeOutputFile_Error

    ' nowBeingModifiedFlag = True
    
    If ioMethodADO = False Then
        ' using the FileSystemObject as it handles EOL with CrLf whereas INPUT LINE does not.
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set outfile = fso.OpenTextFile(fileToUpdate, ForWriting, 0)
        For useloop = thisLineCount To 1 Step -1
            stringToWrite = outputFileArray(useloop)
            If RTrim$(LTrim$(stringToWrite)) <> vbNullString Then
                ' vbLf , vbCrLf & vbCr
                outfile.Write stringToWrite & vbCrLf
            End If
        Next useloop
        outfile.Close
        
    Else

        ' We open two streams, one to write the data with a two-byte BOM (byte order marker), the default
        ' for a utf8 binary streams using ADO whilst the second stream takes the data and rewrites it without
        ' the BOM by changing the position to 3 prior to writing it.
        ' We do all these shennanigans in order to write utf8 data that is compatible with unix systems
        ' as the other client is written using javascript and operates as a unix style app on both Mac osX and Windows.

        Set sendStm = New ADODB.Stream
        Set BinaryStream = New ADODB.Stream
    
        With sendStm
            .Open
            .Type = 2 '2 text
            .Charset = "UTF-8"
            .LineSeparator = -1 ' adCRLF 'vbCrLf
            .Position = 0 ' normal write with BOM
            .Flush ' flush it beforehand
        End With
    
        If (BinaryStream.State And adStateOpen) = adStateOpen Then
            fileToUpdate = fileToUpdate
            'BinaryStream.Close
        End If
                
        wendCount = 0
        While fIsFileAlreadyOpen(fileToUpdate) = True
            ' we wait until the file is fully closed
            'fileToUpdate = fileToUpdate
            wendCount = wendCount + 1
            If wendCount >= 250000 Then ' breakout to prevent eternal looping when the file is open
                
                debugLog "%Err-I-ErrorNumber 23 - ADO Error number 3004, a File Write Error. Dropbox synch. error (backlog). Your internet connection is either very slow or Dropbox is struggling to synchronise."

                'MsgBox ">> ADO Error 3004 File Write Error in procedure sendSomething of Form FireCallMain" ' same error as pushed out by ADO
                Exit Sub
            End If
        Wend
        
        

        With BinaryStream
            .Type = 1
            .Mode = 3 'adModeReadWrite
            .Open
        End With
    
        ' read the output array line by line and write the stream line by line in utf8 with BOM.
        For useloop = thisLineCount To 1 Step -1
            stringToWrite = outputFileArray(useloop)
            If RTrim$(LTrim$(stringToWrite)) <> vbNullString Then
                sendStm.WriteText stringToWrite, adWriteLine    'adWriteLine = line by line
            End If
        Next useloop
    
        sendStm.Position = 3 'Strips BOM (sets to start beyond the first 3 bytes)
        sendStm.CopyTo BinaryStream ' copy to the binary stream witout BOM
        sendStm.Position = 0 ' set back to default
        
        ' check the o/s has already closed the file, when copying/paste-ing multiple lines, VB6 runs so quickly
        ' that the o/s does not always have time to complete the previous write before the program wishes to
        ' write the next line. The wend allows us to check and loop until the external task is complete and the
        ' file is closed.
        
'        wendCount = 0
'        While fIsFileAlreadyOpen(fileToUpdate & "standard") = True
'            ' we wait until the file is fully closed
'            wendCount = wendCount + 1
'            If wendCount >= 100 Then ' breakout to prevent eternal looping when the file is open
'                MsgBox "Error 3004 File Write Error in procedure sendSomething of Form FireCallMain" ' same error as pushed out by ADO
'                Exit Sub
'            End If
'        Wend
        
'        sendStm.SaveToFile fileToUpdate & "standard", 2  ' 1 = no overwrite, 2 = overwrite /adSaveCreateOverWrite
        sendStm.Flush
        sendStm.Close
        
        BinaryStream.SaveToFile fileToUpdate, 2
        BinaryStream.Flush
        BinaryStream.Close
    End If

    On Error GoTo 0
    Exit Sub

writeOutputFile_Error:

    With err
         If .Number <> 0 Then
            debugLog "Error " & err.Number & " (" & err.Description & ") in procedure writeOutputFile of Module modCommon"
            Resume Next
          End If
    End With
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : updateArchiveFile
' Author    : beededea
' Date      : 28/01/2022
' Purpose   : updates the archive file
'             quite similar to writeOutputFile but with minor changes
'---------------------------------------------------------------------------------------
'
Private Function updateArchiveFile(ByRef thisArray() As String, ByVal fileToUpdate As String, ByVal thisLineCount As Long)

    Dim outfile As Object
    Dim wendCount As Long
    Dim useloop As Long
    Dim stringToWrite As String
    
    Dim sendStm As ADODB.Stream
    Dim BinaryStream As New ADODB.Stream
    
    Const ForWriting As Integer = 2
    Const adStateClosed As Long = 0 'Indicates that the object is closed.
    Const adStateOpen As Long = 1 'Indicates that the object is open.
    Const adStateConnecting As Long = 2 'Indicates that the object is connecting.
    Const adStateExecuting As Long = 4 'Indicates that the object is executing a command.
    Const adStateFetching As Long = 8 'Indicates that the rows of the object are being retrieved.
    Const adSaveCreateNotExist As Long = 1
    
    ' now write the new line to the file, we have to rewrite the whole file for each line (!) as the newest data is
    ' always at the beginning and you cannot just write to the beginning of a Windows file in sequential mode as
    ' they are not record orientated.
    
    ' Alternative options:
    ' read/write in binary mode storing each line with a key of some sort using WriteUTF8FileEx <-preferred
    ' write each line to a temporary file and then append the full text file to the new one each time
    
    ' turn off the polling timer - this is a switch that allows the polling timer logic to be disabled
    ' we are modifying the file so we want no-one else to do anything
    'On Error GoTo updateArchiveFile_Error

    ' nowBeingModifiedFlag = True

    
    If ioMethodADO = False Then
        ' using the FileSystemObject as it handles EOL with CrLf whereas INPUT LINE does not.
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set outfile = fso.OpenTextFile(fileToUpdate, ForWriting, 0)
        For useloop = thisLineCount To 1 Step -1
            stringToWrite = thisArray(useloop)
            If RTrim$(LTrim$(stringToWrite)) <> vbNullString Then
                ' vbLf , vbCrLf & vbCr
                outfile.Write stringToWrite & vbCrLf
            End If
        Next useloop
        outfile.Close
        
    Else

        ' We open two streams, one to write the data with a two-byte BOM (byte order marker), the default
        ' for a utf8 binary streams using ADO whilst the second stream takes the data and rewrites it without
        ' the BOM by changing the position to 3 prior to writing it.
        ' We do all these shennanigans in order to write utf8 data that is compatible with unix systems
        ' as the other client is written using javascript and operates as a unix style app on both Mac osX and Windows.

        Set sendStm = New ADODB.Stream
        Set BinaryStream = New ADODB.Stream
    
        With sendStm
            .Open
            .Type = 2 '2 text
            .Charset = "UTF-8"
            .LineSeparator = -1 ' adCRLF 'vbCrLf
            .Position = 0 ' normal write with BOM
            .Flush ' flush it beforehand
        End With
    
        If (BinaryStream.State And adStateOpen) = adStateOpen Then
            fileToUpdate = fileToUpdate
            'BinaryStream.Close
        End If
                
        wendCount = 0
        While fIsFileAlreadyOpen(fileToUpdate) = True
            ' we wait until the file is fully closed
            'fileToUpdate = fileToUpdate
            wendCount = wendCount + 1
            If wendCount >= 250000 Then ' breakout to prevent eternal looping when the file is open
                
                debugLog "%Err-I-ErrorNumber 23 - ADO Error number 3004, a File Write Error. Dropbox synch. error (backlog). Your internet connection is either very slow or Dropbox is struggling to synchronise."

                'MsgBox ">> ADO Error 3004 File Write Error in procedure sendSomething of Form FireCallMain" ' same error as pushed out by ADO
                Exit Function
            End If
        Wend

        With BinaryStream
            .Type = 1
            .Mode = 3 'adModeReadWrite
            .Open
        End With
    
        ' read the output array line by line and write the stream line by line in utf8 with BOM.
        For useloop = thisLineCount To 1 Step -1
            stringToWrite = thisArray(useloop)
            If RTrim$(LTrim$(stringToWrite)) <> vbNullString Then
                sendStm.WriteText stringToWrite, adWriteLine    'adWriteLine = line by line
            End If
        Next useloop
    
        sendStm.Position = 3 'Strips BOM (sets to start beyond the first 3 bytes)
        sendStm.CopyTo BinaryStream ' copy to the binary stream witout BOM
        sendStm.Position = 0 ' set back to default
        
        ' check the o/s has already closed the file, when copying/paste-ing multiple lines, VB6 runs so quickly
        ' that the o/s does not always have time to complete the previous write before the program wishes to
        ' write the next line. The wend allows us to check and loop until the external task is complete and the
        ' file is closed.
        
'        wendCount = 0
'        While fIsFileAlreadyOpen(fileToUpdate & "standard") = True
'            ' we wait until the file is fully closed
'            wendCount = wendCount + 1
'            If wendCount >= 100 Then ' breakout to prevent eternal looping when the file is open
'                MsgBox "Error 3004 File Write Error in procedure sendSomething of Form FireCallMain" ' same error as pushed out by ADO
'                Exit Sub
'            End If
'        Wend
        
'        sendStm.SaveToFile fileToUpdate & "standard", 2  ' 1 = no overwrite, 2 = overwrite /adSaveCreateOverWrite
        sendStm.Flush
        sendStm.Close
        
        BinaryStream.SaveToFile fileToUpdate, adSaveCreateNotExist
        BinaryStream.Flush
        BinaryStream.Close
        
        If fFExists(fileToUpdate) Then updateArchiveFile = True
        
    End If

    On Error GoTo 0
    Exit Function

updateArchiveFile_Error:

    updateArchiveFile = False
    With err
         If .Number <> 0 Then
             MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure updateArchiveFile of Module modCommon"
            Resume Next
          End If
    End With
    
End Function
' Olaf Schmidt
Function Epoch2Date(ByVal E As Currency, Optional msFrac) As Date
Const Estart As Double = #1/1/1970#
   
  msFrac = 0
  If E > 10000000000@ Then E = E * 0.001: msFrac = E - Int(E)
  Epoch2Date = Estart + (E - msFrac) / 86400
End Function

'---------------------------------------------------------------------------------------
' Procedure : fSecondsFromDateString
' Author    : Olaf Schmidt
' Date      : 18/03/2022
' Purpose   :
'---------------------------------------------------------------------------------------
'
Function fSecondsFromDateString(ByVal dateString As String, Optional msFrac As Integer) As Currency
Const Estart As Double = #1/1/1970#

    Dim C As String
    Dim E As String
    Dim d As Date
    
    C = Mid$(dateString, 1, 19)
    d = CDate(C)
    msFrac = Val(Mid$(dateString, 21, 3))
    
    fSecondsFromDateString = CLng((d - Estart) * 86400) '  1643899670

    On Error GoTo 0
    Exit Function

fSecondsFromDateString_Error:

    With err
         If .Number <> 0 Then
            MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fSecondsFromDateString of Module modCommon"
            Resume Next
          End If
    End With
End Function

''need to converge the following four functions
'Private Function fConvertEpochToVB6String(unixTimeStamp As String) As String
'' expects a string in seconds not a timestamp
'     Dim stampMinusSubSecond As Long
'     Dim vb6DateTime As Date
'     Dim formattedString As String
'     Dim bias As Long
'
'     stampMinusSubSecond = Val(unixTimeStamp) / 1000
'     vb6DateTime = fVb6DateFromEpoch(stampMinusSubSecond, bias)
'
'     formattedString = Format$(vb6DateTime, "dd mmm yyyy hh:nn:ss")
'     bias = bias / 60
'     formattedString = DateAdd("h", bias, formattedString)
'
'     fConvertEpochToVB6String = formattedString
'End Function

' qvb6   https://www.vbforums.com/member.php?291519-qvb6
Public Function fVb6DateFromEpoch(ByVal iEpoch As Long, Optional ByRef bias As Long) As Date
    Dim C As Currency
    Dim u As TIME_ZONE_INFORMATION
    Dim ret As Long

    GetMem4 iEpoch, C
    ' 86400 is the number of seconds in a day.
    ' 25569 is the days between 1/1/1970 (Epoch origin) and 12/30/1899 (VB6 origin).
    ret = GetTimeZoneInformation(u)
    If ret = TIME_ZONE_ID_DAYLIGHT Then
        bias = u.bias + u.DaylightBias
    Else
        bias = u.bias
    End If
    fVb6DateFromEpoch = CDbl(C * 10000@) / 86400# + 25569# - (CDbl(bias) / 1440#)
End Function

' Format a date string from a unix time stamp     ' Wed, 30 Jun 2021 14:55:27 GMT
Private Function fConvertEpochToTimeString(unixTimeStamp As String) As String
     Dim stampMinusSubSecond As Long
     Dim vb6DateTime As Date
     Dim formattedString As String
     Dim bias As Long
     
     stampMinusSubSecond = Val(unixTimeStamp) / 1000
     vb6DateTime = fVb6DateFromEpoch(stampMinusSubSecond, bias)
     
     formattedString = Format$(vb6DateTime, "dd mmm yyyy hh:nn:ss")
     bias = bias / 60
     formattedString = DateAdd("h", bias, formattedString)
            
     Dim dayOfWeek As String
     Select Case DatePart("w", vb6DateTime)
         Case vbSunday
             dayOfWeek = "Sun"
         Case vbMonday
             dayOfWeek = "Mon"
         Case vbTuesday
             dayOfWeek = "Tue"
         Case vbWednesday
             dayOfWeek = "Wed"
         Case vbThursday
             dayOfWeek = "Thu"
         Case vbFriday
             dayOfWeek = "Fri"
         Case vbSaturday
             dayOfWeek = "Sat"
     End Select
     
     fConvertEpochToTimeString = dayOfWeek & ", " & formattedString & " GMT" ' & biasString
     
End Function



' universal time format is required for unix systems that we may be chatting with
'fnGetDateInUniversalFormat Austin Hickl http://computer-programming-forum.com/66-vb-controls/6dff1bae05df0a6e.htm
'- formats date in form "1998.12.31 23:59:59.456
Public Function fGetDateInUniversalFormat() As String
  Dim TimeZoneInfo As TIME_ZONE_INFORMATION
  'Dim currentBias As Long
  Dim currentLocaltime As SYSTEMTIME

  'Windows returns the inverse of the bias we need
'  If GetTimeZoneInformation(TimeZoneInfo) = TIME_ZONE_ID_DAYLIGHT Then
'    currentBias = -(TimeZoneInfo.bias + TimeZoneInfo.DaylightBias)
'  Else
'    currentBias = -(TimeZoneInfo.bias + TimeZoneInfo.StandardBias)
'  End If

  GetSystemTime currentLocaltime
  

  With currentLocaltime
    fGetDateInUniversalFormat = Format$(.wYear, "0000") & "-" & Format(.wMonth, "00") & "-" & Format(.wDay, "00") & " " & Format$(.wHour, "00") & ":" & Format(.wMinute, "00") & ":" & Format(.wSecond, "00") & "." & Right$(Format(.wMilliseconds, "000"), 3) '& " " & FormatTimezoneOffset(currentBias)
  End With
End Function 'fnGetDateInUniversalFormat
Public Function fGetDateNoChars() As String
  Dim TimeZoneInfo As TIME_ZONE_INFORMATION
  'Dim currentBias As Long
  Dim currentLocaltime As SYSTEMTIME

  'Windows returns the inverse of the bias we need
'  If GetTimeZoneInformation(TimeZoneInfo) = TIME_ZONE_ID_DAYLIGHT Then
'    currentBias = -(TimeZoneInfo.bias + TimeZoneInfo.DaylightBias)
'  Else
'    currentBias = -(TimeZoneInfo.bias + TimeZoneInfo.StandardBias)
'  End If

  GetSystemTime currentLocaltime
  

  With currentLocaltime
    fGetDateNoChars = Format$(.wYear, "0000") & "" & Format(.wMonth, "00") & "" & Format(.wDay, "00") & "" & Format$(.wHour, "00") & "" & Format(.wMinute, "00") & "" & Format(.wSecond, "00") & "" & Right$(Format(.wMilliseconds, "000"), 3) '& "" & FormatTimezoneOffset(currentBias)
  End With
End Function 'fnGetDateInUniversalFormat

'need to converge the above



''    Dim vbEmailAddress As String
''    Dim vbEmailSubject As String
''    Dim vbEmailBody As String
''
''    vbEmailAddress = FireCallPrefs.txtEmailAddress.Text
''    vbEmailSubject = "FireCall Dropbox Failure Email"
''    vbEmailBody = "Dropbox Sharing is not currently active." & vbCrLf & vbCrLf & vbCrLf
''
''    'vbEmailBody = Replace(vbEmailBody, vbCrLf, "%0D%0A")
''
''    If vbEmailAddress = "" Then Exit Sub
''    ShellExecute FireCallMain.hWnd, "open", "mailto:" & vbEmailAddress & "?subject=" & vbEmailSubject & "&body=" & vbEmailBody & Chr(34), vbNullString, vbNullString, 1
''




'
'---------------------------------------------------------------------------------------
' Procedure : checkLicenceState
' Author    : beededea
' Date      : 20/06/2019
' Purpose   : check the state of the licence
'---------------------------------------------------------------------------------------
'
Public Sub checkLicenceState(slicence As Integer)
    'Dim slicence As Integer

   'On Error GoTo checkLicenceState_Error
   'If debugflg = 1 Then DebugPrint "%" & "checkLicenceState"

    'toolSettingsFile = App.Path & "\settings.ini"
    ' read the tool's own settings file (
    If fFExists(FCWSettingsFile) Then ' does the tool's own settings.ini exist?
        slicence = Val(fGetINISetting("Software\Firecallwin", "Licence", FCWSettingsFile))
        ' if the licence state is not already accepted then display the licence form
        If slicence = 0 Then
            Call LoadFileToTB(licence.txtLicenceTextBox, App.Path & "\licence.txt", False)
            
            licence.Show vbModal ' show the licence screen in VB modal mode (ie. on its own)
            ' on the licence box change the state fo the licence acceptance
        End If
    End If
    
    ' show the licence screen if it has never been run before and set it to be in focus
    If licence.Visible = True Then
        licence.SetFocus
    End If

   On Error GoTo 0
   Exit Sub

checkLicenceState_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure checkLicenceState of Form common"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : fSpecialFolder
' Author    :  si_the_geek vbforums
' Date      : 17/10/2019
' Purpose   : No longer used as the shell object usage causes concern to AV tools
'---------------------------------------------------------------------------------------
'
'Public Function fSpecialFolder(ByRef pFolder As eSpecialFolders) As String
'
'  Dim objShell  As Object
'  Dim objFolder As Object
'
'  ' On Error GoTo fSpecialFolder_Error
'
'  Set objShell = CreateObject("Shell.Application")
'  Set objFolder = objShell.NameSpace(CLng(pFolder))
'
'  If (Not objFolder Is Nothing) Then fSpecialFolder = objFolder.Self.Path
'
'  Set objFolder = Nothing
'  Set objShell = Nothing
'
'  if fSpecialFolder = vbNullString Then err.Raise 513, "fnSpecialFolder", "The folder path could not be detected"
'
'   On Error GoTo 0
'   Exit Function
'
'fnSpecialFolder_Error:
'
'    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure fSpecialFolder of Module Common"
'
'End Function



     
'---------------------------------------------------------------------------------------
' Procedure : fSpecialFolder
' Author    :  si_the_geek vbforums
' Date      : 17/10/2019
' Purpose   : Returns the path to the specified special folder (AppData etc)
'---------------------------------------------------------------------------------------
'
Public Function fSpecialFolder(pfe As FolderEnum) As String
    Const MAX_PATH = 260
    Dim strPath As String
    Dim strBuffer As String
    
    strBuffer = Space$(MAX_PATH)
    If SHGetFolderPath(0, pfe, 0, 0, strBuffer) = 0 Then strPath = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
    If Right$(strPath, 1) = "\" Then strPath = Left$(strPath, Len(strPath) - 1)
    fSpecialFolder = strPath
End Function

'-----------------------------------------------------------
'perform multiple instr on a string. returns true if ANY or ALL instr passes
'-----------------------------------------------------------
Public Function fMultiInstr(sToInspect As String, searchType As String, ParamArray sArrConditions()) As Integer
    Dim i As Integer
    Dim iUpp As Integer
    Dim strLoc As Integer
    
    fMultiInstr = 0
    iUpp = UBound(sArrConditions) 'instr conditions
    
    For i = 0 To iUpp ' loop them
            strLoc = InStr(1, sToInspect, sArrConditions(i))
            
            If searchType = "ANY" And strLoc > 0 Then
                fMultiInstr = strLoc
                Exit Function
            End If
            If searchType = "ALL" And strLoc <= 0 Then
                Exit Function '     if instr returns 0 then exit
            End If
    Next i
    If searchType = "ALL" Then fMultiInstr = strLoc
End Function


' routine called at startup to create or run the two email timers
Public Sub startTheEmailTimers()
    
    Dim emailIntervalMillisecs As Long

    Const lngSecs As Long = 65 ' just used to avoid multiplying two integers
    Const lngThousand As Long = 1000

    ' start the email timer in code
    If fInIDE Then
        ' VB6 timers cannot exceed 65 seconds (65535 ms)
'        lngSecs = 65
'        lngThousand = 1000
        ' when multiplying two integer values and assigning to a long in the IDE it causes a failure as the IDE is handling the two numbers as integers
        ' emailIntervalMillisecs = 65 * 1000 '  < this fails
        emailIntervalMillisecs = lngSecs * lngThousand ' works!
        FireCallMain.emailTimer.Interval = emailIntervalMillisecs
        FireCallMain.emailTimer.Enabled = True
        debugLog "Starting startEmailTimer using VB6 timer at interval of " & emailIntervalMillisecs & "ms", False
    Else
        ' using a timer in code rather than a VB6 timer as VB6 timers cannot exceed 65 seconds (65535 ms)
        ' and if you want a longer timer you have to roll your own.
        ' in addition, unfortunately this code timer method does not work in the IDE
        
        ' stop any possible running timer first
        Call stopEmailTimer
        
        ' using a timer in code rather than a VB6 timer as VB6 timers cannot exceed 65 seconds (65535 ms)
        emailIntervalMillisecs = FCWAdviceIntervalSecs * lngThousand
        
        'MsgBox "FCWAdviceIntervalSecs " & FCWAdviceIntervalSecs
        
        ' final check to prevent starting this timer when working in the IDE, should never get this far
        If Not fInIDE Then
            ' Don't start the timer If it's already running.
            If emailTimerID = 0 Then
                ' this has a callback routine that it jumps to on each interval completion
                emailTimerID = SetTimer(0, 3, emailIntervalMillisecs, AddressOf emailTimer_CodeTimer)
                debugLog "Starting startEmailTimer using API timer, ID = " & emailTimerID & " at interval of " & emailIntervalMillisecs & "ms", False
            End If
        Else
            MsgBox "Please note: Timers in code will not run in the IDE, defaulting to VB6 timers <65secs."
        End If
    End If
End Sub




' The email timer logic itself that is used by both the VB6 standard timer in the IDE and the callback timer during runtime.
' The logic is in a globally accessible separate routine as it is called directly by both the VB6 timer and the callback timer.

'   forward texts by mail on a regular basis containing all recent texts within the timestamp period
'---------------------------------------------------------------------------------------
' Procedure : emailTimer_TimerLogic
' Author    : beededea
' Date      : 23/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub emailTimer_TimerLogic()

    Dim lastInputVar As LASTINPUTINFO
    Dim currentIdleTime As Long



    Dim a As String
    Dim lastEmailTimeInSecs  As Long
    'Dim emailIntervalMillisecs As Long
    Dim oldMessageCnt As Long
    Dim errMessageCnt As Long
    
    Dim thisLine As String
    Dim findStr As Integer
    Dim timestamp As String
    Dim useloop As Long
    Dim stampTimeDiffInSecs As Long
    Dim emailBodyStr As String
    Dim errBodyStr As String
    Dim iFile As Integer
    
    Const lngThousand As Long = 1000
    
    ' initialise vars
    
   On Error GoTo emailTimer_TimerLogic_Error

    a = vbNullString
    useloop = 0
    lastEmailTimeInSecs = 0
    'emailIntervalMillisecs = 0
    oldMessageCnt = 0
    errMessageCnt = 0
    thisLine = ""
    findStr = 0
    timestamp = vbNullString
    stampTimeDiffInSecs = 0
    emailBodyStr = vbNullString
    errBodyStr = vbNullString
    iFile = 0
        
   'MsgBox "FCWAdviceIntervalSecs " & FCWAdviceIntervalSecs
    
    If Val(FCWAdviceIntervalSecs) = 0 Then Exit Sub
    If Val(FCWSendEmails) = 0 And Val(FCWSendErrorEmails) = 0 Then Exit Sub
    
    ' check to see if the app has not been used for a while, ie it has been idle
    lastInputVar.cbSize = Len(lastInputVar)
    Call GetLastInputInfo(lastInputVar)
    currentIdleTime = GetTickCount - lastInputVar.dwTime
    
    ' only allows the function to continue if FCW has been idle for more than 30 secs
    ' reason for this is that the code to send a STARTTLS/CDO email is fairly chunky and results in UI delays.
    If currentIdleTime < 30000 Then Exit Sub
    
    ' check the date/time of the last advice message
    lastEmailTimeInSecs = fSecondsFromDateString(FCWLastEmail) ' eg. FCWLastEmail="2022-02-03 13:18:08.185"
    
    FireCallMain.picWEmailIcon.Visible = True
    ' sometimes the overall processing prevents images from appearing in their expected state
    ' so we give the process a nudge
    DoEvents
    FireCallMain.picWEmailIcon.Refresh
          
    If Val(FCWSendEmails) <> 0 Then           ' advice messages
                    
        ' disable the polling timer - this is a switch that allows the polling timer logic to be disabled
        ' we are modifying the file so we want no-one else to do anything
        nowBeingModifiedFlag = True
        
        While pollingFlag = True  ' flag that indicates that polling is still underway
            ' we wait until the polling is complete, VB6 timers are at least partly asynchronous and so this waits until the polling has complete
        Wend
        
        ' if we have arrived this far then it is definitely not polling for data
              
        ' loop through the input message array to determine those that occurred since the last advice email time
        ' scroll through the lines in the input file and find their timestamps
        For useloop = 1 To UBound(inputFileArray)
            thisLine = inputFileArray(useloop)
            
            findStr = InStr(23, thisLine, "    ")
            timestamp = Mid$(thisLine, 1, 23)
          
            ' use datediff to extract the time in seconds from a timestamp
            
            If IsDate(Mid$(thisLine, 1, 19)) Then
              ' 2022-02-05 08:41:46.056
              stampTimeDiffInSecs = fSecondsFromDateString(timestamp)
            
              ' compare the timestamp of each to the interval specified and build a list - yup
              If stampTimeDiffInSecs > lastEmailTimeInSecs Then
                  emailBodyStr = thisLine & vbCrLf & emailBodyStr
                  oldMessageCnt = oldMessageCnt + 1
              Else
                  Exit For ' when the next set of timestamps are older then the email datestamp no need to loop any further
              End If
            End If
            DoEvents
        Next useloop
        
        nowBeingModifiedFlag = False
        
        ' if we have any messages to send
        If oldMessageCnt > 0 Then
            ' send an email containing the messages
            Call FireCallMain.initiateEmail(emailBodyStr)
        
            'MsgBox "Emailing recent messages"
              
            ' store the last advice timestamp to allow comparison
            FCWLastEmail = fGetDateInUniversalFormat
            If fFExists(FCWSettingsFile) Then
                PutINISetting "Software\FireCallWin", "lastEmail", FCWLastEmail, FCWSettingsFile
            End If
        End If
    End If
    
    If Val(FCWSendErrorEmails) <> 0 Then           ' error messages
    
        ' if any error message has been generated in the interval since the last advice message then append the error message
        '     this requires an error log with timestamps
    
        iFile = FreeFile
    
        Open FCWSettingsDir & "\FCWDebugOutput.log" For Input As iFile
    
        ' loop through the input message array to determine those that occurred since the last advice email time
        ' scroll through the lines in the input file and find their timestamps
        
        Do While Not EOF(iFile)
            Line Input #iFile, thisLine
            
            findStr = InStr(23, thisLine, "    ")
            timestamp = Mid$(thisLine, 1, 23)
          
            ' use datediff to extract the time in seconds from a timestamp
            
            ' 2022-02-05 08:41:46.056
            If IsDate(Mid$(thisLine, 1, 19)) Then
                stampTimeDiffInSecs = fSecondsFromDateString(timestamp)
          
                ' compare the timestamp of each to the interval specified and build a list - yup
                If stampTimeDiffInSecs > lastEmailTimeInSecs Then
                    errBodyStr = errBodyStr + thisLine & vbCrLf
                    errMessageCnt = errMessageCnt + 1
                End If
            End If
            DoEvents
        Loop
        
        Close iFile

        ' if we have any messages to send
        If errMessageCnt > 0 Then
              ' send an email containing the messages
              Call FireCallMain.initiateEmail(errBodyStr)
          
              'MsgBox "Emailing recent messages"
                
              ' store the last advice timestamp to allow comparison
              FCWLastEmail = fGetDateInUniversalFormat
              If fFExists(FCWSettingsFile) Then
                  PutINISetting "Software\FireCallWin", "lastEmail", FCWLastEmail, FCWSettingsFile
              End If
        End If
    End If
    
    FireCallMain.picWEmailIcon.ToolTipText = "If the email icon persists then it means a background email task has failed to connect"
    FireCallMain.emailIconTimer.Enabled = True

   On Error GoTo 0
   Exit Sub

emailTimer_TimerLogic_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure emailTimer_TimerLogic of Module modCommon"
    
End Sub




' The timer that stops the email timer
Private Sub stopEmailTimer()
    ' Don't stop the timer If it isn't running.
    If emailTimerID Then
        KillTimer 0, emailTimerID
        emailTimerID = 0
    End If
End Sub


' Callback routine called byAddressOf used by the email timer. Note: This function only operates at runtime,
' ie. it doesn't work in the IDE because in the IDE everything works in the main thread. Callback functions operate
' in a separate thread and this achieves basic multi threading but may limit some functionality - but basic commands seem to operate correctly

Public Sub emailTimer_CodeTimer()
    Call emailTimer_TimerLogic
End Sub



' Callback routine called byAddressOf used by the houseKeeping timer. Note: This function only operates at runtime,
' ie. it doesn't work in the IDE because in the IDE everything works in the main thread. Callback functions operate
' in a separate thread and this achieves basic multi threading but may limit some functionality - but basic commands seem to operate correctly

Public Sub houseKeepingTimer_CodeTimer()
    Call houseKeepingTimerLogic(False)
End Sub

' The housekeeping timer runs regularly
    
'       loop through the output file
'       determine each lines is older than the date
'       determine the archive location
'       write the line to an archive file
'       remove the line from the output file

'---------------------------------------------------------------------------------------
' Procedure : houseKeepingTimerLogic
' Author    : beededea
' Date      : 23/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub houseKeepingTimerLogic(ByVal bypassIdleCheck As Boolean)
    Dim lastInputVar As LASTINPUTINFO
    Dim currentIdleTime As Long

    Dim lastHouseKeepDateTimeInSecs As Long
    Dim findStr As Integer
    Dim stampTimeInSecs As Long
    Dim useloop As Long
    Dim stampTimeDiffInSecs As Long
    Dim emailBodyStr As String
    Dim errBodyStr As String
    Dim iFile As Integer
    
    Dim maxLineLength As Integer
    Dim outputArrayTimeStamp As String
    Dim archiveTimeInSecs As Long
    Dim nowInSecs As Long
    Dim lastHouseKeepDiff As Long
    Dim stampDiffInSecs As Long
    
    Dim archiveArray() As String
    Dim temporaryArray() As String
    Dim archiveFilePath As String
    
    Dim archiveLoc As Long
    Dim tempLoc As Long

    Dim a As String
    Dim timestamp As String
    Dim testArchiveExists As Boolean
    
    Const lngThousand As Long = 1000
    
    ' initialise vars

   On Error GoTo houseKeepingTimerLogic_Error

    maxLineLength = 0
    archiveTimeInSecs = 0
    outputArrayTimeStamp = ""
    
    useloop = 0
    findStr = 0
    lastHouseKeepDateTimeInSecs = 0
    stampTimeDiffInSecs = 0
    emailBodyStr = vbNullString
    errBodyStr = vbNullString
    iFile = 0
    testArchiveExists = False
            
    If Val(FCWAutomaticHousekeeping) = 0 Then Exit Sub
    
    debugLog "running automatic housekeeping using code timer"
     
    ' check to see if the app has not been used for a while, ie. it has been idle
    lastInputVar.cbSize = Len(lastInputVar)
    Call GetLastInputInfo(lastInputVar)
    currentIdleTime = GetTickCount - lastInputVar.dwTime
    
    ' only allows the function to continue if FCW's user has been idle for more than 30 secs
    If Not bypassIdleCheck = True And currentIdleTime < 30000 Then Exit Sub
    
    ' check the date/time of the last advice message
    lastHouseKeepDateTimeInSecs = fSecondsFromDateString(FCWLastHouseKeep) ' eg. FCWLastEmail="2022-02-03 13:18:08.185"
    nowInSecs = fSecondsFromDateString(Now)
    lastHouseKeepDiff = nowInSecs - lastHouseKeepDateTimeInSecs
    
    archiveTimeInSecs = Val(FCWArchiveDays) * 24 * 3600
    
    ' here we extract the time in seconds from the housekeeping archive period
    If lastHouseKeepDiff > archiveTimeInSecs Then Exit Sub

    While pollingFlag = True  ' flag that indicates that polling is still underway
        ' we wait until the polling is complete, VB6 timers are partially asynchronous and so this waits until the polling has complete
    Wend
    
    ' lock the output file writing process to prevent the array being updated
    nowBeingModifiedFlag = True ' this is a switch set during sendSomething that allows/disallows the polling timer logic to run

    maxLineLength = Val(UBound(outputFileArray))

    'create an archive array of the same size as the output array
    ReDim archiveArray(maxLineLength)
    
    'create an temporary output array of the same size as the old output array
    ReDim temporaryArray(maxLineLength)
    
    archiveLoc = 0
    tempLoc = 0
        
    ' loop through the output array
    For useloop = 1 To maxLineLength
        ' disassemble the timestamp
        outputArrayTimeStamp = Mid$(outputFileArray(useloop), 1, 23)
        stampTimeInSecs = fSecondsFromDateString(outputArrayTimeStamp)
        stampDiffInSecs = nowInSecs - stampTimeInSecs
   
        ' determine whether each lines is older than the date selected
        If stampDiffInSecs > archiveTimeInSecs Then
            '  write the line to an archive array
            archiveLoc = archiveLoc + 1
            archiveArray(archiveLoc) = outputFileArray(useloop)
            
        Else
            ' write a new output array without the older data
            tempLoc = tempLoc + 1
            temporaryArray(tempLoc) = outputFileArray(useloop)
            
        End If
    Next useloop
    
    ' if none to archive then exit
    If archiveLoc = 0 Then
        nowBeingModifiedFlag = False ' this is a switch set during sendSomething that allows/disallows the polling timer logic to run
        Exit Sub
    End If
    
    ' redim the archive array to the new size
    ReDim Preserve archiveArray(archiveLoc)
    
    ' determine a timestamp for the archive filename
    timestamp = fGetDateNoChars ' timestamp with no illegal chars for BinaryStream.SaveToFile to fall over ":"

    ' determine the archive location
    archiveFilePath = FCWArchiveFolder & "\" & "archive" & timestamp & ".txt"
    
    ' write the archive file using the archive array
    testArchiveExists = updateArchiveFile(archiveArray(), archiveFilePath, archiveLoc)
            
    ' check the archive file exists
    If testArchiveExists = False Then
        nowBeingModifiedFlag = False ' this is a switch set during sendSomething that allows/disallows the polling timer logic to run
        Exit Sub
    End If
    
    ' redim the output array to the new size
    ReDim Preserve temporaryArray(tempLoc)
    outputLineCount = tempLoc

    ' loop through the new output array and copy each line to the output array
    For useloop = 1 To outputLineCount
        ' write the new output array without the older data
        'a = temporaryArray(useloop)
        outputFileArray(useloop) = temporaryArray(useloop)
    Next useloop
    
    Call writeOutputFile(FCWSharedOutputFile, outputLineCount)
        
    ' re-read the file chosen as the output file
    ' re populate the array the same length as your output file
    ' update the listbox using the array
    Call readOutputFileWriteArrayWriteListbox(FCWSharedOutputFile)
    
    outputLineCount = fLineCount(FCWSharedOutputFile) ' this might not be required - for testing
        
    ' this next stuff always appears after a call to readOutputFileWriteArrayWriteListbox - DEAN
    ' I may need to tidy or incorporate this into the routine
    
    CTRL_1 = False ' ensuring that automatic click caused by the next few commands does not cause any URL to
                   ' automatically show in the browser (Ctrl+click)
    If Val(FCWLoadBottom) = 1 Then
        FireCallMain.lbxOutputTextArea.ListIndex = FireCallMain.lbxOutputTextArea.ListCount - 1
    Else
        FireCallMain.lbxOutputTextArea.ListIndex = 0
    End If
    
    outputDataChangedFlag = True
    
    ' populate the combined listbox
    If FCWSingleListBox = "1" Then Call populateCombinedBox

    ' turn the polling timer back on - this is a switch that allows the polling timer logic to run
    nowBeingModifiedFlag = False ' this is a switch set during sendSomething that allows/disallows the polling timer logic to run

    ' store the last advice timestamp to allow comparison
    FCWLastHouseKeep = fGetDateInUniversalFormat
    If fFExists(FCWSettingsFile) Then
        PutINISetting "Software\FireCallWin", "lastHouseKeep", FCWLastHouseKeep, FCWSettingsFile
    End If

   On Error GoTo 0
   Exit Sub

houseKeepingTimerLogic_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure houseKeepingTimerLogic of Module modCommon"

End Sub



' routine called at startup to create or run the two HouseKeeping timers
'---------------------------------------------------------------------------------------
' Procedure : startTheHouseKeepingTimers
' Author    : beededea
' Date      : 23/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub startTheHouseKeepingTimers()
    
    Dim HouseKeepingIntervalMillisecs As Long

    Const lngSecs As Long = 65 ' just used to avoid multiplying two integers
    Const lngThousand As Long = 1000

    ' start the HouseKeeping timer in code
   On Error GoTo startTheHouseKeepingTimers_Error

    If fInIDE Then
        ' VB6 timers cannot exceed 65 seconds (65535 ms)
'        lngSecs = 65
'        lngThousand = 1000
        ' when multiplying two integer values and assigning to a long in the IDE it causes a failure as the IDE is handling the two numbers as integers
        ' HouseKeepingIntervalMillisecs = 65 * 1000 '  < this fails
        HouseKeepingIntervalMillisecs = lngSecs * lngThousand ' works!
        FireCallMain.houseKeepingTimer.Interval = HouseKeepingIntervalMillisecs
        FireCallMain.houseKeepingTimer.Enabled = True
        debugLog "Starting startHouseKeepingTimer using VB6 timer, at interval of " & HouseKeepingIntervalMillisecs & "ms", False
    Else
        ' using a timer in code rather than a VB6 timer as VB6 timers cannot exceed 65 seconds (65535 ms)
        ' and if you want a longer timer you have to roll your own.
        ' in addition, unfortunately this code timer method does not work in the IDE
        
        ' stop any possible running timer first
        Call stopHouseKeepingTimer
        
        ' using a timer in code rather than a VB6 timer as VB6 timers cannot exceed 65 seconds (65535 ms)
        HouseKeepingIntervalMillisecs = 3600 * lngThousand ' every hour
        
        'MsgBox "FCWAdviceIntervalSecs " & FCWAdviceIntervalSecs
        
        ' final check to prevent starting this timer when working in the IDE, should never get this far
        If Not fInIDE Then
            ' Don't start the timer If it's already running.
            If houseKeepingTimerID = 0 Then
                ' this has a callback routine that it jumps to on each interval completion
                houseKeepingTimerID = SetTimer(0, 4, HouseKeepingIntervalMillisecs, AddressOf houseKeepingTimer_CodeTimer)
                debugLog "Starting startHouseKeepingTimer using API timer, ID = " & houseKeepingTimerID & " at interval of " & HouseKeepingIntervalMillisecs & "ms", False
            End If
        Else
            debugLog "Please note: Timers in code will not run in the IDE, defaulting to VB6 timers <65secs."
        End If
    End If

   On Error GoTo 0
   Exit Sub

startTheHouseKeepingTimers_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure startTheHouseKeepingTimers of Module modCommon"
End Sub



' The timer that stops the houseKeeping timer
'---------------------------------------------------------------------------------------
' Procedure : stopHouseKeepingTimer
' Author    : beededea
' Date      : 23/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub stopHouseKeepingTimer()
    ' Don't stop the timer If it isn't running.
   On Error GoTo stopHouseKeepingTimer_Error

    If houseKeepingTimerID Then
        KillTimer 0, houseKeepingTimerID
        houseKeepingTimerID = 0
    End If

   On Error GoTo 0
   Exit Sub

stopHouseKeepingTimer_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure stopHouseKeepingTimer of Module modCommon"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : debugLog
' Author    : beededea
' Date      : 23/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub debugLog(inputStr As String, Optional msgBoxOutOverride As Boolean)

    Dim FN As Integer
    Dim timestamp As String

   On Error GoTo debugLog_Error

    FN = FreeFile

    If msgBoxOut = True And Not msgBoxOutOverride = False Then MsgBox inputStr
    
    timestamp = fGetDateInUniversalFormat
    
    ' write the error to the log file
    If msgLogOut = True Then

        Open FCWSettingsDir & "\FCWDebugOutput.log" For Append As FN
        Print #FN, timestamp & " " & inputStr
        Close FN
        
    End If

   On Error GoTo 0
   Exit Sub

debugLog_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure debugLog of Module modCommon"
End Sub


' default positions prior to any resizing/shifting
'---------------------------------------------------------------------------------------
' Procedure : putImageInPlace
' Author    : beededea
' Date      : 23/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub putImageInPlace()
   On Error GoTo putImageInPlace_Error

    FireCallMain.picImagePrintOut.Left = 140
    FireCallMain.picImagePrintOut.Top = 585
    FireCallMain.picImagePrintOut.Width = 2160
    FireCallMain.picImagePrintOut.Height = 2475

'    FireCallMain.picPrintOutShadow.Left = 165
'    FireCallMain.picPrintOutShadow.Top = 620
'    FireCallMain.picPrintOutShadow.Width = 2160
'    FireCallMain.picPrintOutShadow.Height = 2475

   On Error GoTo 0
   Exit Sub

putImageInPlace_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure putImageInPlace of Module modCommon"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : setModernThemeColours
' Author    : beededea
' Date      : 23/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub setModernThemeColours()
         
    ' variables declared
    Dim SysClr As Long
        
    'initialise the dimensioned variables
   On Error GoTo setModernThemeColours_Error

    SysClr = 0
    
    'FireCallPrefs.mnuAuto.Caption = "Auto Theme Selection Cannot be Enabled"

    'MsgBox "Windows Alternate Theme detected"
    SysClr = GetSysColor(COLOR_BTNFACE)
    If SysClr = 13160660 Then
        Call setThemeShade(212, 208, 199)
        FCWSkinTheme = "dark"
    Else ' 15790320
        Call setThemeShade(240, 240, 240)
        FCWSkinTheme = "light"
    End If

   On Error GoTo 0
   Exit Sub

setModernThemeColours_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure setModernThemeColours of Module modCommon"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : stripOut
' Author    : beededea
' Date      : 23/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function stripOut(from As String, What As String) As String

    Dim i As Integer

   On Error GoTo stripOut_Error

    stripOut = from
    For i = 1 To Len(What)
        stripOut = Replace(stripOut, Mid$(What, i, 1), "")
    Next i

   On Error GoTo 0
   Exit Function

stripOut_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure stripOut of Module modCommon"

End Function

'Public Function fBorderSize(frm As Form) As RECT
'
'' this does not work at all
'
'    'returns the size of the borders applied by Windows around the outside/ inside of a Form when Vista+ and Aero type theme is applied
'    ' typically returns negative values in Windows 10
'
'    Dim FrmDims As RECT, BordersizeTemp As RECT
'    Dim lret&
'    Static BorderSizeFixed As RECT, BorderSizeSizable As RECT
'    Static InitFixed As Boolean, InitSizable As Boolean
'
'    Const DWMWA_EXTENDED_FRAME_BOUNDS = 9&
'
'    'return stored values if available for Frm.BorderStyle (Frm does not have to be visible and we avoid calling the API/ Error handler every time which should be quicker)
'    Select Case frm.BorderStyle
'        Case vbBSNone
'            fBorderSize = BordersizeTemp 'borders always zero
'            Exit Function
'        Case vbFixedSingle, vbFixedDouble, vbFixedToolWindow
'            If InitFixed Then       'return precalculated/ stored values
'                fBorderSize = BorderSizeFixed
'                Exit Function
'            End If
'        Case vbSizable, vbSizableToolWindow
'            If InitSizable Then     'return precalculated/ stored values
'                fBorderSize = BorderSizeSizable
'                Exit Function
'            End If
'    End Select
'
'    'following code only fires twice, once to get/ store Fixed form values, once to get/ store Sizable form values
'
'    On Error Resume Next    'API below is not supported in XP and may cause error when called so catch that to keep use under XP sweet
'    'to return the Aero Borders Frm must be Shown/ Visible, otherwize zero is returned for all borders
'    lret = DwmGetWindowAttribute(frm.hwnd, DWMWA_EXTENDED_FRAME_BOUNDS, FrmDims, LenB(FrmDims))
'    'returns lret = 0 if an Aero Theme is active, Aero Themes are optional in Vista and Win7, in Win 8.1 and 10 they are always active without option
'    If lret = 0 And err.Number = 0 Then
'        On Error GoTo 0
'        With BordersizeTemp
'            .Left = (frm.Left - FrmDims.Left * Screen.TwipsPerPixelX)
'            .Top = (frm.Top - FrmDims.Top * Screen.TwipsPerPixelY)
'            .Right = FrmDims.Right * Screen.TwipsPerPixelX - (frm.Left + frm.Width)
'            .Bottom = FrmDims.Bottom * Screen.TwipsPerPixelY - (frm.Top + frm.Height)
'        End With
'    Else
'        'all the borders are returned as zero, it's XP or Aero is switched off
'        On Error GoTo 0
'        InitFixed = True
'        InitSizable = True
'        Exit Function
'    End If
'
'    Select Case frm.BorderStyle
'        Case 1, 3, 4
'            BorderSizeFixed = BordersizeTemp
'            InitFixed = True
'        Case Else
'            BorderSizeSizable = BordersizeTemp
'            InitSizable = True
'    End Select
'
'End Function
' credit wqweto
'---------------------------------------------------------------------------------------
' Procedure : IsValidPath
' Author    : wqweto
' Date      : 23/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
    Public Function IsValidPath(sPath As String) As Boolean
   On Error GoTo IsValidPath_Error

    If sPath = "" Then IsValidPath = False: Exit Function ' this would mean 2 or more \\ together and is not valid
        IsValidPath = (sPath = SanitizePath(sPath))

   On Error GoTo 0
   Exit Function

IsValidPath_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure IsValidPath of Module modCommon"
    End Function
     
'---------------------------------------------------------------------------------------
' Procedure : SanitizePath
' Author    : wqweto
' Date      : 23/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
    Public Function SanitizePath(sPath As String, Optional InvalidChar As String = "?") As String
    Dim vSplit          As Variant
    Dim lIdx            As Long
    
   On Error GoTo SanitizePath_Error

    vSplit = Split(sPath, "\")
    For lIdx = IIf(vSplit(0) Like "?:", 1, 0) To UBound(vSplit)
        vSplit(lIdx) = SanitizeFileName(CStr(vSplit(lIdx)), InvalidChar)
    Next
    SanitizePath = Join(vSplit, "\")

   On Error GoTo 0
   Exit Function

SanitizePath_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure SanitizePath of Module modCommon"
        
    End Function
     
'---------------------------------------------------------------------------------------
' Procedure : SanitizeFileName
' Author    : wqweto
' Date      : 23/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
    Public Function SanitizeFileName(sFilename As String, Optional InvalidChar As String = "?") As String
   On Error GoTo SanitizeFileName_Error

        With CreateObject("VBScript.RegExp")
            .Global = True
            .Pattern = "[\x00-\x1F""<>\|:\*\?\\/]"
            SanitizeFileName = .Replace(sFilename, InvalidChar)
        End With

   On Error GoTo 0
   Exit Function

SanitizeFileName_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure SanitizeFileName of Module modCommon"
    End Function


' function to scale VB picture objects according to DPI settings
'Public Function ScaleStdPicture(ByVal thePic As StdPicture, ParentForm As Object, TwipsPerPixel As Long) As IPicture
'' note: TwipsPerPixel parameter value is relative to the desired DPI to scale to.
'
'    'Private Type PictDesc
'    '    Size As Long
'    '    Type As Long
'    '    hHandle As Long
'    '    lParam1 As Long      for bitmaps/WMF only
'    '                         WMF = extentX, BMP = Palette handle
'    '    lParam2 As Long      for WMF only: extentY
'
'    'End Type
'
'    Dim lpPictDesc(0 To 3) As Long  ' equivalent to a PictDesc structure
'    Dim aGUID(0 To 3) As Long       ' equivalent to GUID
'    Dim hImage As Long
'    Dim cx As Long, cy As Long
'    Const LR_COPYFROMRESOURCE As Long = &H4000
'    Const LR_COPYRETURNORG As Long = &H4
'
'    If thePic Is Nothing Then Exit Function
'
'    On Error Resume Next
'    cx = ParentForm.ScaleX(thePic.Width, vbHimetric, vbPixels)
'        cx = cx * (1440 / TwipsPerPixel) / 96
'    cy = ParentForm.ScaleY(thePic.Height, vbHimetric, vbPixels)
'        cy = cy * (1440 / TwipsPerPixel) / 96
'    If err Then ' something's wrong, passed invalid Object?
'        err.Clear
'    Else
'        Select Case thePic.Type
'            Case vbPicTypeBitmap
'                hImage = CopyImage(thePic.handle, 0&, cx, cy, LR_COPYRETURNORG)
'            Case vbPicTypeIcon
'                hImage = CopyImage(thePic.handle, 1&, cx, cy, LR_COPYFROMRESOURCE Or LR_COPYRETURNORG)
'                If hImage = 0& Then
'                    hImage = CopyImage(thePic.handle, 1&, cx, cy, LR_COPYRETURNORG)
'                End If
'            Case Else
'        End Select
'    End If
'    On Error GoTo 0
'
'    If hImage = 0& Or hImage = thePic.handle Then
'        Set ScaleStdPicture = thePic
'    Else
'        ' fill in PictDesc structure
'        lpPictDesc(0) = 16&
'        lpPictDesc(1) = thePic.Type
'        lpPictDesc(2) = hImage
'        ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
'        aGUID(0) = &H7BF80980
'        aGUID(1) = &H101ABF32
'        aGUID(2) = &HAA00BB8B
'        aGUID(3) = &HAB0C3000
'
'        ' create stdPicture
'        Call OleCreatePictureIndirect(lpPictDesc(0), aGUID(0), True, ScaleStdPicture)
'    End If
'
'End Function





'---------------------------------------------------------------------------------------
' Procedure : savestring
' Author    : beededea
' Date      : 05/07/2019
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub savestring(ByRef hKey As Long, ByRef strPath As String, ByRef strvalue As String, ByRef strData As String)

    Dim keyhand As Long
    Dim R As Long
   On Error GoTo savestring_Error

    R = RegCreateKey(hKey, strPath, keyhand)
    R = RegSetValueEx(keyhand, strvalue, 0, REG_SZ, ByVal strData, Len(strData))
    R = RegCloseKey(keyhand)

   On Error GoTo 0
   Exit Sub

savestring_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure savestring of Module Common"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : centreMainScreen
' Author    : beededea
' Date      : 23/10/2024
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Sub centreMainScreen()
   On Error GoTo centreMainScreen_Error

    FireCallMain.Top = screenHeightTwips / 2 - FireCallMain.Height / 2
    FireCallMain.Left = screenWidthTwips / 2 - FireCallMain.Width / 2

   On Error GoTo 0
   Exit Sub

centreMainScreen_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure centreMainScreen of Module modCommon"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : backupOutputFile
' Author    : beededea
' Date      : 30/05/2019
' Purpose   : Creates an incrementally named backup of the settings.ini
'---------------------------------------------------------------------------------------

Public Sub backupOutputFile(ByVal fileToBackupFullPath As String, backupCommand As String)

    Dim trgtBackupFilename As String
    Dim useloop As Integer
    Dim srchBackupFile As String
    Dim versionNumberAvailable As Integer
    Dim bkpfileFound As Boolean
    
    
    ' set the name of the bkp file
   
   ' On Error GoTo backupOutputFile_Error
      If debugflg = 1 Then Debug.Print "%" & "backupOutputFile"

        trgtBackupFilename = FCWBackupFolder & "\" & backupCommand & "-" & fGetFileNameFromPath(fileToBackupFullPath)
                
        'check for any version of an already existing backup file with the same suffix.
        For useloop = 1 To 32767
            srchBackupFile = trgtBackupFilename & "." & useloop
          
            If fFExists(srchBackupFile) Then
              ' found a file
              bkpfileFound = True
            Else
              ' no file found use this entry
              'GoTo l_exit_bkp_loop
              Exit For
            End If
        Next useloop
        
l_exit_bkp_loop:
        
        If bkpfileFound = True Then
            bkpfileFound = False
            versionNumberAvailable = useloop
            
            'if versionNumberAvailable >= 32767 then
                'versionNumberAvailable = 1
                'if fFExists(trgtBackupFilename) Then
                    'delete trgtBackupFilename
                'endif
            'endif
        Else
             versionNumberAvailable = 1
        End If
        
        trgtBackupFilename = trgtBackupFilename & "." & Trim$(Str$(versionNumberAvailable))
        If Not fFExists(trgtBackupFilename) Then
            ' copy the original settings file to a duplicate that we will keep as a safety backup

                If fFExists(fileToBackupFullPath) Then
                    If fDirExists(FCWBackupFolder) Then
                        FileCopy fileToBackupFullPath, trgtBackupFilename
                    End If
                End If

        End If
        
   On Error GoTo 0
   Exit Sub

backupOutputFile_Error:

    debugLog "Error " & err.Number & " (" & err.Description & ") in procedure backupOutputFile of Form FireCallMain"
        
End Sub
