Attribute VB_Name = "basWinAPI"
'**************************************************
'* basWinAPI version info:
'*  See basGlobals (Globals.bas).
'**************************************************

Private Const ModuleName = "basWinAPI"

Public Const PrsQryInfo = &H400
Public Const StillActive = &H103

Type POINTStruct
   X As Integer
   Y As Integer
End Type

Type MSGStruct
  hwnd As Integer
  iMsg As Integer
  wParam As Integer
  lParam As Long
  lTime As Long
  Point As POINTStruct
End Type

Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type

'********************************************************************
'* Added for window capture routines
'********************************************************************
Public Type PALETTEENTRY
  peRed As Byte
  peGreen As Byte
  peBlue As Byte
  peFlags As Byte
End Type

Public Type LOGPALETTE
  palVersion As Integer
  palNumEntries As Integer
  palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
End Type

Public Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Public Type PicBmp
  Size As Long
  Type As Long
  hBmp As Long
  hPal As Long
  Reserved As Long
End Type

Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal iCapabilitiy As Long) As Long
Public Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long

Public Declare Function CreateCompatibleBitmap Lib "gdi32" ( _
   ByVal hdc As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long) As Long

Public Declare Function GetSystemPaletteEntries Lib "gdi32" ( _
   ByVal hdc As Long, ByVal wStartIndex As Long, _
   ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long

Public Declare Function BitBlt Lib "gdi32" ( _
   ByVal hDCDest As Long, ByVal XDest As Long, _
   ByVal YDest As Long, ByVal nWidth As Long, _
   ByVal nHeight As Long, ByVal hDCSrc As Long, _
   ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Public Declare Function SelectPalette Lib "gdi32" ( _
   ByVal hdc As Long, ByVal hPalette As Long, _
   ByVal bForceBackground As Long) As Long

Public Declare Function OleCreatePictureIndirect _
   Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, _
   ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
   
Public Declare Function OpenProcess Lib "kernel32" _
               (ByVal dwDesiredAccess As Long, _
                ByVal bInheritHandle As Long, _
                ByVal dwProcessId As Long) As Long

Public Declare Function GetExitCodeProcess Lib "kernel32" _
               (ByVal hProcess As Long, _
                lpExitCode As Long) As Long

'********************************************************************
'********************************************************************

Public MsgInfo As MSGStruct

'***************************************************************
'* Since the following constants relate to API calls and are
'* given the following names in the API help documentation,
'* I will not give their names an "smc" prefix.
'***************************************************************
Public Const WM_RBUTTONUP = &H205
Public Const WM_MOUSEMOVE = &H200
Public Const WM_SETREDRAW = &HB
Public Const WM_USER = &H400
Public Const WM_FONTCHANGE = &H1D
Public Const WM_HSCROLL = &H114
Public Const WM_VSCROLL = &H115
Public Const SB_HORZ = 0
Public Const SB_VERT = 1
Public Const SB_CTL = 2
Public Const SB_BOTH = 3
Public Const LB_SETTABSTOPS = (WM_USER + 19)   'Used to set tabstops in list boxes.
Public Const LB_FINDSTRING = (WM_USER + 16)    'This is used for list box string search API
Public Const LB_FINDSTREXACT = (WM_USER + 35)  'This is used for list box string search API
Public Const CB_FINDSTRING = (WM_USER + 12)    'This is used for combo box string search API
Public Const CB_SHOWDROPDOWN = (WM_USER + 15)  'This is used to programatically drop speaker combo. box.
Public Const CB_FINDSTREXACT = (WM_USER + 24)  'This is used for combo box string search API
Public Const EM_GETSEL = WM_USER
Public Const EM_SETSEL = (WM_USER + 1)
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_NORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_MAXIMIZE = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const HWND_TOPMOST = -1
Public Const HWND_BROADCAST = &HFFFF
Public Const MK_RBUTTON = 2
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDNEXT = 2
Public Const HH_DISPLAY_TOC = 1
Public Const HKEY_CURRENT_USER = &H80000001
Public Const KEY_ALL_ACCESS = &H3F 'Combines the STANDARD_RIGHTS_REQUIRED, KEY_QUERY_VALUE, KEY_SET_VALUE, KEY_CREATE_SUB_KEY, KEY_ENUMERATE_SUB_KEYS, KEY_NOTIFY, and KEY_CREATE_LINK access rights.
Public Const REG_SZ As Long = 1

Declare Function GetVersion Lib "kernel32" () As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd%, ByVal wCmd%) As Long
Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Declare Function AddFontResource Lib "gdi32" (ByVal lpFileName As String) As Long
Declare Function GetDoubleClickTime Lib "user32" () As Long

Declare Function GetTempFileName Lib "kernel32" _
                 Alias "GetTempFileNameA" _
                 (ByVal lpszPath$, _
                  ByVal lpPrefixString$, _
                  ByVal wUnique%, _
                  ByVal lpTempFileName$) As Integer
  
Declare Function GetTempPath Lib "kernel32" _
                 Alias "GetTempPathA" _
                 (ByVal nBufferLength As Long, _
                  ByVal lpBuffer As String) As Long

Declare Function GetWindowText Lib "user32" _
                 Alias "GetWindowTextA" _
                 (ByVal hwnd%, _
                  ByVal lpSting$, _
                  ByVal nMaxCount%) As Integer

Declare Function GetWindowsDirectory Lib "kernel32" _
                 Alias "GetWindowsDirectoryA" _
                 (ByVal lpBuffer$, _
                  ByVal nSize As Integer) As Integer

Declare Function GetSystemDirectory Lib "kernel32" _
                 Alias "GetSystemDirectoryA" _
                 (ByVal lpBuffer$, _
                  ByVal nSize%) As Integer

Declare Function SendMessage Lib "user32" _
                 Alias "SendMessageA" _
                 (ByVal hwnd As Integer, _
                  ByVal wMsg As Integer, _
                  ByVal wParam As Integer, _
                  lParam As Any) As Long
                  
Declare Function GetScrollRange Lib "user32" _
                 (ByVal hwnd As Long, _
                  ByVal nBar As Long, _
                  lpMinPos As Long, _
                  lpMaxPos As Long) As Long

Declare Function SetScrollPos Lib "user32" _
                 (ByVal hwnd As Long, _
                  ByVal nBar As Long, _
                  ByVal nPos As Long, _
                  ByVal bRedraw As Long) As Long

Declare Function SetScrollInfo Lib "user32" _
                 (ByVal hwnd As Long, _
                  ByVal n As Long, _
                  lpcScrollInfo As SCROLLINFO, _
                  ByVal bool As Boolean) As Long

Declare Function HtmlHelp Lib "HHCtrl.ocx" _
                 Alias "HtmlHelpA" _
                 (ByVal hWndCaller As Long, _
                 ByVal pszFile As String, _
                 ByVal uCommand As Long, _
                 dwData As Any) As Long

Declare Function GetPrivateProfileString Lib "kernel32" _
                 Alias "GetPrivateProfileStringA" _
                 (ByVal lpSection$, _
                  ByVal lpEntry As Any, ByVal lpDefault$, _
                  ByVal lpReturnedString$, _
                  ByVal nSize%, _
                  ByVal lpFile$) As Integer

Declare Function WritePrivateProfileString Lib "kernel32" _
                 Alias "WritePrivateProfileStringA" _
                 (ByVal lpSection$, _
                  ByVal lpEntry As Any, _
                  ByVal lpValue As Any, _
                  ByVal lplFile$) As Integer

Declare Function GetPrivateProfileSection Lib "kernel32" _
                 Alias "GetPrivateProfileSectionA" _
                 (ByVal lpAppName As String, _
                  ByVal lpReturnedString As String, _
                  ByVal nSize As Long, _
                  ByVal lpFileName As String) As Long
                  
Declare Function RegCreateKey Lib "advapi32.dll" _
                 Alias "RegCreateKeyA" _
                 (ByVal hKey As Long, _
                 ByVal lpSubKey As String, _
                 phkResult As Long) As Long

Declare Function RegOpenKeyEx Lib "advapi32.dll" _
                 Alias "RegOpenKeyExA" _
                 (ByVal hKey As Long, _
                 ByVal lpSubKey As String, _
                 ByVal ulOptions As Long, _
                 ByVal samDesired As Long, _
                 phkResult As Long) As Long

Declare Function RegCloseKey Lib "advapi32.dll" _
                 (ByVal hKey As Long) As Long

Declare Function RegQueryValueExNULL Lib "advapi32.dll" _
                 Alias "RegQueryValueExA" _
                 (ByVal hKey As Long, _
                 ByVal lpValueName As String, _
                 ByVal lpReserved As Long, _
                 lpType As Long, _
                 ByVal lpData As Long, _
                 lpcbData As Long) As Long

Declare Function RegQueryValueExString Lib "advapi32.dll" _
                 Alias "RegQueryValueExA" _
                 (ByVal hKey As Long, _
                 ByVal lpValueName As String, _
                 ByVal lpReserved As Long, _
                 lpType As Long, _
                 ByVal lpData As String, _
                 lpcbData As Long) As Long

Declare Function RegSetValueEx Lib "advapi32.dll" _
                 Alias "RegSetValueExA" _
                 (ByVal hKey As Long, _
                 ByVal lpValueName As String, _
                 ByVal Reserved As Long, _
                 ByVal dwType As Long, _
                 lpData As Any, _
                 ByVal cbData As Long) As Long

Public Function AppActive(iStarthWnd%, sAppWinTitle As String) As Long
      
  '***********************************************
  '* This function will return the window handle
  '* of the application whose title bar contains
  '* sAppWinTitle in the first Len(sAppWinTitle)
  '* characters of its title text. If no windows
  '* contain sAppWinTitle then zero is returned.
  '***********************************************
      
  Dim hwnd As Long
  Dim iLen As Integer
  Dim sTitle As String
  
  On Error GoTo AppActiveErr
  AppActive = 0
  
  '***************************************************
  '* First start off with the window handle passed
  '* to this function and then get the window handle
  '* of the next window in the list.
  '***************************************************
  hwnd = GetWindow(iStarthWnd, GW_HWNDFIRST)
  hwnd = GetWindow(hwnd, GW_HWNDNEXT)
 
  '***************************************************
  '* Now cycle through the active windows and get
  '* each one's title. If any of the window titles
  '* match sAppWinTitle (or sAppWinTitle matches
  '* the first Len(sAppWinTitle) characters) then
  '* return true and get out.
  '***************************************************
  While hwnd <> 0
    iLen = GetWindowTextLength(hwnd)
    sTitle = Space$(iLen + 1)
    If GetWindowText(hwnd, sTitle, iLen + 1) > 0 Then
      If Left$(sTitle, Len(sAppWinTitle)) = sAppWinTitle Then
        AppActive = hwnd
        Exit Function
      End If
    End If
    hwnd = GetWindow(hwnd, GW_HWNDNEXT)
    DoEvents
  Wend

  Exit Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
AppActiveErr:
  MsgBox Err.Description, vbInformation, App.Title

End Function

Public Function CBSearch(CBCtl As ComboBox, sString As String) As Long

  On Error Resume Next
  CBSearch = SendMessage(CBCtl.hwnd, CB_FINDSTREXACT, -1, ByVal sString)

End Function

Public Function GetAllINISettings(ByVal INIPath$, ByVal INISection$) As Variant
  
  Dim iNullCount As Integer
  Dim iNull As Integer
  Dim iEqu As Integer
  Dim sBuff As String * 25000
  Dim strArray() As String
  Dim sDoubleNull As String * 2
  Dim i As Integer
  Dim iLast As Integer
  
  On Error Resume Next
  
  '* initialize buffer to spaces
  sBuff = String$(Len(sBuff), vbNullChar)
  sDoubleNull = String$(2, vbNullChar)
  
  '* Read settings into buffer as null-delimited string
  If (GetPrivateProfileSection(ByVal INISection, ByVal sBuff, _
      Len(sBuff) - 256, ByVal INIPath) = 0) Then
    GetAllINISettings = Null
    Exit Function
  End If

  '* Count Nulls
  iLast = 1
  iNull = 0
  iNullCount = 0
  
  Do
    iNull = InStr(iNull + 1, sBuff, vbNullChar)
    If (Mid$(sBuff, iLast, 1) <> ";") Then iNullCount = iNullCount + 1
    iLast = iNull + 1
  Loop Until (Mid$(sBuff, iLast, 1) = vbNullChar)
  
  '* Parse string into array
  ReDim strArray(0 To iNullCount - 1, 0 To 1)
  i = 0
  iLast = 1
  iNull = 0
  
  Do
    iNull = InStr(iNull + 1, sBuff, vbNullChar)
    If (Mid$(sBuff, iLast, 1) <> ";") Then
      iEqu = InStr(iLast, sBuff, "=")
      If (iEqu > 0) Then
        strArray(i, 0) = Mid$(sBuff, iLast, iEqu - iLast)
        strArray(i, 1) = Mid$(sBuff, iEqu + 1, iNull - iEqu - 1)
        'Debug.Print i; ":  "; strArray(i, 0); "="; strArray(i, 1)
        i = i + 1
      End If
    End If
    
    iLast = iNull + 1
  Loop Until (Mid$(sBuff, iLast, 1) = vbNullChar)
  
'  Do While True
'    iEqu = InStr(iNull + 1, sBuff, "=")
'    If (iEqu = 0) Then Exit Do
'
'    If (Mid$(sBuff, iNull + 1, 1) <> ";") Then
'      strArray(i, 0) = Mid$(sBuff, iNull + 1, (iEqu - iNull) - 1)
'      strArray(i, 1) = Mid$(sBuff, iEqu + 1, InStr(iEqu, sBuff, Chr$(0)) - iEqu - 1)
'      Debug.Print i; ":  "; strArray(i, 0); "="; strArray(i, 1)
'    End If
'    i = i + 1
'    iNull = InStr(iNull + 1, sBuff, Chr$(0))
'  Loop
  
  GetAllINISettings = strArray

End Function

Public Function GetINIEntry$(ByVal Section$, ByVal Entry$, Optional INIPath)

  '*********************************************************
  '* This routine will read an entry from the Speech
  '* Manager INI file. After reading the entry the result
  '* string is scanned for a null terminator. If found then
  '* the result string is trimmed so it contains only
  '* the characters up to the null terminator. VB doesn't
  '* see the null as a terminator and thus thinks the
  '* string may be longer.
  '*********************************************************
  
  Dim i As Integer
  Dim sBuffer As String
  
  On Error Resume Next
  
  sBuffer = Space$(128)                'Allocate  buffer for result string.
  
  Call GetPrivateProfileString _
    (Section, Entry, "", sBuffer, Len(sBuffer), _
    IIf(IsMissing(INIPath), gINIPath, CStr(INIPath)))

  '*******************************
  '* Look for null terminator and
  '* replace it with a space.
  '*******************************
  i = InStr(1, sBuffer, vbNullChar)
  If i > 0 Then Mid$(sBuffer, i, 1) = " "
  GetINIEntry = Trim$(sBuffer)
  
End Function

Public Function GetTmpFileName() As String
  
  '*************************************************
  '* This function will create and return the full
  '* path and file name of a temporary file. Make
  '* sure to delete the temporary file when your
  '* fininshed with it.
  '*************************************************
  
  Dim iRet As Integer
  Dim sBuff As String * 254
   
  On Error Resume Next
  iRet = GetTempFileName(0, "", 0, sBuff)
  If iRet > 0 Then
    GetTmpFileName = Left(sBuff, InStr(sBuff, vbNullChar) - 1)
  Else
    GetTmpFileName = ""
  End If
  
End Function

Public Function GetWindowsDir() As String

  '**********************************
  '* This function returns the full
  '* path of the Windows directory.
  '**********************************

  Dim iRetSz As Integer
  Dim sBuff As String * 254

  On Error GoTo GetWindowsDirErr
  iRetSz = GetWindowsDirectory(sBuff, 254)
  If iRetSz > 0 Then
    GetWindowsDir = Left(sBuff, iRetSz)
    Exit Function
  End If

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GetWindowsDirErr:
  GetWindowsDir = ""
  
End Function

Public Function GetWinVersion() As Single
  
  Dim lVer As Long
  
  On Error Resume Next
  lVer = GetVersion()
  GetWinVersion = (lVer And &HFF&) + (((lVer And &HFF00&) \ 256) * 0.01)

End Function

Public Sub RegisterFont(sFontFile As String)

  '*****************************************
  '* This function assumes that the font
  '* is already in the font directory for
  '* Win95 or the Windows system directory
  '* for windows 3.x.
  '*****************************************
  
  Dim i As Integer
  
  On Error Resume Next
  
  If AddFontResource(sFontFile) Then _
    i = SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, 0)
  
End Sub

Public Function WriteINIEntry%(ByVal sSection$, ByVal sEntry$, ByVal sValue$, Optional vINIPath)

  Call WritePrivateProfileString(sSection, sEntry, _
    sValue, IIf(IsMissing(vINIPath), gINIPath, CStr(vINIPath)))

End Function
