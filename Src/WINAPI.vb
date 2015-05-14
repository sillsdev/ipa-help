Option Strict Off
Option Explicit On
Module basWinAPI
	'**************************************************
	'* basWinAPI version info:
	'*  See basGlobals (Globals.bas).
	'**************************************************
	
	Private Const ModuleName As String = "basWinAPI"
	
	Public Const PrsQryInfo As Integer = &H400
	Public Const StillActive As Integer = &H103
	
	Structure POINTStruct
		Dim X As Short
		Dim Y As Short
	End Structure
	
	Structure MSGStruct
		Dim hwnd As Short
		Dim iMsg As Short
		Dim wParam As Short
		Dim lParam As Integer
		Dim lTime As Integer
		Dim Point As POINTStruct
	End Structure
	
	Structure SCROLLINFO
		Dim cbSize As Integer
		Dim fMask As Integer
		Dim nMin As Integer
		Dim nMax As Integer
		Dim nPage As Integer
		Dim nPos As Integer
		Dim nTrackPos As Integer
	End Structure
	
	'********************************************************************
	'* Added for window capture routines
	'********************************************************************
	Public Structure PALETTEENTRY
		Dim peRed As Byte
		Dim peGreen As Byte
		Dim peBlue As Byte
		Dim peFlags As Byte
	End Structure
	
	Public Structure LOGPALETTE
		Dim palVersion As Short
		Dim palNumEntries As Short
		<VBFixedArray(255)> Dim palPalEntry() As PALETTEENTRY ' Enough for 256 colors.
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim palPalEntry(255)
		End Sub
	End Structure
	
	Public Structure GUID
		Dim Data1 As Integer
		Dim Data2 As Short
		Dim Data3 As Short
		<VBFixedArray(7)> Dim Data4() As Byte
		
		'UPGRADE_TODO: "Initialize" must be called to initialize instances of this structure. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B4BFF9E0-8631-45CF-910E-62AB3970F27B"'
		Public Sub Initialize()
			ReDim Data4(7)
		End Sub
	End Structure
	
	Public Structure PicBmp
		Dim Size As Integer
		Dim Type As Integer
		Dim hBmp As Integer
		Dim hPal As Integer
		Dim Reserved As Integer
	End Structure
	
	Public Structure RECT
		'UPGRADE_NOTE: Left was upgraded to Left_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Left_Renamed As Integer
		Dim Top As Integer
		'UPGRADE_NOTE: Right was upgraded to Right_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Right_Renamed As Integer
		Dim Bottom As Integer
	End Structure
	
	Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Integer) As Integer
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'kg-Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal nPlanes As Integer, ByVal nBitCount As Integer, ByRef lpBits As Any) As Integer
    Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal nPlanes As Integer, ByVal nBitCount As Integer, ByRef lpBits As Integer) As Integer
    Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Integer, ByVal crColor As Integer) As Integer
	Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Integer) As Integer
	Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Integer) As Integer
	Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Integer, ByVal iCapabilitiy As Integer) As Integer
	'UPGRADE_WARNING: Structure LOGPALETTE may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function CreatePalette Lib "gdi32" (ByRef lpLogPalette As LOGPALETTE) As Integer
	Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Integer, ByVal hObject As Integer) As Integer
	Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Integer) As Integer
	Public Declare Function GetForegroundWindow Lib "user32" () As Integer
	Public Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Integer) As Integer
	Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Integer) As Integer
	Public Declare Function GetDC Lib "user32" (ByVal hwnd As Integer) As Integer
	'UPGRADE_WARNING: Structure RECT may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Integer, ByRef lpRect As RECT) As Integer
	Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Integer, ByVal hdc As Integer) As Integer
	Public Declare Function GetDesktopWindow Lib "user32" () As Integer
	
	Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
	
	'UPGRADE_WARNING: Structure PALETTEENTRY may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Integer, ByVal wStartIndex As Integer, ByVal wNumEntries As Integer, ByRef lpPaletteEntries As PALETTEENTRY) As Integer
	
	Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Integer, ByVal XDest As Integer, ByVal YDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hDCSrc As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Integer) As Integer
	
	Public Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Integer, ByVal hPalette As Integer, ByVal bForceBackground As Integer) As Integer
	
	'UPGRADE_WARNING: Structure IPicture may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure GUID may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	'UPGRADE_WARNING: Structure PicBmp may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (ByRef PicDesc As PicBmp, ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Integer, ByRef IPic As System.Drawing.Image) As Integer
	
	Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessId As Integer) As Integer
	
	Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Integer, ByRef lpExitCode As Integer) As Integer
	
	'********************************************************************
	'********************************************************************
	
	Public MsgInfo As MSGStruct
	
	'***************************************************************
	'* Since the following constants relate to API calls and are
	'* given the following names in the API help documentation,
	'* I will not give their names an "smc" prefix.
	'***************************************************************
	Public Const WM_RBUTTONUP As Integer = &H205
	Public Const WM_MOUSEMOVE As Integer = &H200
	Public Const WM_SETREDRAW As Integer = &HB
	Public Const WM_USER As Integer = &H400
	Public Const WM_FONTCHANGE As Integer = &H1D
	Public Const WM_HSCROLL As Integer = &H114
	Public Const WM_VSCROLL As Integer = &H115
	Public Const SB_HORZ As Short = 0
	Public Const SB_VERT As Short = 1
	Public Const SB_CTL As Short = 2
	Public Const SB_BOTH As Short = 3
	Public Const LB_SETTABSTOPS As Decimal = (WM_USER + 19) 'Used to set tabstops in list boxes.
	Public Const LB_FINDSTRING As Decimal = (WM_USER + 16) 'This is used for list box string search API
	Public Const LB_FINDSTREXACT As Decimal = (WM_USER + 35) 'This is used for list box string search API
	Public Const CB_FINDSTRING As Decimal = (WM_USER + 12) 'This is used for combo box string search API
	Public Const CB_SHOWDROPDOWN As Decimal = (WM_USER + 15) 'This is used to programatically drop speaker combo. box.
	Public Const CB_FINDSTREXACT As Decimal = (WM_USER + 24) 'This is used for combo box string search API
	Public Const EM_GETSEL As Integer = WM_USER
	Public Const EM_SETSEL As Decimal = (WM_USER + 1)
	Public Const SWP_NOMOVE As Short = 2
	Public Const SWP_NOSIZE As Short = 1
	Public Const SW_HIDE As Short = 0
	Public Const SW_SHOWNORMAL As Short = 1
	Public Const SW_NORMAL As Short = 1
	Public Const SW_SHOWMINIMIZED As Short = 2
	Public Const SW_SHOWMAXIMIZED As Short = 3
	Public Const SW_MAXIMIZE As Short = 3
	Public Const SW_SHOWNOACTIVATE As Short = 4
	Public Const SW_SHOW As Short = 5
	Public Const SW_MINIMIZE As Short = 6
	Public Const SW_SHOWMINNOACTIVE As Short = 7
	Public Const SW_SHOWNA As Short = 8
	Public Const SW_RESTORE As Short = 9
	Public Const HWND_TOPMOST As Short = -1
    Public Const HWND_BROADCAST As Integer = &HFFFF
	Public Const MK_RBUTTON As Short = 2
	Public Const GW_HWNDFIRST As Short = 0
	Public Const GW_HWNDNEXT As Short = 2
	Public Const HH_DISPLAY_TOC As Short = 1
	Public Const HKEY_CURRENT_USER As Integer = &H80000001
	Public Const KEY_ALL_ACCESS As Integer = &H3F 'Combines the STANDARD_RIGHTS_REQUIRED, KEY_QUERY_VALUE, KEY_SET_VALUE, KEY_CREATE_SUB_KEY, KEY_ENUMERATE_SUB_KEYS, KEY_NOTIFY, and KEY_CREATE_LINK access rights.
	Public Const REG_SZ As Integer = 1
	
	Declare Function GetVersion Lib "kernel32" () As Integer
	Declare Function GetWindow Lib "user32" (ByVal hwnd As Short, ByVal wCmd As Short) As Integer
	Declare Function GetWindowTextLength Lib "user32"  Alias "GetWindowTextLengthA"(ByVal hwnd As Integer) As Integer
	
	Declare Function AddFontResource Lib "gdi32" (ByVal lpFileName As String) As Integer
	Declare Function GetDoubleClickTime Lib "user32" () As Integer
	
	Declare Function GetTempFileName Lib "kernel32"  Alias "GetTempFileNameA"(ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Short, ByVal lpTempFileName As String) As Short
	
	Declare Function GetTempPath Lib "kernel32"  Alias "GetTempPathA"(ByVal nBufferLength As Integer, ByVal lpBuffer As String) As Integer
	
	Declare Function GetWindowText Lib "user32"  Alias "GetWindowTextA"(ByVal hwnd As Short, ByVal lpSting As String, ByVal nMaxCount As Short) As Short
	
	Declare Function GetWindowsDirectory Lib "kernel32"  Alias "GetWindowsDirectoryA"(ByVal lpBuffer As String, ByVal nSize As Short) As Short
	
	Declare Function GetSystemDirectory Lib "kernel32"  Alias "GetSystemDirectoryA"(ByVal lpBuffer As String, ByVal nSize As Short) As Short
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'kg-Declare Function SendMessage Lib "user32"  Alias "SendMessageA"(ByVal hwnd As Short, ByVal wMsg As Short, ByVal wParam As Short, ByRef lParam As Any) As Integer
    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Integer, ByVal wMsg As Short, ByVal wParam As Short, ByRef lParam As Long) As Integer

	Declare Function GetScrollRange Lib "user32" (ByVal hwnd As Integer, ByVal nBar As Integer, ByRef lpMinPos As Integer, ByRef lpMaxPos As Integer) As Integer
	
	Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Integer, ByVal nBar As Integer, ByVal nPos As Integer, ByVal bRedraw As Integer) As Integer
	
	'UPGRADE_WARNING: Structure SCROLLINFO may require marshalling attributes to be passed as an argument in this Declare statement. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="C429C3A5-5D47-4CD9-8F51-74A1616405DC"'
	Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Integer, ByVal n As Integer, ByRef lpcScrollInfo As SCROLLINFO, ByVal bool As Boolean) As Integer
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'kg-Declare Function HtmlHelp Lib "HHCtrl.ocx"  Alias "HtmlHelpA"(ByVal hWndCaller As Integer, ByVal pszFile As String, ByVal uCommand As Integer, ByRef dwData As Any) As Integer
    Declare Function HtmlHelp Lib "HHCtrl.ocx" Alias "HtmlHelpA" (ByVal hWndCaller As Integer, ByVal pszFile As String, ByVal uCommand As Integer, ByRef dwData As Integer) As Integer

	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'kg-Declare Function GetPrivateProfileString Lib "kernel32"  Alias "GetPrivateProfileStringA"(ByVal lpSection As String, ByVal lpEntry As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Short, ByVal lpFile As String) As Short
    Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSection As String, ByVal lpEntry As Integer, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Short, ByVal lpFile As String) As Short

	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'kg-Declare Function WritePrivateProfileString Lib "kernel32"  Alias "WritePrivateProfileStringA"(ByVal lpSection As String, ByVal lpEntry As Any, ByVal lpValue As Any, ByVal lplFile As String) As Short
    Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpSection As String, ByVal lpEntry As Integer, ByVal lpValue As Integer, ByVal lplFile As String) As Short

	Declare Function GetPrivateProfileSection Lib "kernel32"  Alias "GetPrivateProfileSectionA"(ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
	
	Declare Function RegCreateKey Lib "advapi32.dll"  Alias "RegCreateKeyA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByRef phkResult As Integer) As Integer
	
	Declare Function RegOpenKeyEx Lib "advapi32.dll"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
	
	Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
	
	Declare Function RegQueryValueExNULL Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As Integer, ByRef lpcbData As Integer) As Integer
	
	Declare Function RegQueryValueExString Lib "advapi32.dll"  Alias "RegQueryValueExA"(ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As String, ByRef lpcbData As Integer) As Integer
	
	'UPGRADE_ISSUE: Declaring a parameter 'As Any' is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="FAE78A8D-8978-4FD4-8208-5B7324A8F795"'
    'kg-Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal Reserved As Integer, ByVal dwType As Integer, ByRef lpData As Any, ByVal cbData As Integer) As Integer
    Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal Reserved As Integer, ByVal dwType As Integer, ByRef lpData As Integer, ByVal cbData As Integer) As Integer
	
	Public Function AppActive(ByRef iStarthWnd As Short, ByRef sAppWinTitle As String) As Integer
		
		'***********************************************
		'* This function will return the window handle
		'* of the application whose title bar contains
		'* sAppWinTitle in the first Len(sAppWinTitle)
		'* characters of its title text. If no windows
		'* contain sAppWinTitle then zero is returned.
		'***********************************************
		
		Dim hwnd As Integer
		Dim iLen As Short
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
			sTitle = Space(iLen + 1)
			If GetWindowText(hwnd, sTitle, iLen + 1) > 0 Then
				If Left(sTitle, Len(sAppWinTitle)) = sAppWinTitle Then
					AppActive = hwnd
					Exit Function
				End If
			End If
			hwnd = GetWindow(hwnd, GW_HWNDNEXT)
			System.Windows.Forms.Application.DoEvents()
		End While
		
		Exit Function
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
AppActiveErr: 
		MsgBox(Err.Description, MsgBoxStyle.Information, My.Application.Info.Title)
		
	End Function
	
	Public Function CBSearch(ByRef CBCtl As System.Windows.Forms.ComboBox, ByRef sString As String) As Integer
		
		On Error Resume Next
		CBSearch = SendMessage(CBCtl.Handle.ToInt32, CB_FINDSTREXACT, -1, sString)
		
	End Function
	
    Public Function GetAllINISettings(ByVal INIPath As String, ByVal INISection As String) As Object

        Dim iNullCount As Short
        Dim iNull As Short
        Dim iEqu As Short
        Dim sBuff As New VB6.FixedLengthString(25000)
        Dim strArray(,) As String
        Dim sDoubleNull As New VB6.FixedLengthString(2)
        Dim i As Short
        Dim iLast As Short

        On Error Resume Next

        '* initialize buffer to spaces
        sBuff.Value = New String(vbNullChar, Len(sBuff.Value))
        sDoubleNull.Value = New String(vbNullChar, 2)

        '* Read settings into buffer as null-delimited string
        If (GetPrivateProfileSection(INISection, sBuff.Value, Len(sBuff.Value) - 256, INIPath) = 0) Then
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            'UPGRADE_WARNING: Couldn't resolve default property of object GetAllINISettings. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            GetAllINISettings = System.DBNull.Value
            Exit Function
        End If

        '* Count Nulls
        iLast = 1
        iNull = 0
        iNullCount = 0

        Do
            iNull = InStr(iNull + 1, sBuff.Value, vbNullChar)
            If (Mid(sBuff.Value, iLast, 1) <> ";") Then iNullCount = iNullCount + 1
            iLast = iNull + 1
        Loop Until (Mid(sBuff.Value, iLast, 1) = vbNullChar)

        '* Parse string into array
        ReDim strArray(iNullCount - 1, 1)
        i = 0
        iLast = 1
        iNull = 0

        Do
            iNull = InStr(iNull + 1, sBuff.Value, vbNullChar)
            If (Mid(sBuff.Value, iLast, 1) <> ";") Then
                iEqu = InStr(iLast, sBuff.Value, "=")
                If (iEqu > 0) Then
                    strArray(i, 0) = Mid(sBuff.Value, iLast, iEqu - iLast)
                    strArray(i, 1) = Mid(sBuff.Value, iEqu + 1, iNull - iEqu - 1)
                    'Debug.Print i; ":  "; strArray(i, 0); "="; strArray(i, 1)
                    i = i + 1
                End If
            End If

            iLast = iNull + 1
        Loop Until (Mid(sBuff.Value, iLast, 1) = vbNullChar)

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

        'UPGRADE_WARNING: Couldn't resolve default property of object GetAllINISettings. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        GetAllINISettings = VB6.CopyArray(strArray)

    End Function
	
	Public Function GetINIEntry(ByVal Section As String, ByVal Entry As String, Optional ByRef INIPath As Object = Nothing) As String
		
		'*********************************************************
		'* This routine will read an entry from the Speech
		'* Manager INI file. After reading the entry the result
		'* string is scanned for a null terminator. If found then
		'* the result string is trimmed so it contains only
		'* the characters up to the null terminator. VB doesn't
		'* see the null as a terminator and thus thinks the
		'* string may be longer.
		'*********************************************************
		
		Dim i As Short
		Dim sBuffer As String
		
		On Error Resume Next
		
		sBuffer = Space(128) 'Allocate  buffer for result string.
		
		'UPGRADE_WARNING: Couldn't resolve default property of object INIPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		Call GetPrivateProfileString(Section, Entry, "", sBuffer, Len(sBuffer), IIf(IsNothing(INIPath), gINIPath, CStr(INIPath)))
		
		'*******************************
		'* Look for null terminator and
		'* replace it with a space.
		'*******************************
		i = InStr(1, sBuffer, vbNullChar)
		If i > 0 Then Mid(sBuffer, i, 1) = " "
		GetINIEntry = Trim(sBuffer)
		
	End Function
	
	Public Function GetTmpFileName() As String
		
		'*************************************************
		'* This function will create and return the full
		'* path and file name of a temporary file. Make
		'* sure to delete the temporary file when your
		'* fininshed with it.
		'*************************************************
		
		Dim iRet As Short
		Dim sBuff As New VB6.FixedLengthString(254)
		
		On Error Resume Next
		iRet = GetTempFileName(CStr(0), "", 0, sBuff.Value)
		If iRet > 0 Then
			GetTmpFileName = Left(sBuff.Value, InStr(sBuff.Value, vbNullChar) - 1)
		Else
			GetTmpFileName = ""
		End If
		
	End Function
	
	Public Function GetWindowsDir() As String
		
		'**********************************
		'* This function returns the full
		'* path of the Windows directory.
		'**********************************
		
		Dim iRetSz As Short
		Dim sBuff As New VB6.FixedLengthString(254)
		
		On Error GoTo GetWindowsDirErr
		iRetSz = GetWindowsDirectory(sBuff.Value, 254)
		If iRetSz > 0 Then
			GetWindowsDir = Left(sBuff.Value, iRetSz)
			Exit Function
		End If
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GetWindowsDirErr: 
		GetWindowsDir = ""
		
	End Function
	
	Public Function GetWinVersion() As Single
		
		Dim lVer As Integer
		
		On Error Resume Next
		lVer = GetVersion()
		GetWinVersion = CShort(lVer And &HFF) + (((lVer And &HFF00) \ 256) * 0.01)
		
	End Function
	
	Public Sub RegisterFont(ByRef sFontFile As String)
		
		'*****************************************
		'* This function assumes that the font
		'* is already in the font directory for
		'* Win95 or the Windows system directory
		'* for windows 3.x.
		'*****************************************
		
		Dim i As Short
		
		On Error Resume Next
		
		If AddFontResource(sFontFile) Then i = SendMessage(HWND_BROADCAST, WM_FONTCHANGE, 0, 0)
		
	End Sub
	
	Public Function WriteINIEntry(ByVal sSection As String, ByVal sEntry As String, ByVal sValue As String, Optional ByRef vINIPath As Object = Nothing) As Short
        Call WritePrivateProfileString(sSection, sEntry, sValue, IIf(IsNothing(vINIPath), gINIPath, CStr(vINIPath)))
    End Function
End Module