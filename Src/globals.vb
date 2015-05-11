Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Module basGlobals
	'**************************************************************************************************
	'* IPA-HELP                                                                                       *
	'* Version 1.62 16-bit                                                                            *
	'* 5/18/99                                                                                        *
	'* Written by David Olson and Corey Wenger                                                        *
	'**************************************************************************************************
	'*                                                                                                *
	'* REVISION INFORMATION                                                                           *
	'*                                                                                                *
	'*                                                                                                *
	'** Version 1.6 1/27/99 ***************************************************************************
	'*                                                                                                *
	'*  Test                                                                                          *
	'*  1/26/99   Fix indexing problem between gTestGrp and gTestGrpCorrect.                          *
	'*  1/27/99   Mark test item incorrect if not clicked on in time allowed.                         *
	'*                                                                                                *
	'*   2/1/99   frmDispAmerOth2: New                                                                *
	'*   2/2/99   frmDispAmerDia2: New                                                                *
	'*   2/5/99   frmDispAmerVow2: New                                                                *
	'*   2/8/99   frmDispAmerCon: symbols corrected                                                   *
	'*   2/8/99   frmDispAmerVow: symbols corrected                                                   *
	'*  2/16/99   frmDispAmerCon2: New                                                                *
	'*  2/16/99   frmDiagPtArtn: diagram improved (slightly)                                          *
	'*                                                                                                *
	'** Version 1.61 3/31/99 **************************************************************************
	'*  2/25/99   Add new sounds from Dallas                                                          *
	'*  3/31/99   frmListFile: New                                                                    *
	'*            txsil.wl: New                                                                       *
	'*            ipahelp.ini: 'WordLists' section added                                              *
	'*            mdiMain: load wordlist identification info from ipahelp.ini                         *
	'*                     ShowTBWordList sub added                                                   *
	'*            frmMenu: Word List section added                                                    *
	'*                                                                                                *
	'** Version 1.62 5/18/99 **************************************************************************
	'*  4/22/99   Toolbar overhaul:                                                                   *
	'*            3 new buttons added for interface with SA.                                          *
	'*            Toolbar configurations are stored in gTBLayout collection.                          *
	'*            ShowTB is passed the configuration name and arranges the toolbar accordingly.       *
	'*            SyncBttnsWithSelection is called whenever an update to the toolbar is needed.       *
	'*            Configurations for several of the forms were changed.                               *
	'*            mdiMain: ShowTB sub added (general sub for loading all toolbars)                    *
	'*                     All other 'ShowTB_____' subs deleted                                       *
	'*            globals.bas: SyncBttnsWithSelection sub added (general sub for updating toolbar)    *
	'*                         UpdateMCIButtons sub deleted                                           *
	'*            frmTestSetup: gTestSetupActive added to show when the test setup form is active.    *
	'*              This is necessary so that the 'Test' button stays down while that form is active. *
	'*  4/27/99   frmMenu: ShowTB call added                                                          *
	'*            Invalidate recording on form activate.                                              *
	'*            Only enable Test menu in frmMenu, frmCon, frmVow.                                   *
	'*  4/28/99   Toolbar MCI button enable/disable status altered to avoid file open/close problems. *
	'*            [display] section changed to [General Info] in word list file.                      *
	'*            SoundPath key added in [General Info] section.                                      *
	'*  4/29/99   Colors for disabled WordList items changed for better discernability.               *
	'*            frmTestSetup: WordList tab added                                                    *
	'*  4/30/99   frmWordList: SoundDelay = 1.5 seconds                                               *
	'*            frmWordList: Repeat limits = 1 - 9 times                                            *
	'*            mdiHelpCharts: fixed shut-down problem.                                             *
	'*  5/12/99   frmWordList Test: Panel on top of right side of form                                *
	'*  5/18/99   mdiHelpCharts: fixed StopTest bug                                                   *
	'*  5/21/99   Interface with SA (Slow Replay, Pitch, Auto Align)                                  *
	'*  9/16/02   Made changes in several places to allow for no words lists (DDO).                   *
	'*                                                                                                *
	'*** IN PROGRESS: *********************************************************************************
	'*                                                                                                *
	'*** TO DO: ***************************************************************************************
	'*                                                                                                *
	'**************************************************************************************************
	
	
	Public Const INIFile As String = "IPAHelp.ini"
	Public Const cSettingsSect As String = "Settings"
	Public Const cPathsSect As String = "Paths"
	Public Const cSRSpeedEntry As String = "SlowedReplaySpeed"
	Public Const cSAINIEntry As String = "SA"
	Public Const cPlayInitDelayEntry As String = "InitialPlaybackDelay"
	Public Const cPlayRepeatDelayEntry As String = "DelayBetweenRepeatedPlaybacks"
	Public Const cRepeatCountEntry As String = "RepeatCount"
	Public Const SystemType As String = "32-bit"
	Public Const cMCIBusy As String = "Busy"
	Public Const cFontNameEntry As String = "IPAFont"
	Public Const cFontSizeEntry As String = "IPASize"
	Public Const cFontBoldEntry As String = "IPABold"
	Public Const cFontItalicEntry As String = "IPAItalic"
	Public Const cSoundsEntry As String = "Sounds"
	Public Const cLeftEntry As String = "Left"
	Public Const cTopEntry As String = "Top"
	Public Const cWidthEntry As String = "Width"
	Public Const cHeightEntry As String = "Height"
	Public Const cSoundPathEntry As String = "SoundPath"
	Public Const cWLColsEntry As String = "ColWidths"
	Public Const cWLEditModeColsEntry As String = "EditModeColWidths"
	Public Const cNewUserAdviceEntry As String = "NewUserAdvice"
	'Public Const cChartFontName = "ASAP SILManuscript"
	'"ASAP SILDoulos"
	
	'* Toolbar rearranged by CLW 4/22/99
	Public Const c1stBttnLeft As Short = 9 'Left position of first toolbar button in toolbar.
	Public Const cTBHeight As Short = 435 'Height of toolbar in twips.
	Public Const cTBBttnWidth As Short = 42 'Button width in pixels CLW 4/16/99
	Public Const cTBGap As Short = 9 'Gap between toolbar buttons.
	Public Const cInvisPos As Short = -50 'Set left to this to move button off screen.
	Public Const cTBUpRow As Short = 0 'Row in the picture clip control of up button bitmaps.
	Public Const cTBDownRow As Short = 1 'Row in the picture clip control of down button bitmaps.
	Public Const cTBDisRow As Short = 2 'Row in the picture clip control of down button bitmaps.
	Public Const cTBInvis As Short = 3 'Makes button invisible
	
	Public Const cVocBttn As Short = 0 'Indices for each button.
	Public Const cInterVocBttn As Short = 1
	Public Const cSRBttn As Short = 2
	Public Const cRecBttn As Short = 3
	Public Const cMCIStopBttn As Short = 4
	Public Const cPlayBttn As Short = 5
	Public Const cCompBttn As Short = 6
	Public Const cPitchBttn As Short = 7
	Public Const cTestBttn As Short = 8
	Public Const cTestStopBttn As Short = 9
	Public Const cExitBttn As Short = 10
	Public Const cAABttn As Short = 11
	
	Public Const cUser As Short = 0
	Public Const cRandom As Short = 1
	Public Const cKeyCtrlTab As Short = 17 '* Keycode value often used for Ctrl + Tab key. CLW 5/12/99
	Public Const cOurMsg As Integer = &H7FF0 '* Message number for waking SA up via SendMessage. CLW 5/20/99
	Public Const cTerminateSA As Integer = &HE003 '* Message number for terminating SA. CLW 5/21/99
	Public Const cListModeSA As Integer = &HE001 '* List File mode parameter for SendMessage. CLW 5/20/99
	
	Public Const cTestChart As Short = 0 'Index of layout radio button (test setup form).
	Public Const cTestBrief As Short = 1
	Public Const cThisTest As Short = 0 'Index for gTestGrpCorrect
	Public Const cLastTest As Short = 1
	
	Public gAppPathWriteAccess As Boolean 'True if user has write access to app folder
	Public gINIPath As String
	Public gTBLayout As New Collection 'Stores toolbar layout info for various views
	Public gTBLayoutName As String 'Keeps track of current toolbar layout
	Public gBttnSpace As Object 'Space between each toolbar button and its left neighbor
	Public gTestSetupActive As Boolean 'Is the test setup form active? CLW 4/22/99
	Public gTestForm As String 'form that is active for this test. CLW 5/5/99
	Public gTestTag As String 'Tag of form (if relevant) active for this test. CLW 5/5/99
	Public gTestCatChoice As Short 'Whether or not category in Word List is chosen by user or program 5/11/99
	
	Public gMsg As String 'Message strings for messages boxes.
	Public gMsg1 As String
	Public gMsg2 As String
	Public gPixelX As Short 'Mouse position or position on screen.
	Public gPixelY As Short
	Public gToolTip(20) As String 'ToolTip text.
	Public gStatusLine(20) As String 'Status Bar text.
	Public gStatLine As System.Windows.Forms.StatusStrip
    Public WithEvents gMMCtrl As MCI.MMControl 'Shortcut for multi-media control.
	Public gMMStatus As Short 'Status for multi-media control.
	Public gSRSpeed As Short 'Slowed replay speed from INI file.
	Public gWavPath As String 'Path for sound files.
	Public gMstrWavName As String 'Name of template wave file for compare feature.
	Public gTmpWavName As String 'Name of temp wave file for compare.
	Public gTmpWavPath As String 'Path to temp file.
    Public gWordListID(,) As Object 'Wordlist description and filename
	Public gListFilePath As String 'Path to list file for commmunicating with SA.
	Public gAAFilePath As String 'Path to auto-align text file.
	
	Public gPhonGrpColl As New Collection 'Stores phone groups (e.g. Plosives, Front vowels).
	Public gPhonGrpNameArray As Object 'Allows indexing of group names.
	Public gTestGrp As Object 'Current group under test.
	Public gTestGrpName As String 'Name of current test group.
	Public gTestActive As Boolean 'True if in test mode.
	Public gTestLayout As Short 'Index of layout radio button (test setup).
	Public gItemNumber As Object 'Sound item number currently under test.
	Public gFirstSndPlayed As Boolean 'True if first test sound was played.
	Public gTestItemCorrect As Boolean 'Keeps track of correct guess on current test item
	Public gTestGrpCorrect As Object 'Array. Shows if test sound was not played yet
	'(vbGrayed), identified correctly (vbChecked),
	'or identified incorrectly (vbUnchecked).
	Public gRetestList As Object 'List of sounds for retest.
	Public gRetestListEmpty As Boolean 'True if all sounds identified correctly.
	Public gRetestActive As Boolean 'True if currently in retest mode.
	
	Public gRepeatPlay As Short 'Number of times to repeat the current item.
	Public gRecordingExists As Boolean 'Has the user made a recording?
	Public gWordListRepeat As Short 'Number of times to repeat items from word lists
	Public gSAPath As String 'Path to Speech Analyzer executable.
	
	Public Structure CharInfoStruct
		Dim sChar As String
		Dim sEx1 As String
		Dim sEx2 As String
		Dim sName As String
		Dim sDesc As String
		Dim sSoundFile As String
	End Structure
	
	Public Sub CenterForm(ByRef frm As System.Windows.Forms.Form, Optional ByRef RelativeToScreen As Object = Nothing)
		
		'********************************************************
		'* This function will center a form within the bounds
		'* of this program's MDI form unless the flag
		'* RelativeToScreen is set to True. Then the form is
		'* centered relative to the screen. If the MDI is
		'* maximized then there will be no difference.
		'********************************************************
		
		Dim bRelToScn As Boolean
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		bRelToScn = IIf(IsNothing(RelativeToScreen), False, RelativeToScreen)
		
		With frm
			If bRelToScn Then
				.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(.Height)) \ 2) 'Center vertically in screen.
				.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(.Width)) \ 2) 'Center horizontally in screen.
			Else
				.Top = VB6.TwipsToPixelsY(IIf(Not frm.MDIParent Is Nothing, 0, VB6.PixelsToTwipsY(mdiHelpCharts.Top)) + ((VB6.PixelsToTwipsY(mdiHelpCharts.Height) - VB6.PixelsToTwipsY(.Height)) \ 2)) 'Center vertically in MDI.
				.Left = VB6.TwipsToPixelsX(IIf(Not frm.MDIParent Is Nothing, 0, VB6.PixelsToTwipsX(mdiHelpCharts.Left)) + ((VB6.PixelsToTwipsX(mdiHelpCharts.Width) - VB6.PixelsToTwipsX(.Width)) \ 2)) 'Center horizontally in MDI.
			End If
		End With
		
	End Sub
	
	Public Function CheckForNeededFiles() As Boolean
		
		'***********************************************
		'* This function will search for several files
		'* that are needed before this program may
		'* run properly.
		'***********************************************
		'TODO: Change files to 32-bit versions and add things like the
		'  tree control, list view control and the hyperlink control.
		CheckForNeededFiles = True
		Exit Function
		
		Dim iPathsz As Short
		Dim sSysPathBuff As New VB6.FixedLengthString(254)
		Dim sSysPath As String
		Dim sAppPath As String
		
		On Error GoTo CheckForNeededFilesErr
		CheckForNeededFiles = False
		
		'**************************************
		'* Get the Windows system path.
		'**************************************
		iPathsz = GetSystemDirectory(sSysPathBuff.Value, 254)
		If iPathsz = 0 Then
			gMsg = "Error getting Windows system path."
			MsgBox(gMsg, MsgBoxStyle.Information, My.Application.Info.Title)
			Exit Function
		End If
		
		'**************************************
		'* This strips off the null terminator.
		'**************************************
		sSysPath = Trim(Left(sSysPathBuff.Value, iPathsz))
		
		gMsg1 = ""
		gMsg2 = ""
		If Not FileExist(sSysPath & "\" & "MHRUN600.DLL") Then gMsg1 = gMsg1 & "MHRUN600.DLL" & vbCrLf
		If Not FileExist(sSysPath & "\" & "MHSPLIT.VBX") Then gMsg1 = gMsg1 & "MHSPLIT.VBX" & vbCrLf
		If Not FileExist(sSysPath & "\" & "MHTIP.VBX") Then gMsg1 = gMsg1 & "MHTIP.VBX" & vbCrLf
		If Not FileExist(sSysPath & "\" & "PICCLP16.OCX") Then gMsg1 = gMsg1 & "PICCLP16.OCX" & vbCrLf
		If Not FileExist(sSysPath & "\" & "THREED16.OCX") Then gMsg1 = gMsg1 & "THREED16.OCX" & vbCrLf
		If Not FileExist(sSysPath & "\" & "MCI16.OCX") Then gMsg1 = gMsg1 & "MCI16.OCX" & vbCrLf
		If Not FileExist(sSysPath & "\" & "COMDLG16.OCX") Then gMsg1 = gMsg1 & "COMDLG16.OCX" & vbCrLf
		If Not FileExist(sSysPath & "\" & "TABCTL16.OCX") Then gMsg1 = gMsg1 & "TABCTL16.OCX" & vbCrLf
		
		sAppPath = My.Application.Info.DirectoryPath & IIf((Right(My.Application.Info.DirectoryPath, 1) = "\"), "", "\")
		If Not FileExist(sAppPath & "IPAHELP.INI") Then gMsg2 = gMsg2 & "IPAHELP.INI" & vbCrLf
		
		If gMsg1 <> "" Then
			gMsg1 = My.Application.Info.Title & " needs the following" & vbCrLf & "files in the Windows system directory" & vbCrLf & "but they could not be found." & vbCrLf & vbCrLf & gMsg1 & vbCrLf
			MsgBox(gMsg1, MsgBoxStyle.Information, My.Application.Info.Title)
		End If
		
		If gMsg2 <> "" Then
			gMsg2 = My.Application.Info.Title & " needs the following" & vbCrLf & "files in the Application directory" & vbCrLf & "but they could not be found." & vbCrLf & vbCrLf & gMsg2 & vbCrLf
			MsgBox(gMsg2, MsgBoxStyle.Information, My.Application.Info.Title)
		End If
		
		CheckForNeededFiles = (gMsg1 = "" And gMsg2 = "")
		
		Exit Function
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
CheckForNeededFilesErr: 
		MsgBox(Err.Description, MsgBoxStyle.Information, My.Application.Info.Title)
		
	End Function
	
	Public Sub ConvertWLFilesToXML()
		
		Dim i As Short
		Dim hProg As Integer
		Dim hProcess As Integer
		Dim ExitCode As Integer
		Dim xmlWL As clsXMLWordList
		
		For i = 0 To WordListArraySize()
			'********************************************************
			'* If the specified word list file exists, then check to
			'* see if it's in XML format.
			'********************************************************
			'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (FileExist(gWordListID(i, 1))) Then
				xmlWL = New clsXMLWordList
				
				'******************************************************
				'* If the XML parser object can't load the file as an
				'* XML file, then convert it to an XML file by calling
				'* the XML word list conversion program.
				'******************************************************
				'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (xmlWL.Load(CStr(gWordListID(i, 1)))) Then
					If (Len(xmlWL.ID) = 0) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(i, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						xmlWL.ID = gWordListID(i, 0)
					Else
						'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(i, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gWordListID(i, 0) = xmlWL.ID
					End If
					'UPGRADE_NOTE: Object xmlWL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					xmlWL = Nothing
				Else
					'UPGRADE_NOTE: Object xmlWL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
					xmlWL = Nothing
					gStatLine.Text = "Converting word lists to XML format."
					'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(i, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					hProg = Shell(My.Application.Info.DirectoryPath & "\" & "IHWLtoXML.exe /n /t """ & gWordListID(i, 0) & """" & gWordListID(i, 1), AppWinStyle.NormalFocus)
					
					'****************************************************
					'* Go through this loop while the program is still
					'* active and don't quite until it's finished.
					'****************************************************
					hProcess = OpenProcess(PrsQryInfo, False, hProg)
					Do 
						Call GetExitCodeProcess(hProcess, ExitCode)
						System.Windows.Forms.Application.DoEvents()
					Loop While ExitCode = StillActive
					
					'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object GiveFileXMLExtension(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(i, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gWordListID(i, 1) = GiveFileXMLExtension(gWordListID(i, 1))
				End If
			End If
		Next 
		
	End Sub
	
	Public Function FileExist(ByVal sFileSpec As String) As Short
		
		'***********************************************
		'* This function will use the Dir$ function to
		'* determine whether or not a file exists.
		'***********************************************
		
		Dim sRetSpec As String
		
		On Error GoTo BadFileSpec
		
		sFileSpec = Trim(sFileSpec)
		If (Len(sFileSpec) = 0) Then GoTo BadFileSpec
		
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		sRetSpec = Dir(sFileSpec) 'Can file be found?
		On Error GoTo 0
		If (Len(sRetSpec) = 0) Then GoTo BadFileSpec
		FileExist = True
		Exit Function
		
BadFileSpec: 
		FileExist = False
		Exit Function
		
	End Function
	
	Public Function GetCaptionFromTag(ByRef sTag As String) As Object
		
		Dim i As Short
		
		On Error Resume Next
		
		i = InStr(1, sTag, ";")
		If i > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object GetCaptionFromTag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetCaptionFromTag = Left(sTag, i - 1) & " - (" & Mid(sTag, i + 1) & ")"
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object GetCaptionFromTag. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			GetCaptionFromTag = ""
		End If
		
	End Function
	
	Private Function GetCharBytes(ByRef sCodes As String) As String
		
		'*******************************************************************
		'* This routine will parse a string consisting of one or more three
		'* digit numbers separated by commas. Each number is treated as a
		'* character code point and converted to its character. The
		'* returned string is the concatenation of each character.
		'*******************************************************************
		
		Dim i As Short
		Dim j As Short
		
		On Error Resume Next
		
		GetCharBytes = ""
		If (Len(sCodes) = 0) Then Exit Function
		
		j = 1
		Do 
			i = InStr(j, sCodes, ",")
			If (i = 0) Then
				GetCharBytes = GetCharBytes & Chr(CInt(Mid(sCodes, j)))
				Exit Do
			End If
			
			GetCharBytes = GetCharBytes & Chr(CInt(Mid(sCodes, j, i - j)))
			j = i + 1
			i = InStr(j, sCodes, ",")
		Loop 
		
	End Function
	
	Public Sub GetCharFromINIStr(ByRef sINIStr As String, ByRef lblCtrl As System.Windows.Forms.Label)
		
		Dim i As Short
		Dim j As Short
		
		On Error Resume Next
		
		i = InStr(1, sINIStr, ";")
		
		If i > 0 Then
			With lblCtrl
				'.Font.Name = cChartFontName
				.Tag = Mid(sINIStr, i + 1)
				j = 0
				.Text = ""
				Do 
					j = j + 1
					If Mid(sINIStr, j, 1) = "," Then
						.Text = .Text & Chr(Val(Mid(sINIStr, j - 3, 3)))
					ElseIf Mid(sINIStr, j, 1) = ";" Then 
						.Text = .Text & Chr(Val(Mid(sINIStr, j - 3, 3)))
					End If
				Loop Until Mid(sINIStr, j, 1) = ";"
			End With
		End If
		
	End Sub
	
	Public Sub GetINISections(ByRef sINIPath As String, ByRef sINISections() As String)
		
		Dim iFileNum As Short
		Dim sNextLine As String
		Dim iSections As Short
		Dim sSectionName As String
		Dim iTemp As Short
		Dim sTemp As String
		
		On Error Resume Next
		
		'* open List file
		iFileNum = FreeFile
		FileOpen(iFileNum, sINIPath, OpenMode.Input)
		'* find section headings
		iSections = 0
		ReDim sINISections(0)
		While Not EOF(iFileNum)
			sNextLine = LineInput(iFileNum)
			sNextLine = Trim(sNextLine)
			If Left(sNextLine, 1) = "[" And Right(sNextLine, 1) = "]" Then
				iSections = iSections + 1
				ReDim Preserve sINISections(iSections - 1)
				sSectionName = Mid(sNextLine, 2, Len(sNextLine) - 2)
				sINISections(iSections - 1) = sSectionName
				iTemp = UBound(sINISections)
			End If
		End While
		
		FileClose(iFileNum)
		
	End Sub
	
	Private Function GiveFileXMLExtension(ByVal sFileName As String) As Object
		
		'******************************************************
		'* This routine will change the extension on sFileName
		'* to ".XML"
		'******************************************************
		
		Dim i As Short
		
		i = Len(sFileName)
		Do While (i > 0)
			If (Mid(sFileName, i, 1) = ".") Then Exit Do
			i = i - 1
		Loop 
		
		'UPGRADE_WARNING: Couldn't resolve default property of object GiveFileXMLExtension. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		GiveFileXMLExtension = IIf(i > 0, Left(sFileName, i) & "xml", sFileName)
		
	End Function
	
	Public Function MakeFullPath(ByRef Folder As String, ByRef FileName As String) As Object
		
		If (CDbl(Trim(CStr(Len(FileName)))) = 0) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MakeFullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MakeFullPath = ""
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object MakeFullPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			MakeFullPath = Folder & IIf(Right(Folder, 1) = "\", "", "\") & Trim(FileName)
		End If
		
	End Function
	
	Public Function MakeInterVocName() As String
		
        With mdiHelpCharts.GetActiveMdiChild
            Select Case .Name
                Case "frmDispCon", "frmDispVow", "frmDispDia", "frmDispSSeg"
                    'UPGRADE_WARNING: Couldn't resolve default property of object mdiHelpCharts.ActiveForm.CurrIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object mdiHelpCharts.ActiveForm.WavNamePart1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    MakeInterVocName = gWavPath & .WavNamePart1 & "-" & VB6.Format(Trim(Str(.CurrIndex)), "00") & "b.wav"
                Case "frmWordList"
                    'UPGRADE_ISSUE: Control SoundFile could not be resolved because it was within the generic namespace ActiveMDIChild. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                    MakeInterVocName = mdiHelpCharts.GetActiveMdiChild.SoundFile

            End Select
        End With
	End Function
	
	Public Function MakeVocName() As String
		
        With mdiHelpCharts.GetActiveMdiChild
            Select Case .Name
                Case "frmDispCon", "frmDispVow", "frmDispDia", "frmDispSSeg"
                    'UPGRADE_WARNING: Couldn't resolve default property of object mdiHelpCharts.ActiveForm.CurrIndex. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object mdiHelpCharts.ActiveForm.WavNamePart1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    MakeVocName = gWavPath & .WavNamePart1 & "-" & VB6.Format(Trim(Str(.CurrIndex)), "00") & IIf(.Name = "frmDispVow", "w.wav", "a.wav") '"a.wav"
                Case "frmWordList"
                    'UPGRADE_ISSUE: Control SoundFile could not be resolved because it was within the generic namespace ActiveMDIChild. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
                    MakeVocName = mdiHelpCharts.GetActiveMdiChild.SoundFile

            End Select
        End With
		
	End Function
	
	Public Sub MarkItemCorrect(ByRef iMark As Short, Optional ByRef Index As Object = Nothing)
		
		Dim i As Short 'Index for gTestGrp.
		
		'* This line can be reinstated or put on options page if user wants to be
		'* able to continually retest on the same set. CLW 1/26/99
		'If gFirstSndPlayed And Not gRetestActive Then
		'**********************************************************************************************
		'* Search test group for the item under test and the item that was clicked (supplied only if
		'* item was guessed incorrectly). Caller must supply iMark: vbUnchecked if incorrect,
		'* vbChecked if correct. If the item under test was guessed correctly this time, but
		'* previously guessed incorrectly, it will be left as vbUnchecked. Only items that are
		'* guessed correctly every time can be vbChecked.
		'**********************************************************************************************
		For i = 0 To UBound(gTestGrp) 'Find correct item and item clicked.
			'UPGRADE_WARNING: Couldn't resolve default property of object Index. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object gTestGrp(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object gItemNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object gTestGrpCorrect(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (gTestGrp(i) = gItemNumber Or gTestGrp(i) = CShort(Index)) And gTestGrpCorrect(cThisTest, i) <> System.Windows.Forms.CheckState.Unchecked Then gTestGrpCorrect(cThisTest, i) = iMark 'Change mark for that test group item.
		Next i
		'End If
		If iMark = System.Windows.Forms.CheckState.Checked Then gTestItemCorrect = True
		
	End Sub
	
	Public Sub Pause(ByVal dSecs As Double)
		
		'**********************************************
		'* This routine will cause the PC to pause for
		'* dSecs. If the start time is before midnight
		'* and the end time is after then don't pause
		'* since the pause would be for an entire day.
		'* That's my way of dealing with a pause that
		'* spans two days.
		'*
		'* So that a form may stop the pause, the
		'* active form's global variable called
		'* CancelPause is check. If it is true then
		'* the pause loop is exited. If the global
		'* variable doesn't exist then the If is
		'* ignored.
		'**********************************************
		
        'Dim dEnd As Double
        'Dim bCancel As Boolean

        'On Error Resume Next
        'dEnd = VB.Timer() + dSecs
        'If dEnd > 86400# Then Exit Sub

        '      Do While VB.Timer() < dEnd
        '          Err.Clear()
        '          bCancel = System.Windows.Forms.Form.ActiveForm.CancelPause
        '          If Err.Number = 0 Then
        '              If bCancel Then Exit Do
        '          End If
        '          System.Windows.Forms.Application.DoEvents()
        '      Loop

        System.Threading.Thread.Sleep(dSecs)


	End Sub
	
	Public Sub PlayWav(ByVal sWavFile As String)
		
		On Error Resume Next
		
		'*********************************************************
		'* Ay this point in the code, the MCI control should
		'* already be closed. However, this is just a precaution.
		'*********************************************************
		gMMCtrl.Command = "Close"
		
		'If gMMCtrl.Tag <> "Compare" Then .Tag = "PlayWav" '***** Added by CLW 4/28/99
		
		If Mid(sWavFile, 2, 1) = ":" Then
			gMMCtrl.FileName = sWavFile
		Else
			gMMCtrl.FileName = gWavPath & sWavFile 'Tell MCI the WAV path.
		End If
		
		gMMCtrl.Command = "Open" 'Open MCI control w/WAV file.
		
		'*********************************************************
		'* This seems like a kludge and it probably is. However,
		'* it seems to be necessary. I suspect setting the Wait
		'* property only works after the MCI control is "Open".
		'* And since, the Wait property is expected to be set by
		'* the caller, I perform this, seemingly useless,
		'* assignment now, after the "Open" command.
		'*********************************************************
		gMMCtrl.Wait = gMMCtrl.Wait
		
		gMMCtrl.Notify = True
		gMMCtrl.Command = "Play"
		
	End Sub
	
	Public Sub RandomTest()
		
		'****************************************************
		'* This function will either start the timer (during
		'* test mode), or turn the timer off (if not in test
		'* mode). It is necessary to disable the timer at the
		'* end of the alotted time.
		'****************************************************
		With mdiHelpCharts
			Select Case gTestForm '* Added by CLW 5/6/99
				Case "frmDispCon", "frmDispVow"
					If gTestActive Then
						.Timer2.Interval = 5000
						.Timer2.Enabled = True
					Else
						.Timer2.Enabled = False
					End If
				Case "frmWordList" '* Added by CLW 5/6/99
					If gFirstSndPlayed Then .Timer2.Enabled = False
			End Select
		End With
		
	End Sub
	
	Public Sub SetIPAFontInfo(ByRef lblCtrl As System.Windows.Forms.Label, ByRef sINISection As String)
		
		On Error Resume Next
		
		If (Len(GetINIEntry(sINISection, cFontNameEntry, gINIPath)) = 0) Then Call WriteIPAFontInfo(lblCtrl, sINISection)
		
		With lblCtrl.Font
			lblCtrl.Font = VB6.FontChangeName(lblCtrl.Font, GetINIEntry(sINISection, cFontNameEntry, gINIPath))
			lblCtrl.Font = VB6.FontChangeSize(lblCtrl.Font, CDec(GetINIEntry(sINISection, cFontSizeEntry, gINIPath)))
			lblCtrl.Font = VB6.FontChangeBold(lblCtrl.Font, CBool(GetINIEntry(sINISection, cFontBoldEntry, gINIPath)))
			lblCtrl.Font = VB6.FontChangeItalic(lblCtrl.Font, CBool(GetINIEntry(sINISection, cFontItalicEntry, gINIPath)))
		End With
		
	End Sub
	
	Public Sub SetSlowedReplayStatLine()
		
		On Error Resume Next
		gStatusLine(cSRBttn) = "Playback selected item at " & gSRSpeed & "% speed"
		
	End Sub
	
	Private Sub SetupHelpText()
		
		Dim sTmp As String
		
		On Error Resume Next
		
		gToolTip(cSRBttn) = "Slowed Replay"
		gToolTip(cPitchBttn) = "Pitch Graph"
		gToolTip(cAABttn) = "Auto Align"
		gToolTip(cVocBttn) = "Hear Sound Alone (Vocalic)"
		gToolTip(cInterVocBttn) = "Hear Sound in Context (Intervocalic)"
		gToolTip(cRecBttn) = "Record"
		gToolTip(cMCIStopBttn) = "Stop"
		gToolTip(cPlayBttn) = "Play"
		gToolTip(cCompBttn) = "Compare"
		gToolTip(cTestBttn) = "Start Test"
		gToolTip(cTestStopBttn) = "Stop"
		gToolTip(cExitBttn) = "Close Active Window"
		
		Call SetSlowedReplayStatLine()
		gStatusLine(cPitchBttn) = "Display a pitch graph for the selected item"
		gStatusLine(cAABttn) = "Auto-align phonetic text for the selected item"
		
		sTmp = "Listen to a sample of the selected "
		gStatusLine(cVocBttn) = sTmp & "sound by itself"
		gStatusLine(cInterVocBttn) = sTmp & "sound in context"
		gStatusLine(cRecBttn) = "Record Sound for Comparison"
		gStatusLine(cMCIStopBttn) = "Stop Recording or Playback"
		gStatusLine(cPlayBttn) = "Play Recorded Sound"
		gStatusLine(cCompBttn) = "Compare Recorded Sound with Original"
		gStatusLine(cTestBttn) = "Begin Listening Test"
		gStatusLine(cTestStopBttn) = "End Listening Test"
		gStatusLine(cExitBttn) = gToolTip(cExitBttn)
		
	End Sub
	
	Public Function StripOffFileName(ByVal FullFilePath As String) As String
		
		'***************************************************
		'* This function receives a full file path
		'* and returns just the path portion.
		'***************************************************
		
		Dim iPtr As Short
		Dim iSavPtr As Short
		
		On Error Resume Next
		StripOffFileName = ""
		
		iPtr = 0
		Do  'Loop until we find the
			iSavPtr = iPtr + 1 '  last backslash.
			iPtr = InStr(iSavPtr, FullFilePath, "\")
		Loop Until iPtr = 0
		
		StripOffFileName = Left(FullFilePath, iSavPtr - 1)
		
	End Function
	
	Public Function StripOffPath(ByRef FullFilePath As String) As String
		
		'***************************************************
		'* This function receives a full file path and
		'* returns just the file name portion of the path.
		'***************************************************
		
		Dim iPtr As Short
		Dim iSavPtr As Short
		
		On Error Resume Next
		StripOffPath = ""
		If (Len(FullFilePath) = 0) Then Exit Function
		
		iPtr = 0
		Do  'Loop until we find the
			iSavPtr = iPtr + 1 '  last backslash.
			iPtr = InStr(iSavPtr, FullFilePath, "\")
		Loop Until iPtr = 0
		
		StripOffPath = Mid(FullFilePath, iSavPtr)
		
	End Function
	
	Public Sub UpdateRetestList()
		
		'****************************************************
		'* This function updates the list of untested or
		'* incorrectly guessed items from the previous test.
		'* The boolean flag gRetestListEmpty is set to true
		'* if all items were correctly guessed.
		'****************************************************
		
		Dim i As Short
		Dim iNumCorrect As Short
		Dim j As Short
		
		On Error Resume Next
		
		iNumCorrect = 0 '* Keep track of number of correct items
		
		For i = 0 To UBound(gTestGrpCorrect, 2) '* First count number correct.
			Select Case gTestGrpCorrect(cThisTest, i)
				Case System.Windows.Forms.CheckState.Checked : iNumCorrect = iNumCorrect + 1
				Case System.Windows.Forms.CheckState.Unchecked
					'UPGRADE_WARNING: Couldn't resolve default property of object gTestGrpCorrect(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gTestGrpCorrect(cThisTest, i) = System.Windows.Forms.CheckState.Indeterminate
			End Select
		Next i
		
		'ReDim gRetestList(0 To (UBound(gTestGrp) - iNumCorrect))  'Resize retest list to number wrong.
		'j = 0
		'For i = 0 To UBound(gTestGrp)
		'  If gTestGrpCorrect(i) <> vbChecked Then         'If not correct, add to retest list.
		'    gRetestList(j) = gTestGrp(i)
		'    j = j + 1
		'  End If
		'Next i
		
		'* If all correct, no need for retest.
		gRetestListEmpty = (iNumCorrect = UBound(gTestGrpCorrect, 2) + 1) '* CLW mod 1/26/99
		
	End Sub
	
	Public Sub UpdateTestMenu()
		
		'****************************************************
		'* This function updates the test menu's Enabled
		'* property according to whether or not any of the
		'* Vow or Con sound files exist. On forms where the
		'* Test button can be enabled, its status is also
		'* updated.
		'****************************************************
		
		Dim sWavFile As String
		Dim iRow As Short
		Dim iCell As Short
		
		On Error Resume Next
		
		With mdiHelpCharts
			Select Case .ActiveMDIChild.Name
				Case "frmDispCon"
					.mnuTest.Enabled = FileExist(gWavPath & "Con*.wav")
				Case "frmDispVow"
					.mnuTest.Enabled = FileExist(gWavPath & "Vow*.wav")
				Case "frmWordList"
					.mnuTest.Enabled = True
				Case "frmMenu"
					.mnuTest.Enabled = True
				Case Else
					.mnuTest.Enabled = False
			End Select
			
			.TBar.Items.Item("Test").Enabled = .mnuTest.Enabled
		End With
		
	End Sub
	
	Public Function WordListArraySize() As Short
		
		On Error Resume Next
		WordListArraySize = UBound(gWordListID, 1)
		If (Err.Number <> 0) Then WordListArraySize = -1
		
	End Function
	
	Public Sub WriteIPAFontInfo(ByRef lblCtrl As System.Windows.Forms.Label, ByRef sINISection As String)
		
		On Error Resume Next
		
		With lblCtrl.Font
			Call WriteINIEntry(sINISection, cFontNameEntry, .Name, gINIPath)
			Call WriteINIEntry(sINISection, cFontSizeEntry, CStr(.SizeInPoints), gINIPath)
			Call WriteINIEntry(sINISection, cFontBoldEntry, CStr(.Bold), gINIPath)
			Call WriteINIEntry(sINISection, cFontItalicEntry, CStr(.Italic), gINIPath)
		End With
		
	End Sub
	
	Public Sub WriteWordListsToIni(Optional ByRef vWLPaths As Object = Nothing)
		
		Dim i As Short
		Dim iMax As Short
		
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If (IsNothing(vWLPaths)) Then
			iMax = WordListArraySize()
		Else
			iMax = UBound(vWLPaths)
			ReDim gWordListID(UBound(vWLPaths), 1)
		End If
		
		Call WriteINIEntry("WordListPaths", vbNullString, vbNullString, gINIPath)
		
		For i = 0 To iMax
			'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
			If (IsNothing(vWLPaths)) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(i, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call WriteINIEntry("WordListPaths", gWordListID(i, 0), gWordListID(i, 1), gINIPath)
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object vWLPaths(i, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object vWLPaths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				Call WriteINIEntry("WordListPaths", vWLPaths(i, 0), vWLPaths(i, 1), gINIPath)
				'UPGRADE_WARNING: Couldn't resolve default property of object vWLPaths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(i, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gWordListID(i, 0) = vWLPaths(i, 0)
				'UPGRADE_WARNING: Couldn't resolve default property of object vWLPaths(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(i, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gWordListID(i, 1) = vWLPaths(i, 1)
			End If
		Next 
		
	End Sub
End Module