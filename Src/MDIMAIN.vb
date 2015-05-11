Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class mdiHelpCharts
	Inherits System.Windows.Forms.Form
	
	'**************************************************
	'* mdiHelpCharts version info:
	'*  See basGlobals (Globals.bas).
	'**************************************************

	Public bStartTest As Boolean '* Accessed from frmTestSetup.
	Private bctrlkey As Boolean
	Private bFormWasLoaded As Boolean '* Determines whether or not to reopen a form.
	Private bDontRestoreTBAfterPlay As Boolean
	Private EnterCount As Short
	Private dRunningLen As Double '* Used by MCI control.
    Private dRecLength As Double '* Used by MCI control.

    Private WithEvents mciIPA As New MCI.MMControl

	Public Sub CallSA()
		
		Dim sArguments As String
		
		On Error GoTo CallSAErr
		
		'*******************************************************
		'* Show List file if Ctrl + Click
		'*******************************************************
        If (bctrlkey) Then
            Call Shell("Notepad " & gListFilePath, AppWinStyle.NormalFocus)
        End If

		'*******************************************************
		'* Generate Command line string
		'*******************************************************
		gSAPath = GetINIEntry(cPathsSect, cSAINIEntry, gINIPath)
		sArguments = " -l " & gListFilePath
		'*******************************************************
		'* Show Command line if Ctrl + Click
        '*******************************************************
		If (bctrlkey) Then
            If (MsgBox("Command Line: " & vbCrLf & sArguments & vbCrLf & vbCrLf & "Do you want to run this?", MsgBoxStyle.YesNo) = MsgBoxResult.No) Then
                Exit Sub
            End If
        End If
		
		Call Shell(gSAPath & sArguments, AppWinStyle.NormalFocus)
		
		bctrlkey = False
		Exit Sub
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
CallSAErr: 
		If (Err.Number = 53) Then
			MsgBox(gSAPath & " not found.", MsgBoxStyle.Information, My.Application.Info.Title)
		Else
			MsgBox(Err.Description, MsgBoxStyle.Information, My.Application.Info.Title)
		End If
		
	End Sub
	
	Public Sub EnableTBarButtons(ByRef sKeys As String)
		
		Dim i As Short
		Dim j As Short
		
		On Error Resume Next
		
        For i = 1 To TBar.Items.Count
            If (TBar.Items.Item(i).Visible) Then
                j = InStr(sKeys, TBar.Items.Item(i).Name & ";")
                If (InStr(TBar.Items.Item(i).Name, "PlayRec") > 0 And j > 0) Then
                    TBar.Items.Item(i).Enabled = FileExist(gTmpWavPath & gTmpWavName)
                ElseIf (TBar.Items.Item(i).Name = "StopRec" And j > 0) Then
                    TBar.Items.Item(i).Enabled = (CType(TBar.Items.Item("Record"), ToolStripMenuItem).Checked = True Or CType(TBar.Items.Item("PlayRec"), ToolStripMenuItem).Checked = True Or CType(TBar.Items.Item("PlayRecSpeaker"), ToolStripMenuItem).Checked = True)
                Else
                    TBar.Items.Item(i).Enabled = (j > 0)
                End If
            End If
        Next
        TBar.Refresh()

	End Sub
	
	Private Sub LoadWordLists()
		
		Dim i As Short
		Dim vINISettings As Object
		
		'************************************************
		'* Word Lists
		'************************************************
        vINISettings = GetAllINISettings(gINIPath, "WordListPaths")
		
        If Not (IsDBNull(vINISettings)) Then
            ReDim gWordListID(UBound(vINISettings), 1)

            For i = 0 To UBound(vINISettings)
                gWordListID(i, 0) = vINISettings(i, 0)
                gWordListID(i, 1) = vINISettings(i, 1)
            Next

            Call ConvertWLFilesToXML()
            Call WriteWordListsToIni()
            gStatLine.Text = ""
        End If
		
	End Sub
	
	Private Sub MakePhonGrpColl()
		
		Dim CentVowGroup, AllVowGroup, NonPulConGroup, FricConGroup, PlosConGroup, AllConGroup, NTFConGroup, ApprConGroup, OtherConGroup, FrontVowGroup, BackVowGroup As Object
		
		On Error Resume Next
		
        AllConGroup = New Object() {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83}
        PlosConGroup = New Object() {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12}
        NTFConGroup = New Object() {13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24}
        FricConGroup = New Object() {25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48}
        ApprConGroup = New Object() {49, 50, 51, 52, 53, 54, 55, 56, 57}
        NonPulConGroup = New Object() {58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70}
        OtherConGroup = New Object() {72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83}
        AllVowGroup = New Object() {0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27}
        FrontVowGroup = New Object() {0, 1, 6, 7, 9, 10, 16, 17, 22, 24, 25}
        CentVowGroup = New Object() {2, 3, 11, 12, 15, 18, 19, 23}
        BackVowGroup = New Object() {4, 5, 8, 13, 14, 20, 21, 26, 27}
		With gPhonGrpColl
			.Add(AllConGroup, "AllCon")
			.Add(PlosConGroup, "PlosCon")
			.Add(NTFConGroup, "NTFCon")
			.Add(FricConGroup, "FricCon")
			.Add(ApprConGroup, "ApprCon")
			.Add(NonPulConGroup, "NonPulCon")
			.Add(OtherConGroup, "OtherCon")
			.Add(AllVowGroup, "AllVow")
			.Add(FrontVowGroup, "FrontVow")
			.Add(CentVowGroup, "CentVow")
			.Add(BackVowGroup, "BackVow")
		End With
        gPhonGrpNameArray = New Object() {"AllCon", "PlosCon", "NTFCon", "FricCon", "ApprCon", "NonPulCon", "OtherCon", "AllVow", "FrontVow", "CentVow", "BackVow"}
		gTestGrpName = "AllCon"
		gTestLayout = cTestChart
		gRetestListEmpty = True
		
	End Sub
	
	Private Sub PlayRecording(ByRef iButtonState As Short, ByRef bCompare As Boolean)
		
		On Error Resume Next
		
		'****************************************************************
		'* If we're here and the button's state is not pressed, it
		'* means the user pressed the button while it was down in order
		'* to stop playback. Therefore, stop playback.
		'****************************************************************
		If (iButtonState = False) Then
			CType(TBar.Items.Item("PlayRec"), ToolStripButton).Checked = False
			CType(TBar.Items.Item("PlayRecSpeaker"), ToolStripButton).Checked = False
			mnuFile.Enabled = True
			mnuTest.Enabled = True
			mnuWindow.Enabled = True
			mnuHelp.Enabled = True
            Call GetActiveMdiChild.UpdateAfterRecordAndPlayback()
			gMMCtrl.Tag = ""
			Exit Sub
		End If
		
		'****************************************************************
		'* If MCI control is busy then don't allow playback.
		'****************************************************************
		If (Len(gMMCtrl.Tag) > 0) Then
			Beep()
			CType(TBar.Items.Item("PlayRec"), ToolStripButton).Checked = False
			CType(TBar.Items.Item("PlayRecSpeaker"), ToolStripButton).Checked = False
			Exit Sub
		End If
		
		Call EnableTBarButtons("StopRec;" & IIf(bCompare, "PlayRecSpeaker;", "PlayRec;"))
		mnuFile.Enabled = False
		mnuTest.Enabled = False
		mnuWindow.Enabled = False
		mnuHelp.Enabled = False
		
		gMMCtrl.Tag = cMCIBusy
		bDontRestoreTBAfterPlay = True
		Call PlayWav(gTmpWavPath & gTmpWavName)
		gMMCtrl.Tag = ""
		
		If (bCompare) Then
			With gMMCtrl
				While (.Mode <> MCI.ModeConstants.mciModeNotOpen And .Mode <> MCI.ModeConstants.mciModeReady) : System.Windows.Forms.Application.DoEvents() : End While
				.Command = "Close"
			End With
			bDontRestoreTBAfterPlay = False
            Call GetActiveMdiChild.Play(0)
		End If
		
		bDontRestoreTBAfterPlay = False
		
	End Sub
	
	Private Sub RecordUser(ByRef iButtonState As Short)
		
		On Error Resume Next
		
		'****************************************************************
		'* If we're here and the button's state is not pressed, it
		'* means the user pressed the button while it was down in order
		'* to stop recording. Therefore, stop recording.
		'****************************************************************
		If (iButtonState = False) Then
			Call StopRecording()
			Exit Sub
		End If
		
		'****************************************************************
		'* If MCI control is busy then don't allow recoding.
		'****************************************************************
		If (Len(gMMCtrl.Tag) > 0) Then
			Beep()
			CType(TBar.Items.Item("Record"), ToolStripButton).Checked = False
			Exit Sub
		End If
		
		'***********************************************************************
		'* We make use of a Template file (gMstrWavName) which automatically
		'* sets the sampling frequency and data size.
		'***********************************************************************
		Kill(gTmpWavPath & gTmpWavName)
		If FileExist(gWavPath & gMstrWavName) Then FileCopy((gWavPath & gMstrWavName), gTmpWavName)
		Call EnableTBarButtons("Record;StopRec;")
		
		With gMMCtrl
			.Tag = cMCIBusy
			.FileName = gTmpWavPath & gTmpWavName
			.Command = "Open"
			gMMCtrl.Command = "Record"
		End With
		
	End Sub
	
	Public Sub ShowTBarButtons(ByRef sKeys As String)
		
		Dim i As Short
		
		On Error Resume Next
		
		With TBar
			For i = 1 To .Items.Count
				'UPGRADE_WARNING: Lower bound of collection TBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				.Items.Item(i).Visible = (InStr(sKeys, .Items.Item(i).Name & ";") > 0)
			Next 
			
			If Not (.Visible) Then .Visible = True
			
			'*********************************************************************
			'* Remove the following two lines when a link with SA is working.
			'*********************************************************************
			.Items.Item("PlaySlow").Enabled = False
			.Items.Item("Pitch").Enabled = False
		End With
		
	End Sub
	
	Public Sub StartTest()
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		
		Dim frm As System.Windows.Forms.Form
		Dim i As Short
		
		'*********************
		'* Activate test mode
		'*********************
		gTestActive = True
        If (CType(TBar.Items.Item("Test"), ToolStripMenuItem).Checked <> True) Then CType(TBar.Items.Item("Test"), ToolStripButton).Checked = True

        '****************************************************
        '* Show form for test and disable unused chars.
        '* Reworked to accomodate Word Lists by CLW 5/5/99
        '****************************************************
        bFormWasLoaded = False
        For Each frm In My.Application.OpenForms
            If ((frm.Name = gTestForm) And (frm.Tag = gTestTag)) Then
                frm.Show()
                If TypeOf frm Is frmDispCon Then
                    Call CType(frm, frmDispCon).UpdateFormForTest()
                ElseIf TypeOf frm Is frmDispVow Then
                    Call CType(frm, frmDispCon).UpdateFormForTest()
                End If
                bFormWasLoaded = True
                Exit For
            End If
        Next frm

        If Not (bFormWasLoaded) Then
            Select Case gTestForm
                Case "frmDispCon"
                    frmDispCon.Show()
                    System.Windows.Forms.Application.DoEvents()
                    Call frmDispCon.UpdateFormForTest()

                Case "frmDispVow"
                    frmDispVow.Show()
                    System.Windows.Forms.Application.DoEvents()
                    Call frmDispVow.UpdateFormForTest()

                Case "frmWordList"
                    System.Windows.Forms.Application.DoEvents()
                    frmWordList.Initialize(gTestTag)
                    Call frmWordList.UpdateFormForTest()
            End Select
        End If

        System.Windows.Forms.Application.DoEvents()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

        '****************************************
        '* Disable extra menus. Update toobar.
        '****************************************
        mnuFile.Enabled = False
        mnuTestSetup.Enabled = False
        mnuWindow.Enabled = False
        mnuTestStop.Enabled = True
        Call EnableTBarButtons("Test;")

        '*********************************************
        '* 2 second delay before playing first sound.
        '*********************************************
        If (gTestForm <> "frmWordList") Then
            gFirstSndPlayed = False
            With Timer2
                .Interval = 2000
                .Enabled = True
            End With
            Call RandomTest()
        Else
            Call Pause(1)
            Call (CType(frm, frmWordList)).TestPlay()
        End If

    End Sub
	
	Private Sub StopRecording()
		
		On Error Resume Next
		
		With gMMCtrl
			.Command = "Stop"
			.Command = "Save"
			.Command = "Close"
			.Tag = ""
		End With
		
		CType(TBar.Items.Item("Record"), ToolStripButton).Checked = False
        Call GetActiveMdiChild.UpdateAfterRecordAndPlayback()
		
	End Sub
	
	Private Sub StopTest()
		
		Dim sUpdateGrpName As String
		Dim vUpdateGroup As Object
		Dim frm As System.Windows.Forms.Form
		
		On Error Resume Next
		
		'* First, see if user switched forms during the test.
		'* Added by CLW 5/12/99
		If ActiveMDIChild.Name <> gTestForm Then
			For	Each frm In My.Application.OpenForms
				If (frm.Name = gTestForm And frm.Tag = gTestTag) Then Exit For
			Next frm
		End If
		
		'********************************************
		'* If the test is active, that means it was
		'* not aborted while the test setup form
		'* was open. This is necessary to know when
		'* deciding to reopen the active form or not.
		'********************************************
		If gTestActive Then
			If (gTestForm <> "frmWordList") Then '* Added by CLW 5/11/99
				sUpdateGrpName = "All" & VB.Right(gTestGrpName, 3)
				'UPGRADE_WARNING: Couldn't resolve default property of object gPhonGrpColl(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object vUpdateGroup. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				vUpdateGroup = gPhonGrpColl.Item(sUpdateGrpName)
				Call UpdateRetestList() '* Moved by CLW 5/11/99
			End If
			gTestActive = False
			If bFormWasLoaded Then
                Call GetActiveMdiChild.UpdateFormAfterTest()
			Else
				ActiveMDIChild.Close()
			End If
			
			Call RandomTest()
		End If
		
		CType(TBar.Items.Item("Test"), ToolStripButton).Checked = False
		mnuFile.Enabled = True
		mnuTestSetup.Enabled = True
		mnuWindow.Enabled = True
		mnuTestStop.Enabled = False
		
	End Sub
	
    Private Sub mciIPA_Done(ByVal eventSender As System.Object, ByVal eventArgs As AxMCI.DmciEvents_DoneEvent)

        If (Not bDontRestoreTBAfterPlay And (CType(TBar.Items.Item("PlayRec"), ToolStripMenuItem).Checked = True Or CType(TBar.Items.Item("PlayRecSpeaker"), ToolStripMenuItem).Checked = True)) Then
            Call PlayRecording(False, False)
        End If
        gMMCtrl.Command = "Close"

    End Sub
	
    Public Sub mciIPA_StatusUpdate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)
        '************************************************
        '* Keep checking to see if the recording is
        '* 2 seconds long yet. If it is, stop recording.
        '************************************************
        'Debug.Print .length; " "; .Mode
        If (gMMCtrl.Mode = MCI.ModeConstants.mciModeRecord And gMMCtrl.Length = 2000) Then Call StopRecording()

    End Sub
	
	Private Sub mdiHelpCharts_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim i As Short
		Dim sTemp As String
        Dim soundPath As String
        Dim vINISettings As Object
        Dim CentVowGroup, AllVowGroup, NonPulConGroup, FricConGroup, PlosConGroup, AllConGroup, NTFConGroup, ApprConGroup, OtherConGroup, FrontVowGroup, BackVowGroup As Object
        Dim tempPath As String
        Dim iLen As Short
        Dim sAccessTestPath As String
        Dim sRegPath As String
        Dim sRegValueName As String
        Dim sRegValue As String
        Dim lRegDataLen As Integer
        Dim lRetVal As Integer 'Result of RegOpenKeyEx
        Dim hKey As Object
        Dim hSubKey As Integer
        Dim ret As Object

        On Error GoTo MDIFormLoadErr

        If Not CheckForNeededFiles() Then End
        gStatLine = panStatus
        Dim mciIPA As New MCI.MMControl
        gMMCtrl = mciIPA
        gStatLine.Text = ""
        bctrlkey = False
        bDontRestoreTBAfterPlay = False

        '************************************************
        '* INI file and help file stuff.
        '************************************************
        gINIPath = My.Application.Info.DirectoryPath & IIf(Len(My.Application.Info.DirectoryPath) = 3, "", "\") & INIFile

        soundPath = GetINIEntry(cPathsSect, cSoundsEntry, gINIPath)
        gWavPath = My.Application.Info.DirectoryPath & IIf(VB.Right(My.Application.Info.DirectoryPath, 1) = "\", "", "\") 'IPA Help folder is default Sounds path

        '*************************************************
        '* If the Sounds path key was there then set the
        '* global for that value. Otherwise look for one.
        '*************************************************
        If ((soundPath.Length > 0) And FileExist(soundPath)) Then
            gWavPath = soundPath
        Else
            If FileExist(gWavPath & "Sounds\*.*") Then gWavPath = gWavPath & "Sounds\"
        End If

        sTemp = GetINIEntry(cSettingsSect, cSRSpeedEntry, gINIPath)
        gSRSpeed = IIf(Len(sTemp) = 0, 50, Val(sTemp))

        tempPath = Space(255)
        iLen = GetTempPath(255, tempPath)
        tempPath = VB.Left(tempPath, iLen)
        If iLen = 0 Then tempPath = My.Application.Info.DirectoryPath
        tempPath = Trim(tempPath)
        gListFilePath = MakeFullPath(tempPath, "ipa-help.lst")
        gAAFilePath = MakeFullPath(tempPath, "aa.txt")

        '************************************************
        ' Check if user has write access to app path and
        ' disable menu items if not.
        '************************************************
        gAppPathWriteAccess = False
        'UPGRADE_WARNING: Couldn't resolve default property of object MakeFullPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        sAccessTestPath = MakeFullPath(My.Application.Info.DirectoryPath, "AccessTest.ini")
        Call WriteINIEntry("Test", "WriteAccess", "True", sAccessTestPath)
        If (FileExist(sAccessTestPath)) Then
            gAppPathWriteAccess = True
            Kill(sAccessTestPath)
        Else
        End If
        mnuPreferences.Enabled = gAppPathWriteAccess
        mnuEditMode.Enabled = gAppPathWriteAccess

        '************************************************
        ' Make sure the IPA Help location is available in
        ' the registry.
        '************************************************
        sRegPath = "Software\SIL\IPA Help"
        sRegValueName = "Location"
        sRegValue = ""
        lRetVal = RegCreateKey(HKEY_CURRENT_USER, sRegPath, hKey)
        If lRetVal = 0 Then
            lRetVal = RegQueryValueExNULL(hKey, sRegValueName, 0, REG_SZ, 0, lRegDataLen)
            If lRetVal = 0 Then
                sRegValue = New String(Chr(0), lRegDataLen - 1)
                lRetVal = RegQueryValueExString(hKey, sRegValueName, 0, REG_SZ, sRegValue, lRegDataLen)
            End If
            If Not (lRetVal = 0 And Len(sRegValue) > 0) Then
                'No value or the value is empty. Set the value.
                RegSetValueEx(hKey, sRegValueName, 0, REG_SZ, My.Application.Info.DirectoryPath, Len(My.Application.Info.DirectoryPath))
            End If
        End If
        'RegSetValueEx hKey, sRegValueName, 0, REG_SZ, "X", Len("X")
        RegCloseKey(hKey)

        '************************************************
        '* Initialize the MCI control.
        '************************************************
        gMMCtrl.PlayVisible = False
        gMMCtrl.Shareable = False
        gMMCtrl.DeviceType = "WaveAudio"
        gMMCtrl.UpdateInterval = 100
        gMMCtrl.TimeFormat = MCI.FormatConstants.mciFormatMilliseconds

        gMstrWavName = "template.wav"
        gTmpWavPath = tempPath & IIf(VB.Right(tempPath, 1) = "\", "", "\")
        gTmpWavName = "~ipa-wav.tmp"

        On Error Resume Next
        Kill(gTmpWavPath & gTmpWavName)
        gPixelX = VB6.TwipsPerPixelX
        gPixelY = VB6.TwipsPerPixelY

        Height = VB6.TwipsToPixelsY(6645)
        Width = VB6.TwipsToPixelsX(9105)
        TBar.Visible = False
        Show()
        Call LoadWordLists()
        frmMenu.Show()
        Call MakePhonGrpColl()
        gSAPath = ""

        Exit Sub

        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MDIFormLoadErr:
        MsgBox(Err.Description, MsgBoxStyle.Information, My.Application.Info.Title)

	End Sub
	
	'UPGRADE_ISSUE: Form event MDIForm.MouseMove was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="ABD9AF39-7E24-4AFF-AD8D-3675C1AA3054"'
	Private Sub MDIForm_MouseMove(ByRef Button As Short, ByRef Shift As Short, ByRef X As Single, ByRef Y As Single)
		
		On Error Resume Next
		gStatLine.Text = ""
		
	End Sub
	
	Private Sub mdiHelpCharts_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		Dim i As Short
		
		On Error Resume Next
		Kill(gTmpWavPath & gTmpWavName)
		'UPGRADE_NOTE: Object mdiHelpCharts may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Me.Dispose()
		
		eventArgs.Cancel = Cancel
	End Sub
	
	Public Sub mnuAbout_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuAbout.Click
		
		On Error Resume Next
		frmHelpAbout.ShowDialog()
		
	End Sub
	
	Public Sub mnuBkgrdColor_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuBkgrdColor.Click
		Dim Index As Short = mnuBkgrdColor.GetIndex(eventSender)
		
		On Error Resume Next
		'UPGRADE_ISSUE: Control IPAHelpPrint could not be resolved because it was within the generic namespace ActiveMDIChild. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        Call GetActiveMdiChild.IPAHelpPrint(False, Index = 1)
		
    End Sub

    ReadOnly Property GetActiveMdiChild() As Object
        Get
            Return Me.ActiveMdiChild
        End Get
    End Property
	
	Public Sub mnuDeleteCategory_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuDeleteCategory.Click
		
		On Error Resume Next
        Call GetActiveMdiChild.DeleteCategory()
		
	End Sub
	
	Public Sub mnuEditFonts_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditFonts.Click
		
		On Error Resume Next
		'UPGRADE_ISSUE: Control EditFonts could not be resolved because it was within the generic namespace ActiveMDIChild. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        Call GetActiveMdiChild.EditFonts()
		
	End Sub
	
	Public Sub mnuEditMode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditMode.Click
		
		On Error Resume Next
        Call GetActiveMdiChild.EditMode()
		
	End Sub
	
	Public Sub mnuEditSoundPath_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditSoundPath.Click
		
		On Error Resume Next
        Call GetActiveMdiChild.EditSoundPath()
		
	End Sub
	
	Public Sub mnuEditTitle_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuEditTitle.Click
		
		On Error Resume Next
        Call GetActiveMdiChild.EditTitle()
		
	End Sub
	
	Public Sub mnuExit_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuExit.Click
		
		On Error Resume Next
		Me.Close()
		
	End Sub
	
	Private Sub mnuFilePath_Click()
		
		Dim bWavPathUpdate As Boolean
		
		'************************************************
		'* Pass gWavPath as default directory.
		'* Show frmFilePath (vbModal means a value must
		'* be returned before execution will continue).
		'* Check to see if a directory was returned.
		'************************************************
		With frmFilePath
			.sDir = gWavPath
			.ShowDialog()
			If (Len(.sDir) > 0) Then gWavPath = .sDir
			bWavPathUpdate = True
		End With
		frmFilePath.Close()
		
		'************************************************
		'* If there was an update to the Wave file path,
		'* then we need to update the play buttons and
		'* test menu (and button, if applicable).
		'************************************************
		If bWavPathUpdate Then
			With ActiveMDIChild
				Select Case .Name
					Case "frmDispCon", "frmDispDia", "frmDispVow", "frmDispVow" 'On forms that the test button can be enabled,
						'          Call UpdatePlayBttns(.CurrIndex)
					Case Else
						Call UpdateTestMenu()
				End Select
			End With
		End If
		bWavPathUpdate = False
		
	End Sub
	
	Public Sub mnuInsert_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuInsert.Click
		Dim Index As Short = mnuInsert.GetIndex(eventSender)
		
		On Error Resume Next
		'UPGRADE_ISSUE: Control AddNewCategory could not be resolved because it was within the generic namespace ActiveMDIChild. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="084D22AD-ECB1-400F-B4C7-418ECEC5E36E"'
        Call GetActiveMdiChild.AddNewCategory(Index)
		
	End Sub
	
	Public Sub mnuIPA_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuIPA.Click
		Dim Index As Short = mnuIPA.GetIndex(eventSender)
		
		On Error Resume Next
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		
		Select Case Index
			Case 0 : frmDispCon.Show()
			Case 1 : frmDispVow.Show()
			Case 2 : frmDispDia.Show()
			Case 3 : frmDispSSeg.Show()
		End Select
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	Public Sub mnuPlayback_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPlayback.Click
		Dim Index As Short = mnuPlayback.GetIndex(eventSender)
		On Error Resume Next
		If (Index) Then
            Call GetActiveMdiChild.PlaySlow()
		Else
            Call GetActiveMdiChild.Play(0)
		End If
	End Sub
	
	Public Sub mnuPreferences_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPreferences.Click
		
		On Error Resume Next
		frmPreferences.ShowDialog()
		
	End Sub
	
	Public Sub mnuPrint_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuPrint.Click
		Dim Index As Short = mnuPrint.GetIndex(eventSender)
		
		On Error Resume Next
        Call GetActiveMdiChild.IPAHelpPrint(True, False)
		
	End Sub
	
	Public Sub mnuSILConsSILIPA_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSILConsSILIPA.Click
		
		On Error Resume Next
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		frmDispAmerCon2.Show()
		
	End Sub
	
	Public Sub mnuSILConsSILOnly_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSILConsSILOnly.Click
		
		On Error Resume Next
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		frmDispAmerCon.Show()
		
	End Sub
	
	Public Sub mnuSILVowsSILIPA_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSILVowsSILIPA.Click
		
		On Error Resume Next
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		frmDispAmerVow2.Show()
		
	End Sub
	
	Public Sub mnuSILVowsSILOnly_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuSILVowsSILOnly.Click
		
		On Error Resume Next
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		frmDispAmerVow.Show()
		
	End Sub
	
	Public Sub mnuTestSetup_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuTestSetup.Click
		
		On Error Resume Next
		frmTestSetup.ShowDialog()
		System.Windows.Forms.Application.DoEvents()
		If (bStartTest) Then Call StartTest()
		
	End Sub
	
	Public Sub mnuTestStop_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuTestStop.Click
		
		On Error Resume Next
		Call StopTest()
		
	End Sub
	
	Public Sub mnuUsing_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles mnuUsing.Click
		
		On Error Resume Next
        HtmlHelp(Me.Handle.ToInt32, ".\Help_for_IPA_Help.chm", HH_DISPLAY_TOC, 0)
		
	End Sub
	
	Private Sub TBar_ButtonClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles _TBar_Button1.Click, _TBar_Button2.Click, _TBar_Button3.Click, _TBar_Button4.Click, _TBar_Button5.Click, _TBar_Button6.Click, _TBar_Button7.Click, _TBar_Button8.Click, _TBar_Button9.Click, _TBar_Button10.Click, _TBar_Button11.Click, _TBar_Button12.Click, _TBar_Button13.Click, _TBar_Button14.Click, _TBar_Button15.Click, _TBar_Button16.Click

        Dim Button As System.Windows.Forms.ToolStripMenuItem = CType(eventSender, System.Windows.Forms.ToolStripMenuItem)
		
		On Error Resume Next
		
		Select Case Button.Name
            Case "Exit"
                If (My.Application.OpenForms.Count = 1) Then
                    Me.Close()
                Else
                    GetActiveMdiChild.Controls.Close()
                End If
            Case "PlayOnly"
                Call GetActiveMdiChild.Play(0)
            Case "PlayInterVocalic"
                Call GetActiveMdiChild.Play(1)
            Case "PlaySlow"
                Call GetActiveMdiChild.PlaySlow()
            Case "Record" : Call RecordUser((Button.Checked))
            Case "PlayRec" : Call PlayRecording((Button.Checked), False)
            Case "PlayRecSpeaker" : Call PlayRecording((Button.Checked), True)
            Case "SaveEdits"
                Call GetActiveMdiChild.EditMode(False)
            Case "Pitch"
                Call GetActiveMdiChild.ShowPitchPlot()

            Case "StopRec"
                If (gMMCtrl.Mode = MCI.ModeConstants.mciModeRecord) Then
                    Call StopRecording()
                Else
                    Call PlayRecording(False, False)
                End If

            Case "Test"
                If (Button.Checked = True) Then
                    Call mnuTestSetup_Click(mnuTestSetup, New System.EventArgs())
                Else
                    Call StopTest()
                End If
        End Select
		
	End Sub
	
	Private Sub TBar_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles TBar.MouseDown

        Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		On Error Resume Next
		
        'UPGRADE_ISSUE: MSComctlLib.Button property TBar.Buttons.Left was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        If ((X >= TBar.Items.Item("PlaySlow").ContentRectangle.Left And X < (TBar.Items.Item("PlaySlow").ContentRectangle.Left + TBar.Items.Item("PlaySlow").ContentRectangle.Width)) Or
            (X >= TBar.Items.Item("Pitch").ContentRectangle.Left And X < (TBar.Items.Item("Pitch").ContentRectangle.Left + TBar.Items.Item("Pitch").ContentRectangle.Width))) Then
            bctrlkey = ((Shift And VB6.ShiftConstants.CtrlMask) > 0)
        End If

	End Sub
	
	Private Sub TBar_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles TBar.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		Dim i As Short
		
		On Error Resume Next
		
        For i = 0 To TBar.Items.Count
            If (X >= TBar.Items.Item(i).ContentRectangle.Left And X <= (TBar.Items.Item(i).ContentRectangle.Left + TBar.Items.Item(i).ContentRectangle.Width)) Then
                gStatLine.Text = " " & TBar.Items.Item(i).Tag
                Exit Sub
            End If
        Next

        gStatLine.Text = ""

	End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		
		'************************************************
		'* This timer is used for two purposes. One is
		'* to play the selected symbol when a form is
		'* activated (after a 2 second delay). The second
		'* purpose is to delay the playing of the
		'* vocalic sound until after the double-click
		'* time as set in windows. This gives time for
		'* the double-click event to take place.
		'* Otherwise, it would never occur.
		'************************************************
		On Error Resume Next
        Call GetActiveMdiChild.Play(0)
		Timer1.Enabled = False
		
	End Sub
	
	Private Sub Timer2_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer2.Tick
		
		Dim sWavNamePart1 As String
		Dim i As Short
		
		On Error Resume Next
		
		'* See if any clicks were made. If not mark incorrect
		If (gFirstSndPlayed And Not gTestItemCorrect) Then Call MarkItemCorrect(System.Windows.Forms.CheckState.Unchecked) '* CLW 1/27/99
		
		'************************************************
		'* Close previous test item.
		'************************************************
        With GetActiveMdiChild
            'UPGRADE_WARNING: Couldn't resolve default property of object ActiveForm.lblSmile. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .lblSmile.Visible = False
            'UPGRADE_WARNING: Couldn't resolve default property of object ActiveForm.lblFrown. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            .lblFrown.Visible = False
        End With
		
		If gFirstSndPlayed Then
            Select Case GetActiveMdiChild.Name
                Case "frmDispCon"
                    With frmDispCon.Con(gItemNumber)
                        .ForeColor = System.Drawing.SystemColors.ControlText
                        .BackColor = System.Drawing.SystemColors.Control
                    End With
                Case "frmDispVow"
                    With frmDispVow.Vowel(gItemNumber)
                        .ForeColor = System.Drawing.SystemColors.ControlText
                        .BackColor = System.Drawing.SystemColors.Control
                    End With
                Case Else
            End Select
		End If
		
		If gFirstSndPlayed And Timer2.Interval = 5000 And Not gTestItemCorrect Then
			'**********************************************
			'* Show user the correct test item.
			'**********************************************
			Select Case ActiveMDIChild.Name
				Case "frmDispCon"
					With frmDispCon
						.Con(gItemNumber).ForeColor = System.Drawing.SystemColors.ControlText
						.Con(gItemNumber).BackColor = .lblSmile.BackColor
					End With
				Case "frmDispVow"
					With frmDispVow
						.Vowel(gItemNumber).ForeColor = System.Drawing.SystemColors.ControlText
						.Vowel(gItemNumber).BackColor = .lblSmile.BackColor
					End With
				Case Else
			End Select
			Timer2.Interval = 2000
		Else
			'************************************************
			'* Begin new test item.
			'************************************************
			Do 
				'get a random number between lowest and highest index
				i = Int((UBound(gTestGrp) - LBound(gTestGrp) + 1) * Rnd() + LBound(gTestGrp))
				'UPGRADE_WARNING: Couldn't resolve default property of object gTestGrp(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object gItemNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gItemNumber = gTestGrp(i)
				'* Check if valid index CLW 1/26/99
			Loop While (gTestGrpCorrect(cLastTest, i) = System.Windows.Forms.CheckState.Checked)
			
			sWavNamePart1 = VB.Right(gTestGrpName, 3)
			'UPGRADE_WARNING: Couldn't resolve default property of object gItemNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Call PlayWav(sWavNamePart1 & "-" & VB6.Format(Trim(Str(gItemNumber)), "00") & IIf(sWavNamePart1 = "Vow", "W.wav", "A.wav")) ' "A.Wav")
			'ActiveForm.Label2.Caption = gItemNumber
			gFirstSndPlayed = True
			gTestItemCorrect = False '* CLW 1/27/99
			Timer2.Interval = 5000
		End If
		
	End Sub
End Class