Option Strict Off
Option Explicit On
Friend Class frmPreferences
	Inherits System.Windows.Forms.Form
	
	Private Const NoWordList As String = "<none>"
	Private Const LVColWidINIEntry As String = "LVCol"
	Private Const NewTitle As String = "New Title"
	Private Const DefaultPhoneticFont As String = "SILDoulosIPA"
	Private Const DefaultPhoneticFontSize As Short = 12
	
	Private bGetFileSpecAfterTitleEdit As Boolean
	Private bAddExistingWLFile As Boolean
	Private sDlgTitle As String
	Private sFileSpec As String
	
	Private Const WLFileExists As Short = 0
	Private Const WLFileNew As Short = 1
	Private Const WLFileInvalid As Short = 2
	Private Const WLFileBrowseCancel As Short = 3
	Private Const WLDuplicate As Short = 4
	
	Private Sub GetSoundsPath()
		
		On Error Resume Next
		
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
			If Len(.sDir) > 0 Then gWavPath = .sDir
			txtSndsLoc.Text = .sDir
			bWavPathUpdate = True
		End With
		
		frmFilePath.Close()
		
		'************************************************
		'* If there was an update to the Wave file path,
		'* then we need to update the play buttons and
		'* test menu (and button, if applicable).
		'************************************************
		If bWavPathUpdate Then
			With mdiHelpCharts.ActiveMDIChild
				Select Case .Name
					Case "frmDispCon", "frmDispDia", "frmDispVow", "frmDispSSeg", "frmWordList"
					Case Else
						Call UpdateTestMenu()
				End Select
			End With
		End If
		
	End Sub
	
	Private Function GetWLFileSpec() As Short
		
		'***************************************************************************
		'* Returns:
		'*    WLFileExists       - File specifiedexists and is a valid XML word
		'*                         list file.
		'*    WLFileNew          - File specified is new.
		'*    WLFileInvalid      - File specified exists but is not a valid XML
		'*                         word list file.
		'*    WLFileBrowseCancel - User canceled out of the browse dialog.
		'***************************************************************************
		
		On Error GoTo GetWLFileSpecErr
		
		GetWLFileSpec = WLFileBrowseCancel
		
		Dim i As Short
        Dim xmlWL As New clsXMLWordList

        Dim dlg As New OpenFileDialog

		'UPGRADE_WARNING: CommonDialog variable was not upgraded Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="671167DC-EA81-475D-B690-7A40C7BF4A23"'
        If (lvWLLocations.FocusedItem Is Nothing) Then Exit Function
        'UPGRADE_WARNING: Lower bound of collection lvWLLocations.SelectedItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
        dlg.InitialDirectory = StripOffFileName(lvWLLocations.FocusedItem.SubItems(1).Text)
        dlg.Title = "Word List Location For " & lvWLLocations.FocusedItem.Text
        'UPGRADE_WARNING: Filter has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
        dlg.Filter = "Word List (*.xml)|*.xml|All Files (*.*)|*.*"
        dlg.FileName = ""
        dlg.CheckPathExists = True

        Dim result As DialogResult
        result = dlg.ShowDialog()
        If (result = Windows.Forms.DialogResult.Cancel) Then
            GoTo GetWLFileSpecErr
        End If

        dlg.FileName = LCase(dlg.FileName)
        cmdRemove.Enabled = True

        With lvWLLocations

            '**********************************************************
            '* Check if the specified file is already in the list of
            '* word lists.
            '**********************************************************
            For i = 1 To .Items.Count
                'UPGRADE_WARNING: Lower bound of collection lvWLLocations.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                'UPGRADE_WARNING: Lower bound of collection lvWLLocations.ListItems() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                If (.Items.Item(i).SubItems(1).Text = dlgOpen.FileName) Then
                    MsgBox(dlgOpen.FileName & " has already been chosen.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, My.Application.Info.Title)
                    GetWLFileSpec = WLDuplicate
                    Exit Function
                End If
            Next
        End With

        'UPGRADE_WARNING: Lower bound of collection lvWLLocations.SelectedItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
        If lvWLLocations.FocusedItem.SubItems.Count > 1 Then
            lvWLLocations.FocusedItem.SubItems(1).Text = dlg.FileName
        Else
            lvWLLocations.FocusedItem.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, dlg.FileName))
        End If

        '************************************************************
        '* If the user chose a file that doesn't exist then create
        '* a shell of an XML word list file with one empty category.
        '************************************************************
        If Not (FileExist(dlg.FileName)) Then
            With xmlWL
                .LoadNew((dlgOpen.FileName))
                .ID = lvWLLocations.FocusedItem.Text
                .PhoneticFontName = DefaultPhoneticFont
                .PhoneticFontSize = DefaultPhoneticFontSize
                .AddCategory("New Category")
                .Save()
            End With
            GetWLFileSpec = WLFileNew
        Else
            xmlWL.Load(dlg.FileName)
            If (xmlWL.IsFileValidWL) Then
                'UPGRADE_WARNING: Lower bound of collection lvWLLocations.SelectedItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                If lvWLLocations.FocusedItem.SubItems.Count > 1 Then
                    lvWLLocations.FocusedItem.SubItems(1).Text = dlg.FileName
                Else
                    lvWLLocations.FocusedItem.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, dlg.FileName))
                End If
                If (Len(xmlWL.ID) <> 0) Then lvWLLocations.FocusedItem.Text = xmlWL.ID
                GetWLFileSpec = WLFileExists
            Else
                MsgBox("Invalid word list file: " & dlg.FileName, MsgBoxStyle.Information + MsgBoxStyle.OkOnly, My.Application.Info.Title)
                dlg.FileName = ""
                GetWLFileSpec = WLFileInvalid
            End If

            'UPGRADE_NOTE: Object xmlWL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
            xmlWL = Nothing
        End If

        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GetWLFileSpecErr:
        Exit Function

    End Function
	
	Private Sub SelText(ByRef txtBox As System.Windows.Forms.TextBox)
		
		On Error Resume Next
		
		With txtBox
			.SelectionStart = 0
			.SelectionLength = Len(.Text)
		End With
		
	End Sub
	
	Private Function TenTimes(ByRef sngNumber As Single) As Short
		
		TenTimes = sngNumber * 10
		
	End Function
	
	Private Function TitleExists(ByRef sTitle As String, Optional ByRef vItem As Object = Nothing, Optional ByRef vShowMsg As Object = Nothing) As Boolean
		
		'**********************************************************
		'* This routine determines whether or not a title exists in
		'* the list of word lists. If an item is supplied, that
		'* item will be exempted from the check.
		'**********************************************************
		
		Dim bShowMsg As Boolean
		Dim ExemptItem As System.Windows.Forms.ListViewItem
		Dim item As System.Windows.Forms.ListViewItem
		
		On Error Resume Next
		
		TitleExists = False
		bShowMsg = True
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		'UPGRADE_WARNING: Couldn't resolve default property of object vShowMsg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Not (IsNothing(vShowMsg)) Then bShowMsg = vShowMsg
		
		'UPGRADE_NOTE: Object ExemptItem may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ExemptItem = Nothing
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If Not (IsNothing(vItem)) Then ExemptItem = vItem
		
		With lvWLLocations
			For	Each item In .Items
				If Not (item Is ExemptItem) Then
					If (StrComp(item.Text, sTitle, CompareMethod.Text) = 0) Then
						If (bShowMsg) Then MsgBox("'" & sTitle & "' already exists.", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, My.Application.Info.Title)
						TitleExists = True
						Exit Function
					End If
				End If
			Next item
		End With
		
	End Function
	
	Private Sub ValidateNumber(ByRef txtBox As System.Windows.Forms.TextBox, ByRef bMustBeInt As Boolean)
		
		On Error Resume Next
		
		With txtBox
			If (Len(.Text) = 0) Then
				.Text = CStr(0)
				Call SelText(txtBox)
			ElseIf (Not (IsNumeric(.Text)) Or (bMustBeInt And InStr(.Text, ".") > 0)) Then 
				Beep()
				.Text = .Tag
				Call SelText(txtBox)
				Exit Sub
			End If
			
			.Tag = .Text
		End With
		
	End Sub
	
	Private Function ValidateSRSpeed() As Boolean
		
		On Error Resume Next
		
		With txtSRSpeed
			If (Val(.Text) < 10 Or Val(.Text) > 333) Then
				MsgBox("The minimum and maximum values" & vbCrLf & "for this field are 10 and 333.", MsgBoxStyle.Information, My.Application.Info.Title)
				.Focus()
				Call SelText(txtSRSpeed)
				ValidateSRSpeed = False
				Exit Function
			End If
		End With
		
		ValidateSRSpeed = True
		
	End Function
	
	Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
		
		Dim i As Short
		Dim iRet As Short
		Dim sTmpTitle As String
		
		i = 0
		
		'*****************************************************
		'* Find a unique, temporary name to give the added
		'* word list.
		'*****************************************************
		Do 
			sTmpTitle = NewTitle & IIf(i = 0, "", " " & i)
			i = i + 1
		Loop While (TitleExists(sTmpTitle,  , False))
		
		With lvWLLocations
			'***************************************************
			'* Add the temp. name to the list of word lists.
			'***************************************************
			.Items.Add(sTmpTitle)
			'UPGRADE_WARNING: Lower bound of collection lvWLLocations.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			.FocusedItem = .Items.Item(.Items.Count)
			
			'***************************************************
			'* Open a file/open dialog box.
			'***************************************************
			iRet = GetWLFileSpec()
			
			'***************************************************
			'* User canceled when browsing for different or
			'* new word list file, or the file chosen exits
			'* but is invalid.
			'***************************************************
			If (iRet = WLFileBrowseCancel Or iRet = WLFileInvalid Or iRet = WLDuplicate) Then
				.Items.RemoveAt(.FocusedItem.Index)
				Exit Sub
			End If
			
			'UPGRADE_WARNING: Lower bound of collection lvWLLocations.SelectedItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			If .FocusedItem.SubItems.Count > 1 Then
				.FocusedItem.SubItems(1).Text = dlgOpen.FileName
			Else
				.FocusedItem.SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, dlgOpen.FileName))
			End If
			.Focus()
			cmdRemove.Enabled = True
			cmdBrowseWL.Enabled = (.Items.Count > 0)
			
			'***************************************************
			'* File user specified doesn't exist so be helpful
			'* and put the user automatically in the word list
			'* title's edit mode.
			'***************************************************
			If (iRet = WLFileNew) Then
				.Focus()
				.FocusedItem.BeginEdit()
				System.Windows.Forms.SendKeys.Send("{Home}+{End}")
			End If
		End With
		
	End Sub
	
	Private Sub cmdBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowse.Click
		Dim Index As Short = cmdBrowse.GetIndex(eventSender)
		
		On Error GoTo BrowseCancel
		
		If (Index = 0) Then
			Call GetSoundsPath()
			Exit Sub
        Else
            Dim dlg As OpenFileDialog

            If (Index = 0) Then
            Else
                dlg.InitialDirectory = StripOffFileName(txtSASLoc.Text)
                dlg.Title = "Speech Analyzer Server Location (Current: " & txtSASLoc.Text & ")"
                dlg.Filter = "Speech Analyzer Server (*.exe)"
                dlg.FileName = "*.exe;*.pif"
                dlg.CheckFileExists = True
                dlg.CheckPathExists = True
            End If

            Dim result As DialogResult
            result = dlg.ShowDialog()
            If (result = Windows.Forms.DialogResult.Cancel) Then
                Exit Sub
            End If

            dlg.FileName = LCase(dlg.FileName)
            txtSASLoc.Text = dlg.FileName

        End If

        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BrowseCancel:
        Exit Sub

    End Sub
	
	Private Sub cmdBrowseWL_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowseWL.Click
		
		Call GetWLFileSpec()
		
	End Sub
	
	Private Sub cmdOKCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOKCancel.Click
		Dim Index As Short = cmdOKCancel.GetIndex(eventSender)
		
		Dim i As Short
		
		On Error Resume Next
		
		If (Index = 0) Then
			If Not (ValidateSRSpeed()) Then Exit Sub
			
			With lvWLLocations
				If (.Items.Count = 0) Then
					Erase gWordListID
				Else
					ReDim gWordListID(.Items.Count - 1, 1)
					
					For i = 0 To UBound(gWordListID, 1)
						'UPGRADE_WARNING: Lower bound of collection lvWLLocations.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
						'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(i, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gWordListID(i, 0) = .Items.Item(i + 1).Text
						'UPGRADE_WARNING: Lower bound of collection lvWLLocations.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
						'UPGRADE_WARNING: Lower bound of collection lvWLLocations.ListItems() has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
						'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(i, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						gWordListID(i, 1) = .Items.Item(i + 1).SubItems(1).Text
					Next 
					
					Call ConvertWLFilesToXML()
				End If
				
				Call WriteWordListsToIni()
			End With
			
			gWavPath = txtSndsLoc.Text
			Call WriteINIEntry(cPathsSect, cSoundsEntry, gWavPath, gINIPath)
			
			gSAPath = txtSASLoc.Text
			Call WriteINIEntry(cPathsSect, cSAINIEntry, gSAPath, gINIPath)
			
			gSRSpeed = CShort(txtSRSpeed.Text)
			Call SetSlowedReplayStatLine()
			
			Call frmMenu.SetupWordListOptions()
			
			Call WriteINIEntry(cSettingsSect, cPlayInitDelayEntry, txtInitDelay.Text, gINIPath)
			Call WriteINIEntry(cSettingsSect, cRepeatCountEntry, txtRptCount.Text, gINIPath)
			Call WriteINIEntry(cSettingsSect, cPlayRepeatDelayEntry, txtRptDelay.Text, gINIPath)
			Call WriteINIEntry(cSettingsSect, cSRSpeedEntry, CStr(gSRSpeed), gINIPath)
		End If
		
		Me.Close()
		Exit Sub
		
	End Sub
	
	Private Sub cmdRemove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRemove.Click
		
		Dim i As Short
		
		On Error Resume Next
		
		With lvWLLocations
			If Not (.FocusedItem Is Nothing) Then
				.Items.RemoveAt(.FocusedItem.Index)
				If (.Items.Count > 0) Then
					'UPGRADE_WARNING: Lower bound of collection lvWLLocations.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					.FocusedItem = .Items.Item(1)
					cmdRemove.Enabled = True
				Else
					cmdRemove.Enabled = False
				End If
			End If
			
			.Focus()
			cmdBrowseWL.Enabled = (.Items.Count > 0)
		End With
		
	End Sub
	
	Private Sub frmPreferences_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim sTemp As String
		
		On Error Resume Next
		
        txtInitDelay.Text = GetINIEntry(cSettingsSect, cPlayInitDelayEntry, gINIPath)
        If Not (IsNumeric(txtInitDelay.Text)) Then txtInitDelay.Text = CStr(0)
        txtInitDelay.Tag = txtInitDelay.Text
        updnDelay(0).Text = TenTimes(CSng(txtInitDelay.Text))
        
        txtRptCount.Text = GetINIEntry(cSettingsSect, cRepeatCountEntry, gINIPath)
        If Not (IsNumeric(txtRptCount.Text)) Then txtRptCount.Text = CStr(0)
        txtRptCount.Tag = txtRptCount.Text

        txtRptDelay.Text = GetINIEntry(cSettingsSect, cPlayRepeatDelayEntry, gINIPath)
        If Not (IsNumeric(txtRptDelay.Text)) Then txtRptDelay.Text = CStr(0)
        txtRptDelay.Tag = txtRptDelay.Text
        updnDelay(1).Text = TenTimes(CSng(txtRptDelay.Text))
        
        sTemp = GetINIEntry(cSettingsSect, cSRSpeedEntry, gINIPath)
        txtSRSpeed.Text = IIf(Len(sTemp) = 0 Or Not IsNumeric(sTemp), 50, Val(sTemp))
        txtSRSpeed.Tag = txtSRSpeed.Text

        Dim vWLPaths As Object
        Dim i As Short
        Dim j As Short
        vWLPaths = GetAllINISettings(gINIPath, "WordListPaths")
        '**************************************************************
        '* Fill the word list locations from the ini file.
        '**************************************************************
        If Not (IsDBNull(vWLPaths)) Then
            j = 1
            For i = 0 To UBound(vWLPaths)
                sTemp = Trim(IIf(StrComp(Trim(vWLPaths(i, 1)), "none", 1) = 0, "", vWLPaths(i, 1)))
                If (Len(sTemp) > 0) Then
                    lvWLLocations.Items.Add(vWLPaths(i, 0))
                    If lvWLLocations.Items.Item(j).SubItems.Count > 1 Then
                        lvWLLocations.Items.Item(j).SubItems(1).Text = sTemp
                    Else
                        lvWLLocations.Items.Item(j).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, sTemp))
                    End If
                    j = j + 1
                End If
            Next
        End If

        '************************************************************
        '* Select the first item in the list and enable the browse
        '* and remove buttons if there are items in the list.
        '************************************************************
        'UPGRADE_WARNING: Lower bound of collection lvWLLocations.ListItems has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
        lvWLLocations.FocusedItem = lvWLLocations.Items.Item(1)
        cmdBrowseWL.Enabled = (lvWLLocations.Items.Count > 0)
        cmdRemove.Enabled = (lvWLLocations.Items.Count > 0)

        lvWLLocations.Focus()

        '************************************************************
        '* Set the column widths from values from the ini file.
        '************************************************************
        sTemp = GetINIEntry(cSettingsSect, LVColWidINIEntry & "1", gINIPath)
        If (Len(sTemp) > 0) Then lvWLLocations.Columns.Item(1).Width = VB6.TwipsToPixelsX(Val(sTemp))
        sTemp = GetINIEntry(cSettingsSect, LVColWidINIEntry & "2", gINIPath)
        If (Len(sTemp) > 0) Then
            lvWLLocations.Columns.Item(2).Width = VB6.TwipsToPixelsX(Val(sTemp))
        Else
            lvWLLocations.Columns.Item(2).Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(lvWLLocations.Width) - VB6.PixelsToTwipsX(lvWLLocations.Columns.Item(1).Width) - (gPixelX * 6))
        End If
        
        txtSndsLoc.Text = GetINIEntry(cPathsSect, cSoundsEntry, gINIPath)
        txtSASLoc.Text = GetINIEntry(cPathsSect, cSAINIEntry, gINIPath)

        Call CenterForm(Me, True)
        bGetFileSpecAfterTitleEdit = False
        bAddExistingWLFile = False

	End Sub
	
	Private Sub frmPreferences_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		On Error Resume Next
		
		'************************************************************
		'* Write the word list location column widths to ini file.
		'************************************************************
		With lvWLLocations
			'UPGRADE_WARNING: Lower bound of collection lvWLLocations.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			Call WriteINIEntry(cSettingsSect, LVColWidINIEntry & "1", CStr(VB6.PixelsToTwipsX(.Columns.Item(1).Width)), gINIPath)
			'UPGRADE_WARNING: Lower bound of collection lvWLLocations.ColumnHeaders has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			Call WriteINIEntry(cSettingsSect, LVColWidINIEntry & "2", CStr(VB6.PixelsToTwipsX(.Columns.Item(2).Width)), gINIPath)
		End With
		
		'UPGRADE_NOTE: Object frmPreferences may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Me.Dispose()
		
		eventArgs.Cancel = Cancel
	End Sub
	
	Private Sub lvWLLocations_AfterLabelEdit(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.LabelEditEventArgs) Handles lvWLLocations.AfterLabelEdit
		Dim Cancel As Boolean = eventArgs.CancelEdit
		Dim NewString As String = eventArgs.Label
		
		On Error Resume Next
		VB6.SetCancel(cmdOKCancel(1), True)
		
		Dim xmlWL As New clsXMLWordList
		With lvWLLocations
			'UPGRADE_WARNING: Lower bound of collection lvWLLocations.SelectedItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
			xmlWL.Load(.FocusedItem.SubItems(1).Text)
			
			If (xmlWL.IsFileValidWL()) Then
				xmlWL.ID = NewString
				xmlWL.Save()
			Else
				'UPGRADE_WARNING: Lower bound of collection lvWLLocations.SelectedItem has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				MsgBox("Invalid word list file: " & .FocusedItem.SubItems(1).Text, MsgBoxStyle.Information + MsgBoxStyle.OKOnly, My.Application.Info.Title)
				Cancel = True
			End If
			
			'UPGRADE_NOTE: Object xmlWL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
			xmlWL = Nothing
		End With
		
	End Sub
	
	Private Sub lvWLLocations_BeforeLabelEdit(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.LabelEditEventArgs) Handles lvWLLocations.BeforeLabelEdit
		Dim Cancel As Boolean = eventArgs.CancelEdit
		
		On Error Resume Next
		VB6.SetCancel(cmdOKCancel(1), False)
		
	End Sub
	
	Private Sub lvWLLocations_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lvWLLocations.DoubleClick
		
		On Error Resume Next
		Call GetWLFileSpec()
		
	End Sub
	
	Private Sub lvWLLocations_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles lvWLLocations.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		On Error Resume Next
		If (KeyCode = System.Windows.Forms.Keys.F2 And Shift = 0) Then lvWLLocations.FocusedItem.BeginEdit()
		
	End Sub
	
	'UPGRADE_WARNING: Event txtInitDelay.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtInitDelay_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInitDelay.TextChanged
		
		On Error Resume Next
		Call ValidateNumber(txtInitDelay, False)
	    updnDelay(0).Text = TenTimes(CSng(txtInitDelay.Text))
		
	End Sub
	
	Private Sub txtInitDelay_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInitDelay.Enter
		
		On Error Resume Next
		Call SelText(txtInitDelay)
		
	End Sub
	
	'UPGRADE_WARNING: Event txtRptCount.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtRptCount_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRptCount.TextChanged
		
		On Error Resume Next
		Call ValidateNumber(txtRptCount, True)
		
	End Sub
	
	Private Sub txtRptCount_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRptCount.Enter
		
		On Error Resume Next
		Call SelText(txtRptCount)
		
	End Sub
	
	'UPGRADE_WARNING: Event txtRptDelay.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtRptDelay_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRptDelay.TextChanged
		
		On Error Resume Next
		Call ValidateNumber(txtRptDelay, False)
        updnDelay(1).Text = TenTimes(CSng(txtRptDelay.Text))
		
	End Sub
	
	Private Sub txtRptDelay_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtRptDelay.Enter
		
		On Error Resume Next
		Call SelText(txtRptDelay)
		
	End Sub
	
	'UPGRADE_WARNING: Event txtSRSpeed.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub txtSRSpeed_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSRSpeed.TextChanged
		
		On Error Resume Next
		Call ValidateNumber(txtSRSpeed, False)
		
	End Sub
	
	Private Sub txtSRSpeed_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSRSpeed.Enter
		
		On Error Resume Next
		Call SelText(txtSRSpeed)
		
	End Sub
	
	Private Sub txtSRSpeed_Leave(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtSRSpeed.Leave
		
		On Error Resume Next
		Call ValidateSRSpeed()
		
	End Sub
	
    Private Sub updnDelay_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

        Dim Index As Short = updnDelay.GetIndex(eventSender)

        On Error Resume Next

        If (Index = 0) Then
            txtInitDelay.Text = IIf(updnDelay(Index).Text = 0, 0, VB6.Format(updnDelay(Index).Text / 10, "##0.0"))
        Else
            txtRptDelay.Text = IIf(updnDelay(Index).Text = 0, 0, VB6.Format(updnDelay(Index).Text / 10, "##0.0"))
        End If

    End Sub
End Class