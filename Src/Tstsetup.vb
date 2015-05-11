Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmTestSetup
	Inherits System.Windows.Forms.Form
	'**************************************************
	'* frmTestSetup version info:
	'*  See basGlobals (Globals.bas).
	'**************************************************
	
	Private iTestLayoutTmp As Short
	Private sTestGroupTmp As String
	Private iTestGroup As Short
	Private iTestGroupTmp As Short
	Private bRetestTmp As Boolean
	
	Private Function ValidWavFilesInWordListFile(ByRef sWLFileSpec As String) As Boolean
		Dim xmlWL As Object
		
		Dim i As Short
		Dim j As Short
		Dim iCount As Short
		Dim sCategoryNames() As String
		Dim sWLSndPath As String
		Dim vList As Object
		
		On Error Resume Next
		
		ValidWavFilesInWordListFile = False
		
		'**************************************************************
		'* Load xml word list file
		'**************************************************************
		xmlWL = New clsXMLWordList
		'UPGRADE_WARNING: Couldn't resolve default property of object xmlWL.Load. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xmlWL.Load(sWLFileSpec)
		
		'**************************************************************
		'* Get the path for the wave files
		'**************************************************************
		'UPGRADE_WARNING: Couldn't resolve default property of object xmlWL.SoundPath. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sWLSndPath = xmlWL.SoundPath
		sWLSndPath = sWLSndPath & IIf(VB.Right(sWLSndPath, 1) = "\", "", "\")
		
		'**************************************************************
		'* Get the list of word list categories.
		'**************************************************************
		'UPGRADE_WARNING: Couldn't resolve default property of object xmlWL.CategoryNames. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sCategoryNames = VB6.CopyArray(xmlWL.CategoryNames)
		
		Err.Clear()
		i = UBound(sCategoryNames)
		If (Err.Number > 0) Then Exit Function
		
		iCount = 0
		
		For i = 0 To UBound(sCategoryNames)
			'UPGRADE_WARNING: Couldn't resolve default property of object xmlWL.WordsInCategory. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object vList. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vList = xmlWL.WordsInCategory(sCategoryNames(i))
			
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not (IsDbNull(vList)) Then
				For j = 0 To UBound(vList, 1)
					'UPGRADE_WARNING: Couldn't resolve default property of object vList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object MakeFullPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (FileExist(MakeFullPath(sWLSndPath, CStr(vList(j, 5))))) Then
						iCount = iCount + 1
						If (iCount = 3) Then
							ValidWavFilesInWordListFile = True
							'UPGRADE_NOTE: Object xmlWL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
							xmlWL = Nothing
							Exit Function
						End If
					End If
				Next 
			End If
		Next 
		
		'UPGRADE_NOTE: Object xmlWL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		xmlWL = Nothing
		
	End Function
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		
		'* 'Cancel' Button Click
		
		'* Set flag to tell toolbar I'm closing.
		gTestSetupActive = False '* Added by CLW 4/22/99
		
		Call mdiHelpCharts.mnuTestStop_Click(Nothing, New System.EventArgs())
		mdiHelpCharts.bStartTest = False
		Me.Close()
		
	End Sub
	
	Private Sub cmdStart_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdStart.Click
		
		'* 'Start' Button Click
		
		'***********************************************************************
		'* Whether a vowel is visible or not is now determined by its status in
		'* gTestGrpCorrect. This means that gTestGrp does not change depending on
		'* how many the user got correct. This allows the indexing to be
		'* the same for both gTestGrp and gTestGrpCorrect. CLW 1/26/99
		'*
		'* Status of tab control determines which test is started. CLW 4/29/99
		'***********************************************************************
		
		Dim i As Short
		Dim iUpper As Short
		Dim ctl As System.Windows.Forms.Control
		Dim frm As System.Windows.Forms.Form
		
		'* Set flag to tell toolbar I'm closing.
		gTestSetupActive = False '* Added by CLW 4/22/99
		
		If (SSTab1.SelectedIndex = 1) Then
			gTestForm = "frmWordList"
			For i = 0 To 4
				If (Option4(i).Checked) Then
					gTestTag = CStr(i)
					Exit For
				End If
			Next 
		Else
			'************************************************************
			'* Set up the new test group. Initialize the retest list.
			'* Reset gRetestActive to false. Store desired Test Layout.
			'************************************************************
			gTestGrpName = sTestGroupTmp
			'UPGRADE_WARNING: Couldn't resolve default property of object gPhonGrpColl(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object gTestGrp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gTestGrp = gPhonGrpColl.Item(gTestGrpName)
			'ReDim gRetestList(0)
			'gRetestListEmpty = True
			iUpper = UBound(gTestGrp)
			
			If bRetestTmp Then
				For i = 0 To UBound(gTestGrpCorrect, 2)
					'UPGRADE_WARNING: Couldn't resolve default property of object gTestGrpCorrect(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gTestGrpCorrect(cLastTest, i) = gTestGrpCorrect(cThisTest, i)
				Next i
			Else
				ReDim gTestGrpCorrect(1, iUpper)
				For i = 0 To UBound(gTestGrpCorrect, 2)
					'UPGRADE_WARNING: Couldn't resolve default property of object gTestGrpCorrect(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gTestGrpCorrect(cThisTest, i) = System.Windows.Forms.CheckState.Indeterminate
					'UPGRADE_WARNING: Couldn't resolve default property of object gTestGrpCorrect(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					gTestGrpCorrect(cLastTest, i) = System.Windows.Forms.CheckState.Indeterminate
				Next i
			End If
			
			gRetestActive = bRetestTmp
			gTestLayout = iTestLayoutTmp
			gTestForm = "frmDisp" & VB.Right(gTestGrpName, 3)
			gTestTag = ""
		End If
		
		mdiHelpCharts.bStartTest = True
		Me.Close()
		
	End Sub
	
	'UPGRADE_WARNING: Form event frmTestSetup.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmTestSetup_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		On Error Resume Next
		gTestSetupActive = True
		
	End Sub
	
	Private Sub frmTestSetup_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim i As Short
		Dim iCurrOption As Short
		
		On Error Resume Next
		KeyPreview = True
		Call CenterForm(Me)
		
		iTestLayoutTmp = gTestLayout
		sTestGroupTmp = gTestGrpName
		
		'************************************************
		'* Find name of test group from test group array.
		'* Use index of that name to select correct
		'* option button.
		'************************************************
		For i = 0 To 10
			'UPGRADE_WARNING: Couldn't resolve default property of object gPhonGrpNameArray(i). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If gPhonGrpNameArray(i) = sTestGroupTmp Then
				iTestGroup = i
			End If
		Next i
		
		If (gRetestListEmpty) Then
			Option2(iTestGroup).Checked = True
			Option3.Enabled = False
		Else
			Option3.Checked = True
		End If
		Option1(gTestLayout).Checked = True
		
		Dim bGottaValidWordList As Boolean
		If (WordListArraySize() < 0) Then
			SSTab1.TabPages.Item(1).Visible = False
		Else
			bGottaValidWordList = False
			iCurrOption = 0
			
			'************************************************
			'* Load WordList names from IPAHelp.ini.
			'* Added by CLW 4/29/99
			'************************************************
			For i = 0 To 4
				With Option4(i)
					If (Len(gWordListID(i, 0)) > 0) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(i, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						.Text = .Text & " " & gWordListID(i, 0)
						
						'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
						If (Len(gWordListID(i, 1)) = 0 Or Len(Dir(gWordListID(i, 1))) = 0) Then
							.Enabled = False
						Else
							'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							.Enabled = ValidWavFilesInWordListFile(CStr(gWordListID(i, 1)))
						End If
						
						If (WordListArraySize() >= i And .Enabled) Then
							'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(i, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
							If (mdiHelpCharts.ActiveMDIChild.Text = gWordListID(i, 0) & " Word List") Then
								iCurrOption = i
								.Checked = True
							End If
						End If
					Else
						.Visible = False
						.Text = ""
					End If
					
					If (Not bGottaValidWordList And .Enabled) Then bGottaValidWordList = True
				End With
			Next i
			
			If Not (bGottaValidWordList) Then SSTab1.TabPages.Item(1).Visible = False
		End If
		
		If (SSTab1.TabPages.Item(1).Visible And Option4(iCurrOption).Checked And Not Option4(iCurrOption).Enabled) Then
			For i = 0 To 4
				If (Option4(i).Enabled) Then
					Option4(i).Checked = True
					Exit For
				End If
			Next i
		End If
		
		'* Select Category choice method. CLW 5/11/99
		Option5(gTestCatChoice).Checked = True
		
		Select Case mdiHelpCharts.ActiveMDIChild.Name
			Case "frmWordList" : If (SSTab1.TabPages.Item(1).Visible) Then SSTab1.SelectedIndex = 1
			Case "frmDispVow" : Option2(7).Checked = True
			Case Else
		End Select
		
	End Sub
	
	Private Sub frmTestSetup_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		On Error Resume Next
		'UPGRADE_NOTE: Object frmTestSetup may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Me.Dispose()
		
		eventArgs.Cancel = Cancel
	End Sub
	
	'UPGRADE_WARNING: Event Option1.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option1_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option1.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = Option1.GetIndex(eventSender)
			
			Dim i As Short
			
			On Error Resume Next
			iTestLayoutTmp = Index
			
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option2.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option2_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = Option2.GetIndex(eventSender)
			
			On Error Resume Next
			iTestGroupTmp = Index
			'UPGRADE_WARNING: Couldn't resolve default property of object gPhonGrpNameArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sTestGroupTmp = gPhonGrpNameArray(Index)
			
		End If
	End Sub
	
	Private Sub Option2_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option2.Enter
		Dim Index As Short = Option2.GetIndex(eventSender)
		
		On Error Resume Next
		bRetestTmp = False
		Option3.Checked = False
		'UPGRADE_WARNING: Couldn't resolve default property of object gPhonGrpNameArray(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sTestGroupTmp = gPhonGrpNameArray(Index)
		
	End Sub
	
	'UPGRADE_WARNING: Event Option3.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option3_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option3.CheckedChanged
		If eventSender.Checked Then
			
			On Error Resume Next
			Option2(iTestGroup).Checked = True
			bRetestTmp = True
			
		End If
	End Sub
	
	'UPGRADE_WARNING: Event Option5.CheckedChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Option5_CheckedChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Option5.CheckedChanged
		If eventSender.Checked Then
			Dim Index As Short = Option5.GetIndex(eventSender)
			
			On Error Resume Next
			gTestCatChoice = Index '* Added by CLW 5/11/99
			
		End If
	End Sub
End Class