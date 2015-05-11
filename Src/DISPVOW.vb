Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmDispVow
	Inherits System.Windows.Forms.Form
	
	'**************************************************
	'* frmDispVow version info:
	'*  See basGlobals (Globals.bas).
	'**************************************************
	
	
	Public bShowTip As Boolean
	Public WavNamePart1 As String '* 'Vow' (Set in Form_Load)
	Public CurrIndex As Short
	
	Private iMouseButton As Short
	Private Const INISection As String = "Vowels"
	Private Const ShowTipEntry As String = "ShowTip"
	Private Const MaxVowels As Short = 28
	Private Const FrmMaxHeight As Short = 4035
	Private Const FrmMaxWidth As Short = 6690
	Private Const TBarButtons As String = "PlayOnly;PlaySeparator;Record;StopRec;PlayRec;PlayRecSpeaker;RecordSeparator;Test;TestSeparator;Exit;"
	Private Const statusMsg1 As String = " Click on a vowel to select it. "
	Private Const statusMsg2 As String = "Press a position button to " & "see a list of words using " & "the selected vowel."
	
	Public Sub IPAHelpPrint(ByRef bToPrinter As Boolean, ByRef bColorBackground As Boolean)
		
		Dim pic As System.Drawing.Image
		'UPGRADE_NOTE: Capture was upgraded to Capture_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Capture_Renamed As clsCapture
		
		On Error Resume Next
		pic = New System.Drawing.Bitmap(1, 1)
		Capture_Renamed = New clsCapture
		
		If (CurrIndex > -1) Then
			Vowel(CurrIndex).BackColor = System.Drawing.SystemColors.Control
			Vowel(CurrIndex).ForeColor = System.Drawing.SystemColors.ControlText
			System.Windows.Forms.Application.DoEvents()
		End If
		
		If (bToPrinter) Then
			pic = Capture_Renamed.CaptureWindowArea(Me, 50, 40, 430, 230, True)
			Capture_Renamed.PrintChart(pic, "Chart:" & vbTab & vbTab & "IPA Vowels")
		Else
			pic = Capture_Renamed.CaptureWindowArea(Me, 50, 40, 430, 230, IIf(bColorBackground, False, True))
			My.Computer.Clipboard.Clear()
			My.Computer.Clipboard.SetImage(pic)
		End If
		
		'UPGRADE_NOTE: Object Capture_Renamed may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Capture_Renamed = Nothing
		'UPGRADE_NOTE: Object pic may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pic = Nothing
		
		If (CurrIndex > -1) Then
			Vowel(CurrIndex).BackColor = System.Drawing.SystemColors.Highlight
			Vowel(CurrIndex).ForeColor = System.Drawing.SystemColors.HighlightText
		End If
		
	End Sub
	
	Public Sub Play(ByRef iButton As Short)
		
		'**************************************************
		'* This function allows mdiHelpCharts to play the
		'* wave file corresponding to the currently
		'* selected symbol, without knowing which form is
		'* active.
		'**************************************************
		
		If (Len(gMMCtrl.Tag) > 0) Then Exit Sub
		gMMCtrl.Tag = cMCIBusy
		gMMCtrl.Wait = True
		Call PlayWav(WavNamePart1 & "-" & VB6.Format(Trim(Str(CurrIndex)), "00") & "W.wav") ' Changed to "W"
		gMMCtrl.Tag = ""
		
	End Sub
	
	Private Function ReadVowelsFromDB() As Short
		
		'**************************************************
		'* This function loads the characters from the INI
		'* file and displays them on the form.
		'**************************************************
		
		Dim i As Short
        Dim sINIVals(,) As String
		
		On Error GoTo ReadVowelsFromDBErr
		ReadVowelsFromDB = False
		'UPGRADE_WARNING: Couldn't resolve default property of object GetAllINISettings(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sINIVals = GetAllINISettings(gINIPath, INISection)
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If (IsDbNull(sINIVals)) Then Exit Function
		CurrIndex = 0
		
		For i = 0 To UBound(sINIVals, 1)
			With Vowel(i)
				Call GetCharFromINIStr(sINIVals(i, 1), Vowel(i))
				.BackColor = System.Drawing.SystemColors.Control
				.ForeColor = System.Drawing.SystemColors.ControlText
			End With
		Next i
		
		ReadVowelsFromDB = True
		Exit Function
		
ReadVowelsFromDBErr: 
		
	End Function
	
	Public Sub UpdateAfterRecordAndPlayback()
		
		On Error Resume Next
		mdiHelpCharts.EnableTBarButtons(TBarButtons)
		
	End Sub
	
	Public Sub UpdateFormAfterTest()
		
		'**************************************************
		'* This function restores anything changed on the
		'* form for testing purposed (location of symbols,
		'* Visible properties, etc.).
		'**************************************************
		
		Dim i As Short
		Dim iVowelGrpInd As Short
		Dim iGrpInd As Short
		Dim iOldLeft As Short
		Dim iOldTop As Integer
		
		lblSmile.Visible = False
		lblFrown.Visible = False
		With Vowel(gItemNumber)
			.ForeColor = System.Drawing.SystemColors.ControlText
			.BackColor = System.Drawing.SystemColors.Control
		End With
		
		CurrIndex = 0
		With Vowel(CurrIndex)
			.BackColor = System.Drawing.SystemColors.Highlight
			.ForeColor = System.Drawing.SystemColors.HighlightText
		End With
		
		'**********************************************
		'* Restore Vowels to original location (if
		'* Brief test mode), Enabled and Visible state.
		'**********************************************
		iVowelGrpInd = 0
		iGrpInd = 0
		For i = 0 To MaxVowels - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object gPhonGrpColl(AllVow)(iVowelGrpInd). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If i = gPhonGrpColl.Item("AllVow")(iVowelGrpInd) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object gTestGrp(iGrpInd). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If gTestLayout = cTestBrief And i = gTestGrp(iGrpInd) Then
					If gTestGrpCorrect(cLastTest, iGrpInd) <> System.Windows.Forms.CheckState.Checked Then
						'* Restore location.  *
						With Vowel(i)
							.Visible = False
							iOldLeft = Val(Mid(.Tag, 1, 4))
							iOldTop = Val(Mid(.Tag, 5, 4))
							.Left = VB6.TwipsToPixelsX(iOldLeft)
							.Top = VB6.TwipsToPixelsY(iOldTop)
							.Tag = Mid(.Tag, 9)
						End With
					Else
						'* Enable untested vowels (in this group) with sounds *
						Vowel(i).Enabled = True
					End If
					iGrpInd = IIf(iGrpInd = UBound(gTestGrp), iGrpInd, iGrpInd + 1)
				Else
					'* Enable other vowels with sounds *
					Vowel(i).Enabled = True
				End If
				iVowelGrpInd = IIf(iVowelGrpInd = UBound(gPhonGrpColl.Item("AllVow")), iVowelGrpInd, iVowelGrpInd + 1)
			End If
		Next i
		
		If gTestLayout = cTestBrief Then
			'* Make all vowels sounds visible.  *
			For i = 0 To MaxVowels - 1
				Vowel(i).Visible = True
			Next i
		End If
		
		'**********************************************
		'* Replace the scenery if in brief mode
		'**********************************************
		If gTestLayout = cTestBrief Then
			For i = 0 To 10
				Select Case i
					Case 0 To 6
						Shape1(i).Visible = True
						Line1(i).Visible = True
						Label1(i).Visible = True
					Case 7 To 9
						Shape1(i).Visible = True
						Line1(i).Visible = True
					Case 10
						Shape1(i).Visible = True
				End Select
			Next i
		End If
		
		mdiHelpCharts.EnableTBarButtons(TBarButtons)
		
	End Sub
	
	Public Sub UpdateFormForTest()
		
		'***********************************************************************
		'* This function performs any changes to the form
		'* necessary in preparation for testing.
		'*
		'* Whether a vowel is visible or not is now determined by its status in
		'* gTestGrpCorrect. This means that gTestGrp does not change depending on
		'* how many the user got correct. This also allows the indexing to be
		'* the same for both gTestGrp and gTestGrpCorrect. CLW 1/26/99
		'***********************************************************************
		
		Dim i As Short
		Dim iTestCount As Short '* Number of vowels tested on CLW 1/26/99
		Dim iRowLength As Short
		Dim iCol As Short
		Dim iLeftStart As Short
		Dim iTop As Short
		Dim iOldLeft As Short
		Dim iOldTop As Short
		
		With Vowel(CurrIndex)
			.BackColor = System.Drawing.SystemColors.Control
			.ForeColor = System.Drawing.SystemColors.ControlText
		End With
		
		'**********************************************
		'* First disable all Vowels.
		'* If Brief mode, make invisible.
		'**********************************************
		For i = 0 To MaxVowels - 1
			With Vowel(i)
				.Enabled = False
				If gTestLayout = cTestBrief Then .Visible = False
			End With
		Next i
		
		Select Case gTestLayout
			Case cTestBrief
				'********************************************
				'* Clear the extra scenery.
				'********************************************
				For i = 0 To 10
					Select Case i
						Case 0 To 6
							Shape1(i).Visible = False
							Line1(i).Visible = False
							Label1(i).Visible = False
						Case 7 To 9
							Shape1(i).Visible = False
							Line1(i).Visible = False
						Case 10
							Shape1(i).Visible = False
					End Select
				Next i
				'********************************************
				'* Rearrange test Vowels in rows of 10,
				'* and make visible.
				'********************************************
				'* Calculate number of vowels in visible group
				For i = 0 To UBound(gTestGrp)
					If gTestGrpCorrect(cLastTest, i) <> System.Windows.Forms.CheckState.Checked Then iTestCount = iTestCount + 1
				Next i
				iCol = 0
				iRowLength = IIf(iTestCount < 10, iTestCount, 10)
				iLeftStart = (VB6.PixelsToTwipsX(Width) - iRowLength * 315) / 2
				iTop = (VB6.PixelsToTwipsY(Height) - (Int((iTestCount + 10) / 10)) * 425) / 2
				
				For i = 0 To UBound(gTestGrp)
					If gTestGrpCorrect(cLastTest, i) <> System.Windows.Forms.CheckState.Checked Then
						If iCol = 10 Then
							iTop = iTop + 425
							iCol = 0
						End If
						With Vowel(gTestGrp(i))
							iOldLeft = VB6.PixelsToTwipsX(.Left)
							iOldTop = VB6.PixelsToTwipsY(.Top)
							.Tag = VB6.Format(iOldLeft, "0000") & VB6.Format(iOldTop, "0000") & .Tag
							.Left = VB6.TwipsToPixelsX(iLeftStart + 315 * iCol)
							.Top = VB6.TwipsToPixelsY(iTop)
							.Enabled = True
							.Visible = True
							iCol = iCol + 1
						End With
					End If
				Next i
			Case cTestChart
				For i = 0 To UBound(gTestGrp)
					If gTestGrpCorrect(cLastTest, i) <> System.Windows.Forms.CheckState.Checked Then Vowel(gTestGrp(i)).Enabled = True
				Next i
		End Select
		
	End Sub
	
	'UPGRADE_WARNING: Form event frmDispVow.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmDispVow_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		On Error Resume Next
		WindowState = System.Windows.Forms.FormWindowState.Maximized
		gStatLine.Text = ""
		
		Dim vAdvice As Object
		With mdiHelpCharts
			CType(.Controls("mnuExportBitmap"), Object).Visible = True
			CType(.Controls("mnuPrint"), Object)(0).Visible = True
			CType(.Controls("mnuPrint"), Object)(1).Visible = True
			CType(.Controls("panStatus"), Object).Visible = True
			.ShowTBarButtons(TBarButtons)
			.EnableTBarButtons(TBarButtons)
			
			'**************************************************
			'* Give the user advice the first time. Optionally,
			'* (specified in the .ini file) IPAHelp can always
			'* play the selected symbol (by default, this is
			'* the first symbol on the form, when opened).
			'**************************************************
			'UPGRADE_WARNING: Couldn't resolve default property of object vAdvice. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			vAdvice = GetINIEntry(cSettingsSect, cNewUserAdviceEntry, gINIPath)
			Select Case vAdvice
				Case "1" 'Offer advice to a new user the first time only
					gMsg = "IPA Help can play audio samples for the IPA symbols." & vbCrLf & "Use the mouse pointer to select and payback an audio sample."
					MsgBox(gMsg, MsgBoxStyle.Information, My.Application.Info.Title)
					Call WriteINIEntry(cSettingsSect, cNewUserAdviceEntry, "0", gINIPath)
				Case "2" 'Always play the opening sound
					.Timer1.Interval = 2000
					.Timer1.Enabled = True
				Case Else
					'* Do nothing
			End Select
		End With
		
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		If WindowState = vbNormal Then
			Top = VB6.TwipsToPixelsY(-VB6.PixelsToTwipsY(Height))
			Show()
			WindowState = System.Windows.Forms.FormWindowState.Maximized
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Form event frmDispVow.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmDispVow_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
		
		On Error Resume Next
		gStatLine.Text = ""
		CType(mdiHelpCharts.Controls("mnuPrint"), Object)(0).Visible = False
		CType(mdiHelpCharts.Controls("mnuPrint"), Object)(1).Visible = False
		
	End Sub
	
	Private Sub frmDispVow_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		'************************************************************
		'* This routine will monitor keypresses looking for arrow
		'* keys. If one of the arrow keys is pressed then the
		'* current vowel is changed to the next or previous
		'* enabled vowel (depending upon which arrow is pressed).
		'************************************************************
		
		Dim i As Short
		If gTestActive Then Exit Sub
		On Error Resume Next
		
		If Shift > 0 Then Exit Sub
		If KeyCode = System.Windows.Forms.Keys.Return Then Call Play(0)
		
		i = CurrIndex
		If KeyCode = System.Windows.Forms.Keys.Down Or KeyCode = System.Windows.Forms.Keys.Right Then 'Is keypress right or down?
			i = IIf(CurrIndex = MaxVowels - 1, 0, CurrIndex + 1) 'Move to next vowel that is
			While Not Vowel(i).Enabled And i <> CurrIndex '  enabled or until we've
				i = IIf(i = MaxVowels - 1, 0, i + 1) '  gone full circle through
			End While '  all the vowels.
		ElseIf KeyCode = System.Windows.Forms.Keys.Up Or KeyCode = System.Windows.Forms.Keys.Left Then  'Is keypress left or up?
			i = IIf(CurrIndex = 0, MaxVowels - 1, CurrIndex - 1) 'Move to previous vowel that
			While Not Vowel(i).Enabled And i <> CurrIndex '  is enabled or until we've
				i = IIf(i = 0, MaxVowels - 1, i - 1) '  gone full circle through
			End While '  all the vowels.
		End If
		
		If Vowel(i).Enabled Then Call Vowel_MouseDown(Vowel.Item(i), New System.Windows.Forms.MouseEventArgs(&H100000, 0, 0, 0, 0)) 'If we landed on an enabled vowel,
		'  act like we clicked on it.
	End Sub
	
	Private Sub frmDispVow_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		On Error Resume Next
		If KeyAscii = System.Windows.Forms.Keys.Escape Then
			If gTestActive Then
				Call mdiHelpCharts.mnuTestStop_Click(Nothing, New System.EventArgs())
			Else
				Me.Close()
			End If
		End If
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub frmDispVow_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		mdiHelpCharts.panStatus.Visible = True
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		WavNamePart1 = "Vow"
		CurrIndex = -1
		
		If ReadVowelsFromDB() Then
			'If next line is false then there's no vowels in DB.
			If CurrIndex > -1 And gTestActive = False Then
				With Vowel(CurrIndex)
					.BackColor = System.Drawing.SystemColors.Highlight 'Select first enabled vowel.
					.ForeColor = System.Drawing.SystemColors.HighlightText
				End With
			End If
		Else
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			MsgBox("Error reading vowel information from the settings file.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, My.Application.Info.Title)
			Me.Close()
			Exit Sub
		End If
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Top = VB6.TwipsToPixelsY(-VB6.PixelsToTwipsY(Height))
		Show()
		WindowState = System.Windows.Forms.FormWindowState.Maximized
		
	End Sub
	
	Private Sub frmDispVow_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		On Error Resume Next
		If gTestActive Then gStatLine.Text = "Select correct character" Else gStatLine.Text = ""
		Cursor = System.Windows.Forms.Cursors.Default
		
	End Sub
	
	Private Sub frmDispVow_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		On Error Resume Next
		'UPGRADE_WARNING: Form event frmDispVow.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		frmDispVow_Deactivate(Me, New System.EventArgs())
		'UPGRADE_NOTE: Object frmDispVow may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Me.Dispose()
		
		eventArgs.Cancel = Cancel
	End Sub
	
	'Private Sub Form_Resize()
	
	'  On Error Resume Next
	'  If WindowState > vbNormal Then Exit Sub
	'  If Height > FrmMaxHeight Then Height = FrmMaxHeight
	'  If Width > FrmMaxWidth Then Width = FrmMaxWidth
	
	'End Sub
	
	Private Sub Vowel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Vowel.Click
		Dim Index As Short = Vowel.GetIndex(eventSender)
		
		If iMouseButton = VB6.MouseButtonConstants.LeftButton Then
			'*********************************************
			'* If in test mode ...
			'*********************************************
			If gTestActive Then
				'*********************************************
				'* ... check to see if correct IPA character
				'* selected. If so, put smile on top of
				'* character.
				'*********************************************
				'UPGRADE_WARNING: Couldn't resolve default property of object gItemNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If gItemNumber = Index Then
					lblFrown.Visible = False
					With lblSmile
						.Height = Vowel(Index).Height
						.Width = Vowel(Index).Width
						.Top = Vowel(Index).Top
						.Left = Vowel(Index).Left
						.Visible = True
					End With
					Call MarkItemCorrect(System.Windows.Forms.CheckState.Checked)
					'* Make smile visible for 2 seconds, the continue
					With mdiHelpCharts.Timer2
						.Enabled = False
						.Interval = 2000
						.Enabled = True
					End With
					'*********************
					'* If not, put frown.
					'*********************
				ElseIf lblSmile.Visible = False Then 
					With lblFrown
						.Height = Vowel(Index).Height
						.Width = Vowel(Index).Width
						.Top = Vowel(Index).Top
						.Left = Vowel(Index).Left
						.Visible = True
					End With
					Call MarkItemCorrect(System.Windows.Forms.CheckState.Unchecked, Index)
				End If
			Else
				'*********************************************
				'* If not in test mode, play associated sound.
				'*********************************************
				Call Play(0)
			End If
		End If
		
	End Sub
	
	Private Sub Vowel_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Vowel.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = Vowel.GetIndex(eventSender)
		
		'***********************************************
		'* Make the character that CurrIndex points
		'* to not look selected and make the character
		'* just clicked on look selected by changing
		'* its background color.
		'***********************************************
		
		On Error Resume Next
		If gTestActive = False Then
			With Vowel(CurrIndex)
				'* Deselect previous selection.
				.BackColor = System.Drawing.SystemColors.Control
				.ForeColor = System.Drawing.SystemColors.ControlText
			End With
			With Vowel(Index)
				'* Select new selection.
				.BackColor = System.Drawing.SystemColors.Highlight
				.ForeColor = System.Drawing.SystemColors.HighlightText
			End With
		End If
		
		CurrIndex = Index
		iMouseButton = Button
		
	End Sub
	
	Private Sub Vowel_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles Vowel.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = Vowel.GetIndex(eventSender)
		
		On Error Resume Next
		If Not gTestActive Then
			'UPGRADE_WARNING: Couldn't resolve default property of object GetCaptionFromTag(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			gStatLine.Text = GetCaptionFromTag(Vowel(Index).Tag)
		End If
		
	End Sub
End Class