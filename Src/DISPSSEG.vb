Option Strict Off
Option Explicit On
Friend Class frmDispSSeg
	Inherits System.Windows.Forms.Form
	
	'**************************************************
	'* frmDispSseg version info:
	'*  See basGlobals (Globals.bas).
	'**************************************************
	
	
	Public bShowTip As Boolean
	Public WavNamePart1 As String '* 'Sseg' (Set in Form_Load)
	Public CurrIndex As Short
	
	Private iMouseButton As Short
	Private Const INISection As String = "SupraSeg"
	Private Const ShowTipEntry As String = "ShowTip"
	Private Const MaxSSegs As Short = 15
	Private Const FrmMaxHeight As Short = 4080
	Private Const FrmMaxWidth As Short = 8550
	Private Const TBarButtons As String = "PlayOnly;PlaySeparator;Record;StopRec;PlayRec;PlayRecSpeaker;RecordSeparator;Exit;"
	Private Const statusMsg1 As String = " Click on a suprasegmental to select it. "
	Private Const statusMsg2 As String = "Press a position button to " & "see a list of words using " & "the selected suprasegmental."
	
	Public Sub IPAHelpPrint(ByRef bToPrinter As Boolean, ByRef bColorBackground As Boolean)
		
		Dim pic As System.Drawing.Image
		'UPGRADE_NOTE: Capture was upgraded to Capture_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Capture_Renamed As clsCapture
		
		On Error Resume Next
		pic = New System.Drawing.Bitmap(1, 1)
		Capture_Renamed = New clsCapture
		
		If (CurrIndex > -1) Then
			lblss(CurrIndex).BackColor = System.Drawing.SystemColors.Control
			lblss(CurrIndex).ForeColor = System.Drawing.SystemColors.ControlText
			System.Windows.Forms.Application.DoEvents()
		End If
		
		If (bToPrinter) Then
			pic = Capture_Renamed.CaptureWindowArea(Me, 15, 16, 560, 235, True)
			Capture_Renamed.PrintChart(pic, "Chart:" & vbTab & vbTab & "IPA Suprasegmentals, Tones and Word Accents")
		Else
			pic = Capture_Renamed.CaptureWindowArea(Me, 15, 16, 560, 235, IIf(bColorBackground, False, True))
			My.Computer.Clipboard.Clear()
			My.Computer.Clipboard.SetImage(pic)
		End If
		
		'UPGRADE_NOTE: Object Capture_Renamed may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Capture_Renamed = Nothing
		'UPGRADE_NOTE: Object pic may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pic = Nothing
		
		If (CurrIndex > -1) Then
			lblss(CurrIndex).BackColor = System.Drawing.SystemColors.Highlight
			lblss(CurrIndex).ForeColor = System.Drawing.SystemColors.HighlightText
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
		Call PlayWav(WavNamePart1 & "-" & VB6.Format(Trim(Str(CurrIndex)), "00") & "A.wav")
		gMMCtrl.Tag = ""
		
	End Sub
	
	Private Function ReadSSegsFromDB() As Short
		
		'**************************************************
		'* This function loads the characters from the INI
		'* file and displays them on the form.
		'**************************************************
		
		Dim i As Short
        Dim sINIVals(,) As String
		
		On Error GoTo ReadSSegsFromDBErr
		ReadSSegsFromDB = False
		'UPGRADE_WARNING: Couldn't resolve default property of object GetAllINISettings(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sINIVals = GetAllINISettings(gINIPath, INISection)
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If (IsDbNull(sINIVals)) Then Exit Function
		CurrIndex = 0
		
		For i = 0 To UBound(sINIVals, 1)
			With lblss(i)
				Call GetCharFromINIStr(sINIVals(i, 1), lblss(i))
				.BackColor = System.Drawing.SystemColors.Control
				.ForeColor = System.Drawing.SystemColors.ControlText
			End With
		Next i
		
		ReadSSegsFromDB = True
		Exit Function
		
ReadSSegsFromDBErr: 
		
	End Function
	
	Public Sub UpdateAfterRecordAndPlayback()
		
		On Error Resume Next
		mdiHelpCharts.EnableTBarButtons(TBarButtons)
		
	End Sub
	
	'UPGRADE_WARNING: Form event frmDispSSeg.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmDispSSeg_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		On Error Resume Next
		WindowState = System.Windows.Forms.FormWindowState.Maximized
		gStatLine.Text = ""
		
		Dim vAdvice As Object
		With mdiHelpCharts
			.panStatus.Visible = True
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
		
		CType(mdiHelpCharts.Controls("mnuExportBitmap"), Object).Visible = True
		CType(mdiHelpCharts.Controls("mnuPrint"), Object)(0).Visible = True
		CType(mdiHelpCharts.Controls("mnuPrint"), Object)(1).Visible = True
		
	End Sub
	
	'UPGRADE_WARNING: Form event frmDispSSeg.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmDispSSeg_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
		
		On Error Resume Next
		gStatLine.Text = ""
		CType(mdiHelpCharts.Controls("mnuExportBitmap"), Object).Visible = False
		CType(mdiHelpCharts.Controls("mnuPrint"), Object)(0).Visible = False
		CType(mdiHelpCharts.Controls("mnuPrint"), Object)(1).Visible = False
		
	End Sub
	
	Private Sub frmDispSSeg_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		'************************************************************
		'* This routine will monitor keypresses looking for arrow
		'* keys. If one of the arrow keys is pressed then the
		'* current vowel is changed to the next or previous
		'* enabled vowel (depending upon which arrow is pressed).
		'************************************************************
		
		Dim i As Short
		
		On Error Resume Next
		If KeyCode = System.Windows.Forms.Keys.Return Then Call Play(0)
		If Shift > 0 Then Exit Sub
		
		i = CurrIndex
		If KeyCode = System.Windows.Forms.Keys.Down Or KeyCode = System.Windows.Forms.Keys.Right Then 'Is keypress right or down?
			i = IIf(CurrIndex = MaxSSegs - 1, 0, CurrIndex + 1) 'Move to next vowel that is
			While Not lblss(i).Enabled And i <> CurrIndex '  enabled or until we've
				i = IIf(i = MaxSSegs - 1, 0, i + 1) '  gone full circle through
			End While '  all the SSegs.
		ElseIf KeyCode = System.Windows.Forms.Keys.Up Or KeyCode = System.Windows.Forms.Keys.Left Then  'Is keypress left or up?
			i = IIf(CurrIndex = 0, MaxSSegs - 1, CurrIndex - 1) 'Move to previous vowel that
			While Not lblss(i).Enabled And i <> CurrIndex '  is enabled or until we've
				i = IIf(i = 0, MaxSSegs - 1, i - 1) '  gone full circle through
			End While '  all the SSegs.
		End If
		
		If lblss(i).Enabled Then Call lblss_MouseDown(lblss.Item(i), New System.Windows.Forms.MouseEventArgs(&H100000, 0, 0, 0, 0)) 'If we landed on an enabled vowel,
		'  act like we clicked on it.
	End Sub
	
	Private Sub frmDispSSeg_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		On Error Resume Next
		If KeyAscii = System.Windows.Forms.Keys.Escape Then Me.Close()
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub frmDispSSeg_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		mdiHelpCharts.panStatus.Visible = True
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		WavNamePart1 = "Sseg"
		Height = VB6.TwipsToPixelsY(FrmMaxHeight)
		Width = VB6.TwipsToPixelsX(FrmMaxWidth)
		Top = 0
		Left = 0
		CurrIndex = -1
		
		If ReadSSegsFromDB() Then
			If CurrIndex > -1 Then 'If false then there's no SSegs in DB.
				Call lblss_MouseDown(lblss.Item(CurrIndex), New System.Windows.Forms.MouseEventArgs(&H100000, 0, 0, 0, 0))
				'With lblss(CurrIndex)
				'  .BackColor = vbHighlight       'Select first enabled vowel.
				'  .ForeColor = vbHighlightText
				'End With
			End If
		Else
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			MsgBox("Error reading suprasegmental information from the settings file.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, My.Application.Info.Title)
			Me.Close()
			Exit Sub
		End If
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Top = VB6.TwipsToPixelsY(-VB6.PixelsToTwipsY(Height))
		Show()
		WindowState = System.Windows.Forms.FormWindowState.Maximized
		
	End Sub
	
	Private Sub frmDispSSeg_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		On Error Resume Next
		gStatLine.Text = ""
		
	End Sub
	
	Private Sub frmDispSSeg_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		On Error Resume Next
		'UPGRADE_WARNING: Form event frmDispSSeg.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		frmDispSSeg_Deactivate(Me, New System.EventArgs())
		'UPGRADE_NOTE: Object frmDispSSeg may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Me.Dispose()
		
		eventArgs.Cancel = Cancel
	End Sub
	
	'Private Sub Form_Resize()
	
	'On Error Resume Next
	'If WindowState > vbNormal Then Exit Sub
	'If Height > FrmMaxHeight Then Height = FrmMaxHeight
	'If Width > FrmMaxWidth Then Width = FrmMaxWidth
	
	'End Sub
	
	Private Sub lblss_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lblss.Click
		Dim Index As Short = lblss.GetIndex(eventSender)
		
		On Error Resume Next
		If iMouseButton = VB6.MouseButtonConstants.LeftButton Then Call Play(0)
		
	End Sub
	
	Private Sub lblss_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblss.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = lblss.GetIndex(eventSender)
		
		On Error Resume Next
		With lblss(CurrIndex)
			'* Deselect previous selection.
			.BackColor = System.Drawing.SystemColors.Control
			.ForeColor = System.Drawing.SystemColors.ControlText
		End With
		With lblss(Index)
			'* Select new selection.
			.BackColor = System.Drawing.SystemColors.Highlight
			.ForeColor = System.Drawing.SystemColors.HighlightText
		End With
		CurrIndex = Index
		iMouseButton = Button
		
	End Sub
	
	Private Sub lblss_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblss.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = lblss.GetIndex(eventSender)
		
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object GetCaptionFromTag(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gStatLine.Text = GetCaptionFromTag(lblss(Index).Tag)
		
	End Sub
End Class