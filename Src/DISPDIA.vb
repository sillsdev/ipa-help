Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmDispDia
	Inherits System.Windows.Forms.Form
	
	'**************************************************
	'* frmDispDia version info:
	'*  See basGlobals (Globals.bas).
	'**************************************************
	
	
	Public bShowTip As Boolean
	Public WavNamePart1 As String '* 'Dia' (Set in Form_Load)
	Public CurrIndex As Short
	
	Private iMouseButton As Short
	Private Const INISection As String = "Diacritics"
	Private Const ShowTipEntry As String = "ShowTip"
	Private Const MaxDia As Short = 46
	Private Const DiaPerCol As Short = 12
	Private Const FrmMaxHeight As Short = 5010
	Private Const FrmMaxWidth As Short = 8115
	Private Const TBarButtons As String = "PlayOnly;PlayInterVocalic;PlaySeparator;Record;StopRec;PlayRec;PlayRecSpeaker;RecordSeparator;Exit;"
	Private Const statusMsg1 As String = " Click on a diacritic to select it. "
	Private Const statusMsg2 As String = "Press a position button to " & "see a list of words using " & "the selected diacritic."
	
	Public Sub IPAHelpPrint(ByRef bToPrinter As Boolean, ByRef bColorBackground As Boolean)
		
		Dim pic As System.Drawing.Image
		'UPGRADE_NOTE: Capture was upgraded to Capture_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Capture_Renamed As clsCapture
		
		On Error Resume Next
		pic = New System.Drawing.Bitmap(1, 1)
		Capture_Renamed = New clsCapture
		
		If (CurrIndex > -1) Then
			lbldia(CurrIndex).BackColor = System.Drawing.SystemColors.Control
			lbldia(CurrIndex).ForeColor = System.Drawing.SystemColors.ControlText
			System.Windows.Forms.Application.DoEvents()
		End If
		
		If (bToPrinter) Then
			pic = Capture_Renamed.CaptureWindowArea(Me, 34, 17, 529, 301, True)
			Capture_Renamed.PrintChart(pic, "Chart:" & vbTab & vbTab & "IPA Diacritics")
		Else
			pic = Capture_Renamed.CaptureWindowArea(Me, 34, 17, 529, 301, IIf(bColorBackground, False, True))
			My.Computer.Clipboard.Clear()
			My.Computer.Clipboard.SetImage(pic)
		End If
		
		'UPGRADE_NOTE: Object Capture_Renamed may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		Capture_Renamed = Nothing
		'UPGRADE_NOTE: Object pic may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		pic = Nothing
		
		If (CurrIndex > -1) Then
			lbldia(CurrIndex).BackColor = System.Drawing.SystemColors.Highlight
			lbldia(CurrIndex).ForeColor = System.Drawing.SystemColors.HighlightText
		End If
		
	End Sub
	
	Public Sub Play(ByRef iButton As Short)
		
		'**************************************************
		'* This function allows mdiHelpCharts to play the
		'* wave file corresponding to the currently
		'* selected symbol, without knowing which form is
		'* active.
		'**************************************************
		
		On Error Resume Next
		If (Len(gMMCtrl.Tag) > 0) Then Exit Sub
		If (iButton = 1) Then CType(mdiHelpCharts.Controls("Timer1"), Object).Enabled = False
		gMMCtrl.Tag = cMCIBusy
		gMMCtrl.Wait = True
		Call PlayWav(WavNamePart1 & "-" & VB6.Format(Trim(Str(CurrIndex)), "00") & IIf(iButton = 0, "a", "b") & ".wav")
		gMMCtrl.Tag = ""
		
	End Sub
	
	Private Sub LoadTipTriggerShapes()
		
		'*******************************************************
		'* This function will load all the shapes that sit on
		'* top of each IPA diacritic that, when the mouse is
		'* moved over, will trigger the tool tip to appear. The
		'* problem with using the label control of each
		'* character is that their heights are overlap each
		'* other thus making for a confused mess when trying
		'* to decide which character to display in the tooltip.
		'*******************************************************
		
		Dim i As Short
		
		On Error Resume Next
		
		shpTipTrgr(0).BorderStyle = System.Drawing.Drawing2D.DashStyle.Custom
		
		For i = 1 To MaxDia - 1
			shpTipTrgr.Load(i)
			With shpTipTrgr(i)
				.Top = VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(Line1(i Mod DiaPerCol).Y1))
				.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Line2(i \ DiaPerCol).X1))
				.BringToFront()
				.Visible = True
			End With
		Next i
		
	End Sub
	
	Private Function ReadDiacriticsFromDB() As Short
		
		'**************************************************
		'* This function loads the characters from the INI
		'* file and displays them on the form.
		'**************************************************
		
		Dim i As Short
        Dim sINIVals(,) As String
		
		On Error GoTo ReadDiacriticsFromDBErr
		ReadDiacriticsFromDB = False
		'UPGRADE_WARNING: Couldn't resolve default property of object GetAllINISettings(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sINIVals = GetAllINISettings(gINIPath, INISection)
		'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
		If (IsDbNull(sINIVals)) Then Exit Function
		CurrIndex = 0
		
		For i = 0 To UBound(sINIVals, 1)
			With lbldia(i)
				Call GetCharFromINIStr(sINIVals(i, 1), lbldia(i))
				.BackColor = System.Drawing.SystemColors.Control
				.ForeColor = System.Drawing.SystemColors.ControlText
				.SendToBack()
			End With
		Next i
		
		ReadDiacriticsFromDB = True
		Exit Function
		
ReadDiacriticsFromDBErr: 
		
	End Function
	
	Public Sub UpdateAfterRecordAndPlayback()
		
		On Error Resume Next
		mdiHelpCharts.EnableTBarButtons(TBarButtons)
		
	End Sub
	
	'UPGRADE_WARNING: Form event frmDispDia.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmDispDia_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
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
	
	'UPGRADE_WARNING: Form event frmDispDia.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmDispDia_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
		
		On Error Resume Next
		gStatLine.Text = ""
		CType(mdiHelpCharts.Controls("mnuExportBitmap"), Object).Visible = False
		CType(mdiHelpCharts.Controls("mnuPrint"), Object)(0).Visible = False
		CType(mdiHelpCharts.Controls("mnuPrint"), Object)(1).Visible = False
		
	End Sub
	
	Private Sub frmDispDia_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		'************************************************************
		'* This routine will monitor keypresses looking for arrow
		'* keys. If one of the arrow keys is pressed then the
		'* current diacritic is changed to the next or previous
		'* enabled diacritic (depending upon which arrow is pressed).
		'************************************************************
		
		Dim i As Short
		
		On Error Resume Next
		
		If KeyCode = System.Windows.Forms.Keys.Return Then Call Play(IIf((Shift And VB6.ShiftConstants.ShiftMask) = 0, 0, 1))
		If Shift > 0 Then Exit Sub
		
		i = CurrIndex
		If KeyCode = System.Windows.Forms.Keys.Down Or KeyCode = System.Windows.Forms.Keys.Right Then 'Is keypress right or down?
			i = IIf(CurrIndex = MaxDia - 1, 0, CurrIndex + 1) 'Move to next diacritic that is
			While Not lbldia(i).Enabled And i <> CurrIndex '  enabled or until we've
				i = IIf(i = MaxDia - 1, 0, i + 1) '  gone full circle through
			End While '  all the diacritics.
		ElseIf KeyCode = System.Windows.Forms.Keys.Up Or KeyCode = System.Windows.Forms.Keys.Left Then  'Is keypress left or up?
			i = IIf(CurrIndex = 0, MaxDia - 1, CurrIndex - 1) 'Move to previous diacritic that
			While Not lbldia(i).Enabled And i <> CurrIndex '  is enabled or until we've
				i = IIf(i = 0, MaxDia - 1, i - 1) '  gone full circle through
			End While '  all the diacritics.
		End If
		
		If lbldia(i).Enabled Then Call lblDia_MouseDown(lblDia.Item(i), New System.Windows.Forms.MouseEventArgs(&H100000, 0, 0, 0, 0)) 'If we landed on an enabled diacritic,
		'  act like we clicked on it.
	End Sub
	
	Private Sub frmDispDia_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		On Error Resume Next
		If KeyAscii = System.Windows.Forms.Keys.Escape Then Me.Close()
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub frmDispDia_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		mdiHelpCharts.panStatus.Visible = True
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
		WavNamePart1 = "Dia"
		Height = VB6.TwipsToPixelsY(FrmMaxHeight)
		Width = VB6.TwipsToPixelsX(FrmMaxWidth)
		Top = 0
		Left = 0
		CurrIndex = -1
		
		If ReadDiacriticsFromDB() Then
			Call LoadTipTriggerShapes()
			If CurrIndex > -1 Then Call lblDia_MouseDown(lblDia.Item(CurrIndex), New System.Windows.Forms.MouseEventArgs(&H100000, 0, 0, 0, 0))
		Else
			'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
			System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
			MsgBox("Error reading diacritic information from the settings file.", MsgBoxStyle.Information + MsgBoxStyle.OKOnly, My.Application.Info.Title)
			Me.Close()
			Exit Sub
		End If
		
		'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
		Top = VB6.TwipsToPixelsY(-VB6.PixelsToTwipsY(Height))
		Show()
		WindowState = System.Windows.Forms.FormWindowState.Maximized
		
	End Sub
	
	Private Sub frmDispDia_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		On Error Resume Next
		gStatLine.Text = ""
		
	End Sub
	
	Private Sub frmDispDia_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		On Error Resume Next
		'UPGRADE_WARNING: Form event frmDispDia.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		frmDispDia_Deactivate(Me, New System.EventArgs())
		'UPGRADE_NOTE: Object frmDispDia may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Me.Dispose()
		
		eventArgs.Cancel = Cancel
	End Sub
	
	'Private Sub Form_Resize()
	
	'On Error Resume Next
	'If WindowState > vbNormal Then Exit Sub
	'If Height > FrmMaxHeight Then Height = FrmMaxHeight
	'If Width > FrmMaxWidth Then Width = FrmMaxWidth
	
	'End Sub
	
	Private Sub lbldia_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lbldia.Click
		Dim Index As Short = lbldia.GetIndex(eventSender)
		
		'Enable intervocalic button if intervocalic example exists
		Dim intvoc As Boolean
		Dim fileName As String
		'fileName = WavNamePart1 & "-" & Format$(Trim$(Str$(Index)), "00") & "b.wav"
		fileName = MakeInterVocName()
		'UPGRADE_WARNING: Lower bound of collection mdiHelpCharts.TBar.Buttons has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
		mdiHelpCharts.TBar.Items.Item(cInterVocBttn + 1).Enabled = FileExist(fileName)
		
		If iMouseButton = VB6.MouseButtonConstants.LeftButton Then
			'*********************************************
			'* Start double-click timer.
			'*********************************************
			'UPGRADE_WARNING: Timer property .Interval cannot have a value of 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="169ECF4A-1968-402D-B243-16603CC08604"'
			CType(mdiHelpCharts.Controls("Timer1"), Object).Interval = GetDoubleClickTime()
			CType(mdiHelpCharts.Controls("Timer1"), Object).Enabled = True
		End If
		
	End Sub
	
	Private Sub lbldia_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lbldia.DoubleClick
		Dim Index As Short = lbldia.GetIndex(eventSender)
		
		'*********************************************
		'* Double-click means not Single-click so
		'* turn the double-click timer off.
		'*********************************************
		If iMouseButton = VB6.MouseButtonConstants.LeftButton Then Call Play(1)
		
	End Sub
	
	Private Sub lblDia_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lblDia.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = lblDia.GetIndex(eventSender)
		
		On Error Resume Next
		With lbldia(CurrIndex)
			'* Deselect previous selection.
			.ForeColor = System.Drawing.SystemColors.ControlText
			.BackColor = System.Drawing.SystemColors.Control
		End With
		With lbldia(Index)
			'* Select new selection.
			.ForeColor = System.Drawing.SystemColors.HighlightText
			.BackColor = System.Drawing.SystemColors.Highlight
		End With
		
		CurrIndex = Index
		iMouseButton = Button
		
	End Sub
	
	Private Sub lbldia_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles lbldia.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		Dim Index As Short = lbldia.GetIndex(eventSender)
		
		On Error Resume Next
		'UPGRADE_WARNING: Couldn't resolve default property of object GetCaptionFromTag(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		gStatLine.Text = GetCaptionFromTag(lbldia(Index).Tag)
		
	End Sub
End Class