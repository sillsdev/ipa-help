Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmDispAmerDia2
	Inherits System.Windows.Forms.Form
	'**************************************************
	'* frmDispSILAmerDia version info:
	'*  See basGlobals (Globals.bas).
	'**************************************************
	
	
	Dim CurrIndex As Short
	Dim CharDesc() As String
	
	Private Const TBarButtons As String = "Exit;"
	Private Const MaxVowels As Short = 24
	Private Const FrmMaxHeight As Short = 4020
	Private Const FrmMaxWidth As Short = 7830
	
	'UPGRADE_WARNING: Form event frmDispAmerDia2.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmDispAmerDia2_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		On Error Resume Next
		With mdiHelpCharts
			.ShowTBarButtons(TBarButtons)
			.panStatus.Visible = True
			.mnuTest.Enabled = False '* Disable test menu.
		End With
		
		gStatLine.Text = ""
		
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		If WindowState = vbNormal Then
			Top = VB6.TwipsToPixelsY(-VB6.PixelsToTwipsY(Height))
			Show()
			WindowState = System.Windows.Forms.FormWindowState.Maximized
		End If
		
	End Sub
	
	'UPGRADE_WARNING: Form event frmDispAmerDia2.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmDispAmerDia2_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
		
		On Error Resume Next
		gStatLine.Text = ""
		
	End Sub
	
	Private Sub frmDispAmerDia2_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		On Error Resume Next
		Top = VB6.TwipsToPixelsY(-VB6.PixelsToTwipsY(Height))
		Show()
		mdiHelpCharts.panStatus.Visible = True
		
	End Sub
	
	Private Sub frmDispAmerDia2_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		On Error Resume Next
		gStatLine.Text = ""
		
	End Sub
	
	Private Sub frmDispAmerDia2_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		On Error Resume Next
		'UPGRADE_WARNING: Form event frmDispAmerDia2.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		frmDispAmerDia2_Deactivate(Me, New System.EventArgs())
		'UPGRADE_NOTE: Object frmDispAmerDia2 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Me.Dispose()
		Erase CharDesc
		
		eventArgs.Cancel = Cancel
	End Sub
	
	'UPGRADE_WARNING: Event frmDispAmerDia2.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmDispAmerDia2_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		
		On Error Resume Next
		'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
		If WindowState > vbNormal Then Exit Sub
		If VB6.PixelsToTwipsY(Height) > FrmMaxHeight Then Height = VB6.TwipsToPixelsY(FrmMaxHeight)
		If VB6.PixelsToTwipsX(Width) > FrmMaxWidth Then Width = VB6.TwipsToPixelsX(FrmMaxWidth)
		
	End Sub
End Class