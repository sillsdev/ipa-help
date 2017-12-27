Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class frmDispAmerCon2
	Inherits System.Windows.Forms.Form
	
	Private Const TBarButtons As String = "Exit;"
	
	Private sCaption As String ' Title bar caption
	
	'UPGRADE_WARNING: Form event frmDispAmerCon2.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmDispAmerCon2_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		Dim frm As System.Windows.Forms.Form
		
		On Error Resume Next
		
		With mdiHelpCharts
			.Text = My.Application.Info.Title & " - [" & Text & "]"
			.ShowTBarButtons(TBarButtons)
			.panStatus.Visible = True
			.mnuTest.Enabled = False '* Disable test menu.
		End With

        SetBounds(0, VB6.TwipsToPixelsY(15), 0, 0, System.Windows.Forms.BoundsSpecified.X Or System.Windows.Forms.BoundsSpecified.Y)
        gStatLine.Text = ""
		
		For	Each frm In My.Application.OpenForms
			If frm.Name <> "mdiHelpCharts" Then frm.Top = 0
		Next frm
		
		Show()
		
	End Sub
	
	'UPGRADE_WARNING: Form event frmDispAmerCon2.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmDispAmerCon2_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
		
		Dim frm As System.Windows.Forms.Form
		
		On Error Resume Next
		
		mdiHelpCharts.Text = My.Application.Info.Title
		For	Each frm In My.Application.OpenForms
			If frm.Name <> "mdiHelpCharts" Then frm.Top = VB6.TwipsToPixelsY(-VB6.PixelsToTwipsY(Height))
		Next frm
		gStatLine.Text = ""
		
	End Sub
	
	Private Sub frmDispAmerCon2_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		On Error Resume Next
		gStatLine.Text = ""
		
	End Sub
	
	Private Sub frmDispAmerCon2_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		On Error Resume Next
		'UPGRADE_WARNING: Form event frmDispAmerCon2.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		frmDispAmerCon2_Deactivate(Me, New System.EventArgs())
		'UPGRADE_NOTE: Object frmDispAmerCon2 may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Me.Dispose()
		
		eventArgs.Cancel = Cancel
	End Sub
	
	'UPGRADE_WARNING: Event frmDispAmerCon2.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub frmDispAmerCon2_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
		
		On Error Resume Next
		If mdiHelpCharts.WindowState = System.Windows.Forms.FormWindowState.Maximized Then mdiHelpCharts.panStatus.Visible = True
		
	End Sub
End Class