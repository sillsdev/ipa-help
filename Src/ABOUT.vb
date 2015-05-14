Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmHelpAbout
	Inherits System.Windows.Forms.Form
	'**************************************************
	'* frmHelpAbout version info:
	'*  See basGlobals (Globals.bas).
	'**************************************************
	
	
	Private Sub cmdHELPaboutOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdHELPaboutOK.Click
		
		Me.Close()
		
	End Sub
	
	Private Sub frmHelpAbout_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Me.Icon = mdiHelpCharts.Icon
		picHelpAbout.Image = Me.Icon.ToBitmap
		'UPGRADE_ISSUE: App object was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B85A2A7-FE9F-4FBE-AA0C-CF11AC86A305"'
        'With App
        lblHelpAboutVer.Text = "Version " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor
        If My.Application.Info.Version.Revision > 0 Then
            lblHelpAboutVer.Text = lblHelpAboutVer.Text & My.Application.Info.Version.Revision
        End If
        'lblHelpAboutVer.Caption = lblHelpAboutVer.Caption & " (" & SystemType & ")"

        '************************************************
        '* I used the LegalTrademarks property for the
        '* "JAARS - ICS..." message because I wanted if
        '* after the copyright and not with the company
        '* name (i.e. SIL).
        '************************************************
        '    lblCopyright.Caption = Trim$(.CompanyName & vbCrLf & _
        '.LegalCopyright & vbCrLf & _
        '.LegalTrademarks)
        lblCopyright.Text = Trim(My.Application.Info.Copyright)
        'End With
        With lblCopyright
            If VB6.PixelsToTwipsX(.Width) > (VB6.PixelsToTwipsX(Me.Width) - 400) Then
                .Text = VB.Left(.Text, InStr(.Text, ". ") + 1) & vbCrLf & VB.Right(.Text, Len(.Text) - InStr(.Text, ". "))
                .Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(Me.Width) - 400)
                .Left = VB6.TwipsToPixelsX(200)
                'UPGRADE_ISSUE: Label property lblCopyright.WordWrap was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                'kg .WordWrap = True
            End If
            Debug.Print(VB6.PixelsToTwipsX(.Width) & VB6.PixelsToTwipsX(.Left) & VB6.PixelsToTwipsY(.Height))
        End With

        Call CenterForm(Me)
		
	End Sub
	
	Private Sub frmHelpAbout_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		
		'UPGRADE_NOTE: Object frmHelpAbout may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Me.Dispose()
		
	End Sub
End Class