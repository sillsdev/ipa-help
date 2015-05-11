Option Strict Off
Option Explicit On
Friend Class frmRenameWLTitle
	Inherits System.Windows.Forms.Form
	
	Private sOldTitle As String
	
	Public ReadOnly Property Canceled() As Boolean
		Get
			
			On Error Resume Next
			Canceled = (sOldTitle = txtTitle.Text)
			
		End Get
	End Property
	
	
	Public Property Title() As String
		Get
			
			On Error Resume Next
			Title = txtTitle.Text
			
		End Get
		Set(ByVal Value As String)
			
			On Error Resume Next
			sOldTitle = Trim(Value)
			lblInfo.Text = "Rename """ & sOldTitle & """ to:"
			
			With txtTitle
				.Text = sOldTitle
				.SelectionStart = 0
				.SelectionLength = Len(.Text)
			End With
			
		End Set
	End Property
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		
		On Error Resume Next
		Visible = False
		
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		
		Dim i As Short
		Dim sNewTitle As String
		
		On Error Resume Next
		
		With txtTitle
			.Text = Trim(.Text)
			sNewTitle = .Text
			If (Len(sNewTitle) = 0) Then
				MsgBox("You must specify a title.", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, My.Application.Info.Title)
				.Focus()
				Exit Sub
			End If
		End With
		
		For i = 0 To WordListArraySize()
			'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(i, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (gWordListID(i, 0) = sOldTitle) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(i, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gWordListID(i, 0) = sNewTitle
				Call frmMenu.SetupWordListOptions()
			End If
		Next 
		
		Visible = False
		
	End Sub
	
	Private Sub frmRenameWLTitle_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		On Error Resume Next
		Call CenterForm(Me)
		
	End Sub
End Class