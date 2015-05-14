Option Strict Off
Option Explicit On
Friend Class frmGetSoundFile
	Inherits System.Windows.Forms.Form
	
	Private bCanceled As Boolean
	Private sInitialFileName As String
	
	Public ReadOnly Property Canceled() As Boolean
		Get
			
			On Error Resume Next
			Canceled = bCanceled
			
		End Get
	End Property
	
	
	Public Property FileName() As String
		Get
			
			On Error Resume Next
			FileName = File1.FileName
			
		End Get
		Set(ByVal Value As String)
			
			On Error Resume Next
			sInitialFileName = Value
			
		End Set
	End Property
	
	Public WriteOnly Property Folder() As String
		Set(ByVal Value As String)
			
			On Error Resume Next
			File1.Path = Value
			
		End Set
	End Property
	
	'UPGRADE_WARNING: Event cboFileTypes.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub cboFileTypes_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboFileTypes.SelectedIndexChanged
		
		On Error Resume Next
		If (cboFileTypes.SelectedIndex = 0) Then
			File1.Pattern = "*.wav"
		Else
			File1.Pattern = "*.*"
		End If
		
	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		
		On Error Resume Next
		bCanceled = True
		Visible = False
		
	End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		
		On Error Resume Next
		bCanceled = False
		Visible = False
		
	End Sub
	
	Private Sub File1_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles File1.DoubleClick
		
		On Error Resume Next
		Call cmdOK_Click(cmdOK, New System.EventArgs())
		
	End Sub
	
	'UPGRADE_WARNING: Form event frmGetSoundFile.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmGetSoundFile_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		'****************************************************
		'* I do this because I can't get it to work by just
		'* setting the FileListBox's FileName property to
		'* the file name.
		'****************************************************
		
		Dim i As Short
		
		If (Len(sInitialFileName) = 0) Then Exit Sub
		
		With File1
			For i = 0 To .Items.Count - 1
				If (StrComp(.Items(i), sInitialFileName, CompareMethod.Text) = 0) Then
					.SelectedIndex = i
					Exit For
				End If
			Next 
		End With
		
	End Sub
	
	Private Sub frmGetSoundFile_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		On Error Resume Next
		
		With cboFileTypes
			.Items.Add("Wave Files (*.wav)")
			.Items.Add("All Files (*.*)")
			.SelectedIndex = 0
		End With
		
		Call CenterForm(Me)
		Call cboFileTypes_SelectedIndexChanged(cboFileTypes, New System.EventArgs())
		File1.SelectedIndex = 0
		sInitialFileName = ""
		
	End Sub
End Class