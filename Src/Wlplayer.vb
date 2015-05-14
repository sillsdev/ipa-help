Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("clsWordListPlayer_NET.clsWordListPlayer")> Public Class clsWordListPlayer
	
	Private bCancel As Boolean
	
	
	Public Property Cancel() As Boolean
		Get
			
			On Error Resume Next
			Cancel = bCancel
			
		End Get
		Set(ByVal Value As Boolean)
			
			On Error Resume Next
			bCancel = Value
			
		End Set
	End Property
	
	Public Sub Play(ByVal sWavFile As String)
		
		On Error Resume Next
		bCancel = False
		Call PlayWav(sWavFile)
		
	End Sub
	
	'UPGRADE_NOTE: Class_Initialize was upgraded to Class_Initialize_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Private Sub Class_Initialize_Renamed()
		
		On Error Resume Next
		bCancel = False
		
	End Sub
	Public Sub New()
		MyBase.New()
		Class_Initialize_Renamed()
	End Sub
End Class