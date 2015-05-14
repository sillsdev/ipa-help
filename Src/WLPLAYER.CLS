VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWordListPlayer"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Private bCancel As Boolean

Public Property Get Cancel() As Boolean

  On Error Resume Next
  Cancel = bCancel
  
End Property

Public Property Let Cancel(bVal As Boolean)

  On Error Resume Next
  bCancel = bVal
  
End Property

Public Sub Play(ByVal sWavFile$)

  On Error Resume Next
  bCancel = False
  Call PlayWav(sWavFile)

End Sub

Private Sub Class_Initialize()

  On Error Resume Next
  bCancel = False
  
End Sub
