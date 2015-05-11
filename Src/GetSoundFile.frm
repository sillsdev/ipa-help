VERSION 5.00
Begin VB.Form frmGetSoundFile 
   Caption         =   "Choose a Sound File"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4050
   ControlBox      =   0   'False
   Icon            =   "GetSoundFile.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2790
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1980
      TabIndex        =   2
      Top             =   2340
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   2340
      Width           =   975
   End
   Begin VB.ComboBox cboFileTypes 
      Height          =   315
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2385
      Width           =   1830
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3915
   End
End
Attribute VB_Name = "frmGetSoundFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bCanceled As Boolean
Private sInitialFileName As String

Public Property Get Canceled() As Boolean

  On Error Resume Next
  Canceled = bCanceled
  
End Property

Public Property Get FileName() As String

  On Error Resume Next
  FileName = File1.FileName
  
End Property

Public Property Let FileName(sFileName$)

  On Error Resume Next
  sInitialFileName = sFileName

End Property

Public Property Let Folder(sFolder$)

  On Error Resume Next
  File1.Path = sFolder
  
End Property

Private Sub cboFileTypes_Click()
  
  On Error Resume Next
  If (cboFileTypes.ListIndex = 0) Then
    File1.Pattern = "*.wav"
  Else
    File1.Pattern = "*.*"
  End If

End Sub

Private Sub cmdCancel_Click()

  On Error Resume Next
  bCanceled = True
  Visible = False
  
End Sub

Private Sub cmdOK_Click()

  On Error Resume Next
  bCanceled = False
  Visible = False
  
End Sub

Private Sub File1_DblClick()

  On Error Resume Next
  Call cmdOK_Click
  
End Sub

Private Sub Form_Activate()

  '****************************************************
  '* I do this because I can't get it to work by just
  '* setting the FileListBox's FileName property to
  '* the file name.
  '****************************************************

  Dim i As Integer

  If (Len(sInitialFileName) = 0) Then Exit Sub
    
  With File1
    For i = 0 To .ListCount - 1
      If (StrComp(.List(i), sInitialFileName, vbTextCompare) = 0) Then
        .ListIndex = i
        Exit For
      End If
    Next
  End With

End Sub

Private Sub Form_Load()

  On Error Resume Next
  
  With cboFileTypes
    .AddItem "Wave Files (*.wav)"
    .AddItem "All Files (*.*)"
    .ListIndex = 0
  End With
  
  Call CenterForm(Me)
  Call cboFileTypes_Click
  File1.ListIndex = 0
  sInitialFileName = ""
  
End Sub
