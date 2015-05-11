VERSION 5.00
Begin VB.Form frmRenameWLTitle 
   Caption         =   "Rename Word List Title"
   ClientHeight    =   1215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3210
   ControlBox      =   0   'False
   Icon            =   "RenameWLTitle.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1215
   ScaleWidth      =   3210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   780
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1140
      TabIndex        =   2
      Top             =   780
      Width           =   975
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   3075
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "#"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   105
   End
End
Attribute VB_Name = "frmRenameWLTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sOldTitle As String

Public Property Get Canceled() As Boolean

  On Error Resume Next
  Canceled = (sOldTitle = txtTitle.Text)
  
End Property

Public Property Get Title() As String

  On Error Resume Next
  Title = txtTitle.Text
  
End Property

Public Property Let Title(sTitle$)

  On Error Resume Next
  sOldTitle = Trim$(sTitle)
  lblInfo.Caption = "Rename """ & sOldTitle & """ to:"
  
  With txtTitle
    .Text = sOldTitle
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
  
End Property

Private Sub cmdCancel_Click()

  On Error Resume Next
  Visible = False
  
End Sub

Private Sub cmdOK_Click()

  Dim i As Integer
  Dim sNewTitle As String
  
  On Error Resume Next
  
  With txtTitle
    .Text = Trim$(.Text)
    sNewTitle = .Text
    If (Len(sNewTitle) = 0) Then
      MsgBox "You must specify a title.", vbOKOnly + vbInformation, App.Title
      .SetFocus
      Exit Sub
    End If
  End With
  
  For i = 0 To WordListArraySize()
    If (gWordListID(i, 0) = sOldTitle) Then
      gWordListID(i, 0) = sNewTitle
      Call frmMenu.SetupWordListOptions
    End If
  Next
  
  Visible = False

End Sub

Private Sub Form_Load()

  On Error Resume Next
  Call CenterForm(Me)
  
End Sub
