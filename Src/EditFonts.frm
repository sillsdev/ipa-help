VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEditFonts 
   Caption         =   "Edit Fonts"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   ControlBox      =   0   'False
   Icon            =   "EditFonts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1620
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   1500
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3660
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   1620
      TabIndex        =   7
      Top             =   240
      Width           =   3015
      Begin VB.Label lblSample 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sample"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   60
         TabIndex        =   8
         Top             =   180
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "C&hange"
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   1560
      Width           =   1395
   End
   Begin VB.ListBox lstCols 
      Height          =   1080
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   1395
   End
   Begin VB.Label lblFontSpec 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   1200
      Width           =   2955
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Word List &Column"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1755
   End
End
Attribute VB_Name = "frmEditFonts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ColFontInfoStruct
  StyleNum As Integer
  Name As String
  Size As Integer
  Bold As Boolean
  Italic As Boolean
End Type

Dim bCanceled As Boolean
Dim ColFontInfo() As ColFontInfoStruct
Dim frmWL As Form

Public Property Get Canceled() As Boolean

  On Error Resume Next
  Canceled = bCanceled
  
End Property

Private Function StyleNumber(sStyleName$)

  Dim i As Integer
  
  On Error Resume Next
  
  With frmWL!ssGrid
    For i = 0 To .StyleSets.Count - 1
      If (.StyleSets(i).Name = sStyleName) Then
        StyleNumber = i
        Exit Function
      End If
    Next
  End With
  
  StyleNumber = -1
  
End Function

Public Property Set WordlistForm(frm As Form)

  On Error Resume Next
  Set frmWL = frm
  
End Property

Private Sub cmdApply_Click()

  Dim i As Integer
  
  On Error Resume Next
  
  With lstCols
    If (.ListIndex < 0) Then Exit Sub
    i = .ListIndex
  End With
  
  With frmWL!ssGrid
    For i = 0 To UBound(ColFontInfo)
      With .StyleSets(ColFontInfo(i).StyleNum).Font
        .Name = ColFontInfo(i).Name
        .Size = ColFontInfo(i).Size
        .Bold = ColFontInfo(i).Bold
        .Italic = ColFontInfo(i).Italic
      End With
      
      frmWL.ApplyPhoneticFontStyle
      .Refresh
      cmdOK.SetFocus
      cmdApply.Enabled = False
    Next
  End With
  
End Sub

Private Sub cmdCancel_Click()

  On Error Resume Next
  bCanceled = True
  Visible = False

End Sub

Private Sub cmdChange_Click()

  Dim i As Integer
  
  On Error Resume Next
  
  With lstCols
    If (.ListIndex < 0) Then Exit Sub
    dlg.DialogTitle = "Set " & .List(.ListIndex) & " Font"
    i = .ListIndex
  End With
  
  With dlg
    .FontName = ColFontInfo(i).Name
    .FontSize = ColFontInfo(i).Size
    .FontBold = ColFontInfo(i).Bold
    .FontItalic = ColFontInfo(i).Italic
    .Flags = cdlCFScreenFonts
    On Error GoTo cmdChange_Canceled
    .CancelError = True
    .ShowFont
    ColFontInfo(i).Name = .FontName
    ColFontInfo(i).Size = .FontSize
    ColFontInfo(i).Bold = .FontBold
    ColFontInfo(i).Italic = .FontItalic
    Call lstCols_Click
    cmdApply.Enabled = True
  End With
  
cmdChange_Canceled:
    
End Sub

Private Sub cmdOK_Click()

  On Error Resume Next
  If (cmdApply.Enabled) Then Call cmdApply_Click
  bCanceled = False
  Visible = False

End Sub

Private Sub Form_Activate()
  
  Dim i As Integer
  Dim j As Integer
  
  On Error Resume Next
  
  With frmWL!ssGrid
    For i = 0 To .Columns.Count - 1
      If ((.Columns(i).Visible) And Not (.Columns(i).Caption = "WavFile")) Then
        lstCols.AddItem .Columns(i).Caption
        j = lstCols.NewIndex
        ReDim Preserve ColFontInfo(0 To j) As ColFontInfoStruct
        ColFontInfo(j).StyleNum = StyleNumber(.Columns(i).StyleSet)
        
        With .StyleSets(ColFontInfo(j).StyleNum).Font
          ColFontInfo(j).Name = .Name
          ColFontInfo(j).Size = .Size
          ColFontInfo(j).Bold = .Bold
          ColFontInfo(j).Italic = .Italic
        End With
      End If
    Next
  End With
  
  lstCols.ListIndex = 0
  
End Sub

Private Sub Form_Load()

  On Error Resume Next
  Call CenterForm(Me, True)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  On Error Resume Next
  If (UnloadMode = vbFormCode) Then Exit Sub
  Cancel = True
  Call cmdOK_Click
  
End Sub

Private Sub lstCols_Click()

  Dim i As Integer
  
  On Error Resume Next
  
  With lstCols
    If (.ListIndex < 0) Then Exit Sub
    i = .ListIndex
  End With
  
  With lblSample
    .FontName = ColFontInfo(i).Name
    .FontBold = ColFontInfo(i).Bold
    .FontItalic = ColFontInfo(i).Italic
  End With
      
  With ColFontInfo(i)
    lblFontSpec.Caption = .Name & ", " & Round(.Size) & " pt." & _
      IIf(.Bold, ", bold", "") & IIf(.Italic, ", italic", "")
  End With

End Sub
