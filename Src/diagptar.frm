VERSION 5.00
Begin VB.Form frmDiagPtArtn 
   BorderStyle     =   0  'None
   Caption         =   "Place of Articulation"
   ClientHeight    =   5430
   ClientLeft      =   1140
   ClientTop       =   3300
   ClientWidth     =   8985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5430
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   5340
      Left            =   2230
      Picture         =   "DIAGPTAR.frx":0000
      ScaleHeight     =   5310
      ScaleMode       =   0  'User
      ScaleWidth      =   5505
      TabIndex        =   0
      Top             =   -15
      Width           =   5535
   End
End
Attribute VB_Name = "frmDiagPtArtn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
'* frmDiagPtArtn version info:
'*  See basGlobals (Globals.bas).
'**************************************************

Option Explicit

Private Const TBarButtons = "Exit;"

Public Sub IPAHelpPrint(bToPrinter As Boolean, bDummyArgument As Boolean)

  On Error Resume Next
  
  If (bToPrinter) Then
    Dim Capture As clsCapture
    Set Capture = New clsCapture
    Capture.PrintChart Picture1.Picture, "Chart:" & vbTab & vbTab & "Points of Articulation"
    Set Capture = Nothing
  Else
    Clipboard.Clear
    Clipboard.SetData Picture1.Picture, vbCFBitmap
  End If

End Sub

Private Sub Form_Activate()
  
  On Error Resume Next
  
  With mdiHelpCharts
    .ShowTBarButtons TBarButtons
    !panStatus.Visible = True
    !mnuTest.Enabled = False
    !mnuExportBitmap.Visible = True
    !mnuBkgrdColor(1).Enabled = False
    !mnuPrint(0).Visible = True
    !mnuPrint(1).Visible = True
  End With
  
  gStatLine.SimpleText = ""
  If WindowState = vbNormal Then
    Top = -Height
    Show
    WindowState = vbMaximized
  End If
  

End Sub

Private Sub Form_Deactivate()

  On Error Resume Next
  gStatLine.SimpleText = ""
  mdiHelpCharts!mnuExportBitmap.Visible = False
  mdiHelpCharts!mnuBkgrdColor(1).Enabled = True
  mdiHelpCharts!mnuPrint(0).Visible = False
  mdiHelpCharts!mnuPrint(1).Visible = False

End Sub

Private Sub Form_Load()

  mdiHelpCharts.panStatus.Visible = True
  Top = -Height
  Show
  WindowState = vbMaximized
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

  On Error Resume Next
  gStatLine.SimpleText = ""

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
  On Error Resume Next
  Call Form_Deactivate
  Set frmDiagPtArtn = Nothing

End Sub
