VERSION 5.00
Object = "{D1098D84-3ADB-11D4-99CB-E0E24AC10000}#5.0#0"; "avHyperLink.ocx"
Begin VB.Form frmMenu 
   BorderStyle     =   0  'None
   Caption         =   "Main Menu"
   ClientHeight    =   7275
   ClientLeft      =   2415
   ClientTop       =   2820
   ClientWidth     =   9105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "MAINMENU.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7275
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   Begin avHyperLink.HyperLinkLabel lnkChart 
      Height          =   360
      Index           =   4
      Left            =   5625
      TabIndex        =   5
      Top             =   900
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      AutoSize        =   -1  'True
      Caption         =   "#####"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkUpDown 
      Height          =   240
      Index           =   1
      Left            =   5895
      TabIndex        =   10
      ToolTipText     =   "Scroll Down"
      Top             =   2715
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   423
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   0
      Caption         =   "u"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkChart 
      Height          =   360
      Index           =   6
      Left            =   5625
      TabIndex        =   7
      Top             =   1620
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      AutoSize        =   -1  'True
      Caption         =   "#####"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkChart 
      Height          =   360
      Index           =   5
      Left            =   5625
      TabIndex        =   6
      Top             =   1260
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      AutoSize        =   -1  'True
      Caption         =   "#####"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkChart 
      Height          =   360
      Index           =   8
      Left            =   5625
      TabIndex        =   9
      Top             =   2340
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      AutoSize        =   -1  'True
      Caption         =   "#####"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkChart 
      Height          =   360
      Index           =   7
      Left            =   5625
      TabIndex        =   8
      Top             =   1980
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      AutoSize        =   -1  'True
      Caption         =   "#####"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkChart 
      Height          =   360
      Index           =   0
      Left            =   1530
      TabIndex        =   0
      Top             =   900
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   635
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      AutoSize        =   -1  'True
      Caption         =   "Consonants"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkChart 
      Height          =   360
      Index           =   1
      Left            =   1530
      TabIndex        =   1
      Top             =   1260
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   635
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      AutoSize        =   -1  'True
      Caption         =   "Vowels"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkChart 
      Height          =   360
      Index           =   2
      Left            =   1530
      TabIndex        =   2
      Top             =   1620
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   635
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      AutoSize        =   -1  'True
      Caption         =   "Diacritics"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkChart 
      Height          =   360
      Index           =   3
      Left            =   1530
      TabIndex        =   3
      Top             =   1980
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   635
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      AutoSize        =   -1  'True
      Caption         =   "Suprasegmentals"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkUpDown 
      Height          =   240
      Index           =   0
      Left            =   5895
      TabIndex        =   4
      ToolTipText     =   "Scroll Down"
      Top             =   645
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   423
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   0
      Caption         =   "t"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Marlett"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkChart 
      Height          =   360
      Index           =   9
      Left            =   1530
      TabIndex        =   11
      Top             =   3330
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   635
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      AutoSize        =   -1  'True
      Caption         =   "Articulators"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkChart 
      Height          =   360
      Index           =   10
      Left            =   1530
      TabIndex        =   12
      Top             =   3690
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   635
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      AutoSize        =   -1  'True
      Caption         =   "Place of Articulation"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkChart 
      Height          =   360
      Index           =   11
      Left            =   5625
      TabIndex        =   13
      Top             =   3330
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   635
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      AutoSize        =   -1  'True
      Caption         =   "Consonants"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkChart 
      Height          =   360
      Index           =   12
      Left            =   5625
      TabIndex        =   14
      Top             =   3690
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   635
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      AutoSize        =   -1  'True
      Caption         =   "Vowels"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkChart 
      Height          =   360
      Index           =   13
      Left            =   5625
      TabIndex        =   15
      Top             =   4050
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   635
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      AutoSize        =   -1  'True
      Caption         =   "Diacritics"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin avHyperLink.HyperLinkLabel lnkChart 
      Height          =   360
      Index           =   14
      Left            =   5625
      TabIndex        =   16
      Top             =   4410
      Width           =   1920
      _ExtentX        =   3387
      _ExtentY        =   635
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      AutoSize        =   -1  'True
      Caption         =   "Other Symbols"
      BeginProperty HoverFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HoverColour     =   16711680
      URL             =   "dummy"
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   5010
      Picture         =   "MAINMENU.frx":000C
      Top             =   1020
      Width           =   480
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   5580
      X2              =   5580
      Y1              =   930
      Y2              =   2685
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   4980
      Picture         =   "MAINMENU.frx":0316
      Top             =   3450
      Width           =   480
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   5580
      X2              =   5580
      Y1              =   3330
      Y2              =   4770
   End
   Begin VB.Image Image2 
      Height          =   750
      Left            =   705
      Picture         =   "MAINMENU.frx":0758
      Top             =   3330
      Width           =   750
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1485
      X2              =   1485
      Y1              =   3330
      Y2              =   4050
   End
   Begin VB.Image imgIPAIcon 
      Height          =   480
      Left            =   945
      Picture         =   "MAINMENU.frx":0BB7
      Top             =   1020
      Width           =   480
   End
   Begin VB.Image imgCon 
      Height          =   480
      Index           =   0
      Left            =   360
      Picture         =   "MAINMENU.frx":0EC1
      Top             =   2205
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1485
      X2              =   1485
      Y1              =   930
      Y2              =   2325
   End
   Begin VB.Image imgBkg 
      Height          =   1920
      Left            =   135
      Picture         =   "MAINMENU.frx":11CB
      Top             =   5085
      Visible         =   0   'False
      Width           =   5715
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************
'* frmMenu version info:
'*  See basGlobals (Globals.bas).
'**************************************************

Option Explicit

Private Const MaxWordListsVisible = 5

Private iTopWordListName As Integer

Public Sub SetupWordListOptions()

  Dim i As Integer
  Dim iIndex As Integer
  
  On Error GoTo SetupWordListOptionsErr
  
  '* Setup word list tooltips and display
  
  iIndex = iTopWordListName
  
  '**********************************************
  '* Setup word list tooltips and display
  '**********************************************
  For i = 4 To 8
    With lnkChart(i)
      If (iIndex > WordListArraySize()) Then
        .Visible = False
      ElseIf (Len(gWordListID(iIndex, 0)) > 0) Then
        .Caption = gWordListID(iIndex, 0)
        .ToolTipText = "Display " & gWordListID(iIndex, 0) & " Word List"
        .Enabled = FileExist(gWordListID(iIndex, 1))
        .Visible = True
      Else
        .Visible = False
        .Caption = ""
      End If
    End With
    iIndex = iIndex + 1
  Next i
  
  lnkUpDown(0).Visible = (WordListArraySize() >= MaxWordListsVisible)
  lnkUpDown(1).Visible = (WordListArraySize() >= MaxWordListsVisible)
 
 ' Line1(1).Y2 = lblChart(iLastList).Top + lblChart(iLastList).Height
  
  Exit Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
SetupWordListOptionsErr:
  MsgBox Err.Description & " - (SetupWordListOptions)", vbInformation, App.Title
  
End Sub

Private Sub Form_Activate()

  With mdiHelpCharts
    If (!TBar.Visible) Then !TBar.Visible = False
    If (!panStatus.Visible) Then !panStatus.Visible = False
    !mnuTest.Enabled = True
  End With
  
  Call UpdateTestMenu
  
  If WindowState = vbNormal Then
    Top = -Height
    Show
    WindowState = vbMaximized
  End If

End Sub

Private Sub Form_Load()

  Dim i As Integer
    
  On Error Resume Next
  mdiHelpCharts.panStatus.Visible = False
  
  Randomize  ' Set a new seed value
  Top = -Height
  Show
  WindowState = vbMaximized
  Call SetupWordListOptions

  iTopWordListName = 0

  With imgBkg
    For i = 0 To 14
      Set lnkChart(i).Picture = .Picture
    Next

    Set lnkUpDown(0).Picture = .Picture
    Set lnkUpDown(1).Picture = .Picture
  End With

End Sub

Private Sub Form_Paint()
  
  Dim i As Integer
  Dim j As Integer
  
  On Error Resume Next
  With imgBkg
    For i = 0 To ScaleWidth Step .Width - gPixelX
      For j = 0 To ScaleHeight Step .Height
        PaintPicture .Picture, i, j, .Width, .Height
      Next j
    Next i
  End With

End Sub

Private Sub lnkChart_Click(Index As Integer)

  Dim frm As Form
  
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  
  Select Case Index
    Case 0: frmDispCon.Show
    Case 1: frmDispVow.Show
    Case 2: frmDispDia.Show
    Case 3: frmDispSSeg.Show
    Case 9: frmDiagArtrs.Show
    Case 10: frmDiagPtArtn.Show
    Case 13: frmDispAmerDia2.Show
    Case 14: frmDispAmerOther2.Show
    
    Case 4 To 8
      For Each frm In Forms
        If frm.Tag = Str$(Index - 4) Then
          frm.Show
          Screen.MousePointer = vbDefault
          Exit Sub
        End If
      Next frm
      
      Set frm = New frmWordList
      frm.Initialize (Index - 4)
      frm.Show
    
    Case 11
      Screen.MousePointer = vbDefault
      PopupMenu mdiHelpCharts!mnuSILCons
    
    Case 12:
      Screen.MousePointer = vbDefault
      PopupMenu mdiHelpCharts!mnuSILVows
  End Select

  Screen.MousePointer = vbDefault

End Sub

Private Sub lnkUpDown_Click(Index As Integer)

  On Error Resume Next

  If (Index = 0) Then
    If (iTopWordListName = 0) Then Exit Sub
    iTopWordListName = iTopWordListName - 1
  Else
    If ((iTopWordListName + MaxWordListsVisible) > WordListArraySize()) Then Exit Sub
    iTopWordListName = iTopWordListName + 1
  End If

  Call SetupWordListOptions

End Sub
