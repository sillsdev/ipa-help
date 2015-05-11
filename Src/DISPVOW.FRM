VERSION 5.00
Begin VB.Form frmDispVow 
   BorderStyle     =   0  'None
   Caption         =   "IPA Vowels"
   ClientHeight    =   5430
   ClientLeft      =   8175
   ClientTop       =   2055
   ClientWidth     =   8985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5430
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblFrown 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4215
      TabIndex        =   36
      Top             =   4440
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label lblSmile 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3645
      TabIndex        =   35
      Top             =   4440
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1890
      X2              =   3570
      Y1              =   990
      Y2              =   3810
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   7
      Left            =   3555
      Shape           =   3  'Circle
      Top             =   3810
      Width           =   75
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   10
      Left            =   2550
      TabIndex        =   13
      Top             =   1755
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   9
      Left            =   2070
      TabIndex        =   14
      Top             =   1755
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ö"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   4185
      TabIndex        =   26
      Top             =   855
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   1590
      TabIndex        =   0
      Top             =   855
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   2010
      TabIndex        =   34
      Top             =   855
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   8
      Left            =   5820
      TabIndex        =   8
      Top             =   1260
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   3030
      TabIndex        =   11
      Top             =   1260
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   7
      Left            =   3345
      TabIndex        =   10
      Top             =   1260
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   24
      Left            =   3210
      TabIndex        =   16
      Top             =   3660
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "¯"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   25
      Left            =   3630
      TabIndex        =   17
      Top             =   3660
      Width           =   315
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   14
      Left            =   6945
      TabIndex        =   12
      Top             =   1755
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   13
      Left            =   6585
      TabIndex        =   9
      Top             =   1755
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "‚"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   11
      Left            =   4395
      TabIndex        =   30
      Top             =   1755
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "¿"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   17
      Left            =   3105
      TabIndex        =   18
      Top             =   2700
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   16
      Left            =   2640
      TabIndex        =   15
      Top             =   2700
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   22
      Left            =   2955
      TabIndex        =   19
      Top             =   3225
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "µ"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   6555
      TabIndex        =   25
      Top             =   855
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "«"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   15
      Left            =   4740
      TabIndex        =   31
      Top             =   2205
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Œ"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   23
      Left            =   5010
      TabIndex        =   32
      Top             =   3195
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Ï"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   19
      Left            =   5070
      TabIndex        =   28
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   26
      Left            =   6585
      TabIndex        =   23
      Top             =   3660
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Ã"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   20
      Left            =   6585
      TabIndex        =   21
      Top             =   2700
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Î"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   18
      Left            =   4680
      TabIndex        =   33
      Top             =   2700
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   ""
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   27
      Left            =   6945
      TabIndex        =   24
      Top             =   3660
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   ""
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   21
      Left            =   6945
      TabIndex        =   22
      Top             =   2700
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   6945
      TabIndex        =   20
      Top             =   855
      Width           =   270
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   2310
      X2              =   4120
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   2850
      X2              =   4370
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   5340
      X2              =   6520
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   4875
      X2              =   6465
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   5085
      X2              =   6525
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   3420
      X2              =   4630
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Front"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   4
      Left            =   1650
      TabIndex        =   7
      Top             =   615
      Width           =   405
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   810
      TabIndex        =   6
      Top             =   945
      Width           =   780
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Close-mid"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   810
      TabIndex        =   5
      Top             =   1845
      Width           =   780
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Open-mid"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   825
      TabIndex        =   4
      Top             =   2790
      Width           =   780
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   825
      TabIndex        =   3
      Top             =   3735
      Width           =   780
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Central"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   5
      Left            =   4200
      TabIndex        =   2
      Top             =   615
      Width           =   555
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   6
      Left            =   6690
      TabIndex        =   1
      Top             =   615
      Width           =   405
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   3930
      X2              =   6520
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   4485
      X2              =   5265
      Y1              =   1020
      Y2              =   3840
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   6885
      X2              =   6885
      Y1              =   1020
      Y2              =   3840
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   3
      Left            =   2430
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   0
      Left            =   1875
      Shape           =   3  'Circle
      Top             =   990
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   5
      Left            =   2985
      Shape           =   3  'Circle
      Top             =   2850
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   9
      Left            =   4710
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   10
      Left            =   4965
      Shape           =   3  'Circle
      Top             =   2850
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   1
      Left            =   4455
      Shape           =   3  'Circle
      Top             =   990
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   6
      Left            =   6855
      Shape           =   3  'Circle
      Top             =   2850
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   4
      Left            =   6855
      Shape           =   3  'Circle
      Top             =   1920
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   8
      Left            =   6855
      Shape           =   3  'Circle
      Top             =   3810
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   2
      Left            =   6855
      Shape           =   3  'Circle
      Top             =   990
      Width           =   75
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   12
      Left            =   4785
      TabIndex        =   29
      Top             =   1755
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "¬"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   4530
      TabIndex        =   27
      Top             =   855
      Width           =   270
   End
End
Attribute VB_Name = "frmDispVow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************
'* frmDispVow version info:
'*  See basGlobals (Globals.bas).
'**************************************************

Option Explicit

Public bShowTip As Boolean
Public WavNamePart1 As String                       '* 'Vow' (Set in Form_Load)
Public CurrIndex As Integer

Private iMouseButton As Integer
Private Const INISection = "Vowels"
Private Const ShowTipEntry = "ShowTip"
Private Const MaxVowels = 28
Private Const FrmMaxHeight = 4035
Private Const FrmMaxWidth = 6690
Private Const TBarButtons = "PlayOnly;PlaySeparator;Record;StopRec;PlayRec;PlayRecSpeaker;RecordSeparator;Test;TestSeparator;Exit;"
Private Const statusMsg1 = " Click on a vowel to select it. "
Private Const statusMsg2 = "Press a position button to " & _
                           "see a list of words using " & _
                           "the selected vowel."

Public Sub IPAHelpPrint(bToPrinter As Boolean, bColorBackground As Boolean)

  Dim pic As Picture
  Dim Capture As clsCapture
  
  On Error Resume Next
  Set pic = New StdPicture
  Set Capture = New clsCapture
  
  If (CurrIndex > -1) Then
    Vowel(CurrIndex).BackColor = vbButtonFace
    Vowel(CurrIndex).ForeColor = vbButtonText
    DoEvents
  End If
  
  If (bToPrinter) Then
    Set pic = Capture.CaptureWindowArea(Me, 50, 40, 430, 230, True)
    Capture.PrintChart pic, "Chart:" & vbTab & vbTab & "IPA Vowels"
  Else
    Set pic = Capture.CaptureWindowArea(Me, 50, 40, 430, 230, _
              IIf(bColorBackground, False, True))
    Clipboard.Clear
    Clipboard.SetData pic, 2
  End If
  
  Set Capture = Nothing
  Set pic = Nothing
  
  If (CurrIndex > -1) Then
    Vowel(CurrIndex).BackColor = vbHighlight
    Vowel(CurrIndex).ForeColor = vbHighlightText
  End If
  
End Sub

Public Sub Play(iButton As Integer)

'**************************************************
'* This function allows mdiHelpCharts to play the
'* wave file corresponding to the currently
'* selected symbol, without knowing which form is
'* active.
'**************************************************

  If (Len(gMMCtrl.Tag) > 0) Then Exit Sub
  gMMCtrl.Tag = cMCIBusy
  gMMCtrl.Wait = True
  Call PlayWav(WavNamePart1 & "-" & Format$(Trim$(Str$(CurrIndex)), "00") & "W.wav") ' Changed to "W"
  gMMCtrl.Tag = ""
  
End Sub

Private Function ReadVowelsFromDB() As Integer
  
'**************************************************
'* This function loads the characters from the INI
'* file and displays them on the form.
'**************************************************
  
  Dim i As Integer
  Dim sINIVals() As String
  
  On Error GoTo ReadVowelsFromDBErr
  ReadVowelsFromDB = False
  sINIVals = GetAllINISettings(gINIPath, INISection)
  If (IsNull(sINIVals)) Then Exit Function
  CurrIndex = 0
  
  For i = 0 To UBound(sINIVals, 1)
    With Vowel(i)
      Call GetCharFromINIStr(sINIVals(i, 1), Vowel(i))
      .BackColor = vbButtonFace
      .ForeColor = vbButtonText
    End With
  Next i
  
  ReadVowelsFromDB = True
  Exit Function

ReadVowelsFromDBErr:
  
End Function

Public Sub UpdateAfterRecordAndPlayback()

  On Error Resume Next
  mdiHelpCharts.EnableTBarButtons TBarButtons

End Sub

Public Sub UpdateFormAfterTest()
  
'**************************************************
'* This function restores anything changed on the
'* form for testing purposed (location of symbols,
'* Visible properties, etc.).
'**************************************************

  Dim i As Integer
  Dim iVowelGrpInd As Integer
  Dim iGrpInd As Integer
  Dim iOldLeft As Integer
  Dim iOldTop As Long

  lblSmile.Visible = False
  lblFrown.Visible = False
  With Vowel(gItemNumber)
    .ForeColor = vbButtonText
    .BackColor = vbButtonFace
  End With
  
  CurrIndex = 0
  With Vowel(CurrIndex)
    .BackColor = vbHighlight
    .ForeColor = vbHighlightText
  End With
  
  '**********************************************
  '* Restore Vowels to original location (if
  '* Brief test mode), Enabled and Visible state.
  '**********************************************
  iVowelGrpInd = 0
  iGrpInd = 0
  For i = 0 To MaxVowels - 1
    If i = gPhonGrpColl("AllVow")(iVowelGrpInd) Then
      If gTestLayout = cTestBrief _
      And i = gTestGrp(iGrpInd) Then
        If gTestGrpCorrect(cLastTest, iGrpInd) <> vbChecked Then
          '* Restore location.  *
          With Vowel(i)
            .Visible = False
            iOldLeft = Val(Mid$(.Tag, 1, 4))
            iOldTop = Val(Mid$(.Tag, 5, 4))
            .Left = iOldLeft
            .Top = iOldTop
            .Tag = Mid(.Tag, 9)
          End With
        Else
          '* Enable untested vowels (in this group) with sounds *
          Vowel(i).Enabled = True
        End If
        iGrpInd = IIf(iGrpInd = UBound(gTestGrp), _
                      iGrpInd, iGrpInd + 1)
      Else
        '* Enable other vowels with sounds *
        Vowel(i).Enabled = True
      End If
      iVowelGrpInd = IIf(iVowelGrpInd = UBound(gPhonGrpColl("AllVow")), _
                         iVowelGrpInd, iVowelGrpInd + 1)
    End If
  Next i
  
  If gTestLayout = cTestBrief Then
    '* Make all vowels sounds visible.  *
    For i = 0 To MaxVowels - 1
      Vowel(i).Visible = True
    Next i
  End If

  '**********************************************
  '* Replace the scenery if in brief mode
  '**********************************************
  If gTestLayout = cTestBrief Then
    For i = 0 To 10
      Select Case i
        Case 0 To 6
          Shape1(i).Visible = True
          Line1(i).Visible = True
          Label1(i).Visible = True
        Case 7 To 9
          Shape1(i).Visible = True
          Line1(i).Visible = True
        Case 10
          Shape1(i).Visible = True
      End Select
    Next i
  End If
  
  mdiHelpCharts.EnableTBarButtons TBarButtons
  
End Sub

Public Sub UpdateFormForTest()
  
'***********************************************************************
'* This function performs any changes to the form
'* necessary in preparation for testing.
'*
'* Whether a vowel is visible or not is now determined by its status in
'* gTestGrpCorrect. This means that gTestGrp does not change depending on
'* how many the user got correct. This also allows the indexing to be
'* the same for both gTestGrp and gTestGrpCorrect. CLW 1/26/99
'***********************************************************************
  
  Dim i As Integer
  Dim iTestCount As Integer      '* Number of vowels tested on CLW 1/26/99
  Dim iRowLength As Integer
  Dim iCol As Integer
  Dim iLeftStart As Integer
  Dim iTop As Integer
  Dim iOldLeft As Integer
  Dim iOldTop As Integer
  
  With Vowel(CurrIndex)
    .BackColor = vbButtonFace
    .ForeColor = vbButtonText
  End With
  
  '**********************************************
  '* First disable all Vowels.
  '* If Brief mode, make invisible.
  '**********************************************
  For i = 0 To MaxVowels - 1
    With Vowel(i)
      .Enabled = False
      If gTestLayout = cTestBrief _
        Then .Visible = False
    End With
  Next i
  
  Select Case gTestLayout
    Case cTestBrief
      '********************************************
      '* Clear the extra scenery.
      '********************************************
      For i = 0 To 10
        Select Case i
          Case 0 To 6
            Shape1(i).Visible = False
            Line1(i).Visible = False
            Label1(i).Visible = False
          Case 7 To 9
            Shape1(i).Visible = False
            Line1(i).Visible = False
          Case 10
            Shape1(i).Visible = False
        End Select
      Next i
      '********************************************
      '* Rearrange test Vowels in rows of 10,
      '* and make visible.
      '********************************************
      '* Calculate number of vowels in visible group
      For i = 0 To UBound(gTestGrp)
        If gTestGrpCorrect(cLastTest, i) <> vbChecked Then _
          iTestCount = iTestCount + 1
      Next i
      iCol = 0
      iRowLength = IIf(iTestCount < 10, iTestCount, 10)
      iLeftStart = (Width - iRowLength * 315) / 2
      iTop = (Height - (Int((iTestCount + 10) / 10)) * 425) / 2
      
      For i = 0 To UBound(gTestGrp)
        If gTestGrpCorrect(cLastTest, i) <> vbChecked Then
          If iCol = 10 Then
            iTop = iTop + 425
            iCol = 0
          End If
          With Vowel(gTestGrp(i))
            iOldLeft = .Left
            iOldTop = .Top
            .Tag = Format(iOldLeft, "0000") & Format(iOldTop, "0000") & .Tag
            .Left = iLeftStart + 315 * iCol
            .Top = iTop
            .Enabled = True
            .Visible = True
            iCol = iCol + 1
          End With
        End If
      Next i
    Case cTestChart
      For i = 0 To UBound(gTestGrp)
        If gTestGrpCorrect(cLastTest, i) <> vbChecked Then _
          Vowel(gTestGrp(i)).Enabled = True
      Next i
  End Select
  
End Sub

Private Sub Form_Activate()
  
  On Error Resume Next
  WindowState = vbMaximized
  gStatLine.SimpleText = ""
  
  With mdiHelpCharts
    !mnuExportBitmap.Visible = True
    !mnuPrint(0).Visible = True
    !mnuPrint(1).Visible = True
    !panStatus.Visible = True
    .ShowTBarButtons TBarButtons
    .EnableTBarButtons TBarButtons
    
    '**************************************************
    '* Give the user advice the first time. Optionally,
    '* (specified in the .ini file) IPAHelp can always
    '* play the selected symbol (by default, this is
    '* the first symbol on the form, when opened).
    '**************************************************
    Dim vAdvice As Variant
    vAdvice = GetINIEntry(cSettingsSect, cNewUserAdviceEntry, gINIPath)
    Select Case vAdvice
      Case "1" 'Offer advice to a new user the first time only
        gMsg = "IPA Help can play audio samples for the IPA symbols." + vbCrLf + _
               "Use the mouse pointer to select and payback an audio sample."
        MsgBox gMsg, vbInformation, App.Title
        Call WriteINIEntry(cSettingsSect, cNewUserAdviceEntry, "0", gINIPath)
      Case "2" 'Always play the opening sound
        .Timer1.Interval = 2000
        .Timer1.Enabled = True
      Case Else
        '* Do nothing
    End Select
  End With
  
  If WindowState = vbNormal Then
    Top = -Height
    Show
    WindowState = vbMaximized
  End If
  
End Sub

Private Sub Form_Deactivate()
  
  On Error Resume Next
  gStatLine.SimpleText = ""
  mdiHelpCharts!mnuPrint(0).Visible = False
  mdiHelpCharts!mnuPrint(1).Visible = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  '************************************************************
  '* This routine will monitor keypresses looking for arrow
  '* keys. If one of the arrow keys is pressed then the
  '* current vowel is changed to the next or previous
  '* enabled vowel (depending upon which arrow is pressed).
  '************************************************************
  
  Dim i As Integer
  If gTestActive Then Exit Sub
  On Error Resume Next

  If Shift > 0 Then Exit Sub
  If KeyCode = vbKeyReturn Then Call Play(0)
  
  i = CurrIndex
  If KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then    'Is keypress right or down?
    i = IIf(CurrIndex = MaxVowels - 1, 0, CurrIndex + 1) 'Move to next vowel that is
    While Not Vowel(i).Enabled And i <> CurrIndex        '  enabled or until we've
      i = IIf(i = MaxVowels - 1, 0, i + 1)               '  gone full circle through
    Wend                                                 '  all the vowels.
  ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then   'Is keypress left or up?
    i = IIf(CurrIndex = 0, MaxVowels - 1, CurrIndex - 1) 'Move to previous vowel that
    While Not Vowel(i).Enabled And i <> CurrIndex        '  is enabled or until we've
      i = IIf(i = 0, MaxVowels - 1, i - 1)               '  gone full circle through
    Wend                                                 '  all the vowels.
  End If
  
  If Vowel(i).Enabled Then _
    Call Vowel_MouseDown(i, vbLeftButton, 0, 0, 0)       'If we landed on an enabled vowel,
                                                         '  act like we clicked on it.
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  
  On Error Resume Next
  If KeyAscii = vbKeyEscape Then
    If gTestActive Then
      Call mdiHelpCharts.mnuTestStop_Click
    Else
      Unload Me
    End If
  End If

End Sub

Private Sub Form_Load()
  
  mdiHelpCharts.panStatus.Visible = True
  Screen.MousePointer = vbHourglass
  WavNamePart1 = "Vow"
  CurrIndex = -1
  
  If ReadVowelsFromDB() Then
    'If next line is false then there's no vowels in DB.
    If CurrIndex > -1 And _
    gTestActive = False Then
      With Vowel(CurrIndex)
        .BackColor = vbHighlight       'Select first enabled vowel.
        .ForeColor = vbHighlightText
      End With
    End If
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Error reading vowel information from the settings file.", _
      vbInformation + vbOKOnly, App.Title
    Unload Me
    Exit Sub
  End If
  
  Screen.MousePointer = vbDefault
  Top = -Height
  Show
  WindowState = vbMaximized

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
  On Error Resume Next
  If gTestActive _
    Then gStatLine.SimpleText = "Select correct character" _
    Else gStatLine.SimpleText = ""
  MousePointer = 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  On Error Resume Next
  Call Form_Deactivate
  Set frmDispVow = Nothing

End Sub

'Private Sub Form_Resize()

'  On Error Resume Next
'  If WindowState > vbNormal Then Exit Sub
'  If Height > FrmMaxHeight Then Height = FrmMaxHeight
'  If Width > FrmMaxWidth Then Width = FrmMaxWidth
  
'End Sub

Private Sub Vowel_Click(Index As Integer)

  If iMouseButton = vbLeftButton Then
    '*********************************************
    '* If in test mode ...
    '*********************************************
    If gTestActive Then
      '*********************************************
      '* ... check to see if correct IPA character
      '* selected. If so, put smile on top of
      '* character.
      '*********************************************
      If gItemNumber = Index Then
        lblFrown.Visible = False
        With lblSmile
          .Height = Vowel(Index).Height
          .Width = Vowel(Index).Width
          .Top = Vowel(Index).Top
          .Left = Vowel(Index).Left
          .Visible = True
        End With
        Call MarkItemCorrect(vbChecked)
        '* Make smile visible for 2 seconds, the continue
        With mdiHelpCharts.Timer2
          .Enabled = False
          .Interval = 2000
          .Enabled = True
        End With
      '*********************
      '* If not, put frown.
      '*********************
      ElseIf lblSmile.Visible = False Then
        With lblFrown
          .Height = Vowel(Index).Height
          .Width = Vowel(Index).Width
          .Top = Vowel(Index).Top
          .Left = Vowel(Index).Left
          .Visible = True
        End With
        Call MarkItemCorrect(vbUnchecked, Index)
      End If
    Else
      '*********************************************
      '* If not in test mode, play associated sound.
      '*********************************************
      Call Play(0)
    End If
  End If
  
End Sub

Private Sub Vowel_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  '***********************************************
  '* Make the character that CurrIndex points
  '* to not look selected and make the character
  '* just clicked on look selected by changing
  '* its background color.
  '***********************************************
  
  On Error Resume Next
  If gTestActive = False Then
    With Vowel(CurrIndex)
      '* Deselect previous selection.
      .BackColor = vbButtonFace
      .ForeColor = vbButtonText
    End With
    With Vowel(Index)
      '* Select new selection.
      .BackColor = vbHighlight
      .ForeColor = vbHighlightText
    End With
  End If
  
  CurrIndex = Index
  iMouseButton = Button

End Sub

Private Sub Vowel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  On Error Resume Next
  If Not gTestActive Then
    gStatLine.SimpleText = GetCaptionFromTag(Vowel(Index).Tag)
  End If

End Sub

