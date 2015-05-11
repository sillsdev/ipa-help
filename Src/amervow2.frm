VERSION 5.00
Begin VB.Form frmDispAmerVow2 
   BorderStyle     =   0  'None
   Caption         =   "Americanist - IPA Vowels"
   ClientHeight    =   6270
   ClientLeft      =   1065
   ClientTop       =   1170
   ClientWidth     =   9000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Amer Phon SILDoulosL"
      Size            =   8.25
      Charset         =   2
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6270
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[æ]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   38
      Left            =   4995
      TabIndex        =   56
      Top             =   1650
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ä]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   37
      Left            =   5115
      TabIndex        =   55
      Top             =   2250
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[æú]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   36
      Left            =   6120
      TabIndex        =   54
      Top             =   1650
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[eû]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   35
      Left            =   7260
      TabIndex        =   53
      Top             =   2250
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ìû]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   34
      Left            =   4080
      TabIndex        =   52
      Top             =   1650
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[æ]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   33
      Left            =   3750
      TabIndex        =   51
      Top             =   1650
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ì]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   32
      Left            =   6495
      TabIndex        =   50
      Top             =   1650
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ã]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   31
      Left            =   5370
      TabIndex        =   49
      Top             =   3435
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[oû]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   30
      Left            =   3420
      TabIndex        =   48
      Top             =   2250
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[€ú]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   29
      Left            =   7155
      TabIndex        =   47
      Top             =   1020
      Width           =   405
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[uû]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   28
      Left            =   2910
      TabIndex        =   46
      Top             =   1020
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Œ"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   23
      Left            =   5685
      TabIndex        =   18
      Top             =   3795
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "«"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   15
      Left            =   5415
      TabIndex        =   17
      Top             =   2595
      Width           =   270
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Close-mid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   18
      Left            =   825
      TabIndex        =   45
      Top             =   2040
      Width           =   780
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Mid close]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   17
      Left            =   825
      TabIndex        =   44
      Top             =   2340
      Width           =   780
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   16
      Left            =   825
      TabIndex        =   43
      Top             =   2640
      Width           =   780
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Open-mid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   14
      Left            =   825
      TabIndex        =   42
      Top             =   3240
      Width           =   780
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Mid open]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   13
      Left            =   825
      TabIndex        =   41
      Top             =   3540
      Width           =   780
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Near-open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   12
      Left            =   825
      TabIndex        =   40
      Top             =   3855
      Width           =   780
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Low close]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   11
      Left            =   825
      TabIndex        =   39
      Top             =   4140
      Width           =   900
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   10
      Left            =   825
      TabIndex        =   38
      Top             =   4440
      Width           =   780
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Low open]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   9
      Left            =   825
      TabIndex        =   37
      Top             =   4740
      Width           =   900
   End
   Begin VB.Line Line2 
      Index           =   11
      X1              =   600
      X2              =   1905
      Y1              =   4395
      Y2              =   4395
   End
   Begin VB.Line Line2 
      Index           =   10
      X1              =   600
      X2              =   1905
      Y1              =   3795
      Y2              =   3795
   End
   Begin VB.Line Line2 
      Index           =   9
      X1              =   600
      X2              =   1905
      Y1              =   3195
      Y2              =   3195
   End
   Begin VB.Line Line2 
      Index           =   8
      X1              =   600
      X2              =   1905
      Y1              =   2595
      Y2              =   2595
   End
   Begin VB.Line Line2 
      Index           =   7
      X1              =   600
      X2              =   1905
      Y1              =   1995
      Y2              =   1995
   End
   Begin VB.Line Line2 
      Index           =   6
      X1              =   600
      X2              =   1905
      Y1              =   1395
      Y2              =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "[Amer.]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   825
      TabIndex        =   36
      Top             =   510
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "IPA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   825
      TabIndex        =   35
      Top             =   270
      Width           =   495
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   8325
      X2              =   8325
      Y1              =   195
      Y2              =   4995
   End
   Begin VB.Line Line2 
      Index           =   5
      X1              =   600
      X2              =   8330
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Line Line2 
      Index           =   4
      X1              =   1905
      X2              =   1905
      Y1              =   195
      Y2              =   4995
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   600
      X2              =   8330
      Y1              =   195
      Y2              =   195
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   600
      X2              =   8330
      Y1              =   4995
      Y2              =   4995
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   585
      X2              =   585
      Y1              =   195
      Y2              =   4995
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "¬"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   5205
      TabIndex        =   34
      Top             =   795
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   12
      Left            =   5460
      TabIndex        =   33
      Top             =   1995
      Width           =   270
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   2
      Left            =   7530
      Shape           =   3  'Circle
      Top             =   930
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   8
      Left            =   7530
      Shape           =   3  'Circle
      Top             =   4530
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   4
      Left            =   7530
      Shape           =   3  'Circle
      Top             =   2115
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   6
      Left            =   7530
      Shape           =   3  'Circle
      Top             =   3300
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   1
      Left            =   5130
      Shape           =   3  'Circle
      Top             =   930
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   10
      Left            =   5640
      Shape           =   3  'Circle
      Top             =   3300
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   9
      Left            =   5385
      Shape           =   3  'Circle
      Top             =   2115
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   5
      Left            =   3660
      Shape           =   3  'Circle
      Top             =   3315
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   0
      Left            =   2550
      Shape           =   3  'Circle
      Top             =   930
      Width           =   75
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   3
      Left            =   3105
      Shape           =   3  'Circle
      Top             =   2115
      Width           =   75
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   7560
      X2              =   7560
      Y1              =   960
      Y2              =   4550
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   5160
      X2              =   5940
      Y1              =   960
      Y2              =   4550
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   4605
      X2              =   7195
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   6
      Left            =   7365
      TabIndex        =   32
      Top             =   390
      Width           =   405
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Central"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   5
      Left            =   4875
      TabIndex        =   31
      Top             =   390
      Width           =   555
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[High open]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   825
      TabIndex        =   30
      Top             =   1740
      Width           =   900
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Near-close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   825
      TabIndex        =   29
      Top             =   1440
      Width           =   780
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[High close]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   825
      TabIndex        =   28
      Top             =   1140
      Width           =   915
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   825
      TabIndex        =   27
      Top             =   840
      Width           =   780
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Front"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   4
      Left            =   2355
      TabIndex        =   26
      Top             =   390
      Width           =   405
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   4095
      X2              =   5305
      Y1              =   3330
      Y2              =   3330
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   5760
      X2              =   7200
      Y1              =   2145
      Y2              =   2145
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   5550
      X2              =   7140
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   6015
      X2              =   7195
      Y1              =   3330
      Y2              =   3330
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   3525
      X2              =   5045
      Y1              =   2145
      Y2              =   2145
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   2985
      X2              =   4795
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   5
      Left            =   7620
      TabIndex        =   25
      Top             =   795
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   ""
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   21
      Left            =   7620
      TabIndex        =   24
      Top             =   3150
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   ""
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   27
      Left            =   7620
      TabIndex        =   23
      Top             =   4380
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Î"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   18
      Left            =   5340
      TabIndex        =   22
      Top             =   3150
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Ã"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   20
      Left            =   7260
      TabIndex        =   21
      Top             =   3150
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   26
      Left            =   7260
      TabIndex        =   20
      Top             =   4380
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Ï"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   19
      Left            =   5745
      TabIndex        =   19
      Top             =   3150
      Width           =   240
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "µ"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   7230
      TabIndex        =   16
      Top             =   795
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   22
      Left            =   3630
      TabIndex        =   15
      Top             =   3825
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   16
      Left            =   3315
      TabIndex        =   14
      Top             =   3150
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "¿"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   17
      Left            =   3780
      TabIndex        =   13
      Top             =   3150
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "‚"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   11
      Left            =   5070
      TabIndex        =   12
      Top             =   1995
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   13
      Left            =   7260
      TabIndex        =   11
      Top             =   1995
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   14
      Left            =   7620
      TabIndex        =   10
      Top             =   1995
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "¯"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   25
      Left            =   4305
      TabIndex        =   9
      Top             =   4380
      Width           =   315
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   24
      Left            =   3885
      TabIndex        =   8
      Top             =   4380
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   7
      Left            =   4020
      TabIndex        =   7
      Top             =   1395
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   6
      Left            =   3705
      TabIndex        =   6
      Top             =   1395
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   8
      Left            =   6495
      TabIndex        =   5
      Top             =   1395
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   2685
      TabIndex        =   4
      Top             =   795
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   2265
      TabIndex        =   3
      Top             =   795
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ö"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   4860
      TabIndex        =   2
      Top             =   795
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   2745
      TabIndex        =   1
      Top             =   1995
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "ASAP SILDoulos"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   10
      Left            =   3225
      TabIndex        =   0
      Top             =   1995
      Width           =   270
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   75
      Index           =   7
      Left            =   4230
      Shape           =   3  'Circle
      Top             =   4530
      Width           =   75
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2565
      X2              =   4245
      Y1              =   930
      Y2              =   4550
   End
End
Attribute VB_Name = "frmDispAmerVow2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
'* frmDispSILAmerVow version info:
'*  See basGlobals (Globals.bas).
'**************************************************

Option Explicit

Private Const TBarButtons = "Exit;"
Private Const FrmMaxHeight = 4020
Private Const FrmMaxWidth = 7830

Private Sub Form_Activate()

  On Error Resume Next
  With mdiHelpCharts
    .ShowTBarButtons TBarButtons
    .panStatus.Visible = True
    .mnuTest.Enabled = False                        '* Disable test menu.
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

End Sub

Private Sub Form_Load()
  
  On Error Resume Next
  Top = -Height
  Show
  mdiHelpCharts.panStatus.Visible = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

  On Error Resume Next
  gStatLine.SimpleText = ""

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   
  On Error Resume Next
  Call Form_Deactivate
  Set frmDispAmerVow2 = Nothing

End Sub

Private Sub Form_Resize()

  If WindowState > vbNormal Then Exit Sub
  If Height > FrmMaxHeight Then Height = FrmMaxHeight
  If Width > FrmMaxWidth Then Width = FrmMaxWidth

End Sub


