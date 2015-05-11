VERSION 5.00
Begin VB.Form frmDispAmerVow 
   BorderStyle     =   0  'None
   Caption         =   "Americanist Vowels"
   ClientHeight    =   6270
   ClientLeft      =   1470
   ClientTop       =   3120
   ClientWidth     =   8985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6270
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line4 
      X1              =   1365
      X2              =   7515
      Y1              =   4095
      Y2              =   4095
   End
   Begin VB.Line Line3 
      X1              =   1335
      X2              =   1335
      Y1              =   1260
      Y2              =   4125
   End
   Begin VB.Line Line2 
      X1              =   1365
      X2              =   7515
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   1815
      X2              =   1815
      Y1              =   1270
      Y2              =   4090
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2175
      X2              =   3435
      Y1              =   720
      Y2              =   4100
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Index           =   20
      Left            =   1920
      TabIndex        =   41
      Top             =   2345
      Width           =   495
   End
   Begin VB.Label Label5 
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
      Height          =   225
      Index           =   19
      Left            =   1935
      TabIndex        =   40
      Top             =   2800
      Width           =   465
   End
   Begin VB.Label Label5 
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
      Height          =   210
      Index           =   13
      Left            =   1935
      TabIndex        =   39
      Top             =   1410
      Width           =   465
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rd."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   10
      Left            =   7080
      TabIndex        =   38
      Top             =   1035
      Width           =   390
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Unr."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   6285
      TabIndex        =   37
      Top             =   1035
      Width           =   360
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rd."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   5685
      TabIndex        =   36
      Top             =   1035
      Width           =   360
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rd."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   4380
      TabIndex        =   35
      Top             =   1035
      Width           =   330
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Unr."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   4950
      TabIndex        =   34
      Top             =   1035
      Width           =   360
   End
   Begin VB.Label Label5 
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
      Height          =   210
      Index           =   4
      Left            =   6615
      TabIndex        =   33
      Top             =   750
      Width           =   480
   End
   Begin VB.Label Label5 
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
      Height          =   210
      Index           =   3
      Left            =   5205
      TabIndex        =   32
      Top             =   765
      Width           =   510
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2295
      X2              =   7500
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Label Label5 
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
      Height          =   210
      Index           =   2
      Left            =   3345
      TabIndex        =   31
      Top             =   750
      Width           =   480
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Unr."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2640
      TabIndex        =   30
      Top             =   1035
      Width           =   300
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2175
      X2              =   7500
      Y1              =   705
      Y2              =   705
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   6165
      X2              =   6165
      Y1              =   720
      Y2              =   4110
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "å"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   15
      Left            =   3255
      TabIndex        =   29
      Top             =   2625
      Width           =   270
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   4830
      X2              =   4830
      Y1              =   705
      Y2              =   4095
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ëû"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   18
      Left            =   4400
      TabIndex        =   28
      Top             =   3150
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "oû"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   11
      Left            =   4400
      TabIndex        =   27
      Top             =   2200
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ä"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   12
      Left            =   5000
      TabIndex        =   26
      Top             =   2200
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "âû"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   19
      Left            =   6315
      TabIndex        =   25
      Top             =   3150
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "uû"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   1
      Left            =   4400
      TabIndex        =   24
      Top             =   1250
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "uª"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   3
      Left            =   5700
      TabIndex        =   23
      Top             =   1250
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ç"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   2
      Left            =   5000
      TabIndex        =   22
      Top             =   1250
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "€ú"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   4
      Left            =   6315
      TabIndex        =   21
      Top             =   1250
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "á"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   23
      Left            =   7095
      TabIndex        =   20
      Top             =   3630
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   21
      Left            =   3675
      TabIndex        =   19
      Top             =   3630
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ë"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   20
      Left            =   7095
      TabIndex        =   18
      Top             =   3150
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "u"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   5
      Left            =   7095
      TabIndex        =   17
      Top             =   1250
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "à"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   22
      Left            =   5000
      TabIndex        =   16
      Top             =   3630
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "â"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   17
      Left            =   3435
      TabIndex        =   15
      Top             =   3150
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ã"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   16
      Left            =   5000
      TabIndex        =   14
      Top             =   2625
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ì"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   9
      Left            =   7095
      TabIndex        =   13
      Top             =   1725
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   10
      Left            =   3050
      TabIndex        =   12
      Top             =   2200
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "o"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   14
      Left            =   7095
      TabIndex        =   11
      Top             =   2200
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "æ"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   6
      Left            =   2895
      TabIndex        =   10
      Top             =   1725
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ìû"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   7
      Left            =   4400
      TabIndex        =   9
      Top             =   1725
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "eû"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   13
      Left            =   6315
      TabIndex        =   8
      Top             =   2200
      Width           =   270
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "æú"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   8
      Left            =   6315
      TabIndex        =   7
      Top             =   1725
      Width           =   270
   End
   Begin VB.Label Label5 
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
      Height          =   210
      Index           =   0
      Left            =   1935
      TabIndex        =   6
      Top             =   1860
      Width           =   450
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "High"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   14
      Left            =   1410
      TabIndex        =   5
      Top             =   1410
      Width           =   450
   End
   Begin VB.Label Label5 
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
      Height          =   225
      Index           =   15
      Left            =   1410
      TabIndex        =   4
      Top             =   2340
      Width           =   390
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Low"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   16
      Left            =   1410
      TabIndex        =   3
      Top             =   3285
      Width           =   405
   End
   Begin VB.Label Label5 
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
      Height          =   210
      Index           =   17
      Left            =   1935
      TabIndex        =   2
      Top             =   3300
      Width           =   480
   End
   Begin VB.Label Label5 
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
      Height          =   225
      Index           =   18
      Left            =   1935
      TabIndex        =   1
      Top             =   3775
      Width           =   465
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   1350
      X2              =   7515
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   1350
      X2              =   7515
      Y1              =   2190
      Y2              =   2190
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   7500
      X2              =   7500
      Y1              =   705
      Y2              =   4095
   End
   Begin VB.Label Vowel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "i"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   15.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Top             =   1250
      Width           =   270
   End
End
Attribute VB_Name = "frmDispAmerVow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
'* frmDispAmerVow version info:
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
  
  mdiHelpCharts.panStatus.Visible = True
  mdiHelpCharts.MousePointer = vbDefault
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
  Set frmDispAmerVow = Nothing

End Sub

Private Sub Form_Resize()

  If WindowState > vbNormal Then Exit Sub
  If Height > FrmMaxHeight Then Height = FrmMaxHeight
  If Width > FrmMaxWidth Then Width = FrmMaxWidth

End Sub
