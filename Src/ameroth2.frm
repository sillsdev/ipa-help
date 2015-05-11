VERSION 5.00
Begin VB.Form frmDispAmerOther2 
   BorderStyle     =   0  'None
   Caption         =   "Americanist - IPA Other Symbols"
   ClientHeight    =   6270
   ClientLeft      =   1365
   ClientTop       =   2040
   ClientWidth     =   9000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6270
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   91
      Left            =   2790
      TabIndex        =   96
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   90
      Left            =   3450
      TabIndex        =   95
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
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
      Height          =   195
      Index           =   88
      Left            =   3030
      TabIndex        =   94
      Top             =   2100
      Width           =   270
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Simultaneous"
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
      Height          =   195
      Index           =   87
      Left            =   1785
      TabIndex        =   93
      Top             =   2100
      Width           =   945
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   525
      X2              =   4125
      Y1              =   2340
      Y2              =   2340
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Additional mid central vowel"
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
      Height          =   195
      Index           =   65
      Left            =   1800
      TabIndex        =   91
      Top             =   3480
      Width           =   1965
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   540
      X2              =   4140
      Y1              =   3705
      Y2              =   3705
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   " d"
      Height          =   270
      Index           =   64
      Left            =   4980
      TabIndex        =   90
      Top             =   3015
      Width           =   150
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   " r"
      Height          =   270
      Index           =   62
      Left            =   4995
      TabIndex        =   89
      Top             =   2730
      Width           =   150
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Symbols"
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
      Height          =   195
      Index           =   26
      Left            =   570
      TabIndex        =   69
      Top             =   105
      Width           =   1095
   End
   Begin VB.Line Line1 
      Index           =   34
      X1              =   4785
      X2              =   8385
      Y1              =   4575
      Y2              =   4575
   End
   Begin VB.Line Line1 
      Index           =   36
      X1              =   4785
      X2              =   8385
      Y1              =   4005
      Y2              =   4005
   End
   Begin VB.Line Line1 
      Index           =   35
      X1              =   4785
      X2              =   8385
      Y1              =   4290
      Y2              =   4290
   End
   Begin VB.Line Line1 
      Index           =   28
      X1              =   4770
      X2              =   8370
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      Index           =   17
      X1              =   4770
      X2              =   8370
      Y1              =   3150
      Y2              =   3150
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   4770
      X2              =   8370
      Y1              =   3435
      Y2              =   3435
   End
   Begin VB.Line Line1 
      Index           =   33
      X1              =   4785
      X2              =   8385
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Line Line1 
      Index           =   32
      X1              =   4785
      X2              =   8385
      Y1              =   5010
      Y2              =   5010
   End
   Begin VB.Line Line1 
      Index           =   29
      X1              =   4770
      X2              =   8370
      Y1              =   345
      Y2              =   345
   End
   Begin VB.Line Line1 
      Index           =   27
      X1              =   4770
      X2              =   8370
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      Index           =   26
      X1              =   4770
      X2              =   8370
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Line Line1 
      Index           =   25
      X1              =   4770
      X2              =   8370
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      Index           =   24
      X1              =   4770
      X2              =   8370
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Line Line1 
      Index           =   23
      X1              =   4770
      X2              =   8370
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      Index           =   22
      X1              =   4770
      X2              =   8370
      Y1              =   1815
      Y2              =   1815
   End
   Begin VB.Line Line1 
      Index           =   21
      X1              =   4770
      X2              =   8370
      Y1              =   2085
      Y2              =   2085
   End
   Begin VB.Line Line1 
      Index           =   20
      X1              =   4770
      X2              =   8370
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Line Line1 
      Index           =   19
      X1              =   4770
      X2              =   8370
      Y1              =   2610
      Y2              =   2610
   End
   Begin VB.Line Line1 
      Index           =   18
      X1              =   4770
      X2              =   8370
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Suprasegmentals"
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
      Height          =   195
      Index           =   0
      Left            =   4770
      TabIndex        =   0
      Top             =   105
      Width           =   1215
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "È"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   10
      Left            =   5460
      TabIndex        =   1
      Top             =   300
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Ç"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   11
      Left            =   5460
      TabIndex        =   2
      Top             =   495
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ù"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   12
      Left            =   5460
      TabIndex        =   3
      Top             =   750
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   13
      Left            =   5460
      TabIndex        =   4
      Top             =   1020
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " á"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   14
      Left            =   5460
      TabIndex        =   8
      Top             =   1320
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "."
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   15
      Left            =   5460
      TabIndex        =   5
      Top             =   1485
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Ž"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   16
      Left            =   5460
      TabIndex        =   6
      Top             =   1785
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "„"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   17
      Left            =   5460
      TabIndex        =   18
      Top             =   2040
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "í"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   18
      Left            =   5460
      TabIndex        =   7
      Top             =   2190
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Ì"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   55
      Left            =   5475
      TabIndex        =   31
      Top             =   2580
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Í"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   57
      Left            =   5475
      TabIndex        =   32
      Top             =   2850
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "‹"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   42
      Left            =   5430
      TabIndex        =   25
      Top             =   4785
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "›"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   40
      Left            =   5430
      TabIndex        =   24
      Top             =   4560
      Width           =   285
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   " ž"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   9
      Left            =   5490
      TabIndex        =   19
      Top             =   4350
      Width           =   195
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   " ™"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   8
      Left            =   5490
      TabIndex        =   20
      Top             =   4065
      Width           =   195
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   " ”"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   7
      Left            =   5490
      TabIndex        =   21
      Top             =   3735
      Width           =   195
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   " ™"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   6
      Left            =   5475
      TabIndex        =   22
      Top             =   3510
      Width           =   195
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   " ‰"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   5
      Left            =   5490
      TabIndex        =   23
      Top             =   3195
      Width           =   195
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Š"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   24
      Left            =   5820
      TabIndex        =   26
      Top             =   3105
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "‘"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   28
      Left            =   5820
      TabIndex        =   27
      Top             =   3405
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "•"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   31
      Left            =   5820
      TabIndex        =   28
      Top             =   3690
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "š"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   34
      Left            =   5820
      TabIndex        =   29
      Top             =   3975
      Width           =   285
   End
   Begin VB.Line Line1 
      Index           =   31
      X1              =   4770
      X2              =   4770
      Y1              =   345
      Y2              =   5010
   End
   Begin VB.Line Line1 
      Index           =   30
      X1              =   8370
      X2              =   8370
      Y1              =   345
      Y2              =   5010
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "or"
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
      Height          =   195
      Index           =   25
      Left            =   5700
      TabIndex        =   68
      Top             =   4320
      Width           =   135
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "or"
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
      Height          =   195
      Index           =   22
      Left            =   5700
      TabIndex        =   67
      Top             =   4050
      Width           =   135
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "or"
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
      Height          =   195
      Index           =   21
      Left            =   5700
      TabIndex        =   66
      Top             =   3750
      Width           =   135
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "or"
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
      Height          =   195
      Index           =   20
      Left            =   5700
      TabIndex        =   65
      Top             =   3480
      Width           =   135
   End
   Begin VB.Line Line1 
      Index           =   14
      X1              =   540
      X2              =   4140
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   540
      X2              =   4140
      Y1              =   3435
      Y2              =   3435
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   540
      X2              =   4140
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   540
      X2              =   4140
      Y1              =   2610
      Y2              =   2610
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   540
      X2              =   4140
      Y1              =   2070
      Y2              =   2070
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   540
      X2              =   4140
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   540
      X2              =   4140
      Y1              =   1575
      Y2              =   1575
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   540
      X2              =   4140
      Y1              =   1335
      Y2              =   1335
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   540
      X2              =   4140
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   540
      X2              =   4140
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   555
      X2              =   4155
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   540
      X2              =   4140
      Y1              =   345
      Y2              =   345
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   4140
      X2              =   4140
      Y1              =   345
      Y2              =   3705
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   540
      X2              =   540
      Y1              =   345
      Y2              =   3705
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Voiceless labial-velar approx."
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
      Height          =   195
      Index           =   19
      Left            =   1800
      TabIndex        =   64
      Top             =   630
      Width           =   2040
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ü"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   79
      Left            =   1455
      TabIndex        =   63
      Top             =   3105
      Width           =   255
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Voiceless labial-velar fricative"
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
      Height          =   195
      Index           =   79
      Left            =   1800
      TabIndex        =   62
      Top             =   375
      Width           =   2070
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Voiced labial-velar approx."
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
      Height          =   195
      Index           =   80
      Left            =   1800
      TabIndex        =   61
      Top             =   885
      Width           =   1860
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Voiced labial-palatal approx."
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
      Height          =   195
      Index           =   81
      Left            =   1800
      TabIndex        =   60
      Top             =   1110
      Width           =   1980
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Voiceless epiglottal fricative"
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
      Height          =   195
      Index           =   82
      Left            =   1800
      TabIndex        =   59
      Top             =   1365
      Width           =   1950
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Voiced epiglottal fricative"
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
      Height          =   195
      Index           =   83
      Left            =   1800
      TabIndex        =   58
      Top             =   1605
      Width           =   1770
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Epiglottal plosive"
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
      Height          =   195
      Index           =   84
      Left            =   1800
      TabIndex        =   57
      Top             =   1845
      Width           =   1185
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Alveolo-palatal fricatives"
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
      Height          =   195
      Index           =   85
      Left            =   1800
      TabIndex        =   56
      Top             =   3180
      Width           =   1710
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Alveolar lateral flap"
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
      Height          =   195
      Index           =   86
      Left            =   1800
      TabIndex        =   55
      Top             =   2895
      Width           =   1335
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ã"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   72
      Left            =   1215
      TabIndex        =   54
      Top             =   315
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "û"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   77
      Left            =   1215
      TabIndex        =   49
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alveolar lateral click"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   63
      Left            =   1800
      TabIndex        =   44
      Top             =   2655
      Width           =   1410
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Palatoalveolar click"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   69
      Left            =   1800
      TabIndex        =   43
      Top             =   2370
      Width           =   1380
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Falling intonation"
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
      Height          =   195
      Index           =   58
      Left            =   6180
      TabIndex        =   42
      Top             =   2910
      Width           =   1185
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rising intonation"
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
      Height          =   195
      Index           =   56
      Left            =   6180
      TabIndex        =   41
      Top             =   2640
      Width           =   1170
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Upstep"
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
      Height          =   195
      Index           =   43
      Left            =   6180
      TabIndex        =   40
      Top             =   4800
      Width           =   510
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Downstep"
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
      Height          =   195
      Index           =   41
      Left            =   6180
      TabIndex        =   39
      Top             =   4590
      Width           =   720
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Extra low"
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
      Height          =   195
      Index           =   39
      Left            =   6180
      TabIndex        =   38
      Top             =   4335
      Width           =   645
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   36
      Left            =   6180
      TabIndex        =   37
      Top             =   4050
      Width           =   300
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   33
      Left            =   6180
      TabIndex        =   36
      Top             =   3765
      Width           =   255
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   30
      Left            =   6180
      TabIndex        =   35
      Top             =   3480
      Width           =   330
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Extra high"
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
      Height          =   195
      Index           =   27
      Left            =   6180
      TabIndex        =   34
      Top             =   3180
      Width           =   705
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "or"
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
      Height          =   195
      Index           =   23
      Left            =   5700
      TabIndex        =   33
      Top             =   3210
      Width           =   135
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Ÿ"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   37
      Left            =   5820
      TabIndex        =   30
      Top             =   4260
      Width           =   285
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Major (intonation) group"
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
      Height          =   195
      Index           =   8
      Left            =   6180
      TabIndex        =   17
      Top             =   2115
      Width           =   1665
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Linking (absence of a break)"
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
      Height          =   195
      Index           =   9
      Left            =   6180
      TabIndex        =   16
      Top             =   2385
      Width           =   2025
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Minor (foot) group"
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
      Height          =   195
      Index           =   7
      Left            =   6180
      TabIndex        =   15
      Top             =   1845
      Width           =   1245
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Syllable break"
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
      Height          =   195
      Index           =   6
      Left            =   6180
      TabIndex        =   14
      Top             =   1575
      Width           =   990
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Extra-short"
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
      Height          =   195
      Index           =   5
      Left            =   6180
      TabIndex        =   13
      Top             =   1365
      Width           =   750
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Primary stress"
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
      Height          =   195
      Index           =   1
      Left            =   6180
      TabIndex        =   12
      Top             =   360
      Width           =   960
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Half-long"
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
      Height          =   195
      Index           =   4
      Left            =   6180
      TabIndex        =   11
      Top             =   1125
      Width           =   630
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Long"
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
      Height          =   195
      Index           =   3
      Left            =   6180
      TabIndex        =   10
      Top             =   855
      Width           =   360
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Secondary stress"
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
      Height          =   195
      Index           =   2
      Left            =   6180
      TabIndex        =   9
      Top             =   615
      Width           =   1215
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "w8"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1215
      TabIndex        =   70
      Top             =   525
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "w"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   73
      Left            =   1215
      TabIndex        =   53
      Top             =   825
      Width           =   255
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[W]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   51
      Left            =   660
      TabIndex        =   71
      Top             =   270
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[w]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   29
      Left            =   630
      TabIndex        =   72
      Top             =   525
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[tÍ‹]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   32
      Left            =   630
      TabIndex        =   73
      Top             =   2265
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[é‹]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   35
      Left            =   630
      TabIndex        =   74
      Top             =   2550
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[lÒ]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   38
      Left            =   585
      TabIndex        =   75
      Top             =   3105
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[']"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   44
      Left            =   4860
      TabIndex        =   76
      Top             =   270
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[:]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   45
      Left            =   4860
      TabIndex        =   77
      Top             =   765
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[·]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   46
      Left            =   4860
      TabIndex        =   78
      Top             =   1020
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[.]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   47
      Left            =   4860
      TabIndex        =   79
      Top             =   1485
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[\]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   48
      Left            =   4845
      TabIndex        =   80
      Top             =   1785
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[//]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   49
      Left            =   4860
      TabIndex        =   81
      Top             =   2070
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[-]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   50
      Left            =   4860
      TabIndex        =   82
      Top             =   2535
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[-]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   52
      Left            =   4860
      TabIndex        =   83
      Top             =   2805
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[º]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   53
      Left            =   4860
      TabIndex        =   84
      Top             =   3105
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[¥]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   54
      Left            =   4860
      TabIndex        =   85
      Top             =   3390
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[¡]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   59
      Left            =   4860
      TabIndex        =   86
      Top             =   3675
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[•]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   60
      Left            =   4860
      TabIndex        =   87
      Top             =   3960
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[]"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   61
      Left            =   4860
      TabIndex        =   88
      Top             =   4245
      Width           =   405
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ç"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   74
      Left            =   1215
      TabIndex        =   52
      Top             =   1035
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   75
      Left            =   1215
      TabIndex        =   51
      Top             =   1290
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "¹"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   76
      Left            =   1215
      TabIndex        =   50
      Top             =   1545
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Î"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   1
      Left            =   1245
      TabIndex        =   92
      Top             =   3405
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "î"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   81
      Left            =   1215
      TabIndex        =   97
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ä"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   80
      Left            =   1230
      TabIndex        =   48
      Top             =   2850
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "þ"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   78
      Left            =   1230
      TabIndex        =   47
      Top             =   3105
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "’"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   62
      Left            =   1230
      TabIndex        =   45
      Top             =   2610
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "œ"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   61
      Left            =   1230
      TabIndex        =   46
      Top             =   2325
      Width           =   255
   End
End
Attribute VB_Name = "frmDispAmerOther2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
'* frmDispSILAmerOthr version info:
'*  See basGlobals (Globals.bas).
'**************************************************

Option Explicit

Dim CurrIndex As Integer
Dim CharDesc() As String

Private Const TBarButtons = "Exit;"
Private Const MaxVowels = 24
Private Const FrmMaxHeight = 4020
Private Const FrmMaxWidth = 7830

Private Sub Form_Activate()

  On Error Resume Next
  With mdiHelpCharts
    .panStatus.Visible = True
    .ShowTBarButtons TBarButtons
    .mnuTest.Enabled = False
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
  Set frmDispAmerOther2 = Nothing
  Erase CharDesc

End Sub


Private Sub Form_Resize()

  On Error Resume Next
  If WindowState > vbNormal Then Exit Sub
  If Height > FrmMaxHeight Then Height = FrmMaxHeight
  If Width > FrmMaxWidth Then Width = FrmMaxWidth

End Sub


