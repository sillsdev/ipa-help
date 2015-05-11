VERSION 5.00
Begin VB.Form frmDispAmerDia2 
   BorderStyle     =   0  'None
   Caption         =   "Americanist - IPA Diacritics"
   ClientHeight    =   6270
   ClientLeft      =   1350
   ClientTop       =   1995
   ClientWidth     =   9000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6270
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   " ü"
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
      Height          =   255
      Index           =   115
      Left            =   5280
      TabIndex        =   163
      Top             =   3165
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   " ü"
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
      Height          =   255
      Index           =   114
      Left            =   2910
      TabIndex        =   162
      Top             =   3180
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "  Ú"
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
      Height          =   375
      Index           =   113
      Left            =   4710
      TabIndex        =   161
      Top             =   1950
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "  Ú"
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
      Height          =   375
      Index           =   112
      Left            =   2850
      TabIndex        =   160
      Top             =   1965
      Width           =   300
   End
   Begin VB.Label Label2 
      Caption         =   "² Superscript"
      Height          =   195
      Index           =   1
      Left            =   3000
      TabIndex        =   159
      Top             =   4890
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   9.75
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   111
      Left            =   7980
      TabIndex        =   158
      Top             =   2295
      Width           =   195
   End
   Begin VB.Label Label2 
      Caption         =   "¹ Upper Case"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   135
      Top             =   4890
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Pharyngealized"
      Height          =   255
      Index           =   2
      Left            =   3435
      TabIndex        =   113
      Top             =   3090
      Width           =   1200
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   150
      X2              =   150
      Y1              =   105
      Y2              =   4845
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   2865
      X2              =   2865
      Y1              =   105
      Y2              =   4845
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   8790
      X2              =   8790
      Y1              =   105
      Y2              =   4845
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   165
      X2              =   8790
      Y1              =   105
      Y2              =   105
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   150
      X2              =   8790
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   150
      X2              =   8790
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   150
      X2              =   8790
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   150
      X2              =   8790
      Y1              =   1605
      Y2              =   1605
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   150
      X2              =   8790
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   150
      X2              =   8790
      Y1              =   2355
      Y2              =   2355
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   150
      X2              =   8790
      Y1              =   2730
      Y2              =   2730
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   150
      X2              =   8790
      Y1              =   3345
      Y2              =   3345
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   5850
      X2              =   5850
      Y1              =   105
      Y2              =   2730
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   150
      X2              =   8790
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   150
      X2              =   8790
      Y1              =   4095
      Y2              =   4095
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   150
      X2              =   8790
      Y1              =   4470
      Y2              =   4470
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   150
      X2              =   8790
      Y1              =   4845
      Y2              =   4845
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "¨£"
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
      Height          =   390
      Index           =   31
      Left            =   6165
      TabIndex        =   110
      Top             =   3345
      Width           =   285
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d|"
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
      Height          =   390
      Index           =   45
      Left            =   8490
      TabIndex        =   109
      Top             =   2355
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d:"
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
      Height          =   390
      Index           =   44
      Left            =   8490
      TabIndex        =   108
      Top             =   1980
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d<"
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
      Height          =   390
      Index           =   43
      Left            =   8490
      TabIndex        =   107
      Top             =   1605
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "e)"
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
      Height          =   390
      Index           =   42
      Left            =   8490
      TabIndex        =   106
      Top             =   1230
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "t6"
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
      Height          =   390
      Index           =   40
      Left            =   8175
      TabIndex        =   105
      Top             =   825
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "t°"
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
      Height          =   390
      Index           =   38
      Left            =   8175
      TabIndex        =   104
      Top             =   465
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "t0"
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
      Height          =   390
      Index           =   36
      Left            =   8175
      TabIndex        =   103
      Top             =   90
      Width           =   300
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No audible release"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   59
      Left            =   6435
      TabIndex        =   102
      Top             =   2460
      Width           =   1320
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Lateral release"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   58
      Left            =   6435
      TabIndex        =   101
      Top             =   2085
      Width           =   1035
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nasal release"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   57
      Left            =   6435
      TabIndex        =   100
      Top             =   1710
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nasalized"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   56
      Left            =   6435
      TabIndex        =   99
      Top             =   1335
      Width           =   690
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Laminal"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   55
      Left            =   6435
      TabIndex        =   98
      Top             =   960
      Width           =   540
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Apical"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   54
      Left            =   6435
      TabIndex        =   97
      Top             =   585
      Width           =   435
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dental"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   53
      Left            =   6435
      TabIndex        =   96
      Top             =   210
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " ¢"
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
      Index           =   81
      Left            =   3180
      TabIndex        =   95
      Top             =   3720
      Width           =   285
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "eª"
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
      Height          =   405
      Index           =   35
      Left            =   5730
      TabIndex        =   94
      Top             =   4440
      Width           =   255
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "eÁ"
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
      Height          =   405
      Index           =   34
      Left            =   5715
      TabIndex        =   93
      Top             =   4065
      Width           =   270
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "= voiced bilabial approximant)"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   40
      Left            =   6450
      TabIndex        =   92
      Top             =   3825
      Width           =   2085
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "B¢"
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
      Height          =   405
      Index           =   33
      Left            =   6165
      TabIndex        =   91
      Top             =   3720
      Width           =   300
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "("
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   42
      Left            =   6015
      TabIndex        =   90
      Top             =   3825
      Width           =   45
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "e¢"
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
      Height          =   405
      Index           =   32
      Left            =   5715
      TabIndex        =   89
      Top             =   3720
      Width           =   225
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "= voiced alveolar fricative)"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   39
      Left            =   6450
      TabIndex        =   88
      Top             =   3450
      Width           =   1860
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "("
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   43
      Left            =   6015
      TabIndex        =   87
      Top             =   3450
      Width           =   45
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "e£"
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
      Height          =   420
      Index           =   30
      Left            =   5715
      TabIndex        =   86
      Top             =   3330
      Width           =   225
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "lò"
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
      Height          =   390
      Index           =   29
      Left            =   5730
      TabIndex        =   85
      Top             =   2745
      Width           =   240
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "t³"
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
      Height          =   390
      Index           =   27
      Left            =   5160
      TabIndex        =   84
      Top             =   2355
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "tì"
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
      Height          =   390
      Index           =   25
      Left            =   5145
      TabIndex        =   83
      Top             =   1980
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "tJ"
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
      Height          =   390
      Index           =   23
      Left            =   5145
      TabIndex        =   82
      Top             =   1605
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "tW"
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
      Height          =   390
      Index           =   21
      Left            =   5145
      TabIndex        =   81
      Top             =   1230
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "tÑ"
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
      Height          =   390
      Index           =   19
      Left            =   5160
      TabIndex        =   80
      Top             =   840
      Width           =   285
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "b¼"
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
      Height          =   390
      Index           =   17
      Left            =   5160
      TabIndex        =   79
      Top             =   495
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "b-"
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
      Height          =   390
      Index           =   15
      Left            =   5160
      TabIndex        =   78
      Top             =   105
      Width           =   300
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Retracted Tongue Root"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   35
      Left            =   3450
      TabIndex        =   77
      Top             =   4575
      Width           =   1695
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced Tongue Root"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   34
      Left            =   3450
      TabIndex        =   76
      Top             =   4200
      Width           =   1725
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Lowered"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   33
      Left            =   3450
      TabIndex        =   75
      Top             =   3825
      Width           =   615
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Raised"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   32
      Left            =   3450
      TabIndex        =   74
      Top             =   3450
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Velarized or"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   31
      Left            =   3465
      TabIndex        =   73
      Top             =   2820
      Width           =   1200
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pharyngealized"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   30
      Left            =   3450
      TabIndex        =   72
      Top             =   2460
      Width           =   1080
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Velarized"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   29
      Left            =   3450
      TabIndex        =   71
      Top             =   2085
      Width           =   645
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Palatalized"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   28
      Left            =   3450
      TabIndex        =   70
      Top             =   1710
      Width           =   765
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Labialized"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   27
      Left            =   3450
      TabIndex        =   69
      Top             =   1335
      Width           =   705
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Linguolabial"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   26
      Left            =   3450
      TabIndex        =   68
      Top             =   960
      Width           =   840
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Creaky voiced"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   25
      Left            =   3450
      TabIndex        =   67
      Top             =   585
      Width           =   1020
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Breathy voiced"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   24
      Left            =   3450
      TabIndex        =   66
      Top             =   210
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " ~"
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
      Height          =   390
      Index           =   68
      Left            =   465
      TabIndex        =   65
      Top             =   3390
      Width           =   285
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "«Õ"
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
      Height          =   405
      Index           =   14
      Left            =   2580
      TabIndex        =   64
      Top             =   4470
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "e9"
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
      Height          =   405
      Index           =   13
      Left            =   2580
      TabIndex        =   63
      Top             =   4095
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "¨`"
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
      Height          =   420
      Index           =   12
      Left            =   2580
      TabIndex        =   62
      Top             =   3690
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "e~"
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
      Height          =   405
      Index           =   11
      Left            =   2580
      TabIndex        =   61
      Top             =   3345
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "e_"
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
      Height          =   390
      Index           =   10
      Left            =   2580
      TabIndex        =   60
      Top             =   2730
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "i="
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
      Height          =   390
      Index           =   9
      Left            =   2580
      TabIndex        =   59
      Top             =   2355
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "u+"
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
      Height          =   405
      Index           =   8
      Left            =   2580
      TabIndex        =   58
      Top             =   1950
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "¦"
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
      Height          =   405
      Index           =   6
      Left            =   2580
      TabIndex        =   57
      Top             =   1200
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "tH"
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
      Height          =   390
      Index           =   4
      Left            =   2235
      TabIndex        =   56
      Top             =   855
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "s¤"
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
      Height          =   390
      Index           =   2
      Left            =   2235
      TabIndex        =   55
      Top             =   465
      Width           =   285
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rhoticity"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   14
      Left            =   735
      TabIndex        =   54
      Top             =   4575
      Width           =   615
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Non-syllabic"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   13
      Left            =   735
      TabIndex        =   53
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Syllabic"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   12
      Left            =   735
      TabIndex        =   52
      Top             =   3825
      Width           =   540
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mid-centralized"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   11
      Left            =   765
      TabIndex        =   51
      Top             =   3450
      Width           =   1065
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Centralized"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   735
      TabIndex        =   50
      Top             =   2835
      Width           =   780
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Retracted"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   735
      TabIndex        =   49
      Top             =   2460
      Width           =   705
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   735
      TabIndex        =   48
      Top             =   2085
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Less rounded"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   735
      TabIndex        =   47
      Top             =   1710
      Width           =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "More rounded"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   735
      TabIndex        =   46
      Top             =   1335
      Width           =   990
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Aspriated"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   735
      TabIndex        =   45
      Top             =   960
      Width           =   660
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Voiced"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   735
      TabIndex        =   44
      Top             =   585
      Width           =   495
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "n8"
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
      Height          =   390
      Index           =   0
      Left            =   2235
      TabIndex        =   43
      Top             =   90
      Width           =   300
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Voiceless"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   735
      TabIndex        =   42
      Top             =   180
      Width           =   675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "|"
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
      Height          =   360
      Index           =   90
      Left            =   6135
      TabIndex        =   41
      Top             =   2355
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   ":"
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
      Height          =   360
      Index           =   89
      Left            =   6135
      TabIndex        =   40
      Top             =   1995
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "<"
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
      Index           =   88
      Left            =   6135
      TabIndex        =   39
      Top             =   1605
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " )"
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
      Height          =   360
      Index           =   87
      Left            =   6135
      TabIndex        =   38
      Top             =   1245
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " 6"
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
      Height          =   435
      Index           =   86
      Left            =   6135
      TabIndex        =   37
      Top             =   795
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " °"
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
      Height          =   420
      Index           =   85
      Left            =   6135
      TabIndex        =   36
      Top             =   450
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " 0"
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
      Height          =   405
      Index           =   84
      Left            =   6135
      TabIndex        =   35
      Top             =   75
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " ª"
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
      Height          =   450
      Index           =   83
      Left            =   3180
      TabIndex        =   34
      Top             =   4395
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " Á"
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
      Height          =   450
      Index           =   82
      Left            =   3165
      TabIndex        =   33
      Top             =   4035
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " £"
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
      Index           =   80
      Left            =   3180
      TabIndex        =   32
      Top             =   3345
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " ò"
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
      Index           =   79
      Left            =   3165
      TabIndex        =   31
      Top             =   2745
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "³"
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
      Index           =   78
      Left            =   3165
      TabIndex        =   30
      Top             =   2385
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "J"
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
      Height          =   345
      Index           =   76
      Left            =   3165
      TabIndex        =   29
      Top             =   1620
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "W"
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
      Height          =   405
      Index           =   75
      Left            =   3180
      TabIndex        =   28
      Top             =   1215
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " Ñ"
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
      Index           =   74
      Left            =   3165
      TabIndex        =   27
      Top             =   825
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " ¼"
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
      Height          =   480
      Index           =   73
      Left            =   3165
      TabIndex        =   26
      Top             =   450
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " -"
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
      Height          =   390
      Index           =   72
      Left            =   3165
      TabIndex        =   25
      Top             =   75
      Width           =   285
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      Caption         =   " Õ "
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
      Height          =   390
      Index           =   71
      Left            =   480
      TabIndex        =   24
      Top             =   4485
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " 9"
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
      Height          =   435
      Index           =   70
      Left            =   480
      TabIndex        =   23
      Top             =   4050
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " `"
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
      Height          =   435
      Index           =   69
      Left            =   480
      TabIndex        =   22
      Top             =   3675
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " _"
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
      Index           =   67
      Left            =   480
      TabIndex        =   21
      Top             =   2745
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " ="
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
      Index           =   66
      Left            =   480
      TabIndex        =   20
      Top             =   2340
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " +"
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
      Height          =   390
      Index           =   65
      Left            =   480
      TabIndex        =   19
      Top             =   1920
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " 7"
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
      Height          =   540
      Index           =   64
      Left            =   480
      TabIndex        =   18
      Top             =   1545
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " ¦"
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
      Height          =   390
      Index           =   63
      Left            =   480
      TabIndex        =   17
      Top             =   1170
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "H"
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
      Height          =   360
      Index           =   62
      Left            =   480
      TabIndex        =   16
      Top             =   870
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " ¤"
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
      Index           =   61
      Left            =   480
      TabIndex        =   15
      Top             =   450
      Width           =   285
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "7"
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
      Height          =   405
      Index           =   7
      Left            =   2580
      TabIndex        =   14
      Top             =   1590
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ì"
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
      Height          =   345
      Index           =   77
      Left            =   3165
      TabIndex        =   13
      Top             =   2010
      Width           =   285
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d8"
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
      Height          =   390
      Index           =   1
      Left            =   2580
      TabIndex        =   12
      Top             =   90
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "t¤"
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
      Height          =   390
      Index           =   3
      Left            =   2580
      TabIndex        =   11
      Top             =   465
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "dH"
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
      Height          =   390
      Index           =   5
      Left            =   2580
      TabIndex        =   10
      Top             =   855
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "a-"
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
      Height          =   390
      Index           =   16
      Left            =   5520
      TabIndex        =   9
      Top             =   105
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "a¼"
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
      Height          =   390
      Index           =   18
      Left            =   5520
      TabIndex        =   8
      Top             =   495
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "dÑ"
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
      Height          =   390
      Index           =   20
      Left            =   5520
      TabIndex        =   7
      Top             =   840
      Width           =   195
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "dW"
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
      Height          =   390
      Index           =   22
      Left            =   5520
      TabIndex        =   6
      Top             =   1230
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "dJ"
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
      Height          =   390
      Index           =   24
      Left            =   5520
      TabIndex        =   5
      Top             =   1605
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "dì"
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
      Height          =   390
      Index           =   26
      Left            =   5520
      TabIndex        =   4
      Top             =   1980
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d³"
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
      Height          =   390
      Index           =   28
      Left            =   5520
      TabIndex        =   3
      Top             =   2355
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d0"
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
      Height          =   390
      Index           =   37
      Left            =   8490
      TabIndex        =   2
      Top             =   90
      Width           =   285
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d°"
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
      Height          =   390
      Index           =   39
      Left            =   8490
      TabIndex        =   1
      Top             =   465
      Width           =   285
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d6"
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
      Height          =   390
      Index           =   41
      Left            =   8490
      TabIndex        =   0
      Top             =   825
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[]¹"
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
      Height          =   375
      Index           =   0
      Left            =   180
      TabIndex        =   112
      Top             =   135
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " 8"
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
      Index           =   60
      Left            =   405
      TabIndex        =   111
      Top             =   105
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ Í]"
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
      Height          =   375
      Index           =   21
      Left            =   180
      TabIndex        =   119
      Top             =   3720
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[›]"
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
      Height          =   375
      Index           =   19
      Left            =   180
      TabIndex        =   118
      Top             =   2340
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[‹]"
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
      Height          =   375
      Index           =   18
      Left            =   180
      TabIndex        =   117
      Top             =   1950
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ œ]"
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
      Height          =   375
      Index           =   16
      Left            =   165
      TabIndex        =   115
      Top             =   1245
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[H]"
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
      Height          =   375
      Index           =   15
      Left            =   180
      TabIndex        =   114
      Top             =   840
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[}]"
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
      Height          =   375
      Index           =   47
      Left            =   2895
      TabIndex        =   129
      Top             =   3705
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[{]"
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
      Height          =   375
      Index           =   46
      Left            =   2895
      TabIndex        =   128
      Top             =   3345
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ £]"
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
      Height          =   375
      Index           =   44
      Left            =   2880
      TabIndex        =   126
      Top             =   2685
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[P]"
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
      Height          =   375
      Index           =   41
      Left            =   2880
      TabIndex        =   125
      Top             =   2340
      Width           =   300
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   38
      Left            =   2880
      TabIndex        =   124
      Top             =   1995
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[Z]"
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
      Height          =   375
      Index           =   37
      Left            =   2880
      TabIndex        =   123
      Top             =   1620
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[X]"
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
      Height          =   375
      Index           =   36
      Left            =   2895
      TabIndex        =   122
      Top             =   1245
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ ¨]"
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
      Height          =   375
      Index           =   17
      Left            =   2880
      TabIndex        =   116
      Top             =   4455
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ ]"
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
      Height          =   375
      Index           =   45
      Left            =   2880
      TabIndex        =   127
      Top             =   3000
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ Ž]"
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
      Height          =   375
      Index           =   20
      Left            =   180
      TabIndex        =   131
      Top             =   4470
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ ù]"
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
      Height          =   375
      Index           =   22
      Left            =   2880
      TabIndex        =   121
      Top             =   480
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ ^]"
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
      Height          =   375
      Index           =   3
      Left            =   2895
      TabIndex        =   120
      Top             =   105
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[]²"
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
      Height          =   375
      Index           =   50
      Left            =   5865
      TabIndex        =   134
      Top             =   2370
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ ™]"
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
      Height          =   375
      Index           =   49
      Left            =   5880
      TabIndex        =   133
      Top             =   1245
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ Œ]"
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
      Height          =   375
      Index           =   23
      Left            =   5880
      TabIndex        =   132
      Top             =   120
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ ¦]"
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
      Height          =   405
      Index           =   48
      Left            =   2880
      TabIndex        =   130
      Top             =   4065
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[N]"
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
      Height          =   375
      Index           =   51
      Left            =   1830
      TabIndex        =   136
      Top             =   75
      Width           =   330
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[ä]"
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
      Height          =   375
      Index           =   95
      Left            =   1860
      TabIndex        =   142
      Top             =   4440
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[tH]"
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
      Height          =   375
      Index           =   94
      Left            =   1815
      TabIndex        =   141
      Top             =   825
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[r]"
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
      Height          =   375
      Index           =   93
      Left            =   1830
      TabIndex        =   140
      Top             =   1200
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[u<]"
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
      Height          =   375
      Index           =   92
      Left            =   1800
      TabIndex        =   139
      Top             =   1935
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[i>]"
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
      Height          =   375
      Index           =   91
      Left            =   1800
      TabIndex        =   138
      Top             =   2325
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[rÎ]"
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
      Height          =   375
      Index           =   52
      Left            =   1845
      TabIndex        =   137
      Top             =   3675
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[e¦]"
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
      Height          =   405
      Index           =   107
      Left            =   5265
      TabIndex        =   154
      Top             =   4050
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[a_]"
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
      Height          =   375
      Index           =   106
      Left            =   4695
      TabIndex        =   153
      Top             =   75
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[aÙ]"
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
      Height          =   375
      Index           =   105
      Left            =   4680
      TabIndex        =   152
      Top             =   465
      Width           =   315
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[e¨]"
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
      Height          =   375
      Index           =   103
      Left            =   5265
      TabIndex        =   150
      Top             =   4425
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[tX]"
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
      Height          =   375
      Index           =   102
      Left            =   4650
      TabIndex        =   149
      Top             =   1200
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[tZ]"
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
      Height          =   375
      Index           =   101
      Left            =   4650
      TabIndex        =   148
      Top             =   1575
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[t]"
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
      Height          =   375
      Index           =   100
      Left            =   4680
      TabIndex        =   147
      Top             =   1980
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[tP]"
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
      Height          =   375
      Index           =   99
      Left            =   4680
      TabIndex        =   146
      Top             =   2310
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[é]"
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
      Height          =   375
      Index           =   98
      Left            =   5250
      TabIndex        =   145
      Top             =   2700
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[e{]"
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
      Height          =   375
      Index           =   97
      Left            =   5190
      TabIndex        =   144
      Top             =   3330
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[e}]"
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
      Height          =   375
      Index           =   96
      Left            =   5190
      TabIndex        =   143
      Top             =   3705
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[t]"
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
      Height          =   375
      Index           =   104
      Left            =   5250
      TabIndex        =   151
      Top             =   2985
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[tŒ]"
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
      Height          =   375
      Index           =   110
      Left            =   7800
      TabIndex        =   157
      Top             =   60
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[e™]"
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
      Height          =   375
      Index           =   109
      Left            =   7800
      TabIndex        =   156
      Top             =   1215
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "[a ]"
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
      Height          =   375
      Index           =   108
      Left            =   7830
      TabIndex        =   155
      Top             =   2340
      Width           =   405
   End
End
Attribute VB_Name = "frmDispAmerDia2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
'* frmDispSILAmerDia version info:
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
  Set frmDispAmerDia2 = Nothing
  Erase CharDesc

End Sub

Private Sub Form_Resize()

  On Error Resume Next
  If WindowState > vbNormal Then Exit Sub
  If Height > FrmMaxHeight Then Height = FrmMaxHeight
  If Width > FrmMaxWidth Then Width = FrmMaxWidth

End Sub


