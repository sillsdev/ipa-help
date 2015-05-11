VERSION 5.00
Begin VB.Form frmDispCon 
   BorderStyle     =   0  'None
   Caption         =   "IPA Consonants"
   ClientHeight    =   6270
   ClientLeft      =   1455
   ClientTop       =   3105
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6270
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   " (NON-PULMONIC) "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   45
      TabIndex        =   78
      Top             =   3060
      Width           =   3855
      Begin VB.Label Con 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "p'"
         BeginProperty Font 
            Name            =   "ASAP SILManuscript"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   68
         Left            =   2760
         TabIndex        =   79
         Top             =   675
         Width           =   315
      End
      Begin VB.Label Con 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "×"
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
         Index           =   65
         Left            =   1305
         TabIndex        =   80
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Con 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "©"
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
         Index           =   66
         Left            =   1305
         TabIndex        =   81
         Top             =   1245
         Width           =   255
      End
      Begin VB.Label Con 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "ÿ'"
         BeginProperty Font 
            Name            =   "ASAP SILManuscript"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   69
         Left            =   2760
         TabIndex        =   82
         Top             =   930
         Width           =   315
      End
      Begin VB.Label Con 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "‡"
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
         Index           =   58
         Left            =   75
         TabIndex        =   111
         Top             =   420
         Width           =   255
      End
      Begin VB.Label Con 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "s'"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ASAP SILManuscript"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   71
         Left            =   2760
         TabIndex        =   109
         Top             =   1470
         Width           =   315
      End
      Begin VB.Label Con 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "k'"
         BeginProperty Font 
            Name            =   "ASAP SILManuscript"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   70
         Left            =   2760
         TabIndex        =   108
         Top             =   1200
         Width           =   315
      End
      Begin VB.Label Con 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "–"
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
         Index           =   59
         Left            =   75
         TabIndex        =   107
         Top             =   690
         Width           =   255
      End
      Begin VB.Label Con 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "º"
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
         Index           =   63
         Left            =   1305
         TabIndex        =   106
         Top             =   420
         Width           =   255
      End
      Begin VB.Label Con 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "—"
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
         Index           =   60
         Left            =   75
         TabIndex        =   105
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Con 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "ë"
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
         Index           =   64
         Left            =   1305
         TabIndex        =   104
         Top             =   690
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
         Height          =   315
         Index           =   61
         Left            =   75
         TabIndex        =   103
         Top             =   1245
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
         Left            =   75
         TabIndex        =   102
         Top             =   1530
         Width           =   255
      End
      Begin VB.Label Con 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "ý"
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
         Index           =   67
         Left            =   1305
         TabIndex        =   101
         Top             =   1530
         Width           =   255
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bilabial"
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
         Index           =   31
         Left            =   1605
         TabIndex        =   100
         Top             =   510
         Width           =   495
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Bilabial"
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
         Left            =   3120
         TabIndex        =   99
         Top             =   750
         Width           =   495
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Velar"
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
         Index           =   28
         Left            =   1605
         TabIndex        =   98
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Velar"
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
         Index           =   38
         Left            =   3120
         TabIndex        =   97
         Top             =   1275
         Width           =   360
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Alveolar fricative"
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
         Height          =   390
         Index           =   39
         Left            =   3120
         TabIndex        =   96
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alvlr lateral"
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
         Index           =   26
         Left            =   375
         TabIndex        =   95
         Top             =   1620
         Width           =   765
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Uvular"
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
         Left            =   1605
         TabIndex        =   94
         Top             =   1620
         Width           =   465
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dntl/alvlr"
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
         Left            =   1605
         TabIndex        =   93
         Top             =   780
         Width           =   645
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Palatal"
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
         Index           =   29
         Left            =   1605
         TabIndex        =   92
         Top             =   1050
         Width           =   480
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Dntl/alvlr"
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
         Index           =   37
         Left            =   3120
         TabIndex        =   91
         Top             =   1005
         Width           =   645
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Post) alvlr"
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
         Index           =   24
         Left            =   375
         TabIndex        =   90
         Top             =   1035
         Width           =   735
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Palatoalvlr"
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
         Index           =   25
         Left            =   375
         TabIndex        =   89
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bilabial"
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
         Index           =   22
         Left            =   375
         TabIndex        =   88
         Top             =   495
         Width           =   495
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dental"
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
         Index           =   23
         Left            =   375
         TabIndex        =   87
         Top             =   765
         Width           =   465
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "as in:"
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
         Index           =   35
         Left            =   3120
         TabIndex        =   86
         Top             =   495
         Width           =   375
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Clicks"
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
         Left            =   75
         TabIndex        =   85
         Top             =   240
         Width           =   420
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Voiced Implosives"
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
         Index           =   32
         Left            =   1305
         TabIndex        =   84
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Ejectives"
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
         Left            =   2760
         TabIndex        =   83
         Top             =   240
         Width           =   645
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "'"
         BeginProperty Font 
            Name            =   "ASAP SILManuscript"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   34
         Left            =   2775
         TabIndex        =   110
         Top             =   420
         Width           =   165
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " OTHER SYMBOLS "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3945
      TabIndex        =   112
      Top             =   3060
      Width           =   4965
      Begin VB.Label label1 
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
         Index           =   50
         Left            =   2985
         TabIndex        =   140
         Top             =   1725
         Width           =   135
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "é"
         BeginProperty Font 
            Name            =   "ASAP SILManuscript"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   52
         Left            =   3315
         TabIndex        =   139
         Top             =   1615
         Width           =   255
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
         Left            =   2655
         TabIndex        =   138
         Top             =   180
         Width           =   255
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Voiceless labial-velar fric."
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
         Index           =   47
         Left            =   360
         TabIndex        =   137
         Top             =   240
         Width           =   1770
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
         Index           =   44
         Left            =   360
         TabIndex        =   136
         Top             =   510
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
         Index           =   43
         Left            =   360
         TabIndex        =   135
         Top             =   780
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
         Index           =   42
         Left            =   360
         TabIndex        =   134
         Top             =   1050
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
         Index           =   41
         Left            =   360
         TabIndex        =   133
         Top             =   1320
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
         Index           =   40
         Left            =   360
         TabIndex        =   132
         Top             =   1590
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
         Index           =   45
         Left            =   2970
         TabIndex        =   131
         Top             =   240
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
         Index           =   46
         Left            =   2910
         TabIndex        =   130
         Top             =   510
         Width           =   1335
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
         Index           =   48
         Left            =   2910
         TabIndex        =   129
         Top             =   780
         Width           =   945
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
         Index           =   20
         Left            =   4200
         TabIndex        =   128
         Top             =   780
         Width           =   270
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Affricates and double articulations can be represented by two symbols joined by a tie bar if necessary."
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
         Height          =   660
         Index           =   49
         Left            =   2420
         TabIndex        =   127
         Top             =   1050
         Width           =   2500
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
         Left            =   30
         TabIndex        =   126
         Top             =   180
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
         Left            =   30
         TabIndex        =   125
         Top             =   450
         Width           =   255
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
         Height          =   330
         Index           =   74
         Left            =   30
         TabIndex        =   124
         Top             =   720
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
         Left            =   30
         TabIndex        =   123
         Top             =   990
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
         Left            =   30
         TabIndex        =   122
         Top             =   1260
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
         Height          =   330
         Index           =   77
         Left            =   30
         TabIndex        =   121
         Top             =   1530
         Width           =   255
      End
      Begin VB.Label Con 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "tés"
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
         Index           =   83
         Left            =   4440
         TabIndex        =   120
         Top             =   1650
         Width           =   375
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
         Index           =   54
         Left            =   4620
         TabIndex        =   118
         Top             =   720
         Width           =   255
      End
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
         Index           =   53
         Left            =   3960
         TabIndex        =   117
         Top             =   720
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
         Left            =   2520
         TabIndex        =   116
         Top             =   720
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
         Height          =   300
         Index           =   80
         Left            =   2520
         TabIndex        =   115
         Top             =   450
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
         Left            =   2400
         TabIndex        =   114
         Top             =   180
         Width           =   255
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "ƒ"
         BeginProperty Font 
            Name            =   "ASAP SILManuscript"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   51
         Left            =   2520
         TabIndex        =   113
         Top             =   1725
         Width           =   255
      End
      Begin VB.Label Con 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "kƒp"
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
         Height          =   330
         Index           =   82
         Left            =   4080
         TabIndex        =   119
         Top             =   1680
         Width           =   315
      End
   End
   Begin VB.Label lblSmile 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   5520
      TabIndex        =   142
      Top             =   5700
      Visible         =   0   'False
      Width           =   270
   End
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
      Left            =   5970
      TabIndex        =   141
      Top             =   5715
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   0
      X2              =   8925
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   0
      X2              =   8925
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   9
      Left            =   6270
      Top             =   960
      Width           =   540
   End
   Begin VB.Line Line1 
      Index           =   17
      X1              =   8280
      X2              =   8280
      Y1              =   -60
      Y2              =   3015
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   330
      Index           =   8
      Left            =   8655
      Top             =   285
      Width           =   285
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   330
      Index           =   7
      Left            =   7860
      Top             =   285
      Width           =   435
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   6
      Left            =   8295
      Top             =   960
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   5
      Left            =   8295
      Top             =   2340
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   330
      Index           =   4
      Left            =   7440
      Top             =   2685
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   3
      Left            =   7440
      Top             =   1995
      Width           =   1500
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   2
      Left            =   7440
      Top             =   615
      Width           =   1500
   End
   Begin VB.Line Line1 
      Index           =   24
      X1              =   7845
      X2              =   7845
      Y1              =   270
      Y2              =   600
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   1785
      X2              =   1785
      Y1              =   -60
      Y2              =   3015
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   330
      Index           =   1
      Left            =   1125
      Top             =   2685
      Width           =   1515
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   345
      Index           =   0
      Left            =   1125
      Top             =   1995
      Width           =   1515
   End
   Begin VB.Line Line1 
      Index           =   23
      X1              =   8925
      X2              =   8925
      Y1              =   -60
      Y2              =   3015
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   8925
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   0
      X2              =   8925
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   0
      X2              =   8925
      Y1              =   1635
      Y2              =   1635
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   0
      X2              =   8925
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   9
      X1              =   0
      X2              =   8925
      Y1              =   3015
      Y2              =   3015
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   0
      X2              =   8925
      Y1              =   2670
      Y2              =   2670
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   0
      X2              =   8925
      Y1              =   2325
      Y2              =   2325
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   1110
      X2              =   1110
      Y1              =   -30
      Y2              =   3015
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   7425
      X2              =   7425
      Y1              =   -60
      Y2              =   3015
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   2625
      X2              =   2625
      Y1              =   -75
      Y2              =   3015
   End
   Begin VB.Line Line1 
      Index           =   14
      X1              =   6255
      X2              =   6255
      Y1              =   -60
      Y2              =   3015
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   4710
      X2              =   4710
      Y1              =   -60
      Y2              =   3015
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   6795
      X2              =   6795
      Y1              =   -60
      Y2              =   3015
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   5580
      X2              =   5580
      Y1              =   -60
      Y2              =   3015
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(PULMONIC)"
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
      Left            =   0
      TabIndex        =   77
      Top             =   30
      Width           =   930
   End
   Begin VB.Line Line1 
      Index           =   18
      X1              =   8640
      X2              =   8640
      Y1              =   270
      Y2              =   600
   End
   Begin VB.Line Line1 
      Index           =   22
      X1              =   4035
      X2              =   4035
      Y1              =   -60
      Y2              =   270
   End
   Begin VB.Line Line1 
      Index           =   21
      X1              =   3270
      X2              =   3270
      Y1              =   -60
      Y2              =   270
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tap or Flap"
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
      Index           =   15
      Left            =   45
      TabIndex        =   0
      Top             =   1365
      Width           =   810
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Glottal"
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
      Index           =   11
      Left            =   8385
      TabIndex        =   32
      Top             =   30
      Width           =   465
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pharyngl"
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
      Index           =   10
      Left            =   7530
      TabIndex        =   1
      Top             =   30
      Width           =   645
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Lat. Approx."
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
      Left            =   45
      TabIndex        =   2
      Top             =   2745
      Width           =   855
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Approximant"
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
      Index           =   18
      Left            =   45
      TabIndex        =   3
      Top             =   2400
      Width           =   870
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fricative"
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
      Index           =   16
      Left            =   45
      TabIndex        =   5
      Top             =   1725
      Width           =   600
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Trill"
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
      Index           =   14
      Left            =   45
      TabIndex        =   6
      Top             =   1020
      Width           =   240
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nasal"
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
      Index           =   13
      Left            =   45
      TabIndex        =   7
      Top             =   675
      Width           =   405
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Plosive"
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
      Index           =   12
      Left            =   45
      TabIndex        =   8
      Top             =   345
      Width           =   510
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Uvular"
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
      Left            =   6870
      TabIndex        =   9
      Top             =   30
      Width           =   495
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Velar"
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
      Left            =   6345
      TabIndex        =   10
      Top             =   30
      Width           =   375
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Palatal"
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
      Left            =   5685
      TabIndex        =   11
      Top             =   30
      Width           =   495
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Postalv"
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
      Left            =   4095
      TabIndex        =   12
      Top             =   30
      Width           =   555
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Alveolar"
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
      Left            =   3360
      TabIndex        =   13
      Top             =   30
      Width           =   585
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dental"
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
      Left            =   2700
      TabIndex        =   14
      Top             =   30
      Width           =   495
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Labioden"
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
      Left            =   1860
      TabIndex        =   15
      Top             =   30
      Width           =   660
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bilabial"
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
      Left            =   1170
      TabIndex        =   16
      Top             =   30
      Width           =   525
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "p"
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
      Index           =   0
      Left            =   1200
      TabIndex        =   17
      Top             =   285
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "b"
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
      Index           =   1
      Left            =   1515
      TabIndex        =   18
      Top             =   285
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ú"
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
      Index           =   46
      Left            =   8640
      TabIndex        =   19
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "h"
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
      Index           =   45
      Left            =   8325
      TabIndex        =   20
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "À"
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
      Index           =   44
      Left            =   7920
      TabIndex        =   21
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ð"
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
      Index           =   43
      Left            =   7560
      TabIndex        =   22
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Ò"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   42
      Left            =   7155
      TabIndex        =   23
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   41
      Left            =   6885
      TabIndex        =   24
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Ä"
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
      Index           =   40
      Left            =   6540
      TabIndex        =   25
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   315
      Index           =   39
      Left            =   6270
      TabIndex        =   26
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Æ"
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
      Index           =   38
      Left            =   5985
      TabIndex        =   76
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "C"
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
      Index           =   37
      Left            =   5670
      TabIndex        =   75
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   ";"
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
      Index           =   57
      Left            =   6540
      TabIndex        =   74
      Top             =   2685
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "´"
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
      Index           =   56
      Left            =   5985
      TabIndex        =   73
      Top             =   2685
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "l"
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
      Index           =   54
      Left            =   3720
      TabIndex        =   72
      Top             =   2685
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "å"
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
      Index           =   53
      Left            =   6540
      TabIndex        =   71
      Top             =   2340
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "j"
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
      Index           =   52
      Left            =   5985
      TabIndex        =   70
      Top             =   2340
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "¨"
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
      Index           =   50
      Left            =   3720
      TabIndex        =   69
      Top             =   2340
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "V"
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
      Index           =   49
      Left            =   2295
      TabIndex        =   68
      Top             =   2340
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "L"
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
      Index           =   48
      Left            =   3720
      TabIndex        =   67
      Top             =   1995
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Â"
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
      Index           =   47
      Left            =   3360
      TabIndex        =   66
      Top             =   1995
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Z"
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
      Index           =   34
      Left            =   4425
      TabIndex        =   65
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   315
      Index           =   33
      Left            =   4110
      TabIndex        =   64
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "z"
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
      Index           =   32
      Left            =   3720
      TabIndex        =   63
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "s"
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
      Index           =   31
      Left            =   3360
      TabIndex        =   62
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "D"
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
      Index           =   30
      Left            =   3000
      TabIndex        =   61
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "T"
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
      Index           =   29
      Left            =   2640
      TabIndex        =   60
      Top             =   1650
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "v"
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
      Index           =   28
      Left            =   2295
      TabIndex        =   59
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "f"
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
      Index           =   27
      Left            =   1935
      TabIndex        =   58
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "B"
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
      Index           =   26
      Left            =   1515
      TabIndex        =   57
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "R"
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
      Index           =   23
      Left            =   3765
      TabIndex        =   55
      Top             =   1305
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "}"
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
      Index           =   24
      Left            =   5220
      TabIndex        =   54
      Top             =   1305
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   16
      Left            =   5220
      TabIndex        =   53
      Top             =   620
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "{"
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
      Index           =   22
      Left            =   7155
      TabIndex        =   52
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "r"
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
      Index           =   21
      Left            =   3765
      TabIndex        =   51
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "õ"
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
      Index           =   20
      Left            =   1515
      TabIndex        =   50
      Top             =   960
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ê"
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
      Index           =   5
      Left            =   5220
      TabIndex        =   49
      Top             =   285
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "²"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   19
      Left            =   7155
      TabIndex        =   48
      Top             =   620
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   18
      Left            =   6540
      TabIndex        =   47
      Top             =   620
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ø"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   17
      Left            =   5985
      TabIndex        =   46
      Top             =   620
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   15
      Left            =   3750
      TabIndex        =   45
      Top             =   620
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   14
      Left            =   2295
      TabIndex        =   44
      Top             =   620
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Index           =   13
      Left            =   1515
      TabIndex        =   43
      Top             =   620
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ÿ"
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
      Index           =   4
      Left            =   4860
      TabIndex        =   42
      Top             =   285
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "?"
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
      Index           =   12
      Left            =   8370
      TabIndex        =   41
      Top             =   285
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "G"
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
      Index           =   11
      Left            =   7155
      TabIndex        =   40
      Top             =   285
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "q"
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
      Index           =   10
      Left            =   6885
      TabIndex        =   39
      Top             =   285
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "g"
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
      Index           =   9
      Left            =   6540
      TabIndex        =   38
      Top             =   285
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "k"
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
      Index           =   8
      Left            =   6270
      TabIndex        =   37
      Top             =   285
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ï"
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
      Index           =   7
      Left            =   5985
      TabIndex        =   36
      Top             =   285
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "c"
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
      Index           =   6
      Left            =   5670
      TabIndex        =   35
      Top             =   285
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "d"
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
      Index           =   3
      Left            =   3750
      TabIndex        =   34
      Top             =   285
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "t"
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
      Index           =   2
      Left            =   3405
      TabIndex        =   33
      Top             =   285
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "½"
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
      Index           =   36
      Left            =   5220
      TabIndex        =   30
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "§"
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
      Index           =   35
      Left            =   4860
      TabIndex        =   29
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Ó"
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
      Index           =   51
      Left            =   5220
      TabIndex        =   28
      Top             =   2340
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "ñ"
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
      Index           =   55
      Left            =   5220
      TabIndex        =   27
      Top             =   2685
      Width           =   255
   End
   Begin VB.Line Line1 
      Index           =   19
      X1              =   4035
      X2              =   4035
      Y1              =   1635
      Y2              =   1995
   End
   Begin VB.Line Line1 
      Index           =   20
      X1              =   3270
      X2              =   3270
      Y1              =   1635
      Y2              =   1995
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Retroflex"
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
      Left            =   4830
      TabIndex        =   31
      Top             =   30
      Width           =   645
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Lat. Fricative"
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
      Index           =   17
      Left            =   45
      TabIndex        =   4
      Top             =   2070
      Width           =   915
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "¸"
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   25
      Left            =   1200
      TabIndex        =   56
      Top             =   1650
      Width           =   255
   End
End
Attribute VB_Name = "frmDispCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************
'* frmDispCon version info:
'*  See basGlobals (Globals.bas).
'**************************************************

Option Explicit

Public bShowTip As Boolean
Public WavNamePart1 As String                       '* 'Con' (Set in Form_Load)
Public CurrIndex As Integer

Private iMouseButton As Integer
Private Const INISection = "Consonants"
Private Const ShowTipEntry = "ShowTip"
Private Const TBarButtons = "PlayOnly;PlayInterVocalic;PlaySeparator;Record;StopRec;PlayRec;PlayRecSpeaker;RecordSeparator;Test;TestSeparator;Exit;"
Private Const MaxCons = 84
Private Const FrmMaxWidth = 9045
Private Const FrmMaxHeight = 5550
Private Const statusMsg1 = " Click on a consonant to select it. "
Private Const statusMsg2 = "Press a position button to " & _
                           "see a list of words using " & _
                           "the selected consonant."

Public Sub IPAHelpPrint(bToPrinter As Boolean, bColorBackground As Boolean)

  Dim pic As Picture
  Dim Capture As clsCapture
  
  On Error Resume Next
  Set pic = New StdPicture
  Set Capture = New clsCapture
  
  If (CurrIndex > -1) Then
    Con(CurrIndex).BackColor = vbButtonFace
    Con(CurrIndex).ForeColor = vbButtonText
    DoEvents
  End If
  
  If (bToPrinter) Then
    Set pic = Capture.CaptureWindowArea(Me, 0&, 4&, 595&, 342&, True)
    Capture.PrintChart pic, "Chart:" & vbTab & vbTab & "IPA Consonants"
  Else
    Set pic = Capture.CaptureWindowArea(Me, 0&, 0&, 595&, 342&, _
              IIf(bColorBackground, False, True))
    Clipboard.Clear
    Clipboard.SetData pic, 2
  End If
  
  Set Capture = Nothing
  Set pic = Nothing
  
  If (CurrIndex > -1) Then
    Con(CurrIndex).BackColor = vbHighlight
    Con(CurrIndex).ForeColor = vbHighlightText
  End If
  
End Sub

Public Sub Play(iButton As Integer)

  '**************************************************
  '* This function allows mdiHelpCharts to play the
  '* wave file corresponding to the currently
  '* selected symbol, without knowing which form is
  '* active.
  '**************************************************

  On Error Resume Next
  If (Len(gMMCtrl.Tag) > 0) Then Exit Sub
  If (iButton = 1) Then mdiHelpCharts!Timer1.Enabled = False
  gMMCtrl.Tag = cMCIBusy
  gMMCtrl.Wait = True
  Call PlayWav(WavNamePart1 & "-" & Format$(Trim$(Str$(CurrIndex)), "00") & _
               IIf(iButton = 0, "a", "b") & ".wav")
  gMMCtrl.Tag = ""

End Sub

Private Function ReadConsFromDB()
  
'**************************************************
'* This function loads the characters from the INI
'* file and displays them on the form.
'**************************************************
  
  Dim i As Integer
  Dim sINIVals() As String
  
  On Error GoTo ReadConsFromDBErr
  ReadConsFromDB = False
  sINIVals = GetAllINISettings(gINIPath, INISection)
  CurrIndex = 0
  If (IsNull(sINIVals)) Then Exit Function
  
  For i = 0 To UBound(sINIVals, 1)
    With Con(i)
      Call GetCharFromINIStr(sINIVals(i, 1), Con(i))
      .BackColor = vbButtonFace
      .ForeColor = vbButtonText
      .ZOrder vbSendToBack
    End With
  Next i
  
  ReadConsFromDB = True
  Exit Function

ReadConsFromDBErr:

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
  Dim iConGrpInd As Integer
  Dim iGrpInd As Integer
  Dim iOldLeft As Integer
  Dim iOldTop As Long

  lblSmile.Visible = False
  lblFrown.Visible = False
  With Con(gItemNumber)
    .ForeColor = vbButtonText
    .BackColor = vbButtonFace
  End With
  
  CurrIndex = 0
  With Con(CurrIndex)
    .BackColor = vbHighlight
    .ForeColor = vbHighlightText
  End With
  
  '**********************************************
  '* Restore Cons to original location (if Brief
  '* test mode), Enabled and Visible state.
  '**********************************************
  iConGrpInd = 0
  iGrpInd = 0
  For i = 0 To MaxCons - 1
    If i = gPhonGrpColl("AllCon")(iConGrpInd) Then
      If gTestLayout = cTestBrief _
      And i = gTestGrp(iGrpInd) Then
        If gTestGrpCorrect(cLastTest, iGrpInd) <> vbChecked Then
          '* Restore location.  *
          With Con(i)
            .Visible = False
            iOldLeft = Val(Mid$(.Tag, 1, 4))
            iOldTop = Val(Mid$(.Tag, 5, 4))
            .Left = iOldLeft
            .Top = iOldTop
            .Tag = Mid(.Tag, 9)
          End With
        Else
          '* Enable untested cons (in this group) with sounds *
          Con(i).Enabled = True
        End If
        iGrpInd = IIf(iGrpInd = UBound(gTestGrp), _
                      iGrpInd, iGrpInd + 1)
      Else
        '* Enable other Cons with sounds. *
        Con(i).Enabled = True
      End If
      iConGrpInd = IIf(iConGrpInd = UBound(gPhonGrpColl("AllCon")), _
                       iConGrpInd, iConGrpInd + 1)
    End If
  Next i
  
  If gTestLayout = cTestBrief Then
    '* Make all Cons visible.  *
    For i = 0 To MaxCons - 1
      Con(i).Visible = True
    Next i
  End If

  '**********************************************
  '* If in brief mode, replace the scenery.
  '**********************************************
  If gTestLayout = cTestBrief Then
    For i = 0 To 54
      Select Case i
        Case 0 To 9
          Line1(i).Visible = True
          Label1(i).Visible = True
          Shape1(i).Visible = True
        Case 10 To 24
          Line1(i).Visible = True
          Label1(i).Visible = True
        Case 25 To 54
          Label1(i).Visible = True
      End Select
    Next i
    With Frame1
      .Left = 45
      .Top = 3060
      .Visible = True
    End With
    With Frame2
      .Left = 3945
      .Top = 3060
      .Visible = True
    End With
  End If
  
  mdiHelpCharts.EnableTBarButtons TBarButtons
  
End Sub

Public Sub UpdateFormForTest()
  
'***********************************************************************
'* This function performs any changes to the form
'* necessary in preparation for testing.
'*
'* Whether a con is visible or not is now determined by its status in
'* gTestGrpCorrect. This means that gTestGrp does not change depending on
'* how many the user got correct. This also allows the indexing to be
'* the same for both gTestGrp and gTestGrpCorrect. CLW 1/27/99
'***********************************************************************

  Dim i As Integer
  Dim iPrev As Integer
  Dim iTestCount As Integer      '* Number of cons tested on CLW 1/27/99
  Dim iGrpInd As Integer
  Dim iMainCon As Integer
  Dim iNonPulCon As Integer
  Dim iOtherCon As Integer
  Dim iMainRowLength As Integer
  Dim iNonPulRowLength As Integer
  Dim iOtherRowLength As Integer
  Dim iCol As Integer
  Dim iLeftStart As Integer
  Dim iTop As Integer
  Dim iOldLeft As Integer
  Dim iOldTop As Integer
  
  iMainCon = 0
  iNonPulCon = 0
  iOtherCon = 0
  
  With Con(CurrIndex)
    .BackColor = vbButtonFace
    .ForeColor = vbButtonText
  End With
  
  '**********************************************
  '* First disable all Cons.
  '* If Brief mode, make invisible.
  '**********************************************
  For i = 0 To MaxCons - 1
    With Con(i)
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
      For i = 0 To 54
        Select Case i
          Case 0 To 9
            Line1(i).Visible = False
            Label1(i).Visible = False
            Shape1(i).Visible = False
          Case 10 To 24
            Line1(i).Visible = False
            Label1(i).Visible = False
          Case 25 To 54
            Label1(i).Visible = False
        End Select
      Next i
      Select Case gTestGrpName
        Case "AllCon"
          ' Find frame divisions in test group
          For i = 0 To UBound(gTestGrp)
            If gTestGrpCorrect(cLastTest, i) <> vbChecked Then
              Select Case gTestGrp(i)
                Case 0 To 57:  iMainCon = iMainCon + 1
                Case 58 To 71: iNonPulCon = iNonPulCon + 1
                Case 72 To 83: iOtherCon = iOtherCon + 1
              End Select
            End If
          Next i
          
          iLeftStart = (Width - iMainRowLength * 315) / 2
          iTop = (Height - Frame1.Height - (Int((iMainCon + 9) / 10)) * 425) / 2

        Case "NonPulCon"
          Frame2.Visible = False
          With Frame1
            .Top = Frame1.Top / 2
            .Left = Frame2.Width / 2
          End With
          '* Calculate number of NonPulCons in visible group
          For i = 0 To UBound(gTestGrp)
            If gTestGrpCorrect(cLastTest, i) <> vbChecked Then _
              iNonPulCon = iNonPulCon + 1
          Next i
        
        Case "OtherCon"
          Frame1.Visible = False
          With Frame2
            .Top = Frame2.Top / 2
            .Left = Frame1.Width / 2
          End With
          '* Calculate number of OtherCons in visible group
          For i = 0 To UBound(gTestGrp)
            If gTestGrpCorrect(cLastTest, i) <> vbChecked Then _
              iOtherCon = iOtherCon + 1
          Next i
        
        Case Else
          Frame1.Visible = False
          Frame2.Visible = False
          '* Calculate number of MainCons in visible group
          For i = 0 To UBound(gTestGrp)
            If gTestGrpCorrect(cLastTest, i) <> vbChecked Then _
              iMainCon = iMainCon + 1
          Next i
          iTop = (Height - (Int((iMainCon + 9) / 10)) * 425) / 2
      End Select
      
      '********************************************
      '* Rearrange test Cons in rows of 10,
      '* and make visible.
      '********************************************
      iMainRowLength = IIf(iMainCon < 10, iMainCon, 10)
      iNonPulRowLength = IIf(iNonPulCon < 10, iNonPulCon, 10)
      iOtherRowLength = IIf(iOtherCon < 10, iOtherCon, 10)
      
      For iGrpInd = 0 To UBound(gTestGrp)
        i = gTestGrp(iGrpInd)
        If iGrpInd = 0 Then iPrev = 0 _
        Else: iPrev = gTestGrp(iGrpInd - 1)
        'Initialize new container
        
        If iGrpInd = 0 Or Con(i).Container.Name <> Con(iPrev).Container.Name Then
          iCol = 0
          Select Case i
            Case 0 To 57:  iLeftStart = (Width - iMainRowLength * 315) / 2
            Case 58 To 71: iLeftStart = (Frame1.Width - iNonPulRowLength * 315) / 2
                           iTop = (Frame1.Height - (Int((iNonPulCon + 9) / 10)) * 425) / 2
            Case 72 To 83: iLeftStart = (Frame2.Width - iOtherRowLength * 315) / 2
                           iTop = (Frame2.Height - (Int((iOtherCon + 9) / 10)) * 425) / 2
          End Select
        End If
        
        If gTestGrpCorrect(cLastTest, iGrpInd) <> vbChecked Then
          'Begin new line
          If iCol = 10 Then
            iTop = iTop + 425
            iCol = 0
          End If
          'Position Cons
          With Con(i)
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
      Next iGrpInd
    
    Case cTestChart
      For iGrpInd = 0 To UBound(gTestGrp)
        If gTestGrpCorrect(cLastTest, iGrpInd) <> vbChecked Then
          i = gTestGrp(iGrpInd)
          Con(i).Enabled = True
        End If
      Next iGrpInd
  End Select
  
End Sub

Private Sub Con_Click(Index As Integer)

  If iMouseButton = vbLeftButton Then
    '**********************************************
    '* If in test mode ...
    '**********************************************
    If gTestActive Then
      '*********************************************
      '* ... check to see if correct IPA character
      '* selected. If so, put smile on top of
      '* character.
      '*********************************************
      If gItemNumber = Index Then
        lblFrown.Visible = False
        With lblSmile
          .Visible = False
          '****************************************
          '* Update container when switching
          '* between 'main consonant',
          '* 'nonpulmonic', and 'other' frames.
          '****************************************
          If Not (.Container Is Con(Index).Container) Then
            Set .Container = Con(Index).Container
          End If
          .Height = Con(Index).Height
          .Width = Con(Index).Width
          .Top = Con(Index).Top
          .Left = Con(Index).Left
          .Visible = True
          .ZOrder vbBringToFront
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
          .Visible = False
          If Not (.Container Is Con(Index).Container) Then
            Set .Container = Con(Index).Container
          End If
          .Height = Con(Index).Height
          .Width = Con(Index).Width
          .Top = Con(Index).Top
          .Left = Con(Index).Left
          .Visible = True
          .ZOrder vbBringToFront
        End With
        Call MarkItemCorrect(vbUnchecked, Index)
      End If
    Else
      '*********************************************
      '* If not in test mode, start double-click
      '* timer.
      '*********************************************
      With mdiHelpCharts!Timer1
        .Interval = GetDoubleClickTime()
        .Enabled = True
      End With
    End If
  End If
  
End Sub

Private Sub Con_DblClick(Index As Integer)
  
  On Error Resume Next
  If iMouseButton = vbLeftButton Then Call Play(1)
  
End Sub

Private Sub Con_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  '***********************************************
  '* Make the character that CurrIndex points
  '* to not look selected and make the character
  '* just clicked on look selected by changing
  '* its background color.
  '***********************************************
  
  On Error Resume Next
  If gTestActive = False Then
    With Con(CurrIndex)
      '* Deselect previous selection.
      .BackColor = vbButtonFace
      .ForeColor = vbButtonText
    End With
    With Con(Index)
      '* Select new selection.
      .BackColor = vbHighlight
      .ForeColor = vbHighlightText
    End With
  End If
  CurrIndex = Index
  iMouseButton = Button

End Sub

Private Sub Con_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  On Error Resume Next
  If Not gTestActive Then gStatLine.SimpleText = GetCaptionFromTag(Con(Index).Tag)

End Sub

Private Sub Form_Activate()

  On Error Resume Next
  Show
  WindowState = vbMaximized
  gStatLine.SimpleText = ""
  
  With mdiHelpCharts
    !panStatus.Visible = True
    .ShowTBarButtons TBarButtons
    .EnableTBarButtons TBarButtons
    !mnuExportBitmap.Visible = True
    !mnuPrint(0).Visible = True
    !mnuPrint(1).Visible = True
    
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
    WindowState = vbMaximized
  End If

End Sub

Private Sub Form_Deactivate()

  On Error Resume Next
  gStatLine.SimpleText = ""
  mdiHelpCharts!mnuExportBitmap.Visible = False
  mdiHelpCharts!mnuPrint(0).Visible = False
  mdiHelpCharts!mnuPrint(1).Visible = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  '************************************************************
  '* This routine will monitor keypresses looking for arrow
  '* keys. If one of the arrow keys is pressed then the
  '* current consonant is changed to the next or previous
  '* enabled consonant (depending upon which arrow is pressed).
  '************************************************************
  
  Dim i As Integer
  
  On Error Resume Next
  If gTestActive Then Exit Sub
  
  i = CurrIndex
  Select Case KeyCode
    Case vbKeyDown, vbKeyRight                             'Is keypress right or down?
      i = IIf(CurrIndex = MaxCons - 1, 0, CurrIndex + 1)   'Move to next consonant that
      While Not Con(i).Enabled And i <> CurrIndex          '  is enabled or until we've
        i = IIf(i = MaxCons - 1, 0, i + 1)                 '  gone full circle through
      Wend                                                 '  all the consonants.
  
    Case vbKeyUp, vbKeyLeft                                'Is keypress left or up?
      i = IIf(CurrIndex = 0, MaxCons - 1, CurrIndex - 1)   'Move to previous consonant
      While Not Con(i).Enabled And i <> CurrIndex          '  that's enabled or until
        i = IIf(i = 0, MaxCons - 1, i - 1)                 '  we've gone full circle
      Wend                                                 '  through all the consonants.
  
    Case vbKeyReturn: Call Play(IIf((Shift And vbShiftMask) = 0, 0, 1))
  
  End Select
  
  '**************************************************************
  '* If we landed on an enabled Con  act like we clicked on it.
  '**************************************************************
  If Con(i).Enabled Then Call Con_MouseDown(i, vbLeftButton, 0, 0, 0)
                                                         
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
  WavNamePart1 = "Con"
  
  If WindowState = vbNormal Then
    Height = FrmMaxHeight
    Width = FrmMaxWidth
    Top = 0
    Left = 0
  End If
  CurrIndex = -1

  If ReadConsFromDB() Then
    'If next line is false then there's no consonants in DB.
    If (CurrIndex > -1 And Not gTestActive) Then
      With Con(CurrIndex)
        .BackColor = vbHighlight                 'Select first enabled consonant.
        .ForeColor = vbHighlightText
      End With
    End If
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Error reading consonant information from the settings file.", _
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

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  On Error Resume Next
  Call Form_Deactivate
  Set frmDispCon = Nothing

End Sub

