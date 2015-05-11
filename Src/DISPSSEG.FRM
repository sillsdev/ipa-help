VERSION 5.00
Begin VB.Form frmDispSSeg 
   BorderStyle     =   0  'None
   Caption         =   "IPA Suprasegmentals"
   ClientHeight    =   5370
   ClientLeft      =   3660
   ClientTop       =   2370
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5370
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin VB.Label lblss 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "¨i.Qkt"
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
      Height          =   390
      Index           =   4
      Left            =   2055
      TabIndex        =   43
      Top             =   2340
      Width           =   990
   End
   Begin VB.Label lblss 
      Appearance      =   0  'Flat
      Caption         =   "eá"
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
      Height          =   390
      Index           =   3
      Left            =   2430
      TabIndex        =   44
      Top             =   1980
      Width           =   330
   End
   Begin VB.Label lblss 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "e>"
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
      Height          =   390
      Index           =   2
      Left            =   2430
      TabIndex        =   45
      Top             =   1650
      Width           =   330
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "ež"
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
      Height          =   390
      Index           =   9
      Left            =   4320
      TabIndex        =   35
      Top             =   2100
      Width           =   195
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "e™"
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
      Height          =   390
      Index           =   8
      Left            =   4320
      TabIndex        =   36
      Top             =   1770
      Width           =   195
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "e”"
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
      Height          =   390
      Index           =   7
      Left            =   4320
      TabIndex        =   37
      Top             =   1440
      Width           =   195
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "e"
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
      Height          =   390
      Index           =   6
      Left            =   4320
      TabIndex        =   38
      Top             =   1110
      Width           =   195
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "e("
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
      Height          =   390
      Index           =   14
      Left            =   6780
      TabIndex        =   23
      Top             =   2100
      Width           =   225
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "eˆ"
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
      Height          =   330
      Index           =   13
      Left            =   6780
      TabIndex        =   24
      Top             =   1785
      Width           =   210
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "e"
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
      Height          =   390
      Index           =   12
      Left            =   6795
      TabIndex        =   25
      Top             =   1440
      Width           =   180
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "eæ"
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
      Index           =   11
      Left            =   6795
      TabIndex        =   26
      Top             =   1095
      Width           =   180
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "1"
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
      Index           =   38
      Left            =   5310
      TabIndex        =   72
      Top             =   2100
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "2"
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
      Index           =   35
      Left            =   5310
      TabIndex        =   71
      Top             =   1770
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "3"
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
      Index           =   32
      Left            =   5310
      TabIndex        =   70
      Top             =   1440
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "4"
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
      Index           =   29
      Left            =   5310
      TabIndex        =   69
      Top             =   1110
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "5"
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
      Index           =   26
      Left            =   5310
      TabIndex        =   68
      Top             =   780
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Left            =   5115
      TabIndex        =   67
      Top             =   900
      Width           =   135
   End
   Begin VB.Label Label1 
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
      Index           =   44
      Left            =   7080
      TabIndex        =   66
      Top             =   900
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SUPRASEGMENTALS"
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
      Left            =   300
      TabIndex        =   65
      Top             =   420
      Width           =   1650
   End
   Begin VB.Label Label1 
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
      Height          =   345
      Index           =   10
      Left            =   195
      TabIndex        =   64
      Top             =   660
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Height          =   420
      Index           =   11
      Left            =   195
      TabIndex        =   63
      Top             =   945
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Height          =   345
      Index           =   12
      Left            =   195
      TabIndex        =   62
      Top             =   1320
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Height          =   345
      Index           =   13
      Left            =   195
      TabIndex        =   61
      Top             =   1650
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   15
      Left            =   195
      TabIndex        =   60
      Top             =   2280
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Height          =   345
      Index           =   16
      Left            =   195
      TabIndex        =   59
      Top             =   2715
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Left            =   195
      TabIndex        =   58
      Top             =   3330
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Height          =   345
      Index           =   14
      Left            =   195
      TabIndex        =   57
      Top             =   2040
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Left            =   675
      TabIndex        =   56
      Top             =   1110
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Left            =   675
      TabIndex        =   55
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label Label1 
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
      Left            =   675
      TabIndex        =   54
      Top             =   1770
      Width           =   630
   End
   Begin VB.Label Label1 
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
      Left            =   675
      TabIndex        =   53
      Top             =   780
      Width           =   960
   End
   Begin VB.Label Label1 
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
      Left            =   675
      TabIndex        =   52
      Top             =   2100
      Width           =   750
   End
   Begin VB.Label Label1 
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
      Left            =   675
      TabIndex        =   51
      Top             =   2430
      Width           =   990
   End
   Begin VB.Label Label1 
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
      Left            =   675
      TabIndex        =   50
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Label lblss 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "ÇfoUn«ÈtIS«n"
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
      Index           =   0
      Left            =   2025
      TabIndex        =   49
      Top             =   900
      Width           =   1800
   End
   Begin VB.Label Label1 
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
      Left            =   675
      TabIndex        =   48
      Top             =   3420
      Width           =   2025
   End
   Begin VB.Label Label1 
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
      Left            =   675
      TabIndex        =   47
      Top             =   3090
      Width           =   1665
   End
   Begin VB.Label lblss 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "eù"
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
      Height          =   390
      Index           =   1
      Left            =   2430
      TabIndex        =   46
      Top             =   1320
      Width           =   330
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "TONES && WORD ACCENTS"
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
      Left            =   5295
      TabIndex        =   42
      Top             =   255
      Width           =   2055
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LEVEL"
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
      Left            =   4785
      TabIndex        =   41
      Top             =   540
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CONTOUR"
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
      Left            =   7170
      TabIndex        =   40
      Top             =   540
      Width           =   810
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "e‰"
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
      Height          =   390
      Index           =   5
      Left            =   4320
      TabIndex        =   39
      Top             =   765
      Width           =   195
   End
   Begin VB.Label Label1 
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
      Left            =   4260
      TabIndex        =   34
      Top             =   2430
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Left            =   4260
      TabIndex        =   33
      Top             =   2760
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Left            =   4830
      TabIndex        =   32
      Top             =   765
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Left            =   4830
      TabIndex        =   31
      Top             =   1110
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Left            =   4830
      TabIndex        =   30
      Top             =   1440
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Left            =   4830
      TabIndex        =   29
      Top             =   1770
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Height          =   345
      Index           =   37
      Left            =   4830
      TabIndex        =   28
      Top             =   2100
      Width           =   285
   End
   Begin VB.Label lblss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "eó"
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
      Height          =   390
      Index           =   10
      Left            =   6795
      TabIndex        =   27
      Top             =   780
      Width           =   180
   End
   Begin VB.Label Label1 
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
      Left            =   6705
      TabIndex        =   22
      Top             =   2430
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Left            =   6705
      TabIndex        =   21
      Top             =   2760
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "è"
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
      Index           =   45
      Left            =   7245
      TabIndex        =   20
      Top             =   780
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "†"
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
      Index           =   47
      Left            =   7245
      TabIndex        =   19
      Top             =   1110
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Ø"
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
      Index           =   49
      Left            =   7245
      TabIndex        =   18
      Top             =   1440
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "€"
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
      Index           =   51
      Left            =   7245
      TabIndex        =   17
      Top             =   1770
      Width           =   285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Ê"
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
      Height          =   285
      Index           =   53
      Left            =   7245
      TabIndex        =   16
      Top             =   2100
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Height          =   345
      Index           =   17
      Left            =   195
      TabIndex        =   15
      Top             =   3060
      Width           =   285
   End
   Begin VB.Label Label1 
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
      Left            =   4635
      TabIndex        =   14
      Top             =   900
      Width           =   135
   End
   Begin VB.Label Label1 
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
      Left            =   5670
      TabIndex        =   13
      Top             =   900
      Width           =   705
   End
   Begin VB.Label Label1 
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
      Left            =   5670
      TabIndex        =   12
      Top             =   1230
      Width           =   330
   End
   Begin VB.Label Label1 
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
      Left            =   5670
      TabIndex        =   11
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label1 
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
      Left            =   5670
      TabIndex        =   10
      Top             =   1890
      Width           =   300
   End
   Begin VB.Label Label1 
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
      Left            =   5670
      TabIndex        =   9
      Top             =   2220
      Width           =   645
   End
   Begin VB.Label Label1 
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
      Left            =   4620
      TabIndex        =   8
      Top             =   2550
      Width           =   720
   End
   Begin VB.Label Label1 
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
      Left            =   4620
      TabIndex        =   7
      Top             =   2880
      Width           =   510
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rising"
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
      Left            =   7665
      TabIndex        =   6
      Top             =   900
      Width           =   435
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Falling"
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
      Left            =   7665
      TabIndex        =   5
      Top             =   1230
      Width           =   450
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "High rising"
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
      Left            =   7665
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Low rising"
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
      Index           =   52
      Left            =   7665
      TabIndex        =   3
      Top             =   1890
      Width           =   705
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rising-falling"
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
      Index           =   54
      Left            =   7665
      TabIndex        =   2
      Top             =   2220
      Width           =   885
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Global rise"
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
      Left            =   7245
      TabIndex        =   1
      Top             =   2550
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Global fall"
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
      Left            =   7245
      TabIndex        =   0
      Top             =   2880
      Width           =   690
   End
End
Attribute VB_Name = "frmDispSSeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************
'* frmDispSseg version info:
'*  See basGlobals (Globals.bas).
'**************************************************

Option Explicit

Public bShowTip As Boolean
Public WavNamePart1 As String                       '* 'Sseg' (Set in Form_Load)
Public CurrIndex As Integer

Private iMouseButton As Integer
Private Const INISection = "SupraSeg"
Private Const ShowTipEntry = "ShowTip"
Private Const MaxSSegs = 15
Private Const FrmMaxHeight = 4080
Private Const FrmMaxWidth = 8550
Private Const TBarButtons = "PlayOnly;PlaySeparator;Record;StopRec;PlayRec;PlayRecSpeaker;RecordSeparator;Exit;"
Private Const statusMsg1 = " Click on a suprasegmental to select it. "
Private Const statusMsg2 = "Press a position button to " & _
                           "see a list of words using " & _
                           "the selected suprasegmental."

Public Sub IPAHelpPrint(bToPrinter As Boolean, bColorBackground As Boolean)

  Dim pic As Picture
  Dim Capture As clsCapture
  
  On Error Resume Next
  Set pic = New StdPicture
  Set Capture = New clsCapture
  
  If (CurrIndex > -1) Then
    lblss(CurrIndex).BackColor = vbButtonFace
    lblss(CurrIndex).ForeColor = vbButtonText
    DoEvents
  End If
  
  If (bToPrinter) Then
    Set pic = Capture.CaptureWindowArea(Me, 15, 16, 560, 235, True)
    Capture.PrintChart pic, "Chart:" & vbTab & vbTab & "IPA Suprasegmentals, Tones and Word Accents"
  Else
    Set pic = Capture.CaptureWindowArea(Me, 15, 16, 560, 235, _
              IIf(bColorBackground, False, True))
    Clipboard.Clear
    Clipboard.SetData pic, 2
  End If
  
  Set Capture = Nothing
  Set pic = Nothing
  
  If (CurrIndex > -1) Then
    lblss(CurrIndex).BackColor = vbHighlight
    lblss(CurrIndex).ForeColor = vbHighlightText
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
  Call PlayWav(WavNamePart1 & "-" & Format$(Trim$(Str$(CurrIndex)), "00") & "A.wav")
  gMMCtrl.Tag = ""

End Sub

Private Function ReadSSegsFromDB() As Integer
  
'**************************************************
'* This function loads the characters from the INI
'* file and displays them on the form.
'**************************************************
  
  Dim i As Integer
  Dim sINIVals() As String
  
  On Error GoTo ReadSSegsFromDBErr
  ReadSSegsFromDB = False
  sINIVals = GetAllINISettings(gINIPath, INISection)
  If (IsNull(sINIVals)) Then Exit Function
  CurrIndex = 0
  
  For i = 0 To UBound(sINIVals, 1)
    With lblss(i)
      Call GetCharFromINIStr(sINIVals(i, 1), lblss(i))
      .BackColor = vbButtonFace
      .ForeColor = vbButtonText
    End With
  Next i
  
  ReadSSegsFromDB = True
  Exit Function

ReadSSegsFromDBErr:

End Function

Public Sub UpdateAfterRecordAndPlayback()

  On Error Resume Next
  mdiHelpCharts.EnableTBarButtons TBarButtons

End Sub

Private Sub Form_Activate()
  
  On Error Resume Next
  WindowState = vbMaximized
  gStatLine.SimpleText = ""
  
  With mdiHelpCharts
    .panStatus.Visible = True
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
  
  mdiHelpCharts!mnuExportBitmap.Visible = True
  mdiHelpCharts!mnuPrint(0).Visible = True
  mdiHelpCharts!mnuPrint(1).Visible = True
  
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
  '* current vowel is changed to the next or previous
  '* enabled vowel (depending upon which arrow is pressed).
  '************************************************************
  
  Dim i As Integer
  
  On Error Resume Next
  If KeyCode = vbKeyReturn Then Call Play(0)
  If Shift > 0 Then Exit Sub
  
  i = CurrIndex
  If KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then    'Is keypress right or down?
    i = IIf(CurrIndex = MaxSSegs - 1, 0, CurrIndex + 1)  'Move to next vowel that is
    While Not lblss(i).Enabled And i <> CurrIndex        '  enabled or until we've
      i = IIf(i = MaxSSegs - 1, 0, i + 1)                '  gone full circle through
    Wend                                                 '  all the SSegs.
  ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then   'Is keypress left or up?
    i = IIf(CurrIndex = 0, MaxSSegs - 1, CurrIndex - 1)  'Move to previous vowel that
    While Not lblss(i).Enabled And i <> CurrIndex        '  is enabled or until we've
      i = IIf(i = 0, MaxSSegs - 1, i - 1)                '  gone full circle through
    Wend                                                 '  all the SSegs.
  End If
  
  If lblss(i).Enabled Then _
    Call lblss_MouseDown(i, vbLeftButton, 0, 0, 0)       'If we landed on an enabled vowel,
                                                         '  act like we clicked on it.
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  
  On Error Resume Next
  If KeyAscii = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_Load()
  
  mdiHelpCharts.panStatus.Visible = True
  Screen.MousePointer = vbHourglass
  WavNamePart1 = "Sseg"
  Height = FrmMaxHeight
  Width = FrmMaxWidth
  Top = 0
  Left = 0
  CurrIndex = -1
  
  If ReadSSegsFromDB() Then
    If CurrIndex > -1 Then                           'If false then there's no SSegs in DB.
      Call lblss_MouseDown(CurrIndex, vbLeftButton, 0, 0, 0)
      'With lblss(CurrIndex)
      '  .BackColor = vbHighlight       'Select first enabled vowel.
      '  .ForeColor = vbHighlightText
      'End With
    End If
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Error reading suprasegmental information from the settings file.", _
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
  gStatLine.SimpleText = ""
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  On Error Resume Next
  Call Form_Deactivate
  Set frmDispSSeg = Nothing

End Sub

'Private Sub Form_Resize()

  'On Error Resume Next
  'If WindowState > vbNormal Then Exit Sub
  'If Height > FrmMaxHeight Then Height = FrmMaxHeight
  'If Width > FrmMaxWidth Then Width = FrmMaxWidth
  
'End Sub

Private Sub lblss_Click(Index As Integer)

  On Error Resume Next
  If iMouseButton = vbLeftButton Then Call Play(0)
  
End Sub

Private Sub lblss_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  On Error Resume Next
  With lblss(CurrIndex)
    '* Deselect previous selection.
    .BackColor = vbButtonFace
    .ForeColor = vbButtonText
  End With
  With lblss(Index)
    '* Select new selection.
    .BackColor = vbHighlight
    .ForeColor = vbHighlightText
  End With
  CurrIndex = Index
  iMouseButton = Button

End Sub

Private Sub lblss_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  On Error Resume Next
  gStatLine.SimpleText = GetCaptionFromTag(lblss(Index).Tag)

End Sub
