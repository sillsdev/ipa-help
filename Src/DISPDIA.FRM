VERSION 5.00
Begin VB.Form frmDispDia 
   BorderStyle     =   0  'None
   Caption         =   "IPA Diacritics"
   ClientHeight    =   7185
   ClientLeft      =   4950
   ClientTop       =   2715
   ClientWidth     =   8985
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ASAP SILManuscript"
      Size            =   8.25
      Charset         =   2
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
   ScaleHeight     =   7185
   ScaleWidth      =   8985
   ShowInTaskbar   =   0   'False
   Begin VB.Line Line1 
      Index           =   0
      X1              =   510
      X2              =   8430
      Y1              =   255
      Y2              =   255
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   510
      X2              =   8430
      Y1              =   3255
      Y2              =   3255
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   510
      X2              =   8445
      Y1              =   4755
      Y2              =   4755
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   510
      X2              =   8430
      Y1              =   4380
      Y2              =   4380
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   510
      X2              =   8430
      Y1              =   4005
      Y2              =   4005
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   510
      X2              =   8430
      Y1              =   3630
      Y2              =   3630
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   510
      X2              =   8430
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   510
      X2              =   8430
      Y1              =   2505
      Y2              =   2505
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   510
      X2              =   8430
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   510
      X2              =   8430
      Y1              =   1755
      Y2              =   1755
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   510
      X2              =   8430
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   510
      X2              =   8430
      Y1              =   1005
      Y2              =   1005
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   510
      X2              =   8430
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d6"
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
      Index           =   41
      Left            =   8040
      TabIndex        =   68
      Top             =   975
      Width           =   285
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d∞"
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
      Index           =   39
      Left            =   8040
      TabIndex        =   65
      Top             =   615
      Width           =   285
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d0"
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
      Index           =   37
      Left            =   8040
      TabIndex        =   62
      Top             =   240
      Width           =   285
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d≥"
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
      Index           =   28
      Left            =   5385
      TabIndex        =   47
      Top             =   2505
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "dÏ"
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
      Index           =   26
      Left            =   5385
      TabIndex        =   44
      Top             =   2130
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "dJ"
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
      Index           =   24
      Left            =   5385
      TabIndex        =   41
      Top             =   1755
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "dW"
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
      Index           =   22
      Left            =   5385
      TabIndex        =   38
      Top             =   1380
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d—"
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
      Index           =   20
      Left            =   5385
      TabIndex        =   35
      Top             =   990
      Width           =   195
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "aº"
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
      Index           =   18
      Left            =   5385
      TabIndex        =   32
      Top             =   615
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "a-"
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
      Index           =   16
      Left            =   5385
      TabIndex        =   29
      Top             =   255
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "dH"
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
      Left            =   2550
      TabIndex        =   8
      Top             =   1005
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "t§"
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
      Left            =   2550
      TabIndex        =   5
      Top             =   615
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d8"
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
      Left            =   2550
      TabIndex        =   2
      Top             =   240
      Width           =   300
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " 8"
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
      Index           =   60
      Left            =   525
      TabIndex        =   0
      Top             =   225
      Width           =   435
   End
   Begin VB.Shape shpTipTrgr 
      Height          =   390
      Index           =   0
      Left            =   510
      Top             =   255
      Width           =   435
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Ï"
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
      Index           =   77
      Left            =   2985
      TabIndex        =   42
      Top             =   2160
      Width           =   285
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "ç7"
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
      Left            =   2550
      TabIndex        =   12
      Top             =   1740
      Width           =   300
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " §"
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
      Index           =   61
      Left            =   585
      TabIndex        =   3
      Top             =   600
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "H"
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
      Height          =   360
      Index           =   62
      Left            =   585
      TabIndex        =   6
      Top             =   1020
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " ¶"
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
      Index           =   63
      Left            =   585
      TabIndex        =   9
      Top             =   1320
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " 7"
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
      Height          =   540
      Index           =   64
      Left            =   585
      TabIndex        =   11
      Top             =   1695
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " +"
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
      Index           =   65
      Left            =   585
      TabIndex        =   13
      Top             =   2070
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " ="
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
      Index           =   66
      Left            =   585
      TabIndex        =   15
      Top             =   2490
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " _"
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
      Index           =   67
      Left            =   585
      TabIndex        =   17
      Top             =   2895
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " `"
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
      Index           =   69
      Left            =   585
      TabIndex        =   21
      Top             =   3585
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " 9"
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
      Index           =   70
      Left            =   585
      TabIndex        =   23
      Top             =   3960
      Width           =   285
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      Caption         =   " ’ "
      BeginProperty Font 
         Name            =   "ASAP SILManuscript"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   71
      Left            =   585
      TabIndex        =   25
      Top             =   4395
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " -"
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
      Index           =   72
      Left            =   2985
      TabIndex        =   27
      Top             =   225
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " º"
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
      Height          =   480
      Index           =   73
      Left            =   2985
      TabIndex        =   30
      Top             =   600
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " —"
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
      Index           =   74
      Left            =   2985
      TabIndex        =   33
      Top             =   975
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "W"
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
      Index           =   75
      Left            =   2985
      TabIndex        =   36
      Top             =   1365
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "J"
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
      Index           =   76
      Left            =   2985
      TabIndex        =   39
      Top             =   1770
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "≥"
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
      Index           =   78
      Left            =   2985
      TabIndex        =   45
      Top             =   2535
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " Ú"
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
      Index           =   79
      Left            =   2985
      TabIndex        =   48
      Top             =   2910
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " £"
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
      Index           =   80
      Left            =   2985
      TabIndex        =   50
      Top             =   3210
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " ¡"
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
      Height          =   435
      Index           =   82
      Left            =   2985
      TabIndex        =   56
      Top             =   3945
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   435
      Index           =   83
      Left            =   3000
      TabIndex        =   58
      Top             =   4305
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " 0"
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
      Index           =   84
      Left            =   5910
      TabIndex        =   60
      Top             =   225
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " ∞"
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
      Index           =   85
      Left            =   5910
      TabIndex        =   63
      Top             =   600
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " 6"
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
      Height          =   435
      Index           =   86
      Left            =   5910
      TabIndex        =   66
      Top             =   945
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " )"
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
      Height          =   360
      Index           =   87
      Left            =   5910
      TabIndex        =   69
      Top             =   1395
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "<"
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
      Index           =   88
      Left            =   5910
      TabIndex        =   71
      Top             =   1755
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   ":"
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
      Height          =   360
      Index           =   89
      Left            =   5910
      TabIndex        =   73
      Top             =   2145
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "|"
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
      Height          =   360
      Index           =   90
      Left            =   5910
      TabIndex        =   75
      Top             =   2505
      Width           =   285
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Voiceless"
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
      Left            =   1110
      TabIndex        =   111
      Top             =   360
      Width           =   675
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "n8"
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
      Index           =   0
      Left            =   2190
      TabIndex        =   1
      Top             =   240
      Width           =   300
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Voiced"
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
      Left            =   1110
      TabIndex        =   110
      Top             =   735
      Width           =   495
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Aspirated"
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
      Left            =   1110
      TabIndex        =   109
      Top             =   1110
      Width           =   660
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "More rounded"
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
      Left            =   1110
      TabIndex        =   108
      Top             =   1485
      Width           =   990
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Less rounded"
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
      Left            =   1110
      TabIndex        =   107
      Top             =   1860
      Width           =   960
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced"
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
      Left            =   1110
      TabIndex        =   106
      Top             =   2235
      Width           =   735
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Retracted"
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
      Left            =   1110
      TabIndex        =   105
      Top             =   2610
      Width           =   705
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Centralized"
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
      Left            =   1110
      TabIndex        =   104
      Top             =   2985
      Width           =   780
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mid-centralized"
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
      Left            =   1110
      TabIndex        =   103
      Top             =   3360
      Width           =   1065
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Syllabic"
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
      Left            =   1110
      TabIndex        =   102
      Top             =   3735
      Width           =   540
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Non-syllabic"
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
      Left            =   1110
      TabIndex        =   101
      Top             =   4110
      Width           =   855
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Rhoticity"
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
      Left            =   1110
      TabIndex        =   100
      Top             =   4485
      Width           =   615
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "s§"
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
      Left            =   2190
      TabIndex        =   4
      Top             =   615
      Width           =   285
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "tH"
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
      Left            =   2190
      TabIndex        =   7
      Top             =   1005
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "ç¶"
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
      Left            =   2550
      TabIndex        =   10
      Top             =   1350
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "u+"
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
      Left            =   2550
      TabIndex        =   14
      Top             =   2100
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "i="
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
      Left            =   2550
      TabIndex        =   16
      Top             =   2505
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "e_"
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
      Left            =   2550
      TabIndex        =   18
      Top             =   2880
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "e~"
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
      Index           =   11
      Left            =   2550
      TabIndex        =   20
      Top             =   3255
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "®`"
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
      Index           =   12
      Left            =   2550
      TabIndex        =   22
      Top             =   3600
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "e9"
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
      Index           =   13
      Left            =   2550
      TabIndex        =   24
      Top             =   4005
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "´’"
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
      Left            =   2550
      TabIndex        =   26
      Top             =   4380
      Width           =   300
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " ~"
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
      Index           =   68
      Left            =   570
      TabIndex        =   19
      Top             =   3300
      Width           =   285
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Breathy voiced"
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
      Index           =   24
      Left            =   3450
      TabIndex        =   99
      Top             =   360
      Width           =   1065
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Creaky voiced"
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
      Left            =   3450
      TabIndex        =   98
      Top             =   735
      Width           =   1020
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Linguolabial"
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
      Left            =   3450
      TabIndex        =   97
      Top             =   1110
      Width           =   840
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Labialized"
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
      Left            =   3450
      TabIndex        =   96
      Top             =   1485
      Width           =   705
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Palatalized"
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
      Left            =   3450
      TabIndex        =   95
      Top             =   1860
      Width           =   765
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Velarized"
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
      Left            =   3450
      TabIndex        =   94
      Top             =   2235
      Width           =   645
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pharyngealized"
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
      Left            =   3450
      TabIndex        =   93
      Top             =   2610
      Width           =   1080
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Velarized or Pharyngealized"
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
      Left            =   3450
      TabIndex        =   92
      Top             =   2985
      Width           =   1950
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Raised"
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
      Left            =   3450
      TabIndex        =   91
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Lowered"
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
      Left            =   3450
      TabIndex        =   90
      Top             =   3735
      Width           =   615
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced Tongue Root"
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
      Index           =   34
      Left            =   3450
      TabIndex        =   89
      Top             =   4110
      Width           =   1725
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Retracted Tongue Root"
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
      Left            =   3450
      TabIndex        =   88
      Top             =   4485
      Width           =   1695
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "b-"
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
      Index           =   15
      Left            =   4890
      TabIndex        =   28
      Top             =   255
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "bº"
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
      Index           =   17
      Left            =   4890
      TabIndex        =   31
      Top             =   615
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "t—"
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
      Index           =   19
      Left            =   4890
      TabIndex        =   34
      Top             =   990
      Width           =   285
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "tW"
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
      Index           =   21
      Left            =   4875
      TabIndex        =   37
      Top             =   1380
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "tJ"
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
      Index           =   23
      Left            =   4875
      TabIndex        =   40
      Top             =   1755
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "tÏ"
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
      Index           =   25
      Left            =   4875
      TabIndex        =   43
      Top             =   2130
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "t≥"
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
      Index           =   27
      Left            =   4890
      TabIndex        =   46
      Top             =   2505
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "lÚ"
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
      Index           =   29
      Left            =   5580
      TabIndex        =   49
      Top             =   2880
      Width           =   315
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "e£"
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
      Index           =   30
      Left            =   4710
      TabIndex        =   51
      Top             =   3240
      Width           =   225
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "("
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
      Left            =   5385
      TabIndex        =   87
      Top             =   3360
      Width           =   45
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "= voiced alveolar fricative)"
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
      Left            =   5805
      TabIndex        =   86
      Top             =   3360
      Width           =   1860
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "e¢"
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
      Index           =   32
      Left            =   4710
      TabIndex        =   54
      Top             =   3630
      Width           =   225
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "("
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
      Left            =   5385
      TabIndex        =   85
      Top             =   3735
      Width           =   45
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "B¢"
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
      Index           =   33
      Left            =   5505
      TabIndex        =   55
      Top             =   3630
      Width           =   300
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "= voiced bilabial approximant)"
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
      Left            =   5805
      TabIndex        =   84
      Top             =   3735
      Width           =   2085
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "e¡"
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
      Index           =   34
      Left            =   5505
      TabIndex        =   57
      Top             =   3975
      Width           =   285
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
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
      Index           =   35
      Left            =   5505
      TabIndex        =   59
      Top             =   4350
      Width           =   285
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   " ¢"
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
      Index           =   81
      Left            =   2985
      TabIndex        =   53
      Top             =   3585
      Width           =   285
   End
   Begin VB.Label label1 
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
      Index           =   53
      Left            =   6345
      TabIndex        =   83
      Top             =   360
      Width           =   465
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Apical"
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
      Left            =   6345
      TabIndex        =   82
      Top             =   735
      Width           =   435
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Laminal"
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
      Index           =   55
      Left            =   6345
      TabIndex        =   81
      Top             =   1110
      Width           =   540
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nasalized"
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
      Left            =   6345
      TabIndex        =   80
      Top             =   1485
      Width           =   690
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nasal release"
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
      Index           =   57
      Left            =   6345
      TabIndex        =   79
      Top             =   1860
      Width           =   960
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Lateral release"
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
      Left            =   6345
      TabIndex        =   78
      Top             =   2235
      Width           =   1035
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "No audible release"
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
      Index           =   59
      Left            =   6345
      TabIndex        =   77
      Top             =   2610
      Width           =   1320
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "t0"
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
      Index           =   36
      Left            =   7710
      TabIndex        =   61
      Top             =   240
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "t∞"
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
      Index           =   38
      Left            =   7710
      TabIndex        =   64
      Top             =   615
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "t6"
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
      Index           =   40
      Left            =   7710
      TabIndex        =   67
      Top             =   975
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "e)"
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
      Index           =   42
      Left            =   8040
      TabIndex        =   70
      Top             =   1380
      Width           =   300
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d<"
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
      Index           =   43
      Left            =   8040
      TabIndex        =   72
      Top             =   1755
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d:"
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
      Index           =   44
      Left            =   8040
      TabIndex        =   74
      Top             =   2130
      Width           =   330
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "d|"
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
      Index           =   45
      Left            =   8040
      TabIndex        =   76
      Top             =   2505
      Width           =   330
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   8430
      X2              =   8430
      Y1              =   255
      Y2              =   4755
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   2910
      X2              =   2910
      Y1              =   255
      Y2              =   4755
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   510
      X2              =   510
      Y1              =   255
      Y2              =   4755
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   5850
      X2              =   5850
      Y1              =   255
      Y2              =   2880
   End
   Begin VB.Label lbldia 
      Appearance      =   0  'Flat
      Caption         =   "®£"
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
      Index           =   31
      Left            =   5505
      TabIndex        =   52
      Top             =   3255
      Width           =   285
   End
End
Attribute VB_Name = "frmDispDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************
'* frmDispDia version info:
'*  See basGlobals (Globals.bas).
'**************************************************

Option Explicit

Public bShowTip As Boolean
Public WavNamePart1 As String                       '* 'Dia' (Set in Form_Load)
Public CurrIndex As Integer

Private iMouseButton As Integer
Private Const INISection = "Diacritics"
Private Const ShowTipEntry = "ShowTip"
Private Const MaxDia = 46
Private Const DiaPerCol = 12
Private Const FrmMaxHeight = 5010
Private Const FrmMaxWidth = 8115
Private Const TBarButtons = "PlayOnly;PlayInterVocalic;PlaySeparator;Record;StopRec;PlayRec;PlayRecSpeaker;RecordSeparator;Exit;"
Private Const statusMsg1 = " Click on a diacritic to select it. "
Private Const statusMsg2 = "Press a position button to " & _
                           "see a list of words using " & _
                           "the selected diacritic."

Public Sub IPAHelpPrint(bToPrinter As Boolean, bColorBackground As Boolean)

  Dim pic As Picture
  Dim Capture As clsCapture
  
  On Error Resume Next
  Set pic = New StdPicture
  Set Capture = New clsCapture
  
  If (CurrIndex > -1) Then
    lbldia(CurrIndex).BackColor = vbButtonFace
    lbldia(CurrIndex).ForeColor = vbButtonText
    DoEvents
  End If
  
  If (bToPrinter) Then
    Set pic = Capture.CaptureWindowArea(Me, 34, 17, 529, 301, True)
    Capture.PrintChart pic, "Chart:" & vbTab & vbTab & "IPA Diacritics"
  Else
    Set pic = Capture.CaptureWindowArea(Me, 34, 17, 529, 301, _
              IIf(bColorBackground, False, True))
    Clipboard.Clear
    Clipboard.SetData pic, 2
  End If
  
  Set Capture = Nothing
  Set pic = Nothing
  
  If (CurrIndex > -1) Then
    lbldia(CurrIndex).BackColor = vbHighlight
    lbldia(CurrIndex).ForeColor = vbHighlightText
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

Private Sub LoadTipTriggerShapes()

  '*******************************************************
  '* This function will load all the shapes that sit on
  '* top of each IPA diacritic that, when the mouse is
  '* moved over, will trigger the tool tip to appear. The
  '* problem with using the label control of each
  '* character is that their heights are overlap each
  '* other thus making for a confused mess when trying
  '* to decide which character to display in the tooltip.
  '*******************************************************
  
  Dim i As Integer

  On Error Resume Next
  
  shpTipTrgr(0).BorderStyle = vbTransparent
  
  For i = 1 To MaxDia - 1
    Load shpTipTrgr(i)
    With shpTipTrgr(i)
      .Top = Line1(i Mod DiaPerCol).Y1
      .Left = Line2(i \ DiaPerCol).X1
      .ZOrder vbBringToFront
      .Visible = True
    End With
  Next i
  
End Sub

Private Function ReadDiacriticsFromDB() As Integer
  
'**************************************************
'* This function loads the characters from the INI
'* file and displays them on the form.
'**************************************************
  
  Dim i As Integer
  Dim sINIVals() As String
  
  On Error GoTo ReadDiacriticsFromDBErr
  ReadDiacriticsFromDB = False
  sINIVals = GetAllINISettings(gINIPath, INISection)
  If (IsNull(sINIVals)) Then Exit Function
  CurrIndex = 0
  
  For i = 0 To UBound(sINIVals, 1)
    With lbldia(i)
      Call GetCharFromINIStr(sINIVals(i, 1), lbldia(i))
      .BackColor = vbButtonFace
      .ForeColor = vbButtonText
      .ZOrder vbSendToBack
    End With
  Next i
  
  ReadDiacriticsFromDB = True
  Exit Function

ReadDiacriticsFromDBErr:
  
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
  '* current diacritic is changed to the next or previous
  '* enabled diacritic (depending upon which arrow is pressed).
  '************************************************************
  
  Dim i As Integer
  
  On Error Resume Next
  
  If KeyCode = vbKeyReturn Then Call Play(IIf((Shift And vbShiftMask) = 0, 0, 1))
  If Shift > 0 Then Exit Sub
  
  i = CurrIndex
  If KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then   'Is keypress right or down?
    i = IIf(CurrIndex = MaxDia - 1, 0, CurrIndex + 1)   'Move to next diacritic that is
    While Not lbldia(i).Enabled And i <> CurrIndex      '  enabled or until we've
      i = IIf(i = MaxDia - 1, 0, i + 1)                 '  gone full circle through
    Wend                                                '  all the diacritics.
  ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyLeft Then  'Is keypress left or up?
    i = IIf(CurrIndex = 0, MaxDia - 1, CurrIndex - 1)   'Move to previous diacritic that
    While Not lbldia(i).Enabled And i <> CurrIndex      '  is enabled or until we've
      i = IIf(i = 0, MaxDia - 1, i - 1)                 '  gone full circle through
    Wend                                                '  all the diacritics.
  End If
  
  If lbldia(i).Enabled Then _
    Call lblDia_MouseDown(i, vbLeftButton, 0, 0, 0)     'If we landed on an enabled diacritic,
                                                        '  act like we clicked on it.
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  
  On Error Resume Next
  If KeyAscii = vbKeyEscape Then Unload Me

End Sub

Private Sub Form_Load()
  
  mdiHelpCharts.panStatus.Visible = True
  Screen.MousePointer = vbHourglass
  WavNamePart1 = "Dia"
  Height = FrmMaxHeight
  Width = FrmMaxWidth
  Top = 0
  Left = 0
  CurrIndex = -1
  
  If ReadDiacriticsFromDB() Then
    Call LoadTipTriggerShapes
    If CurrIndex > -1 Then _
      Call lblDia_MouseDown(CurrIndex, vbLeftButton, 0, 0, 0)
  Else
    Screen.MousePointer = vbDefault
    MsgBox "Error reading diacritic information from the settings file.", _
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
  Set frmDispDia = Nothing

End Sub

'Private Sub Form_Resize()

  'On Error Resume Next
  'If WindowState > vbNormal Then Exit Sub
  'If Height > FrmMaxHeight Then Height = FrmMaxHeight
  'If Width > FrmMaxWidth Then Width = FrmMaxWidth
  
'End Sub

Private Sub lbldia_Click(Index As Integer)

  'Enable intervocalic button if intervocalic example exists
  Dim intvoc As Boolean
  Dim fileName As String
  'fileName = WavNamePart1 & "-" & Format$(Trim$(Str$(Index)), "00") & "b.wav"
  fileName = MakeInterVocName()
  mdiHelpCharts.TBar.Buttons(cInterVocBttn + 1).Enabled = FileExist(fileName)
  
  If iMouseButton = vbLeftButton Then
    '*********************************************
    '* Start double-click timer.
    '*********************************************
    mdiHelpCharts!Timer1.Interval = GetDoubleClickTime()
    mdiHelpCharts!Timer1.Enabled = True
  End If
  
End Sub

Private Sub lbldia_DblClick(Index As Integer)
  
  '*********************************************
  '* Double-click means not Single-click so
  '* turn the double-click timer off.
  '*********************************************
  If iMouseButton = vbLeftButton Then Call Play(1)

End Sub

Private Sub lblDia_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

  On Error Resume Next
  With lbldia(CurrIndex)
    '* Deselect previous selection.
    .ForeColor = vbButtonText
    .BackColor = vbButtonFace
  End With
  With lbldia(Index)
    '* Select new selection.
    .ForeColor = vbHighlightText
    .BackColor = vbHighlight
  End With
  
  CurrIndex = Index
  iMouseButton = Button
  
End Sub

Private Sub lbldia_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  On Error Resume Next
  gStatLine.SimpleText = GetCaptionFromTag(lbldia(Index).Tag)

End Sub

