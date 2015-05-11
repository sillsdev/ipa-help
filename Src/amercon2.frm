VERSION 5.00
Begin VB.Form frmDispAmerCon2 
   BorderStyle     =   0  'None
   Caption         =   "Americanist - IPA Consonants"
   ClientHeight    =   10710
   ClientLeft      =   0
   ClientTop       =   1800
   ClientWidth     =   11475
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10710
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   1380
      Index           =   22
      Left            =   7995
      Top             =   2310
      Width           =   735
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   127
      Left            =   8790
      TabIndex        =   0
      Top             =   3675
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÔÉ∆"
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
      Index           =   325
      Left            =   7635
      TabIndex        =   330
      Top             =   5790
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s“Æ"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   43
      Left            =   6570
      TabIndex        =   329
      Top             =   4650
      Width           =   120
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pÉf"
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
      Index           =   115
      Left            =   1665
      TabIndex        =   328
      Top             =   5805
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dÆzÆ“"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   158
      Left            =   6885
      TabIndex        =   327
      Top             =   6780
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tÆsÆ“"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   316
      Left            =   6600
      TabIndex        =   326
      Top             =   6765
      Width           =   195
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Height          =   345
      Index           =   40
      Left            =   6660
      TabIndex        =   325
      Top             =   4320
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   44
      Left            =   6585
      TabIndex        =   324
      Top             =   4350
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   239
      Left            =   3030
      TabIndex        =   323
      Top             =   9945
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gÉá"
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
      Index           =   231
      Left            =   1335
      TabIndex        =   322
      Top             =   9915
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kå÷"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   200
      Left            =   7335
      TabIndex        =   321
      Top             =   9495
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "l"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   209
      Left            =   6255
      TabIndex        =   320
      Top             =   8160
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "∆"
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
      Index           =   130
      Left            =   7605
      TabIndex        =   319
      Top             =   3645
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   " é"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   165
      Left            =   9795
      TabIndex        =   318
      Top             =   1875
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "g™"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   317
      Left            =   9060
      TabIndex        =   317
      Top             =   3900
      Width           =   300
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   270
      Left            =   8745
      TabIndex        =   316
      Top             =   3975
      Width           =   345
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "g™"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   7650
      TabIndex        =   315
      Top             =   3915
      Width           =   120
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   129
      Left            =   7365
      TabIndex        =   314
      Top             =   3960
      Width           =   120
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kÆ÷"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   251
      Left            =   8820
      TabIndex        =   313
      Top             =   9510
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "t,"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   163
      Left            =   5940
      TabIndex        =   312
      Top             =   1215
      Width           =   195
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   " Ø"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   8850
      TabIndex        =   311
      Top             =   3975
      Width           =   195
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   1
      Left            =   990
      Top             =   7860
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   0
      Left            =   990
      Top             =   5070
      Width           =   1335
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IPA"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   310
      Top             =   195
      Width           =   255
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Nasal]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   15
      Left            =   120
      TabIndex        =   309
      Top             =   2085
      Width           =   495
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Fricative]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   19
      Left            =   135
      TabIndex        =   308
      Top             =   4185
      Width           =   690
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fricative"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   18
      Left            =   150
      TabIndex        =   307
      Top             =   3840
      Width           =   600
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Trill"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   16
      Left            =   120
      TabIndex        =   306
      Top             =   2445
      Width           =   600
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nasal"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   14
      Left            =   120
      TabIndex        =   305
      Top             =   1740
      Width           =   405
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Stop]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   13
      Left            =   120
      TabIndex        =   304
      Top             =   1395
      Width           =   420
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Plosive"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   12
      Left            =   105
      TabIndex        =   303
      Top             =   1065
      Width           =   510
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Retro. Post."
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   8
      Left            =   6630
      TabIndex        =   302
      Top             =   60
      Width           =   615
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Post."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   5880
      TabIndex        =   301
      Top             =   75
      Width           =   375
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Adv. Post."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   5085
      TabIndex        =   300
      Top             =   135
      Width           =   750
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Alveolar"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   3705
      TabIndex        =   299
      Top             =   150
      Width           =   585
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dental"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   2445
      TabIndex        =   298
      Top             =   150
      Width           =   495
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Labio."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   1755
      TabIndex        =   297
      Top             =   150
      Width           =   435
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bilabial"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   1050
      TabIndex        =   296
      Top             =   135
      Width           =   525
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   0
      Left            =   1080
      TabIndex        =   295
      Top             =   945
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   1
      Left            =   1380
      TabIndex        =   294
      Top             =   960
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "l8"
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
      Left            =   3705
      TabIndex        =   293
      Top             =   7860
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   39
      Left            =   8055
      TabIndex        =   292
      Top             =   3960
      Width           =   345
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "pÉ∏"
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
      Index           =   38
      Left            =   1020
      TabIndex        =   291
      Top             =   5805
      Width           =   285
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   37
      Left            =   4110
      TabIndex        =   290
      Top             =   2325
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   57
      Left            =   4125
      TabIndex        =   289
      Top             =   9240
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍÉZ"
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
      Index           =   56
      Left            =   6885
      TabIndex        =   288
      Top             =   6495
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "È^"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   54
      Left            =   3660
      TabIndex        =   287
      Top             =   5400
      Width           =   405
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dÆã"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   53
      Left            =   4725
      TabIndex        =   286
      Top             =   10155
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "©8"
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
      Left            =   8160
      TabIndex        =   285
      Top             =   9195
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   315
      Index           =   50
      Left            =   5880
      TabIndex        =   284
      Top             =   4335
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "v"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   49
      Left            =   2025
      TabIndex        =   283
      Top             =   4020
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "t¨s¨"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   48
      Left            =   3030
      TabIndex        =   282
      Top             =   6780
      Width           =   195
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "l^"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   47
      Left            =   4155
      TabIndex        =   281
      Top             =   5400
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r£"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   34
      Left            =   4200
      TabIndex        =   280
      Top             =   2625
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r£\"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   33
      Left            =   3825
      TabIndex        =   279
      Top             =   2625
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "zå"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   32
      Left            =   3315
      TabIndex        =   278
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Left            =   2745
      TabIndex        =   277
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Left            =   2370
      TabIndex        =   276
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "v“"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   28
      Left            =   1980
      TabIndex        =   275
      Top             =   3330
      Width           =   360
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Left            =   1695
      TabIndex        =   274
      Top             =   3690
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Left            =   1275
      TabIndex        =   273
      Top             =   3705
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   24
      Left            =   4485
      TabIndex        =   272
      Top             =   2940
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "nå"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   2790
      TabIndex        =   271
      Top             =   1905
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "råè"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   22
      Left            =   3120
      TabIndex        =   270
      Top             =   3300
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "ı"
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
      Left            =   1035
      TabIndex        =   269
      Top             =   2310
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Í"
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
      Index           =   5
      Left            =   4785
      TabIndex        =   268
      Top             =   930
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "né"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   4785
      TabIndex        =   267
      Top             =   1905
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "d™"
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
      Index           =   18
      Left            =   2730
      TabIndex        =   266
      Top             =   4035
      Width           =   285
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "f"
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
      Index           =   17
      Left            =   1680
      TabIndex        =   265
      Top             =   4050
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "n0"
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
      Index           =   15
      Left            =   2775
      TabIndex        =   264
      Top             =   1605
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   14
      Left            =   2100
      TabIndex        =   263
      Top             =   1590
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   13
      Left            =   1410
      TabIndex        =   262
      Top             =   1605
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ˇ"
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
      Index           =   4
      Left            =   4440
      TabIndex        =   261
      Top             =   945
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   9
      Left            =   8445
      TabIndex        =   260
      Top             =   915
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   6
      Left            =   7365
      TabIndex        =   259
      Top             =   915
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   3
      Left            =   4065
      TabIndex        =   258
      Top             =   945
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   2
      Left            =   3780
      TabIndex        =   257
      Top             =   945
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "r8"
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
      Left            =   3795
      TabIndex        =   256
      Top             =   2325
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "”"
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
      Left            =   4725
      TabIndex        =   255
      Top             =   7170
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "l"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   55
      Left            =   4140
      TabIndex        =   254
      Top             =   8160
      Width           =   105
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Retro."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   4515
      TabIndex        =   253
      Top             =   150
      Width           =   465
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Trill]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   17
      Left            =   120
      TabIndex        =   252
      Top             =   2790
      Width           =   330
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "∏"
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
      Left            =   1005
      TabIndex        =   251
      Top             =   3690
      Width           =   255
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Glott. Stop]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   20
      Left            =   90
      TabIndex        =   250
      Top             =   8985
      Width           =   840
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Flap]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   21
      Left            =   135
      TabIndex        =   249
      Top             =   3465
      Width           =   390
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Approx."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   22
      Left            =   120
      TabIndex        =   248
      Top             =   7290
      Width           =   540
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Affricate"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   23
      Left            =   105
      TabIndex        =   247
      Top             =   6615
      Width           =   600
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Eject. Stop"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   24
      Left            =   120
      TabIndex        =   246
      Top             =   8700
      Width           =   780
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Gr. Fric.]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   25
      Left            =   90
      TabIndex        =   245
      Top             =   4860
      Width           =   645
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Fricative"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   26
      Left            =   150
      TabIndex        =   244
      Top             =   4515
      Width           =   600
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "p™"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   58
      Left            =   1005
      TabIndex        =   243
      Top             =   4005
      Width           =   285
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "b™"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   59
      Left            =   1275
      TabIndex        =   242
      Top             =   4020
      Width           =   300
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "z“"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   64
      Left            =   6210
      TabIndex        =   241
      Top             =   4710
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ó"
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
      Index           =   65
      Left            =   5865
      TabIndex        =   240
      Top             =   9930
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dã"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   67
      Left            =   4065
      TabIndex        =   239
      Top             =   10200
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tã"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   68
      Left            =   3765
      TabIndex        =   238
      Top             =   10200
      Width           =   195
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Î8"
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
      Index           =   69
      Left            =   3780
      TabIndex        =   237
      Top             =   9210
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "bv"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   71
      Left            =   2010
      TabIndex        =   236
      Top             =   6105
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "bb™"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   72
      Left            =   1350
      TabIndex        =   235
      Top             =   6105
      Width           =   240
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "bã"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   73
      Left            =   1350
      TabIndex        =   234
      Top             =   10185
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dd™"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   74
      Left            =   2685
      TabIndex        =   233
      Top             =   6090
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pp™"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   75
      Left            =   1005
      TabIndex        =   232
      Top             =   6090
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "b÷"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   76
      Left            =   1350
      TabIndex        =   231
      Top             =   9540
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   77
      Left            =   4095
      TabIndex        =   230
      Top             =   5100
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "¬"
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
      Index           =   78
      Left            =   3735
      TabIndex        =   229
      Top             =   5100
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   79
      Left            =   5535
      TabIndex        =   228
      Top             =   4335
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cÉC"
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
      Index           =   80
      Left            =   7305
      TabIndex        =   227
      Top             =   5790
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   81
      Left            =   3345
      TabIndex        =   226
      Top             =   4395
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   82
      Left            =   3075
      TabIndex        =   225
      Top             =   4395
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "d¨ã"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   83
      Left            =   3360
      TabIndex        =   224
      Top             =   10185
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   85
      Left            =   2040
      TabIndex        =   223
      Top             =   3690
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "f≥"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   86
      Left            =   1665
      TabIndex        =   222
      Top             =   3345
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "p»"
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
      Index           =   87
      Left            =   1020
      TabIndex        =   221
      Top             =   8595
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "t»"
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
      Index           =   88
      Left            =   3720
      TabIndex        =   220
      Top             =   8595
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   89
      Left            =   4815
      TabIndex        =   219
      Top             =   2940
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "˜"
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
      Index           =   90
      Left            =   4800
      TabIndex        =   218
      Top             =   1575
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r“å"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   91
      Left            =   3435
      TabIndex        =   217
      Top             =   3315
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "t0És0"
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
      Index           =   92
      Left            =   3015
      TabIndex        =   216
      Top             =   6435
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pã"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   93
      Left            =   1050
      TabIndex        =   215
      Top             =   10185
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "t0ÉT"
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
      Index           =   94
      Left            =   2340
      TabIndex        =   214
      Top             =   5805
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   98
      Left            =   3810
      TabIndex        =   213
      Top             =   1590
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   99
      Left            =   3765
      TabIndex        =   212
      Top             =   1920
      Width           =   195
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   100
      Left            =   1050
      TabIndex        =   211
      Top             =   1905
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "sÆ"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   101
      Left            =   4455
      TabIndex        =   210
      Top             =   4665
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Í"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   102
      Left            =   8475
      TabIndex        =   209
      Top             =   1860
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ˇÉS"
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
      Index           =   103
      Left            =   6555
      TabIndex        =   208
      Top             =   6480
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍÉΩ"
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
      Index           =   107
      Left            =   4740
      TabIndex        =   207
      Top             =   6465
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "˝8"
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
      Index           =   108
      Left            =   8835
      TabIndex        =   206
      Top             =   9225
      Width           =   150
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   109
      Left            =   3390
      TabIndex        =   205
      Top             =   930
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   110
      Left            =   3090
      TabIndex        =   204
      Top             =   915
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Ω"
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
      Index           =   111
      Left            =   4725
      TabIndex        =   203
      Top             =   4380
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "bÉv"
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
      Index           =   113
      Left            =   2010
      TabIndex        =   202
      Top             =   5805
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pf"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   114
      Left            =   1695
      TabIndex        =   201
      Top             =   6075
      Width           =   225
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Gr. Affric.]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   27
      Left            =   90
      TabIndex        =   200
      Top             =   6945
      Width           =   750
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Fric. Lat.]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   28
      Left            =   105
      TabIndex        =   199
      Top             =   5550
      Width           =   705
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Semivow."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   29
      Left            =   105
      TabIndex        =   198
      Top             =   7650
      Width           =   690
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tap or Flap"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   30
      Left            =   120
      TabIndex        =   197
      Top             =   3150
      Width           =   810
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Affricate"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   31
      Left            =   135
      TabIndex        =   196
      Top             =   5895
      Width           =   600
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Lat. Fric."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   32
      Left            =   105
      TabIndex        =   195
      Top             =   5205
      Width           =   615
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Lateral]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   33
      Left            =   135
      TabIndex        =   194
      Top             =   8295
      Width           =   570
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Lat. Approx."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   34
      Left            =   105
      TabIndex        =   193
      Top             =   7980
      Width           =   855
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "p£"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   120
      Left            =   1080
      TabIndex        =   192
      Top             =   2610
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "b£"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   121
      Left            =   1395
      TabIndex        =   191
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   123
      Left            =   10785
      TabIndex        =   190
      Top             =   3705
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "ˆ"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   124
      Left            =   10500
      TabIndex        =   189
      Top             =   4005
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "d÷"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   125
      Left            =   4065
      TabIndex        =   188
      Top             =   9540
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   " å"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   126
      Left            =   7605
      TabIndex        =   187
      Top             =   4005
      Width           =   210
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Height          =   315
      Index           =   128
      Left            =   8385
      TabIndex        =   186
      Top             =   3645
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   131
      Left            =   7320
      TabIndex        =   185
      Top             =   3660
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gÉñ"
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
      Index           =   132
      Left            =   3300
      TabIndex        =   184
      Top             =   9945
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gÉó"
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
      Index           =   133
      Left            =   4050
      TabIndex        =   183
      Top             =   9915
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   134
      Left            =   4095
      TabIndex        =   182
      Top             =   7875
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   135
      Left            =   8415
      TabIndex        =   181
      Top             =   7155
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   136
      Left            =   7620
      TabIndex        =   180
      Top             =   7170
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "®"
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
      Index           =   137
      Left            =   3690
      TabIndex        =   179
      Top             =   7230
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   138
      Left            =   2025
      TabIndex        =   178
      Top             =   7200
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "p÷"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   139
      Left            =   1035
      TabIndex        =   177
      Top             =   9525
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "∫"
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
      Index           =   140
      Left            =   1410
      TabIndex        =   176
      Top             =   9270
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   141
      Left            =   6195
      TabIndex        =   175
      Top             =   4365
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   315
      Index           =   142
      Left            =   5085
      TabIndex        =   174
      Top             =   4395
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   143
      Left            =   4050
      TabIndex        =   173
      Top             =   4395
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "så"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   144
      Left            =   3045
      TabIndex        =   172
      Top             =   4665
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "p◊"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   145
      Left            =   1020
      TabIndex        =   171
      Top             =   8910
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   146
      Left            =   2370
      TabIndex        =   170
      Top             =   4050
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "vÛ"
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
      Index           =   147
      Left            =   2100
      TabIndex        =   169
      Top             =   3015
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "fÛ"
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
      Index           =   148
      Left            =   1725
      TabIndex        =   168
      Top             =   3015
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tÆsÆ"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   149
      Left            =   4455
      TabIndex        =   167
      Top             =   6765
      Width           =   195
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R8"
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
      Index           =   150
      Left            =   3780
      TabIndex        =   166
      Top             =   2970
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r\é“"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   151
      Left            =   4500
      TabIndex        =   165
      Top             =   3315
      Width           =   90
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dz"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   152
      Left            =   4020
      TabIndex        =   164
      Top             =   6780
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r“"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   153
      Left            =   4140
      TabIndex        =   163
      Top             =   3315
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tÉs"
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
      Index           =   154
      Left            =   3720
      TabIndex        =   162
      Top             =   6495
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "ı"
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
      Index           =   155
      Left            =   1320
      TabIndex        =   161
      Top             =   2325
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "d,"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   156
      Left            =   6180
      TabIndex        =   160
      Top             =   1245
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ts"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   160
      Left            =   3720
      TabIndex        =   159
      Top             =   6780
      Width           =   195
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "m"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   161
      Left            =   1350
      TabIndex        =   158
      Top             =   1875
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "m8"
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
      Index           =   162
      Left            =   1065
      TabIndex        =   157
      Top             =   1590
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dÉZ"
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
      Index           =   164
      Left            =   6135
      TabIndex        =   156
      Top             =   6495
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s“å"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   166
      Left            =   5115
      TabIndex        =   155
      Top             =   4710
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "k"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   168
      Left            =   8160
      TabIndex        =   154
      Top             =   1215
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ˇÉß"
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
      Index           =   169
      Left            =   4395
      TabIndex        =   153
      Top             =   6480
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dÉz"
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
      Index           =   170
      Left            =   4035
      TabIndex        =   152
      Top             =   6495
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dç"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   171
      Left            =   3405
      TabIndex        =   151
      Top             =   1230
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tå"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   172
      Left            =   3090
      TabIndex        =   150
      Top             =   1215
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "t◊"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   173
      Left            =   3750
      TabIndex        =   149
      Top             =   8865
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "È"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   175
      Left            =   3765
      TabIndex        =   148
      Top             =   8175
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Height          =   315
      Index           =   176
      Left            =   4770
      TabIndex        =   147
      Top             =   7935
      Width           =   255
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Affricate]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   35
      Left            =   120
      TabIndex        =   146
      Top             =   6255
      Width           =   675
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "bÉB"
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
      Index           =   177
      Left            =   1320
      TabIndex        =   145
      Top             =   5835
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "d0ÉD"
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
      Index           =   70
      Left            =   2685
      TabIndex        =   144
      Top             =   5805
      Width           =   315
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Amer.]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   36
      Left            =   270
      TabIndex        =   143
      Top             =   555
      Width           =   495
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Alveo.]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   41
      Left            =   5835
      TabIndex        =   142
      Top             =   555
      Width           =   555
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Front. Alveo.]"
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   42
      Left            =   5085
      TabIndex        =   141
      Top             =   495
      Width           =   585
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Alveolar]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   43
      Left            =   3690
      TabIndex        =   140
      Top             =   585
      Width           =   675
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Inter.]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   44
      Left            =   2460
      TabIndex        =   139
      Top             =   585
      Width           =   465
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Labio.]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   45
      Left            =   1710
      TabIndex        =   138
      Top             =   585
      Width           =   525
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Bilabial]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   46
      Left            =   1035
      TabIndex        =   137
      Top             =   570
      Width           =   615
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Ret. Alv.]"
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   47
      Left            =   4335
      TabIndex        =   136
      Top             =   480
      Width           =   690
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "p"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   116
      Left            =   1095
      TabIndex        =   135
      Top             =   1230
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   117
      Left            =   1395
      TabIndex        =   134
      Top             =   1245
      Width           =   135
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Implosive"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   48
      Left            =   135
      TabIndex        =   133
      Top             =   9375
      Width           =   660
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Implosive]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   49
      Left            =   120
      TabIndex        =   132
      Top             =   9675
      Width           =   750
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Click"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   50
      Left            =   135
      TabIndex        =   131
      Top             =   10020
      Width           =   345
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Click]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   51
      Left            =   120
      TabIndex        =   130
      Top             =   10350
      Width           =   435
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Bk. Velar]"
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   9
      Left            =   8790
      TabIndex        =   129
      Top             =   510
      Width           =   660
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Palatal]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   11
      Left            =   7305
      TabIndex        =   128
      Top             =   540
      Width           =   585
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Velar]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   37
      Left            =   8130
      TabIndex        =   127
      Top             =   570
      Width           =   465
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Uvular]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   38
      Left            =   9480
      TabIndex        =   126
      Top             =   600
      Width           =   585
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Phar.]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   39
      Left            =   10170
      TabIndex        =   125
      Top             =   570
      Width           =   495
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Glottal]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   52
      Left            =   10785
      TabIndex        =   124
      Top             =   570
      Width           =   540
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "k◊"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   119
      Left            =   8115
      TabIndex        =   123
      Top             =   8850
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "k»"
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
      Index           =   178
      Left            =   8100
      TabIndex        =   122
      Top             =   8580
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "zÆ"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   180
      Left            =   4740
      TabIndex        =   121
      Top             =   4665
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tÆ"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   181
      Left            =   4485
      TabIndex        =   120
      Top             =   1230
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dÆ"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   182
      Left            =   4800
      TabIndex        =   119
      Top             =   1215
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   315
      Index           =   183
      Left            =   8115
      TabIndex        =   118
      Top             =   3690
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "k≠"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   185
      Left            =   7365
      TabIndex        =   117
      Top             =   1185
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "g≠"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   186
      Left            =   7665
      TabIndex        =   116
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "t¨ã"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   187
      Left            =   3030
      TabIndex        =   115
      Top             =   10155
      Width           =   195
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   192
      Left            =   4155
      TabIndex        =   114
      Top             =   1590
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "nÑ"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   194
      Left            =   6225
      TabIndex        =   113
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   195
      Left            =   8430
      TabIndex        =   112
      Top             =   1515
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   197
      Left            =   4755
      TabIndex        =   111
      Top             =   7485
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "n"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   199
      Left            =   4140
      TabIndex        =   110
      Top             =   1920
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   205
      Left            =   3735
      TabIndex        =   109
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tÆã"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   206
      Left            =   4440
      TabIndex        =   108
      Top             =   10170
      Width           =   195
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "∫8"
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
      Index           =   208
      Left            =   1065
      TabIndex        =   107
      Top             =   9255
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "◊"
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
      Index           =   212
      Left            =   7620
      TabIndex        =   106
      Top             =   9210
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "q»"
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
      Index           =   213
      Left            =   8790
      TabIndex        =   105
      Top             =   8550
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ó"
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
      Index           =   214
      Left            =   3780
      TabIndex        =   104
      Top             =   9915
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "¥"
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
      Index           =   216
      Left            =   7575
      TabIndex        =   103
      Top             =   7890
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   217
      Left            =   8400
      TabIndex        =   102
      Top             =   7845
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "k÷"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   218
      Left            =   8115
      TabIndex        =   101
      Top             =   9525
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "t÷"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   220
      Left            =   3780
      TabIndex        =   100
      Top             =   9525
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kÆ◊"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   222
      Left            =   8805
      TabIndex        =   99
      Top             =   8835
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "g÷"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   223
      Left            =   8400
      TabIndex        =   98
      Top             =   9495
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "◊8"
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
      Index           =   224
      Left            =   7335
      TabIndex        =   97
      Top             =   9210
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "¿"
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
      Index           =   225
      Left            =   10485
      TabIndex        =   96
      Top             =   3690
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "c»"
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
      Index           =   227
      Left            =   7335
      TabIndex        =   95
      Top             =   8505
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gg™"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   228
      Left            =   8385
      TabIndex        =   94
      Top             =   6045
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kÉx"
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
      Index           =   229
      Left            =   8055
      TabIndex        =   93
      Top             =   5790
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "t"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   232
      Left            =   3810
      TabIndex        =   92
      Top             =   1230
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "d"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   233
      Left            =   4110
      TabIndex        =   91
      Top             =   1230
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   " Ø"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   234
      Left            =   9135
      TabIndex        =   90
      Top             =   3960
      Width           =   195
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ké"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   236
      Left            =   8880
      TabIndex        =   89
      Top             =   1305
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   237
      Left            =   8475
      TabIndex        =   88
      Top             =   1185
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   238
      Left            =   8415
      TabIndex        =   87
      Top             =   9195
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "≤"
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
      Index           =   246
      Left            =   9780
      TabIndex        =   86
      Top             =   1590
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gÉó"
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
      Index           =   247
      Left            =   6090
      TabIndex        =   85
      Top             =   9945
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "{8"
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
      Index           =   249
      Left            =   9480
      TabIndex        =   84
      Top             =   2295
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "g™"
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
      Index           =   250
      Left            =   8325
      TabIndex        =   83
      Top             =   3930
      Width           =   300
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r\“"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   252
      Left            =   3810
      TabIndex        =   82
      Top             =   3315
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "ß"
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
      Index           =   256
      Left            =   4425
      TabIndex        =   81
      Top             =   4395
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   259
      Left            =   6915
      TabIndex        =   80
      Top             =   4275
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kx"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   262
      Left            =   8085
      TabIndex        =   79
      Top             =   6075
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gÉƒ"
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
      Index           =   263
      Left            =   8385
      TabIndex        =   78
      Top             =   5760
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GÉ“"
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
      Index           =   264
      Left            =   9120
      TabIndex        =   77
      Top             =   5805
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kÆxÆ"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   265
      Left            =   8790
      TabIndex        =   76
      Top             =   6060
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gågå™"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   266
      Left            =   7665
      TabIndex        =   75
      Top             =   6060
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gÆgÆ™"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   267
      Left            =   9135
      TabIndex        =   74
      Top             =   6045
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "˝"
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
      Index           =   273
      Left            =   9135
      TabIndex        =   73
      Top             =   9225
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   ""
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
      Index           =   274
      Left            =   10110
      TabIndex        =   72
      Top             =   3690
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "˙"
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
      Index           =   277
      Left            =   11130
      TabIndex        =   71
      Top             =   3705
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   2160
      Index           =   14
      Left            =   10740
      Top             =   5715
      Width           =   645
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Uvular"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   53
      Left            =   8850
      TabIndex        =   70
      Top             =   180
      Width           =   495
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kåxå"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   281
      Left            =   7335
      TabIndex        =   69
      Top             =   6075
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "á"
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
      Index           =   284
      Left            =   1050
      TabIndex        =   68
      Top             =   9930
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   286
      Left            =   8145
      TabIndex        =   67
      Top             =   945
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gè"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   287
      Left            =   9180
      TabIndex        =   66
      Top             =   1200
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   288
      Left            =   8880
      TabIndex        =   65
      Top             =   1020
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   289
      Left            =   9180
      TabIndex        =   64
      Top             =   960
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   290
      Left            =   10830
      TabIndex        =   63
      Top             =   960
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÑ"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   295
      Left            =   5925
      TabIndex        =   62
      Top             =   1935
      Width           =   195
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R0"
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
      Index           =   298
      Left            =   3090
      TabIndex        =   61
      Top             =   2925
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   299
      Left            =   9825
      TabIndex        =   60
      Top             =   2385
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   300
      Left            =   8130
      TabIndex        =   59
      Top             =   1890
      Width           =   195
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r“"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   301
      Left            =   4830
      TabIndex        =   58
      Top             =   3300
      Width           =   90
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   302
      Left            =   4080
      TabIndex        =   57
      Top             =   2970
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ts“"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   305
      Left            =   5895
      TabIndex        =   56
      Top             =   6795
      Width           =   195
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Index           =   306
      Left            =   3750
      TabIndex        =   55
      Top             =   4395
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "z"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   307
      Left            =   4050
      TabIndex        =   54
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tÉS"
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
      Index           =   310
      Left            =   5820
      TabIndex        =   53
      Top             =   6480
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tT"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   311
      Left            =   2370
      TabIndex        =   52
      Top             =   6105
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   314
      Left            =   7620
      TabIndex        =   51
      Top             =   930
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   315
      Left            =   9825
      TabIndex        =   50
      Top             =   2670
      Width           =   195
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rò"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   318
      Left            =   9510
      TabIndex        =   49
      Top             =   2610
      Width           =   195
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gå÷"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   322
      Left            =   7605
      TabIndex        =   48
      Top             =   9510
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "gÆ÷"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   323
      Left            =   9120
      TabIndex        =   47
      Top             =   9510
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "h"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   327
      Left            =   10725
      TabIndex        =   46
      Top             =   4020
      Width           =   360
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "y"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   328
      Left            =   7665
      TabIndex        =   45
      Top             =   7470
      Width           =   135
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Palatal"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   55
      Left            =   7320
      TabIndex        =   44
      Top             =   150
      Width           =   495
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Velar"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   56
      Left            =   8205
      TabIndex        =   43
      Top             =   135
      Width           =   375
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Uvular"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   57
      Left            =   9510
      TabIndex        =   42
      Top             =   180
      Width           =   480
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Phar."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   58
      Left            =   10245
      TabIndex        =   41
      Top             =   165
      Width           =   405
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Glottal"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   59
      Left            =   10815
      TabIndex        =   40
      Top             =   165
      Width           =   450
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   18
      Left            =   10110
      Top             =   1620
      Width           =   1275
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   19
      Left            =   10110
      Top             =   7860
      Width           =   1275
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   1380
      Index           =   21
      Left            =   10110
      Top             =   4380
      Width           =   1275
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   1395
      Index           =   23
      Left            =   10740
      Top             =   8535
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   24
      Left            =   11070
      Top             =   945
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   1395
      Index           =   25
      Left            =   10740
      Top             =   2295
      Width           =   645
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dental"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   3075
      TabIndex        =   39
      Top             =   135
      Width           =   495
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Dental]"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   54
      Left            =   3060
      TabIndex        =   38
      Top             =   555
      Width           =   585
   End
   Begin VB.Line Line2 
      X1              =   10410
      X2              =   10410
      Y1              =   945
      Y2              =   1605
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   705
      Index           =   2
      Left            =   10110
      Top             =   9915
      Width           =   1275
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "s“"
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
      Left            =   5910
      TabIndex        =   37
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "z“Æ"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   46
      Left            =   6900
      TabIndex        =   36
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "z“å"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   96
      Left            =   5490
      TabIndex        =   35
      Top             =   4710
      Width           =   135
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "[Retro. Alveo.]"
      ForeColor       =   &H80000008&
      Height          =   360
      Index           =   40
      Left            =   6540
      TabIndex        =   34
      Top             =   480
      Width           =   780
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "qÉX"
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
      Index           =   255
      Left            =   8775
      TabIndex        =   33
      Top             =   5775
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "d0Éz0"
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
      Index           =   105
      Left            =   3360
      TabIndex        =   32
      Top             =   6480
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "d¨z¨"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   3360
      TabIndex        =   31
      Top             =   6750
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dÆzÆ"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   112
      Left            =   4740
      TabIndex        =   30
      Top             =   6750
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dz“"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   122
      Left            =   6120
      TabIndex        =   29
      Top             =   6825
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ˇ»"
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
      Index           =   10
      Left            =   4425
      TabIndex        =   28
      Top             =   8550
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "tÆ◊"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   45
      Left            =   4470
      TabIndex        =   27
      Top             =   8850
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "kå◊"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   60
      Left            =   7395
      TabIndex        =   26
      Top             =   8820
      Width           =   225
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Height          =   360
      Index           =   279
      Left            =   6900
      TabIndex        =   25
      Top             =   4305
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   257
      Left            =   5535
      TabIndex        =   24
      Top             =   4380
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Height          =   315
      Index           =   104
      Left            =   4680
      TabIndex        =   23
      Top             =   3015
      Width           =   360
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "l6"
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
      Index           =   202
      Left            =   6180
      TabIndex        =   22
      Top             =   7830
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   " ,"
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
      Index           =   291
      Left            =   6210
      TabIndex        =   21
      Top             =   8175
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   " å"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   7320
      TabIndex        =   20
      Top             =   4005
      Width           =   210
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H80000011&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   13
      Left            =   10425
      Top             =   945
      Width           =   315
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   282
      Left            =   5940
      TabIndex        =   19
      Top             =   930
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   283
      Left            =   6165
      TabIndex        =   18
      Top             =   945
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   240
      Left            =   10815
      TabIndex        =   17
      Top             =   1260
      Width           =   135
   End
   Begin VB.Line Line1 
      Index           =   40
      X1              =   11055
      X2              =   11055
      Y1              =   915
      Y2              =   1620
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
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
      Height          =   315
      Index           =   326
      Left            =   8085
      TabIndex        =   16
      Top             =   1590
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Height          =   345
      Index           =   245
      Left            =   8130
      TabIndex        =   15
      Top             =   1515
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8 "
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
      Index           =   95
      Left            =   6015
      TabIndex        =   14
      Top             =   1620
      Width           =   330
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "n6"
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
      Index           =   292
      Left            =   6195
      TabIndex        =   13
      Top             =   1560
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "n6"
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
      Index           =   62
      Left            =   5940
      TabIndex        =   12
      Top             =   1530
      Width           =   150
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   " ç"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   215
      Left            =   7590
      TabIndex        =   11
      Top             =   1935
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Í"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   293
      Left            =   7635
      TabIndex        =   10
      Top             =   1830
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¯"
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
      Index           =   97
      Left            =   7620
      TabIndex        =   9
      Top             =   1575
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Í"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   242
      Left            =   9840
      TabIndex        =   8
      Top             =   1830
      Width           =   135
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R0"
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
      Index           =   248
      Left            =   3390
      TabIndex        =   7
      Top             =   2925
      Width           =   165
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "8"
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
      Left            =   3435
      TabIndex        =   6
      Top             =   2985
      Width           =   255
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "rå£"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   23
      Left            =   3390
      TabIndex        =   5
      Top             =   2610
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r£å"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   21
      Left            =   3135
      TabIndex        =   4
      Top             =   2610
      Width           =   105
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "“"
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
      Index           =   42
      Left            =   9105
      TabIndex        =   3
      Top             =   3690
      Width           =   255
   End
   Begin VB.Line Line1 
      Index           =   23
      X1              =   10095
      X2              =   10095
      Y1              =   60
      Y2              =   10610
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   16
      X1              =   11385
      X2              =   11385
      Y1              =   90
      Y2              =   10590
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   20
      X1              =   75
      X2              =   11385
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   60
      X2              =   11385
      Y1              =   2985
      Y2              =   2985
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   90
      X2              =   11385
      Y1              =   3675
      Y2              =   3675
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   75
      X2              =   11385
      Y1              =   1605
      Y2              =   1605
   End
   Begin VB.Line Line1 
      Index           =   18
      X1              =   10725
      X2              =   10725
      Y1              =   60
      Y2              =   10630
   End
   Begin VB.Line Line1 
      Index           =   31
      X1              =   75
      X2              =   11385
      Y1              =   4365
      Y2              =   4365
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   75
      X2              =   11385
      Y1              =   2295
      Y2              =   2295
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   1635
      X2              =   1635
      Y1              =   75
      Y2              =   10575
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   58
      X1              =   975
      X2              =   975
      Y1              =   90
      Y2              =   10590
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2310
      X2              =   2310
      Y1              =   60
      Y2              =   10625
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   56
      X1              =   75
      X2              =   11385
      Y1              =   10620
      Y2              =   10620
   End
   Begin VB.Line Line1 
      Index           =   19
      X1              =   75
      X2              =   11385
      Y1              =   9900
      Y2              =   9900
   End
   Begin VB.Line Line1 
      Index           =   55
      X1              =   75
      X2              =   11385
      Y1              =   7845
      Y2              =   7845
   End
   Begin VB.Line Line1 
      Index           =   57
      X1              =   75
      X2              =   11385
      Y1              =   8535
      Y2              =   8535
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   3675
      X2              =   3675
      Y1              =   60
      Y2              =   10575
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   2985
      X2              =   2985
      Y1              =   60
      Y2              =   10625
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   90
      X2              =   11385
      Y1              =   6450
      Y2              =   6450
   End
   Begin VB.Line Line1 
      Index           =   21
      X1              =   105
      X2              =   11385
      Y1              =   9225
      Y2              =   9225
   End
   Begin VB.Line Line1 
      Index           =   32
      X1              =   75
      X2              =   11385
      Y1              =   5055
      Y2              =   5055
   End
   Begin VB.Line Line1 
      Index           =   54
      X1              =   75
      X2              =   11385
      Y1              =   7155
      Y2              =   7155
   End
   Begin VB.Line Line1 
      Index           =   36
      X1              =   75
      X2              =   11385
      Y1              =   5745
      Y2              =   5745
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   75
      X2              =   11385
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   5790
      X2              =   5790
      Y1              =   60
      Y2              =   10625
   End
   Begin VB.Line Line1 
      Index           =   17
      X1              =   6510
      X2              =   6510
      Y1              =   60
      Y2              =   10620
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   75
      X2              =   75
      Y1              =   90
      Y2              =   10590
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   22
      X1              =   75
      X2              =   11385
      Y1              =   75
      Y2              =   75
   End
   Begin VB.Line Line1 
      Index           =   24
      X1              =   9435
      X2              =   9435
      Y1              =   60
      Y2              =   10625
   End
   Begin VB.Line Line1 
      Index           =   26
      X1              =   7980
      X2              =   7980
      Y1              =   60
      Y2              =   10625
   End
   Begin VB.Line Line1 
      Index           =   25
      X1              =   8715
      X2              =   8715
      Y1              =   60
      Y2              =   10625
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   5070
      X2              =   5070
      Y1              =   60
      Y2              =   10610
   End
   Begin VB.Line Line1 
      Index           =   27
      X1              =   7260
      X2              =   7260
      Y1              =   60
      Y2              =   10575
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   4365
      X2              =   4365
      Y1              =   60
      Y2              =   10625
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " \"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   61
      Left            =   3150
      TabIndex        =   2
      Top             =   2670
      Width           =   90
   End
   Begin VB.Label Con 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " \"
      BeginProperty Font 
         Name            =   "Amer Phon SILDoulosL"
         Size            =   12
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   63
      Left            =   3135
      TabIndex        =   1
      Top             =   3360
      Width           =   90
   End
End
Attribute VB_Name = "frmDispAmerCon2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TBarButtons = "Exit;"

Private sCaption As String              ' Title bar caption

Private Sub Form_Activate()

  Dim frm As Form
  
  On Error Resume Next
  
  With mdiHelpCharts
    .Caption = App.Title & " - [" & Caption & "]"
    .ShowTBarButtons TBarButtons
    .panStatus.Visible = True
    .mnuTest.Enabled = False                        '* Disable test menu.
  End With
  
  Move 0, 15
  gStatLine.SimpleText = ""
  
  For Each frm In Forms
    If frm.Name <> "mdiHelpCharts" Then frm.Top = 0
  Next
  
  Show
  
End Sub

Private Sub Form_Deactivate()

  Dim frm As Form
  
  On Error Resume Next

  mdiHelpCharts.Caption = App.Title
  For Each frm In Forms
    If frm.Name <> "mdiHelpCharts" Then frm.Top = -Height
  Next frm
  gStatLine.SimpleText = ""
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)

  On Error Resume Next
  gStatLine.SimpleText = ""

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  On Error Resume Next
  Call Form_Deactivate
  Set frmDispAmerCon2 = Nothing

End Sub

Private Sub Form_Resize()

  On Error Resume Next
  If mdiHelpCharts.WindowState = vbMaximized Then _
    mdiHelpCharts.panStatus.Visible = True

End Sub
