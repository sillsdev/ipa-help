VERSION 5.00
Begin VB.Form frmHelpAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About IPA Help"
   ClientHeight    =   6765
   ClientLeft      =   6960
   ClientTop       =   3555
   ClientWidth     =   3735
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
   Icon            =   "ABOUT.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6765
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "IPA Help uses the following references:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   60
      TabIndex        =   8
      Top             =   2640
      Width           =   3615
      Begin VB.Label lblRef5 
         Caption         =   "� The IPA phonetic sounds are copyrighted material. They are not in the public domain. See License.rtf for details."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   170
         TabIndex        =   13
         Top             =   2800
         Width           =   3195
      End
      Begin VB.Label lblRef4 
         Caption         =   "� The sample word lists were supplied by SIL."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   190
         Left            =   170
         TabIndex        =   12
         Top             =   2509
         Width           =   3195
      End
      Begin VB.Label lblRef1 
         Caption         =   "� Floyd, Rick, compiler. 1986. Manual for articulatory phonetics. Dallas: Summer Institute of Linguistics."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Index           =   4
         Left            =   165
         TabIndex        =   11
         Top             =   315
         Width           =   3195
      End
      Begin VB.Label lblRef2 
         Caption         =   $"ABOUT.frx":030A
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   165
         TabIndex        =   10
         Top             =   978
         Width           =   3195
      End
      Begin VB.Label lblRef3 
         Caption         =   "� The IPA Consonant and Vowel Charts are based on charts from the International Phonetic Association."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   165
         TabIndex        =   9
         Top             =   1841
         Width           =   3195
      End
   End
   Begin VB.PictureBox picHelpAbout 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   1612
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   1275
      Width           =   480
   End
   Begin VB.CommandButton cmdHelpAboutOK 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   6270
      Width           =   1095
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "#"
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
      Left            =   1785
      TabIndex        =   7
      Top             =   1950
      Width           =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IPA Help"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   0
      Left            =   -15
      TabIndex        =   4
      Top             =   120
      Width           =   3675
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IPA Help"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Index           =   1
      Left            =   15
      TabIndex        =   5
      Top             =   120
      Width           =   3675
   End
   Begin VB.Label lblHelpAboutVer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "(version #)"
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
      Left            =   -15
      TabIndex        =   3
      Top             =   945
      Width           =   3750
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "A Phonetics LearningTool"
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
      Left            =   -15
      TabIndex        =   1
      Top             =   615
      Width           =   3750
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "IPA Help"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Index           =   2
      Left            =   45
      TabIndex        =   6
      Top             =   135
      Width           =   3675
   End
End
Attribute VB_Name = "frmHelpAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
'* frmHelpAbout version info:
'*  See basGlobals (Globals.bas).
'**************************************************

Option Explicit

Private Sub cmdHELPaboutOK_Click()

   Unload frmHelpAbout

End Sub

Private Sub Form_Load()

  Me.Icon = mdiHelpCharts.Icon
  picHelpAbout.Picture = Me.Icon
  With App
    lblHelpAboutVer.Caption = "Version " & _
                            .Major & "." & _
                            .Minor
    If .Revision > 0 Then
        lblHelpAboutVer.Caption = lblHelpAboutVer.Caption & .Revision
    End If
    'lblHelpAboutVer.Caption = lblHelpAboutVer.Caption & " (" & SystemType & ")"
  
    '************************************************
    '* I used the LegalTrademarks property for the
    '* "JAARS - ICS..." message because I wanted if
    '* after the copyright and not with the company
    '* name (i.e. SIL).
    '************************************************
'    lblCopyright.Caption = Trim$(.CompanyName & vbCrLf & _
                         .LegalCopyright & vbCrLf & _
                         .LegalTrademarks)
    lblCopyright.Caption = Trim$(.LegalCopyright)
  End With
  With lblCopyright
    If .Width > (frmHelpAbout.Width - 400) Then
        .Caption = Left$(.Caption, InStr(.Caption, ". ") + 1) _
                   + vbCrLf _
                   + Right$(.Caption, Len(.Caption) - InStr(.Caption, ". "))
        .Width = frmHelpAbout.Width - 400
        .Left = 200
        .WordWrap = True
    End If
    Debug.Print .Width; .Left; .Height
  End With
  
  Call CenterForm(Me)

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set frmHelpAbout = Nothing

End Sub

