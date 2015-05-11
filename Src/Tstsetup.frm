VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTestSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Test Setup"
   ClientHeight    =   4455
   ClientLeft      =   5685
   ClientTop       =   3675
   ClientWidth     =   5355
   Icon            =   "Tstsetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4455
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4125
      TabIndex        =   1
      Top             =   4005
      Width           =   1125
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Default         =   -1  'True
      Height          =   345
      Left            =   2910
      TabIndex        =   0
      Top             =   4005
      Width           =   1125
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3810
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   6720
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "IPA Sounds"
      TabPicture(0)   =   "Tstsetup.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Word List "
      TabPicture(1)   =   "Tstsetup.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame5 
         Caption         =   "Category"
         Height          =   1095
         Left            =   -71895
         TabIndex        =   26
         Top             =   495
         Width           =   1755
         Begin VB.OptionButton Option5 
            Caption         =   "&Random"
            Height          =   330
            Index           =   1
            Left            =   165
            TabIndex        =   28
            Top             =   660
            Width           =   1305
         End
         Begin VB.OptionButton Option5 
            Caption         =   "&User Selected"
            Height          =   330
            Index           =   0
            Left            =   165
            TabIndex        =   27
            Top             =   240
            Value           =   -1  'True
            Width           =   1305
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Word List"
         Height          =   3060
         Left            =   -74745
         TabIndex        =   20
         Top             =   495
         Width           =   2505
         Begin VB.OptionButton Option4 
            Caption         =   "&1"
            Height          =   330
            Index           =   0
            Left            =   165
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton Option4 
            Caption         =   "&2"
            Height          =   330
            Index           =   1
            Left            =   165
            TabIndex        =   24
            Top             =   660
            Width           =   2295
         End
         Begin VB.OptionButton Option4 
            Caption         =   "&3"
            Height          =   330
            Index           =   2
            Left            =   165
            TabIndex        =   23
            Top             =   1065
            Width           =   2250
         End
         Begin VB.OptionButton Option4 
            Caption         =   "&4"
            Height          =   330
            Index           =   3
            Left            =   165
            TabIndex        =   22
            Top             =   1470
            Width           =   2220
         End
         Begin VB.OptionButton Option4 
            Caption         =   "&5"
            Height          =   330
            Index           =   4
            Left            =   165
            TabIndex        =   21
            Top             =   1890
            Width           =   2265
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Previous Group"
         Height          =   735
         Left            =   3615
         TabIndex        =   18
         Top             =   1875
         Width           =   1395
         Begin VB.OptionButton Option3 
            Caption         =   "&Retest"
            Height          =   195
            Left            =   240
            TabIndex        =   19
            Top             =   330
            Width           =   810
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Layout"
         Height          =   1395
         Left            =   3615
         TabIndex        =   15
         Top             =   360
         Width           =   1395
         Begin VB.OptionButton Option1 
            Caption         =   "&IPA Chart"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   420
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Brie&f"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   16
            Top             =   840
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Sounds"
         Height          =   3315
         Left            =   105
         TabIndex        =   3
         Top             =   360
         Width           =   3395
         Begin VB.OptionButton Option2 
            Caption         =   "All &Consonants"
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   14
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&Plosives"
            Height          =   315
            Index           =   1
            Left            =   180
            TabIndex        =   13
            Top             =   770
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&Nasals, Trills, Flaps"
            Height          =   315
            Index           =   2
            Left            =   180
            TabIndex        =   12
            Top             =   1180
            Width           =   1680
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&Fricatives"
            Height          =   315
            Index           =   3
            Left            =   180
            TabIndex        =   11
            Top             =   1590
            Width           =   1035
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&Approximants"
            Height          =   315
            Index           =   4
            Left            =   180
            TabIndex        =   10
            Top             =   2000
            Width           =   1455
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Non P&ulmonic"
            Height          =   315
            Index           =   5
            Left            =   180
            TabIndex        =   9
            Top             =   2410
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Other &Symbols"
            Height          =   315
            Index           =   6
            Left            =   180
            TabIndex        =   8
            Top             =   2820
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "All &Vowels"
            Height          =   315
            Index           =   7
            Left            =   2140
            TabIndex        =   7
            Top             =   360
            Width           =   1035
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Fro&nt"
            Height          =   315
            Index           =   8
            Left            =   2140
            TabIndex        =   6
            Top             =   770
            Width           =   1035
         End
         Begin VB.OptionButton Option2 
            Caption         =   "C&entral"
            Height          =   315
            Index           =   9
            Left            =   2140
            TabIndex        =   5
            Top             =   1180
            Width           =   1035
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&Back"
            Height          =   315
            Index           =   10
            Left            =   2140
            TabIndex        =   4
            Top             =   1590
            Width           =   1035
         End
      End
   End
End
Attribute VB_Name = "frmTestSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
'* frmTestSetup version info:
'*  See basGlobals (Globals.bas).
'**************************************************

Private iTestLayoutTmp As Integer
Private sTestGroupTmp As String
Private iTestGroup As Integer
Private iTestGroupTmp As Integer
Private bRetestTmp As Boolean

Private Function ValidWavFilesInWordListFile(sWLFileSpec$) As Boolean

  Dim i As Integer
  Dim j As Integer
  Dim iCount As Integer
  Dim sCategoryNames() As String
  Dim sWLSndPath As String
  Dim vList As Variant
  
  On Error Resume Next
  
  ValidWavFilesInWordListFile = False
  
  '**************************************************************
  '* Load xml word list file
  '**************************************************************
  Set xmlWL = New clsXMLWordList
  xmlWL.Load sWLFileSpec
  
  '**************************************************************
  '* Get the path for the wave files
  '**************************************************************
  sWLSndPath = xmlWL.SoundPath
  sWLSndPath = sWLSndPath & _
               IIf(Right$(sWLSndPath, 1) = "\", "", "\")
  
  '**************************************************************
  '* Get the list of word list categories.
  '**************************************************************
  sCategoryNames = xmlWL.CategoryNames
  
  Err.Clear
  i = UBound(sCategoryNames)
  If (Err.Number > 0) Then Exit Function
  
  iCount = 0
  
  For i = 0 To UBound(sCategoryNames)
    vList = xmlWL.WordsInCategory(sCategoryNames(i))
    
    If Not (IsNull(vList)) Then
      For j = 0 To UBound(vList, 1)
        If (FileExist(MakeFullPath(sWLSndPath, CStr(vList(j, 5))))) Then
          iCount = iCount + 1
          If (iCount = 3) Then
            ValidWavFilesInWordListFile = True
            Set xmlWL = Nothing
            Exit Function
          End If
        End If
      Next
    End If
  Next
  
  Set xmlWL = Nothing
  
End Function

Private Sub cmdCancel_Click()

'* 'Cancel' Button Click
  
  '* Set flag to tell toolbar I'm closing.
  gTestSetupActive = False  '* Added by CLW 4/22/99
  
  Call mdiHelpCharts.mnuTestStop_Click
  mdiHelpCharts.bStartTest = False
  Unload Me

End Sub

Private Sub cmdStart_Click()

'* 'Start' Button Click

'***********************************************************************
'* Whether a vowel is visible or not is now determined by its status in
'* gTestGrpCorrect. This means that gTestGrp does not change depending on
'* how many the user got correct. This allows the indexing to be
'* the same for both gTestGrp and gTestGrpCorrect. CLW 1/26/99
'*
'* Status of tab control determines which test is started. CLW 4/29/99
'***********************************************************************
  
  Dim i As Integer
  Dim iUpper As Integer
  Dim ctl As Control
  Dim frm As Form
  
  '* Set flag to tell toolbar I'm closing.
  gTestSetupActive = False  '* Added by CLW 4/22/99
  
  If (SSTab1.Tab = 1) Then
    gTestForm = "frmWordList"
    For i = 0 To 4
      If (Option4(i).Value) Then
        gTestTag = i
        Exit For
      End If
    Next
  Else
    '************************************************************
    '* Set up the new test group. Initialize the retest list.
    '* Reset gRetestActive to false. Store desired Test Layout.
    '************************************************************
    gTestGrpName = sTestGroupTmp
    gTestGrp = gPhonGrpColl(gTestGrpName)
    'ReDim gRetestList(0)
    'gRetestListEmpty = True
    iUpper = UBound(gTestGrp)
      
    If bRetestTmp Then
      For i = 0 To UBound(gTestGrpCorrect, 2)
        gTestGrpCorrect(cLastTest, i) = gTestGrpCorrect(cThisTest, i)
      Next i
    Else
      ReDim gTestGrpCorrect(0 To 1, 0 To iUpper)
      For i = 0 To UBound(gTestGrpCorrect, 2)
        gTestGrpCorrect(cThisTest, i) = vbGrayed
        gTestGrpCorrect(cLastTest, i) = vbGrayed
      Next i
    End If
      
    gRetestActive = bRetestTmp
    gTestLayout = iTestLayoutTmp
    gTestForm = "frmDisp" & Right(gTestGrpName, 3)
    gTestTag = ""
  End If
  
  mdiHelpCharts.bStartTest = True
  Unload Me
  
End Sub

Private Sub Form_Activate()

  On Error Resume Next
  gTestSetupActive = True
  
End Sub

Private Sub Form_Load()

  Dim i As Integer
  Dim iCurrOption As Integer
  
  On Error Resume Next
  KeyPreview = True
  Call CenterForm(Me)
  
  iTestLayoutTmp = gTestLayout
  sTestGroupTmp = gTestGrpName
  
  '************************************************
  '* Find name of test group from test group array.
  '* Use index of that name to select correct
  '* option button.
  '************************************************
  For i = 0 To 10
    If gPhonGrpNameArray(i) = sTestGroupTmp Then
      iTestGroup = i
    End If
  Next i
  
  If (gRetestListEmpty) Then
    Option2(iTestGroup) = True
    Option3.Enabled = False
  Else
    Option3 = True
  End If
  Option1(gTestLayout) = True
  
  If (WordListArraySize() < 0) Then
    SSTab1.TabVisible(1) = False
  Else
    Dim bGottaValidWordList As Boolean
    bGottaValidWordList = False
    iCurrOption = 0
  
    '************************************************
    '* Load WordList names from IPAHelp.ini.
    '* Added by CLW 4/29/99
    '************************************************
    For i = 0 To 4
      With Option4(i)
        If (Len(gWordListID(i, 0)) > 0) Then
          .Caption = .Caption & " " & gWordListID(i, 0)
          
          If (Len(gWordListID(i, 1)) = 0 Or Len(Dir(gWordListID(i, 1))) = 0) Then
            .Enabled = False
          Else
            .Enabled = ValidWavFilesInWordListFile(CStr(gWordListID(i, 1)))
          End If
          
          If (WordListArraySize() >= i And .Enabled) Then
            If (mdiHelpCharts.ActiveForm.Caption = gWordListID(i, 0) & " Word List") Then
              iCurrOption = i
              .Value = True
            End If
          End If
        Else
          .Visible = False
          .Caption = ""
        End If
      
        If (Not bGottaValidWordList And .Enabled) Then bGottaValidWordList = True
      End With
    Next i
  
    If Not (bGottaValidWordList) Then SSTab1.TabVisible(1) = False
  End If
  
  If (SSTab1.TabVisible(1) And Option4(iCurrOption).Value And Not Option4(iCurrOption).Enabled) Then
    For i = 0 To 4
      If (Option4(i).Enabled) Then
        Option4(i).Value = True
        Exit For
      End If
    Next i
  End If
  
  '* Select Category choice method. CLW 5/11/99
  Option5(gTestCatChoice) = True
  
  Select Case mdiHelpCharts.ActiveForm.Name
    Case "frmWordList": If (SSTab1.TabVisible(1)) Then SSTab1.Tab = 1
    Case "frmDispVow":  Option2(7).Value = True
    Case Else
  End Select
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  On Error Resume Next
  Set frmTestSetup = Nothing

End Sub

Private Sub Option1_Click(Index As Integer)

  Dim i As Integer

  On Error Resume Next
  iTestLayoutTmp = Index
    
End Sub

Private Sub Option2_Click(Index As Integer)

  On Error Resume Next
  iTestGroupTmp = Index
  sTestGroupTmp = gPhonGrpNameArray(Index)

End Sub

Private Sub Option2_GotFocus(Index As Integer)

  On Error Resume Next
  bRetestTmp = False
  Option3.Value = False
  sTestGroupTmp = gPhonGrpNameArray(Index)
  
End Sub

Private Sub Option3_Click()

  On Error Resume Next
  Option2(iTestGroup).Value = True
  bRetestTmp = True
  
End Sub

Private Sub Option5_Click(Index As Integer)

  On Error Resume Next
  gTestCatChoice = Index '* Added by CLW 5/11/99
  
End Sub
