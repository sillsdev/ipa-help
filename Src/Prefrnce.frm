VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPreferences 
   Caption         =   "IPA Help Preferences"
   ClientHeight    =   3600
   ClientLeft      =   1470
   ClientTop       =   6225
   ClientWidth     =   8655
   Icon            =   "Prefrnce.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3600
   ScaleWidth      =   8655
   Begin VB.CommandButton cmdOKCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   7695
      TabIndex        =   20
      Top             =   3105
      Width           =   915
   End
   Begin VB.CommandButton cmdOKCancel 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   6705
      TabIndex        =   19
      Top             =   3105
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "Other Locations"
      Height          =   1620
      Index           =   1
      Left            =   30
      TabIndex        =   24
      Top             =   1920
      Width           =   5490
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   330
         Index           =   1
         Left            =   4500
         TabIndex        =   10
         Top             =   1155
         Width           =   885
      End
      Begin VB.TextBox txtSASLoc 
         Height          =   315
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1155
         Width           =   4305
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   330
         Index           =   0
         Left            =   4500
         TabIndex        =   7
         Top             =   525
         Width           =   885
      End
      Begin VB.TextBox txtSndsLoc 
         Height          =   315
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   525
         Width           =   4305
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Speech Analy&zer Server Location"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   165
         TabIndex        =   8
         Top             =   915
         Width           =   2370
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Phonetic Sound Files Location"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   165
         TabIndex        =   5
         Top             =   285
         Width           =   2160
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Word List Locations"
      Height          =   1770
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   5475
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   4455
         TabIndex        =   3
         Top             =   765
         Width           =   885
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   4455
         TabIndex        =   4
         Top             =   1230
         Width           =   885
      End
      Begin VB.CommandButton cmdBrowseWL 
         Caption         =   "&Browse..."
         Height          =   375
         Left            =   4455
         TabIndex        =   2
         Top             =   315
         Width           =   885
      End
      Begin MSComctlLib.ListView lvWLLocations 
         Height          =   1410
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   2487
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "Title"
            Text            =   "Word List Title"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "File"
            Text            =   "File"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Word List Audio Playback Settings"
      Height          =   2925
      Index           =   2
      Left            =   5595
      TabIndex        =   21
      Top             =   60
      Width           =   3045
      Begin MSComCtl2.UpDown updnSRSpeed 
         Height          =   300
         Left            =   2310
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2340
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   10
         BuddyControl    =   "txtSRSpeed"
         BuddyDispid     =   196618
         OrigLeft        =   2385
         OrigTop         =   2220
         OrigRight       =   2625
         OrigBottom      =   2505
         Increment       =   5
         Max             =   333
         Min             =   10
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updnDelay 
         Height          =   300
         Index           =   1
         Left            =   2310
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1080
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtRptDelay"
         BuddyDispid     =   196619
         OrigLeft        =   2310
         OrigTop         =   1065
         OrigRight       =   2550
         OrigBottom      =   1410
         Max             =   300
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updnPlaybackTimes 
         Height          =   300
         Left            =   2310
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1710
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtRptCount"
         BuddyDispid     =   196620
         OrigLeft        =   2475
         OrigTop         =   1620
         OrigRight       =   2715
         OrigBottom      =   2055
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtSRSpeed 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1755
         TabIndex        =   18
         Top             =   2355
         Width           =   555
      End
      Begin VB.TextBox txtRptDelay 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1755
         TabIndex        =   14
         Top             =   1095
         Width           =   555
      End
      Begin VB.TextBox txtRptCount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1755
         TabIndex        =   16
         Top             =   1725
         Width           =   555
      End
      Begin VB.TextBox txtInitDelay 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1755
         TabIndex        =   12
         Top             =   465
         Width           =   555
      End
      Begin MSComCtl2.UpDown updnDelay 
         Height          =   300
         Index           =   0
         Left            =   2310
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   450
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtRptDelay"
         BuddyDispid     =   196619
         OrigLeft        =   2310
         OrigTop         =   1065
         OrigRight       =   2550
         OrigBottom      =   1410
         Max             =   300
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         Height          =   195
         Index           =   7
         Left            =   2625
         TabIndex        =   25
         Top             =   2400
         Width           =   450
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Slowed replay speed (as % of normal):"
         Height          =   405
         Index           =   8
         Left            =   180
         TabIndex        =   17
         Top             =   2235
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sec."
         Height          =   195
         Index           =   4
         Left            =   2625
         TabIndex        =   23
         Top             =   1140
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sec."
         Height          =   195
         Index           =   0
         Left            =   2625
         TabIndex        =   22
         Top             =   510
         Width           =   330
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Delay &between repeated playbacks:"
         Height          =   390
         Index           =   3
         Left            =   180
         TabIndex        =   13
         Top             =   975
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "&Times to playback selected word(s):"
         Height          =   390
         Index           =   2
         Left            =   180
         TabIndex        =   15
         Top             =   1605
         Width           =   1425
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Delay before playback begins:"
         ForeColor       =   &H80000008&
         Height          =   390
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   345
         Width           =   1260
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   5955
      Top             =   3090
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const NoWordList = "<none>"
Private Const LVColWidINIEntry = "LVCol"
Private Const NewTitle = "New Title"
Private Const DefaultPhoneticFont = "SILDoulosIPA"
Private Const DefaultPhoneticFontSize = 12

Private bGetFileSpecAfterTitleEdit As Boolean
Private bAddExistingWLFile As Boolean
Private sDlgTitle As String
Private sFileSpec As String

Private Const WLFileExists = 0
Private Const WLFileNew = 1
Private Const WLFileInvalid = 2
Private Const WLFileBrowseCancel = 3
Private Const WLDuplicate = 4

Private Sub GetSoundsPath()

  On Error Resume Next
  
  Dim bWavPathUpdate As Boolean
  
  '************************************************
  '* Pass gWavPath as default directory.
  '* Show frmFilePath (vbModal means a value must
  '* be returned before execution will continue).
  '* Check to see if a directory was returned.
  '************************************************
  With frmFilePath
    .sDir = gWavPath
    .Show vbModal
    If Len(.sDir) > 0 Then gWavPath = .sDir
    txtSndsLoc.Text = .sDir
    bWavPathUpdate = True
  End With
  
  Unload frmFilePath
  
  '************************************************
  '* If there was an update to the Wave file path,
  '* then we need to update the play buttons and
  '* test menu (and button, if applicable).
  '************************************************
  If bWavPathUpdate Then
    With mdiHelpCharts.ActiveForm
      Select Case .Name
        Case "frmDispCon", "frmDispDia", "frmDispVow", "frmDispSSeg", "frmWordList"
        Case Else
          Call UpdateTestMenu
      End Select
    End With
  End If
  
End Sub

Private Function GetWLFileSpec() As Integer

  '***************************************************************************
  '* Returns:
  '*    WLFileExists       - File specifiedexists and is a valid XML word
  '*                         list file.
  '*    WLFileNew          - File specified is new.
  '*    WLFileInvalid      - File specified exists but is not a valid XML
  '*                         word list file.
  '*    WLFileBrowseCancel - User canceled out of the browse dialog.
  '***************************************************************************

  On Error GoTo GetWLFileSpecErr
  
  GetWLFileSpec = WLFileBrowseCancel
  
  With dlg
    If (lvWLLocations.SelectedItem Is Nothing) Then Exit Function
    .InitDir = StripOffFileName(lvWLLocations.SelectedItem.SubItems(1))
    .DialogTitle = "Word List Location For " & lvWLLocations.SelectedItem
    .Filter = "Word List (*.xml)|*.xml|All Files (*.*)|*.*"
    .FileName = ""
    .Flags = cdlOFNPathMustExist
      
    .CancelError = True
    .ShowOpen
    .FileName = LCase$(.FileName)
    cmdRemove.Enabled = True
      
    With lvWLLocations
      Dim i As Integer
      
      '**********************************************************
      '* Check if the specified file is already in the list of
      '* word lists.
      '**********************************************************
      For i = 1 To .ListItems.Count
        If (.ListItems(i).SubItems(1) = dlg.FileName) Then
          MsgBox dlg.FileName & " has already been chosen.", vbInformation + vbOKOnly, App.Title
          GetWLFileSpec = WLDuplicate
          Exit Function
        End If
      Next
    End With
    
    lvWLLocations.SelectedItem.SubItems(1) = .FileName
    
    Dim xmlWL As New clsXMLWordList
    
    '************************************************************
    '* If the user chose a file that doesn't exist then create
    '* a shell of an XML word list file with one empty category.
    '************************************************************
    If Not (FileExist(.FileName)) Then
      With xmlWL
        .LoadNew dlg.FileName
        .ID = lvWLLocations.SelectedItem
        .PhoneticFontName = DefaultPhoneticFont
        .PhoneticFontSize = DefaultPhoneticFontSize
        .AddCategory "New Category"
        .Save
      End With
      GetWLFileSpec = WLFileNew
    Else
      xmlWL.Load .FileName
      If (xmlWL.IsFileValidWL) Then
        lvWLLocations.SelectedItem.SubItems(1) = .FileName
        If (Len(xmlWL.ID) <> 0) Then lvWLLocations.SelectedItem.Text = xmlWL.ID
        GetWLFileSpec = WLFileExists
      Else
        MsgBox "Invalid word list file: " & .FileName, vbInformation + vbOKOnly, App.Title
        .FileName = ""
        GetWLFileSpec = WLFileInvalid
      End If
      
      Set xmlWL = Nothing
    End If
  End With

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GetWLFileSpecErr:
  Exit Function
  
End Function

Private Sub SelText(txtBox As TextBox)

  On Error Resume Next
  
  With txtBox
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
  
End Sub

Private Function TenTimes(sngNumber As Single) As Integer

  TenTimes = sngNumber * 10

End Function

Private Function TitleExists(sTitle$, Optional vItem, Optional vShowMsg) As Boolean

  '**********************************************************
  '* This routine determines whether or not a title exists in
  '* the list of word lists. If an item is supplied, that
  '* item will be exempted from the check.
  '**********************************************************
  
  Dim bShowMsg As Boolean
  Dim ExemptItem As ListItem
  Dim item As ListItem
  
  On Error Resume Next
  
  TitleExists = False
  bShowMsg = True
  If Not (IsMissing(vShowMsg)) Then bShowMsg = vShowMsg
  
  Set ExemptItem = Nothing
  If Not (IsMissing(vItem)) Then Set ExemptItem = vItem
  
  With lvWLLocations
    For Each item In .ListItems
      If Not (item Is ExemptItem) Then
        If (StrComp(item.Text, sTitle, vbTextCompare) = 0) Then
          If (bShowMsg) Then _
            MsgBox "'" & sTitle & "' already exists.", vbOKOnly + vbInformation, App.Title
          TitleExists = True
          Exit Function
        End If
      End If
    Next
  End With
  
End Function

Private Sub ValidateNumber(txtBox As TextBox, bMustBeInt As Boolean)

  On Error Resume Next
  
  With txtBox
    If (Len(.Text) = 0) Then
      .Text = 0
      Call SelText(txtBox)
    ElseIf (Not (IsNumeric(.Text)) Or (bMustBeInt And InStr(.Text, ".") > 0)) Then
      Beep
      .Text = .Tag
      Call SelText(txtBox)
      Exit Sub
    End If
      
    .Tag = .Text
  End With

End Sub

Private Function ValidateSRSpeed() As Boolean

  On Error Resume Next
  
  With txtSRSpeed
    If (Val(.Text) < 10 Or Val(.Text) > 333) Then
      MsgBox "The minimum and maximum values" & vbCrLf & _
             "for this field are 10 and 333.", vbInformation, App.Title
      .SetFocus
      Call SelText(txtSRSpeed)
      ValidateSRSpeed = False
      Exit Function
    End If
  End With

  ValidateSRSpeed = True

End Function

Private Sub cmdAdd_Click()

  Dim i As Integer
  Dim iRet As Integer
  Dim sTmpTitle As String
  
  i = 0
  
  '*****************************************************
  '* Find a unique, temporary name to give the added
  '* word list.
  '*****************************************************
  Do
    sTmpTitle = NewTitle & IIf(i = 0, "", " " & i)
    i = i + 1
  Loop While (TitleExists(sTmpTitle, , False))
  
  With lvWLLocations
    '***************************************************
    '* Add the temp. name to the list of word lists.
    '***************************************************
    .ListItems.Add , , sTmpTitle
    Set .SelectedItem = .ListItems(.ListItems.Count)
  
    '***************************************************
    '* Open a file/open dialog box.
    '***************************************************
    iRet = GetWLFileSpec()
    
    '***************************************************
    '* User canceled when browsing for different or
    '* new word list file, or the file chosen exits
    '* but is invalid.
    '***************************************************
    If (iRet = WLFileBrowseCancel Or iRet = WLFileInvalid Or iRet = WLDuplicate) Then
      .ListItems.Remove .SelectedItem.Index
      Exit Sub
    End If
  
    .SelectedItem.SubItems(1) = dlg.FileName
    .SetFocus
    cmdRemove.Enabled = True
    cmdBrowseWL.Enabled = (.ListItems.Count > 0)
    
    '***************************************************
    '* File user specified doesn't exist so be helpful
    '* and put the user automatically in the word list
    '* title's edit mode.
    '***************************************************
    If (iRet = WLFileNew) Then
      .SetFocus
      .StartLabelEdit
      SendKeys "{Home}+{End}"
    End If
  End With

End Sub

Private Sub cmdBrowse_Click(Index As Integer)

  On Error GoTo BrowseCancel
  
  If (Index = 0) Then
    Call GetSoundsPath
    Exit Sub
  Else
    With dlg
      If (Index = 0) Then
      Else
        .InitDir = StripOffFileName(txtSASLoc.Text)
        .DialogTitle = "Speech Analyzer Server Location (Current: " & txtSASLoc.Text & ")"
        .Filter = "Speech Analyzer Server (*.exe)"
        .FileName = "*.exe;*.pif"
        .Flags = cdlOFNFileMustExist + cdlOFNPathMustExist
      End If
      
      .CancelError = True
      .ShowOpen
      .FileName = LCase$(.FileName)
      
      txtSASLoc.Text = .FileName
    End With
  End If
  
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
BrowseCancel:
  Exit Sub

End Sub

Private Sub cmdBrowseWL_Click()
    
  Call GetWLFileSpec
  
End Sub

Private Sub cmdOKCancel_Click(Index As Integer)

  Dim i As Integer
  
  On Error Resume Next
  
  If (Index = 0) Then
    If Not (ValidateSRSpeed()) Then Exit Sub
    
    With lvWLLocations
      If (.ListItems.Count = 0) Then
        Erase gWordListID
      Else
        ReDim gWordListID(0 To .ListItems.Count - 1, 0 To 1)
        
        For i = 0 To UBound(gWordListID, 1)
          gWordListID(i, 0) = .ListItems(i + 1).Text
          gWordListID(i, 1) = .ListItems(i + 1).SubItems(1)
        Next
        
        Call ConvertWLFilesToXML
      End If
      
      Call WriteWordListsToIni
    End With
    
    gWavPath = txtSndsLoc.Text
    Call WriteINIEntry(cPathsSect, cSoundsEntry, gWavPath, gINIPath)
    
    gSAPath = txtSASLoc.Text
    Call WriteINIEntry(cPathsSect, cSAINIEntry, gSAPath, gINIPath)
  
    gSRSpeed = txtSRSpeed.Text
    Call SetSlowedReplayStatLine
    
    Call frmMenu.SetupWordListOptions
    
    Call WriteINIEntry(cSettingsSect, cPlayInitDelayEntry, txtInitDelay.Text, gINIPath)
    Call WriteINIEntry(cSettingsSect, cRepeatCountEntry, txtRptCount.Text, gINIPath)
    Call WriteINIEntry(cSettingsSect, cPlayRepeatDelayEntry, txtRptDelay.Text, gINIPath)
    Call WriteINIEntry(cSettingsSect, cSRSpeedEntry, gSRSpeed, gINIPath)
  End If
  
  Unload Me
  Exit Sub
  
End Sub

Private Sub cmdRemove_Click()

  Dim i As Integer
  
  On Error Resume Next
  
  With lvWLLocations
    If Not (.SelectedItem Is Nothing) Then
      .ListItems.Remove .SelectedItem.Index
      If (.ListItems.Count > 0) Then
        Set .SelectedItem = .ListItems(1)
        cmdRemove.Enabled = True
      Else
        cmdRemove.Enabled = False
      End If
    End If
  
    .SetFocus
    cmdBrowseWL.Enabled = (.ListItems.Count > 0)
  End With
  
End Sub

Private Sub Form_Load()

  Dim sTemp As String
  
  On Error Resume Next
  
  With txtInitDelay
    .Text = GetINIEntry$(cSettingsSect, cPlayInitDelayEntry, gINIPath)
    If Not (IsNumeric(.Text)) Then .Text = 0
    .Tag = .Text
    updnDelay(0).Value = TenTimes(.Text)
  End With
  
  With txtRptCount
    .Text = GetINIEntry$(cSettingsSect, cRepeatCountEntry, gINIPath)
    If Not (IsNumeric(.Text)) Then .Text = 0
    .Tag = .Text
  End With
  
  With txtRptDelay
    .Text = GetINIEntry$(cSettingsSect, cPlayRepeatDelayEntry, gINIPath)
    If Not (IsNumeric(.Text)) Then .Text = 0
    .Tag = .Text
    updnDelay(1).Value = TenTimes(.Text)
  End With
  
  With txtSRSpeed
    sTemp = GetINIEntry(cSettingsSect, cSRSpeedEntry, gINIPath)
    .Text = IIf(Len(sTemp) = 0 Or Not IsNumeric(sTemp), 50, Val(sTemp))
    .Tag = .Text
  End With
  
  With lvWLLocations
    Dim vWLPaths As Variant
    
    vWLPaths = GetAllINISettings(gINIPath, "WordListPaths")
    
    '**************************************************************
    '* Fill the word list locations from the ini file.
    '**************************************************************
    If Not (IsNull(vWLPaths)) Then
      Dim i As Integer
      Dim j As Integer
      
      j = 1
      For i = 0 To UBound(vWLPaths)
        sTemp = Trim$(IIf(StrComp(Trim$(vWLPaths(i, 1)), "none", 1) = 0, "", vWLPaths(i, 1)))
        If (Len(sTemp) > 0) Then
          .ListItems.Add , , vWLPaths(i, 0)
          .ListItems(j).SubItems(1) = sTemp
          j = j + 1
        End If
      Next
    End If
    
    '************************************************************
    '* Select the first item in the list and enable the browse
    '* and remove buttons if there are items in the list.
    '************************************************************
    Set .SelectedItem = .ListItems(1)
    cmdBrowseWL.Enabled = (.ListItems.Count > 0)
    cmdRemove.Enabled = (.ListItems.Count > 0)
  
    .SetFocus
  
    '************************************************************
    '* Set the column widths from values from the ini file.
    '************************************************************
    sTemp = GetINIEntry(cSettingsSect, LVColWidINIEntry & "1", gINIPath)
    If (Len(sTemp) > 0) Then .ColumnHeaders(1).Width = Val(sTemp)
    sTemp = GetINIEntry(cSettingsSect, LVColWidINIEntry & "2", gINIPath)
    If (Len(sTemp) > 0) Then
      .ColumnHeaders(2).Width = Val(sTemp)
    Else
      .ColumnHeaders(2).Width = .Width - .ColumnHeaders(1).Width - (gPixelX * 6)
    End If
  End With
  
  txtSndsLoc.Text = GetINIEntry$(cPathsSect, cSoundsEntry, gINIPath)
  txtSASLoc.Text = GetINIEntry$(cPathsSect, cSAINIEntry, gINIPath)
  
  Call CenterForm(Me, True)
  bGetFileSpecAfterTitleEdit = False
  bAddExistingWLFile = False
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  On Error Resume Next
  
  '************************************************************
  '* Write the word list location column widths to ini file.
  '************************************************************
  With lvWLLocations
    Call WriteINIEntry(cSettingsSect, LVColWidINIEntry & "1", .ColumnHeaders(1).Width, gINIPath)
    Call WriteINIEntry(cSettingsSect, LVColWidINIEntry & "2", .ColumnHeaders(2).Width, gINIPath)
  End With
  
  Set frmPreferences = Nothing
  
End Sub

Private Sub lvWLLocations_AfterLabelEdit(Cancel As Integer, NewString As String)

  On Error Resume Next
  cmdOKCancel(1).Cancel = True
  
  With lvWLLocations
    Dim xmlWL As New clsXMLWordList
    xmlWL.Load .SelectedItem.SubItems(1)
    
    If (xmlWL.IsFileValidWL()) Then
      xmlWL.ID = NewString
      xmlWL.Save
    Else
      MsgBox "Invalid word list file: " & .SelectedItem.SubItems(1), _
        vbInformation + vbOKOnly, App.Title
      Cancel = True
    End If
    
    Set xmlWL = Nothing
  End With
  
End Sub

Private Sub lvWLLocations_BeforeLabelEdit(Cancel As Integer)

  On Error Resume Next
  cmdOKCancel(1).Cancel = False
  
End Sub

Private Sub lvWLLocations_DblClick()

  On Error Resume Next
  Call GetWLFileSpec

End Sub

Private Sub lvWLLocations_KeyDown(KeyCode As Integer, Shift As Integer)

  On Error Resume Next
  If (KeyCode = vbKeyF2 And Shift = 0) Then lvWLLocations.StartLabelEdit

End Sub

Private Sub txtInitDelay_Change()

  On Error Resume Next
  Call ValidateNumber(txtInitDelay, False)
  updnDelay(0).Value = TenTimes(txtInitDelay.Text)

End Sub

Private Sub txtInitDelay_GotFocus()

  On Error Resume Next
  Call SelText(txtInitDelay)

End Sub

Private Sub txtRptCount_Change()

  On Error Resume Next
  Call ValidateNumber(txtRptCount, True)
  
End Sub

Private Sub txtRptCount_GotFocus()

  On Error Resume Next
  Call SelText(txtRptCount)

End Sub

Private Sub txtRptDelay_Change()

  On Error Resume Next
  Call ValidateNumber(txtRptDelay, False)
  updnDelay(1).Value = TenTimes(txtRptDelay.Text)

End Sub

Private Sub txtRptDelay_GotFocus()

  On Error Resume Next
  Call SelText(txtRptDelay)

End Sub

Private Sub txtSRSpeed_Change()

  On Error Resume Next
  Call ValidateNumber(txtSRSpeed, False)

End Sub

Private Sub txtSRSpeed_GotFocus()

  On Error Resume Next
  Call SelText(txtSRSpeed)

End Sub

Private Sub txtSRSpeed_LostFocus()

  On Error Resume Next
  Call ValidateSRSpeed

End Sub

Private Sub updnDelay_Change(Index As Integer)

  On Error Resume Next
  
  With updnDelay(Index)
    If (Index = 0) Then
      txtInitDelay.Text = IIf(.Value = 0, 0, Format$(.Value / 10, "##0.0"))
    Else
      txtRptDelay.Text = IIf(.Value = 0, 0, Format$(.Value / 10, "##0.0"))
    End If
  End With

End Sub
