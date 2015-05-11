VERSION 5.00
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "mci32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.MDIForm mdiHelpCharts 
   BackColor       =   &H8000000C&
   Caption         =   "IPA Help"
   ClientHeight    =   5310
   ClientLeft      =   3345
   ClientTop       =   2190
   ClientWidth     =   8250
   Icon            =   "MDIMAIN.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar panStatus 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   4935
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ilDisabled 
      Left            =   2370
      Top             =   2655
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   37
      ImageHeight     =   16
      MaskColor       =   16776960
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":04DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":06B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":0886
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":0A5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":0C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":0E02
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":0FD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":11AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":137E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   635
      ButtonWidth     =   1164
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ilEnabled"
      DisabledImageList=   "ilDisabled"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PlayOnly"
            Object.ToolTipText     =   "Play Sample Sound"
            Object.Tag             =   "Listen to a sample of the selected sound by itself"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PlayInterVocalic"
            Object.ToolTipText     =   "Play Sound in Context"
            Object.Tag             =   "Listen to a sample of the selected sound in a sample context"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PlaySlow"
            Object.ToolTipText     =   "Slowed Replay"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PlaySeparator"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Record"
            Object.ToolTipText     =   "Record For Comparison"
            Object.Tag             =   "Record yourself for comparison with prerecorded sound"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "StopRec"
            Object.ToolTipText     =   "Stop"
            Object.Tag             =   "Stop recording or playback"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PlayRec"
            Object.ToolTipText     =   "Play Recording"
            Object.Tag             =   "Play recording of your sound"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PlayRecSpeaker"
            Object.ToolTipText     =   "Compare Recording w/Prerecording"
            Object.Tag             =   "Compare your recording with the prerecorded sound"
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "RecSeparator"
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pitch"
            Object.ToolTipText     =   "Pitch Graph"
            Object.Tag             =   "Show a pitch graph of the selected item(s)"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "PitchSeparator"
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Test"
            Object.ToolTipText     =   "Listening Test"
            Object.Tag             =   "Start a listening test"
            ImageIndex      =   8
            Style           =   1
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TestSeparator"
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveEdits"
            Object.ToolTipText     =   "Save and Exit Edit Mode"
            Object.Tag             =   "Saves changes and exits edit mode"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SaveEditSeparator"
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            Object.ToolTipText     =   "Close Current Window"
            Object.Tag             =   "Close the current window"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.TextBox txtEditModeIndicator 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   6975
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "EDIT MODE"
         Top             =   75
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MCI.MMControl mciIPA 
         Height          =   330
         Left            =   7080
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   582
         _Version        =   393216
         DeviceType      =   ""
         FileName        =   ""
      End
   End
   Begin VB.Timer Timer2 
      Left            =   330
      Top             =   3585
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   495
      Top             =   3045
   End
   Begin MSComctlLib.ImageList ilEnabled 
      Left            =   1575
      Top             =   2595
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   37
      ImageHeight     =   16
      MaskColor       =   16776960
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":1552
            Key             =   "Record"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":1726
            Key             =   "StopRecord"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":18FA
            Key             =   "PlayRecord"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":1ACE
            Key             =   "PlayRecordSpeaker"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":1CA2
            Key             =   "Play"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":1E76
            Key             =   "PlayInterVocalic"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":204A
            Key             =   "Pitch"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":221E
            Key             =   "Test"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":23F2
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":25C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":279A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":296E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":2A72
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMAIN.frx":2B76
            Key             =   "SaveEdits"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPlayback 
         Caption         =   "&Playback"
         Index           =   0
         Shortcut        =   {F4}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPlayback 
         Caption         =   "&Slowed Playback"
         Index           =   1
         Shortcut        =   {F5}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPlayback 
         Caption         =   "-"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExportBitmap 
         Caption         =   "&Export as Bitmap to Clipboard"
         Visible         =   0   'False
         Begin VB.Menu mnuBkgrdColor 
            Caption         =   "&White Background"
            Index           =   0
         End
         Begin VB.Menu mnuBkgrdColor 
            Caption         =   "&Colored Background"
            Index           =   1
         End
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPreferences 
         Caption         =   "Pre&ferences"
      End
      Begin VB.Menu Space10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuIPAParent 
      Caption         =   "&IPA"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuIPA 
         Caption         =   "&Consonants"
         Index           =   0
      End
      Begin VB.Menu mnuIPA 
         Caption         =   "&Vowels"
         Index           =   1
      End
      Begin VB.Menu mnuIPA 
         Caption         =   "&Diacritics"
         Index           =   2
      End
      Begin VB.Menu mnuIPA 
         Caption         =   "&Suprasegmentals"
         Index           =   3
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Visible         =   0   'False
      Begin VB.Menu mnuEditMode 
         Caption         =   "&Enter Edit Mode"
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFonts 
         Caption         =   "&Fonts..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditTitle 
         Caption         =   "Word List &Title..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditSoundPath 
         Caption         =   "Word List's &Sound File Path..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAddNewCategory 
         Caption         =   "&Add New Category"
         Enabled         =   0   'False
         Begin VB.Menu mnuInsert 
            Caption         =   "Insert &Before Current"
            Index           =   0
         End
         Begin VB.Menu mnuInsert 
            Caption         =   "Insert &After Current"
            Index           =   1
         End
      End
      Begin VB.Menu mnuDeleteCategory 
         Caption         =   "&Delete Current Category"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTest 
      Caption         =   "&Test"
      Begin VB.Menu mnuTestSetup 
         Caption         =   "Se&tup"
      End
      Begin VB.Menu mnuTestStop 
         Caption         =   "&Stop"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAmer 
      Caption         =   "&Americanist"
      Visible         =   0   'False
      Begin VB.Menu mnuAmerCon 
         Caption         =   "&Consonants"
      End
      Begin VB.Menu mnuAmerVow 
         Caption         =   "&Vowels"
      End
   End
   Begin VB.Menu mnuConversion 
      Caption         =   "&Conversions"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuUsing 
         Caption         =   "&Help for IPA Help..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About IPA Help..."
      End
   End
   Begin VB.Menu mnuSILCons 
      Caption         =   "SILCons"
      Visible         =   0   'False
      Begin VB.Menu mnuSILConsSILIPA 
         Caption         =   "Americanist - &IPA"
      End
      Begin VB.Menu mnuSILConsSILOnly 
         Caption         =   "&Americanist Only"
      End
   End
   Begin VB.Menu mnuSILVows 
      Caption         =   "SILVows"
      Visible         =   0   'False
      Begin VB.Menu mnuSILVowsSILIPA 
         Caption         =   "Americanist - &IPA"
      End
      Begin VB.Menu mnuSILVowsSILOnly 
         Caption         =   "&Americanist Only"
      End
   End
End
Attribute VB_Name = "mdiHelpCharts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**************************************************
'* mdiHelpCharts version info:
'*  See basGlobals (Globals.bas).
'**************************************************

Option Explicit

Public bStartTest As Boolean                        '* Accessed from frmTestSetup.
Private bctrlkey As Boolean
Private bFormWasLoaded As Boolean                   '* Determines whether or not to reopen a form.
Private bDontRestoreTBAfterPlay As Boolean
Private EnterCount As Integer
Private dRunningLen As Double                       '* Used by MCI control.
Private dRecLength As Double                        '* Used by MCI control.

Public Sub CallSA()
  
  Dim sArguments As String
  
  On Error GoTo CallSAErr
  
  '*******************************************************
  '* Show List file if Ctrl + Click
  '*******************************************************
  If (bctrlkey) Then Call Shell("Notepad " & gListFilePath, vbNormalFocus)
  
  '*******************************************************
  '* Generate Command line string
  '*******************************************************
  gSAPath = GetINIEntry$(cPathsSect, cSAINIEntry, gINIPath)
  sArguments = " -l " & gListFilePath
  '*******************************************************
  '* Show Command line if Ctrl + Click
  '*******************************************************
  If (bctrlkey) Then
    If (MsgBox("Command Line: " & vbCrLf & sArguments & vbCrLf & vbCrLf & _
        "Do you want to run this?", vbYesNo) = vbNo) Then Exit Sub
  End If
  
  Call Shell(gSAPath & sArguments, vbNormalFocus)

  bctrlkey = False
  Exit Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
CallSAErr:
  If (Err.Number = 53) Then
    MsgBox gSAPath & " not found.", vbInformation, App.Title
  Else
    MsgBox Err.Description, vbInformation, App.Title
  End If

End Sub

Public Sub EnableTBarButtons(sKeys$)

  Dim i As Integer
  Dim j As Integer
  
  On Error Resume Next
  
  With TBar
    For i = 1 To .Buttons.Count
      If (.Buttons(i).Style <> tbrSeparator And .Buttons(i).Visible) Then
        j = InStr(sKeys, .Buttons(i).Key & ";")
      
        If (InStr(.Buttons(i).Key, "PlayRec") > 0 And j > 0) Then
          .Buttons(i).Enabled = FileExist(gTmpWavPath & gTmpWavName)
        
        ElseIf (.Buttons(i).Key = "StopRec" And j > 0) Then
          .Buttons(i).Enabled = (.Buttons("Record").Value = tbrPressed Or _
                                 .Buttons("PlayRec").Value = tbrPressed Or _
                                 .Buttons("PlayRecSpeaker").Value = tbrPressed)
        Else
          .Buttons(i).Enabled = (j > 0)
        End If
      End If
    Next
    
    .Refresh
  End With

End Sub

Private Sub LoadWordLists()

  Dim i As Integer
  Dim vINISettings As Variant
  
  '************************************************
  '* Word Lists
  '************************************************
  vINISettings = GetAllINISettings(gINIPath, "WordListPaths")
  
  If Not (IsNull(vINISettings)) Then
    ReDim gWordListID(0 To UBound(vINISettings), 0 To 1) As Variant
  
    For i = 0 To UBound(vINISettings)
      gWordListID(i, 0) = vINISettings(i, 0)
      gWordListID(i, 1) = vINISettings(i, 1)
    Next
  
    Call ConvertWLFilesToXML
    Call WriteWordListsToIni
    gStatLine.SimpleText = ""
  End If
  
End Sub

Private Sub MakePhonGrpColl()

  Dim AllConGroup, PlosConGroup, NTFConGroup, _
    FricConGroup, ApprConGroup, NonPulConGroup, _
    OtherConGroup, AllVowGroup, FrontVowGroup, _
    CentVowGroup, BackVowGroup As Variant
  
  On Error Resume Next
  
  AllConGroup = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, _
                    14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, _
                    25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, _
                    36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, _
                    47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, _
                    58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, _
                    69, 70, 72, 73, 74, 75, 76, 77, 78, 79, 80, _
                    81, 82, 83)
  PlosConGroup = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
  NTFConGroup = Array(13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24)
  FricConGroup = Array(25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, _
                    36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48)
  ApprConGroup = Array(49, 50, 51, 52, 53, 54, 55, 56, 57)
  NonPulConGroup = Array(58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70)
  OtherConGroup = Array(72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83)
  AllVowGroup = Array(0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, _
                    15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27)
  FrontVowGroup = Array(0, 1, 6, 7, 9, 10, 16, 17, 22, 24, 25)
  CentVowGroup = Array(2, 3, 11, 12, 15, 18, 19, 23)
  BackVowGroup = Array(4, 5, 8, 13, 14, 20, 21, 26, 27)
  With gPhonGrpColl
    .Add item:=AllConGroup, Key:="AllCon"
    .Add item:=PlosConGroup, Key:="PlosCon"
    .Add item:=NTFConGroup, Key:="NTFCon"
    .Add item:=FricConGroup, Key:="FricCon"
    .Add item:=ApprConGroup, Key:="ApprCon"
    .Add item:=NonPulConGroup, Key:="NonPulCon"
    .Add item:=OtherConGroup, Key:="OtherCon"
    .Add item:=AllVowGroup, Key:="AllVow"
    .Add item:=FrontVowGroup, Key:="FrontVow"
    .Add item:=CentVowGroup, Key:="CentVow"
    .Add item:=BackVowGroup, Key:="BackVow"
  End With
  gPhonGrpNameArray = Array("AllCon", "PlosCon", "NTFCon", _
            "FricCon", "ApprCon", "NonPulCon", "OtherCon", _
            "AllVow", "FrontVow", "CentVow", "BackVow")
  gTestGrpName = "AllCon"
  gTestLayout = cTestChart
  gRetestListEmpty = True
  
End Sub

Private Sub PlayRecording(iButtonState%, bCompare As Boolean)

  On Error Resume Next
  
  '****************************************************************
  '* If we're here and the button's state is not pressed, it
  '* means the user pressed the button while it was down in order
  '* to stop playback. Therefore, stop playback.
  '****************************************************************
  If (iButtonState = tbrUnpressed) Then
    TBar.Buttons("PlayRec").Value = tbrUnpressed
    TBar.Buttons("PlayRecSpeaker").Value = tbrUnpressed
    mnuFile.Enabled = True
    mnuTest.Enabled = True
    mnuWindow.Enabled = True
    mnuHelp.Enabled = True
    Call ActiveForm.UpdateAfterRecordAndPlayback
    gMMCtrl.Tag = ""
    Exit Sub
  End If
    
  '****************************************************************
  '* If MCI control is busy then don't allow playback.
  '****************************************************************
  If (Len(gMMCtrl.Tag) > 0) Then
    Beep
    TBar.Buttons("PlayRec").Value = tbrUnpressed
    TBar.Buttons("PlayRecSpeaker").Value = tbrUnpressed
    Exit Sub
  End If
  
  Call EnableTBarButtons("StopRec;" & IIf(bCompare, "PlayRecSpeaker;", "PlayRec;"))
  mnuFile.Enabled = False
  mnuTest.Enabled = False
  mnuWindow.Enabled = False
  mnuHelp.Enabled = False
  
  gMMCtrl.Tag = cMCIBusy
  bDontRestoreTBAfterPlay = True
  Call PlayWav(gTmpWavPath & gTmpWavName)
  gMMCtrl.Tag = ""
  
  If (bCompare) Then
    With gMMCtrl
      While (.Mode <> mciModeNotOpen And .Mode <> mciModeReady): DoEvents: Wend
      .Command = "Close"
    End With
    bDontRestoreTBAfterPlay = False
    Call ActiveForm.Play(0)
  End If

  bDontRestoreTBAfterPlay = False

End Sub

Private Sub RecordUser(iButtonState%)

  On Error Resume Next
      
  '****************************************************************
  '* If we're here and the button's state is not pressed, it
  '* means the user pressed the button while it was down in order
  '* to stop recording. Therefore, stop recording.
  '****************************************************************
  If (iButtonState = tbrUnpressed) Then
    Call StopRecording
    Exit Sub
  End If
  
  '****************************************************************
  '* If MCI control is busy then don't allow recoding.
  '****************************************************************
  If (Len(gMMCtrl.Tag) > 0) Then
    Beep
    TBar.Buttons("Record").Value = tbrUnpressed
    Exit Sub
  End If
  
  '***********************************************************************
  '* We make use of a Template file (gMstrWavName) which automatically
  '* sets the sampling frequency and data size.
  '***********************************************************************
  Kill gTmpWavPath & gTmpWavName
  If FileExist(gWavPath & gMstrWavName) Then FileCopy (gWavPath & gMstrWavName), gTmpWavName
  Call EnableTBarButtons("Record;StopRec;")
        
  With gMMCtrl
    .Tag = cMCIBusy
    .FileName = gTmpWavPath & gTmpWavName
    .Command = "Open"
    gMMCtrl.Command = "Record"
  End With
 
End Sub

Public Sub ShowTBarButtons(sKeys As String)

  Dim i As Integer
  
  On Error Resume Next
  
  With TBar
    For i = 1 To .Buttons.Count
      .Buttons(i).Visible = (InStr(sKeys, .Buttons(i).Key & ";") > 0)
    Next

    If Not (.Visible) Then .Visible = True

    '*********************************************************************
    '* Remove the following two lines when a link with SA is working.
    '*********************************************************************
    .Buttons("PlaySlow").Enabled = False
    .Buttons("Pitch").Enabled = False
  End With

End Sub

Public Sub StartTest()

  Screen.MousePointer = vbHourglass

  Dim frm As Form
  Dim i As Integer

  '*********************
  '* Activate test mode
  '*********************
  gTestActive = True
  If (TBar.Buttons("Test").Value <> tbrPressed) Then TBar.Buttons("Test").Value = tbrPressed
  
  '****************************************************
  '* Show form for test and disable unused chars.
  '* Reworked to accomodate Word Lists by CLW 5/5/99
  '****************************************************
  bFormWasLoaded = False
  For Each frm In Forms
    With frm
      If (.Name = gTestForm And .Tag = gTestTag) Then
        .Show
        Call .UpdateFormForTest
        bFormWasLoaded = True
        Exit For
      End If
    End With
  Next frm
  
  If Not (bFormWasLoaded) Then
    Select Case gTestForm
      Case "frmDispCon":  frmDispCon.Show
                          DoEvents
                          Call frmDispCon.UpdateFormForTest
      
      Case "frmDispVow":  frmDispVow.Show
                          DoEvents
                          Call frmDispVow.UpdateFormForTest
      
      Case "frmWordList": Set frm = New frmWordList
                          DoEvents
                          frm.Initialize gTestTag
                          Call frm.UpdateFormForTest
    End Select
  End If
  
  DoEvents
  Screen.MousePointer = vbDefault
  
  '****************************************
  '* Disable extra menus. Update toobar.
  '****************************************
  mnuFile.Enabled = False
  mnuTestSetup.Enabled = False
  mnuWindow.Enabled = False
  mnuTestStop.Enabled = True
  Call EnableTBarButtons("Test;")
  
  '*********************************************
  '* 2 second delay before playing first sound.
  '*********************************************
  If (gTestForm <> "frmWordList") Then 'Or gTestCatChoice = cRandom Then '* Added by CLW 5/11/99
    gFirstSndPlayed = False
    With Timer2
      .Interval = 2000
      .Enabled = True
    End With
    Call RandomTest
  Else
    Call Pause(1)
    Call frm.TestPlay
  End If

End Sub

Private Sub StopRecording()

  On Error Resume Next
    
  With gMMCtrl
    .Command = "Stop"
    .Command = "Save"
    .Command = "Close"
    .Tag = ""
  End With
      
  TBar.Buttons("Record").Value = tbrUnpressed
  Call ActiveForm.UpdateAfterRecordAndPlayback

End Sub

Private Sub StopTest()
  
  Dim sUpdateGrpName As String
  Dim vUpdateGroup As Variant
  Dim frm As Form
  
  On Error Resume Next
  
  '* First, see if user switched forms during the test.
  '* Added by CLW 5/12/99
  If ActiveForm.Name <> gTestForm Then
    For Each frm In Forms
      If (frm.Name = gTestForm And frm.Tag = gTestTag) Then Exit For
    Next frm
  End If
  
  '********************************************
  '* If the test is active, that means it was
  '* not aborted while the test setup form
  '* was open. This is necessary to know when
  '* deciding to reopen the active form or not.
  '********************************************
  If gTestActive Then
    If (gTestForm <> "frmWordList") Then            '* Added by CLW 5/11/99
      sUpdateGrpName = "All" & Right(gTestGrpName, 3)
      vUpdateGroup = gPhonGrpColl(sUpdateGrpName)
      Call UpdateRetestList                         '* Moved by CLW 5/11/99
    End If
    gTestActive = False
    If bFormWasLoaded Then
      Call ActiveForm.UpdateFormAfterTest
    Else
      Unload ActiveForm
    End If
    
    Call RandomTest
  End If
  
  TBar.Buttons("Test").Value = tbrUnpressed
  mnuFile.Enabled = True
  mnuTestSetup.Enabled = True
  mnuWindow.Enabled = True
  mnuTestStop.Enabled = False

End Sub

Private Sub mciIPA_Done(NotifyCode As Integer)

  With gMMCtrl
    If (Not bDontRestoreTBAfterPlay And _
       (TBar.Buttons("PlayRec").Value = tbrPressed Or _
        TBar.Buttons("PlayRecSpeaker").Value = tbrPressed)) Then
      
      Call PlayRecording(tbrUnpressed, False)
    End If
    
    .Command = "Close"
  End With

End Sub

Private Sub mciIPA_StatusUpdate()

  '************************************************
  '* Keep checking to see if the recording is
  '* 2 seconds long yet. If it is, stop recording.
  '************************************************
  With gMMCtrl
    'Debug.Print .length; " "; .Mode
    If (.Mode = mciModeRecord And .length = 2000) Then _
      Call StopRecording
  End With
  
End Sub

Private Sub MDIForm_Load()

  Dim i As Integer
  Dim sTemp As String
  Dim vSoundPath As Variant
  Dim vINISettings As Variant
  Dim AllConGroup, PlosConGroup, NTFConGroup, _
      FricConGroup, ApprConGroup, NonPulConGroup, _
      OtherConGroup, AllVowGroup, FrontVowGroup, _
      CentVowGroup, BackVowGroup As Variant
  Dim sTempPath As String
  Dim iLen As Integer
  Dim sAccessTestPath As String
  Dim sRegPath As String
  Dim sRegValueName As String
  Dim sRegValue As String
  Dim lRegDataLen As Long
  Dim lRetVal As Long 'Result of RegOpenKeyEx
  Dim hKey, hSubKey As Long
  Dim ret
  
  On Error GoTo MDIFormLoadErr
    
  If Not CheckForNeededFiles() Then End
  Set gStatLine = panStatus
  Set gMMCtrl = mciIPA
  gStatLine.SimpleText = ""
  bctrlkey = False
  bDontRestoreTBAfterPlay = False

  '************************************************
  '* INI file and help file stuff.
  '************************************************
  gINIPath = App.Path & IIf(Len(App.Path) = 3, "", "\") & INIFile
  
  vSoundPath = GetINIEntry(cPathsSect, cSoundsEntry, gINIPath)
  gWavPath = App.Path & IIf(Right$(App.Path, 1) = "\", "", "\")        'IPA Help folder is default Sounds path
  
  '*************************************************
  '* If the Sounds path key was there then set the
  '* global for that value. Otherwise look for one.
  '*************************************************
  If (Len(vSoundPath > 0) And FileExist(vSoundPath)) Then
    gWavPath = vSoundPath
  Else
    If FileExist(gWavPath & "Sounds\*.*") Then gWavPath = gWavPath & "Sounds\"
  End If

  sTemp = GetINIEntry(cSettingsSect, cSRSpeedEntry, gINIPath)
  gSRSpeed = IIf(Len(sTemp) = 0, 50, Val(sTemp))

  sTempPath = Space(255)
  iLen = GetTempPath(255, sTempPath)
  sTempPath = Left(sTempPath, iLen)
  If iLen = 0 Then _
    sTempPath = App.Path
  sTempPath = Trim$(sTempPath)
  gListFilePath = MakeFullPath(sTempPath, "ipa-help.lst")
  gAAFilePath = MakeFullPath(sTempPath, "aa.txt")
 
  '************************************************
  ' Check if user has write access to app path and
  ' disable menu items if not.
  '************************************************
  gAppPathWriteAccess = False
  sAccessTestPath = MakeFullPath(App.Path, "AccessTest.ini")
  Call WriteINIEntry("Test", "WriteAccess", "True", sAccessTestPath)
  If (FileExist(sAccessTestPath)) Then
    gAppPathWriteAccess = True
    Kill sAccessTestPath
  Else
  End If
  mnuPreferences.Enabled = gAppPathWriteAccess
  mnuEditMode.Enabled = gAppPathWriteAccess
  
  '************************************************
  ' Make sure the IPA Help location is available in
  ' the registry.
  '************************************************
  sRegPath = "Software\SIL\IPA Help"
  sRegValueName = "Location"
  sRegValue = ""
  lRetVal = RegCreateKey(HKEY_CURRENT_USER, sRegPath, hKey)
  If lRetVal = 0 Then
    lRetVal = RegQueryValueExNULL(hKey, sRegValueName, 0&, REG_SZ, 0&, lRegDataLen)
    If lRetVal = 0 Then
      sRegValue = String(lRegDataLen - 1, 0)
      lRetVal = RegQueryValueExString(hKey, sRegValueName, 0&, REG_SZ, sRegValue, lRegDataLen)
    End If
    If Not (lRetVal = 0 And Len(sRegValue) > 0) Then
      'No value or the value is empty. Set the value.
      RegSetValueEx hKey, sRegValueName, 0, REG_SZ, ByVal App.Path, Len(App.Path)
    End If
  End If
  'RegSetValueEx hKey, sRegValueName, 0, REG_SZ, "X", Len("X")
  RegCloseKey hKey
  
  '************************************************
  '* Initialize the MCI control.
  '************************************************
  With gMMCtrl
    .PlayVisible = False
    .Shareable = False
    .DeviceType = "WaveAudio"
    .UpdateInterval = 100
    .TimeFormat = mciFormatMilliseconds
  End With
  
  gMstrWavName = "template.wav"
  gTmpWavPath = sTempPath & IIf(Right$(sTempPath, 1) = "\", "", "\")
  gTmpWavName = "~ipa-wav.tmp"
  
  On Error Resume Next
  Kill gTmpWavPath & gTmpWavName
  gPixelX = Screen.TwipsPerPixelX
  gPixelY = Screen.TwipsPerPixelY
  
  Height = 6645
  Width = 9105
  TBar.Visible = False
  Show
  Call LoadWordLists
  frmMenu.Show
  Call MakePhonGrpColl
  gSAPath = ""
  
  Exit Sub
  
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MDIFormLoadErr:
  MsgBox Err.Description, vbInformation, App.Title

End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  On Error Resume Next
  gStatLine.SimpleText = ""

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  Dim i As Integer
  
  On Error Resume Next
  Kill gTmpWavPath & gTmpWavName
  Set mdiHelpCharts = Nothing

End Sub

Private Sub mnuAbout_Click()

  On Error Resume Next
  frmHelpAbout.Show vbModal
  
End Sub

Private Sub mnuBkgrdColor_Click(Index As Integer)

  On Error Resume Next
  Call ActiveForm.IPAHelpPrint(False, Index = 1)

End Sub

Private Sub mnuDeleteCategory_Click()

  On Error Resume Next
  Call ActiveForm.DeleteCategory
  
End Sub

Private Sub mnuEditFonts_Click()

  On Error Resume Next
  Call ActiveForm.EditFonts
  
End Sub

Private Sub mnuEditMode_Click()

  On Error Resume Next
  Call ActiveForm.EditMode
  
End Sub

Private Sub mnuEditSoundPath_Click()

  On Error Resume Next
  Call ActiveForm.EditSoundPath
  
End Sub

Private Sub mnuEditTitle_Click()

  On Error Resume Next
  Call ActiveForm.EditTitle
  
End Sub

Private Sub mnuExit_Click()

  On Error Resume Next
  Unload Me

End Sub

Private Sub mnuFilePath_Click()

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
    If (Len(.sDir) > 0) Then gWavPath = .sDir
    bWavPathUpdate = True
  End With
  Unload frmFilePath
  
  '************************************************
  '* If there was an update to the Wave file path,
  '* then we need to update the play buttons and
  '* test menu (and button, if applicable).
  '************************************************
  If bWavPathUpdate Then
    With ActiveForm
      Select Case .Name
        Case "frmDispCon", "frmDispDia", "frmDispVow", _
        "frmDispVow"                                    'On forms that the test button can be enabled,
'          Call UpdatePlayBttns(.CurrIndex)
        Case Else
          Call UpdateTestMenu
      End Select
    End With
  End If
  bWavPathUpdate = False

End Sub

Private Sub mnuInsert_Click(Index As Integer)

  On Error Resume Next
  Call ActiveForm.AddNewCategory(Index)

End Sub

Public Sub mnuIPA_Click(Index As Integer)

  On Error Resume Next
  
  Screen.MousePointer = vbHourglass
  
  Select Case Index
    Case 0: frmDispCon.Show
    Case 1: frmDispVow.Show
    Case 2: frmDispDia.Show
    Case 3: frmDispSSeg.Show
  End Select

  Screen.MousePointer = vbDefault

End Sub

Private Sub mnuPlayback_Click(Index As Integer)
  On Error Resume Next
  If (Index) Then
    Call ActiveForm.PlaySlow
  Else
    Call ActiveForm.Play(0)
  End If
End Sub

Private Sub mnuPreferences_Click()

  On Error Resume Next
  frmPreferences.Show vbModal

End Sub

Private Sub mnuPrint_Click(Index As Integer)

  On Error Resume Next
  Call ActiveForm.IPAHelpPrint(True, False)

End Sub

Private Sub mnuSILConsSILIPA_Click()
  
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  frmDispAmerCon2.Show

End Sub

Private Sub mnuSILConsSILOnly_Click()
  
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  frmDispAmerCon.Show

End Sub

Private Sub mnuSILVowsSILIPA_Click()
  
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  frmDispAmerVow2.Show

End Sub

Private Sub mnuSILVowsSILOnly_Click()
  
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  frmDispAmerVow.Show

End Sub

Private Sub mnuTestSetup_Click()

  On Error Resume Next
  frmTestSetup.Show vbModal
  DoEvents
  If (bStartTest) Then Call StartTest

End Sub

Public Sub mnuTestStop_Click()

  On Error Resume Next
  Call StopTest
  
End Sub

Private Sub mnuUsing_Click()
  
  On Error Resume Next
  HtmlHelp Me.hwnd, App.HelpFile, HH_DISPLAY_TOC, 0

End Sub

Private Sub TBar_ButtonClick(ByVal Button As MSComctlLib.Button)

  On Error Resume Next
  
  Select Case Button.Key
    Case "Exit":             Unload IIf(Forms.Count = 1, Me, ActiveForm)
    Case "PlayOnly":         Call ActiveForm.Play(0)
    Case "PlayInterVocalic": Call ActiveForm.Play(1)
    Case "PlaySlow":         Call ActiveForm.PlaySlow
    Case "Record":           Call RecordUser(Button.Value)
    Case "PlayRec":          Call PlayRecording(Button.Value, False)
    Case "PlayRecSpeaker":   Call PlayRecording(Button.Value, True)
    Case "SaveEdits":        Call ActiveForm.EditMode(False)
    Case "Pitch":            Call ActiveForm.ShowPitchPlot
    
    Case "StopRec"
      If (gMMCtrl.Mode = mciModeRecord) Then
        Call StopRecording
      Else
        Call PlayRecording(tbrUnpressed, False)
      End If

    Case "Test"
      If (Button.Value = tbrPressed) Then
        Call mnuTestSetup_Click
      Else
        Call StopTest
      End If
  End Select

End Sub

Private Sub TBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  On Error Resume Next
  
  With TBar
    If ((X >= .Buttons("PlaySlow").Left And X < (.Buttons("PlaySlow").Left + .Buttons("PlaySlow").Width)) Or _
        (X >= .Buttons("Pitch").Left And X < (.Buttons("Pitch").Left + .Buttons("Pitch").Width))) Then _
      bctrlkey = ((Shift And vbCtrlMask) > 0)
  End With

End Sub

Private Sub TBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim i As Integer
  
  On Error Resume Next

  With TBar
    For i = 1 To .Buttons.Count
      With .Buttons(i)
        If (.Visible And .Style <> tbrSeparator And X >= .Left And X <= (.Left + .Width)) Then
          gStatLine.SimpleText = " " & TBar.Buttons(i).Tag
          Exit Sub
        End If
      End With
    Next
  End With
    
  gStatLine.SimpleText = ""

End Sub

Private Sub Timer1_Timer()
  
  '************************************************
  '* This timer is used for two purposes. One is
  '* to play the selected symbol when a form is
  '* activated (after a 2 second delay). The second
  '* purpose is to delay the playing of the
  '* vocalic sound until after the double-click
  '* time as set in windows. This gives time for
  '* the double-click event to take place.
  '* Otherwise, it would never occur.
  '************************************************
  On Error Resume Next
  Call ActiveForm.Play(0)
  Timer1.Enabled = False
  
End Sub

Private Sub Timer2_Timer()
  
  Dim sWavNamePart1 As String
  Dim i As Integer

  On Error Resume Next
  
  '* See if any clicks were made. If not mark incorrect
  If (gFirstSndPlayed And Not gTestItemCorrect) Then _
    Call MarkItemCorrect(vbUnchecked) '* CLW 1/27/99
    
  '************************************************
  '* Close previous test item.
  '************************************************
  With ActiveForm
    .lblSmile.Visible = False
    .lblFrown.Visible = False
  End With
  
  If gFirstSndPlayed Then
    Select Case ActiveForm.Name
      Case "frmDispCon"
        With frmDispCon.Con(gItemNumber)
          .ForeColor = vbButtonText
          .BackColor = vbButtonFace
        End With
      Case "frmDispVow"
        With frmDispVow.Vowel(gItemNumber)
          .ForeColor = vbButtonText
          .BackColor = vbButtonFace
        End With
      Case Else
    End Select
  End If
  
  If gFirstSndPlayed And _
  Timer2.Interval = 5000 And _
  Not gTestItemCorrect Then
    '**********************************************
    '* Show user the correct test item.
    '**********************************************
    Select Case ActiveForm.Name
      Case "frmDispCon"
        With frmDispCon
          .Con(gItemNumber).ForeColor = vbButtonText
          .Con(gItemNumber).BackColor = .lblSmile.BackColor
        End With
      Case "frmDispVow"
        With frmDispVow
          .Vowel(gItemNumber).ForeColor = vbButtonText
          .Vowel(gItemNumber).BackColor = .lblSmile.BackColor
        End With
      Case Else
    End Select
    Timer2.Interval = 2000
  Else
    '************************************************
    '* Begin new test item.
    '************************************************
    Do
      'get a random number between lowest and highest index
      i = Int((UBound(gTestGrp) - LBound(gTestGrp) + 1) * Rnd + LBound(gTestGrp))
      gItemNumber = gTestGrp(i)
    '* Check if valid index CLW 1/26/99
    Loop While (gTestGrpCorrect(cLastTest, i) = vbChecked)
      
    sWavNamePart1 = Right(gTestGrpName, 3)
    Call PlayWav(sWavNamePart1 & "-" & Format$(Trim$(Str$(gItemNumber)), "00") & IIf(sWavNamePart1 = "Vow", "W.wav", "A.wav")) ' "A.Wav")
    'ActiveForm.Label2.Caption = gItemNumber
    gFirstSndPlayed = True
    gTestItemCorrect = False '* CLW 1/27/99
    Timer2.Interval = 5000
  End If

End Sub
