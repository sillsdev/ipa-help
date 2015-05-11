VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BC496AED-9B4E-11CE-A6D5-0000C0BE9395}#2.0#0"; "ssdatb32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmWordList 
   Caption         =   "Word List"
   ClientHeight    =   5985
   ClientLeft      =   3375
   ClientTop       =   5040
   ClientWidth     =   9135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5985
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView tvSections 
      Height          =   1785
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   3149
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   4
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   2400
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picDragBar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   2760
      MouseIcon       =   "wordlist.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   1050
      ScaleWidth      =   180
      TabIndex        =   11
      Top             =   2520
      Width           =   180
   End
   Begin VB.PictureBox picSplitter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   2760
      MouseIcon       =   "wordlist.frx":0152
      MousePointer    =   99  'Custom
      ScaleHeight     =   1005
      ScaleWidth      =   165
      TabIndex        =   10
      Top             =   3720
      Width           =   165
   End
   Begin VB.PictureBox picTestMode 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3645
      Left            =   3000
      ScaleHeight     =   3615
      ScaleWidth      =   5805
      TabIndex        =   0
      Top             =   1875
      Visible         =   0   'False
      Width           =   5835
      Begin VB.TextBox txtUserTr 
         Height          =   390
         Left            =   75
         TabIndex        =   5
         Top             =   1305
         Width           =   5600
      End
      Begin VB.CommandButton cmdVerify 
         Caption         =   "&Verify"
         Height          =   360
         Left            =   2145
         TabIndex        =   4
         Top             =   2910
         Width           =   1200
      End
      Begin VB.TextBox txtCorrectTr 
         BackColor       =   &H8000000F&
         ForeColor       =   &H80000012&
         Height          =   390
         Left            =   75
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2220
         Width           =   5600
      End
      Begin VB.CommandButton cmdReplay 
         Caption         =   "&Replay"
         Default         =   -1  'True
         Height          =   360
         Left            =   810
         TabIndex        =   2
         Top             =   2910
         Width           =   1200
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next Word >"
         Height          =   360
         Left            =   3465
         TabIndex        =   1
         Top             =   2910
         Width           =   1200
      End
      Begin VB.Label lblTestMode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "TEST MODE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   2100
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Enter Transcription Here:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1770
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "&Correct Transcription:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   1995
         Width           =   1515
      End
   End
   Begin VB.Timer Timer1 
      Left            =   1440
      Top             =   1920
   End
   Begin SSDataWidgets_B.SSDBGrid ssGrid 
      Height          =   1680
      Left            =   3030
      TabIndex        =   9
      Top             =   60
      Width           =   5970
      _Version        =   131078
      DataMode        =   2
      BorderStyle     =   0
      Col.Count       =   7
      stylesets.count =   7
      stylesets(0).Name=   "Ortho"
      stylesets(0).ForeColor=   -2147483640
      stylesets(0).BackColor=   -2147483643
      stylesets(0).HasFont=   -1  'True
      BeginProperty stylesets(0).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(0).Picture=   "wordlist.frx":02A4
      stylesets(0).AlignmentText=   1
      stylesets(1).Name=   "Dialect"
      stylesets(1).ForeColor=   -2147483640
      stylesets(1).BackColor=   -2147483643
      stylesets(1).HasFont=   -1  'True
      BeginProperty stylesets(1).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(1).Picture=   "wordlist.frx":02C0
      stylesets(2).Name=   "Gloss"
      stylesets(2).ForeColor=   -2147483640
      stylesets(2).BackColor=   -2147483643
      stylesets(2).HasFont=   -1  'True
      BeginProperty stylesets(2).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(2).Picture=   "wordlist.frx":02DC
      stylesets(2).AlignmentText=   1
      stylesets(3).Name=   "IPA"
      stylesets(3).ForeColor=   -2147483631
      stylesets(3).BackColor=   -2147483633
      stylesets(3).Picture=   "wordlist.frx":02F8
      stylesets(4).Name=   "Normal"
      stylesets(4).ForeColor=   -2147483640
      stylesets(4).BackColor=   -2147483643
      stylesets(4).HasFont=   -1  'True
      BeginProperty stylesets(4).Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      stylesets(4).Picture=   "wordlist.frx":0314
      stylesets(4).AlignmentText=   1
      stylesets(5).Name=   "IPAAudio"
      stylesets(5).ForeColor=   -2147483640
      stylesets(5).BackColor=   -2147483643
      stylesets(5).Picture=   "wordlist.frx":0330
      stylesets(6).Name=   "BrowseAudioFile"
      stylesets(6).Picture=   "wordlist.frx":034C
      stylesets(6).AlignmentPicture=   0
      AllowUpdate     =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      BalloonHelp     =   0   'False
      ForeColorEven   =   -2147483640
      ForeColorOdd    =   -2147483640
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      Columns.Count   =   7
      Columns(0).Width=   3200
      Columns(0).Caption=   "Phonetic"
      Columns(0).Name =   "Phonetic"
      Columns(0).CaptionAlignment=   0
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   8
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(0).HasForeColor=   -1  'True
      Columns(0).HasBackColor=   -1  'True
      Columns(0).ForeColor=   -2147483640
      Columns(0).BackColor=   -2147483640
      Columns(0).StyleSet=   "IPA"
      Columns(1).Width=   3200
      Columns(1).Caption=   "Orthographic"
      Columns(1).Name =   "Ortho"
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).Locked=   -1  'True
      Columns(1).HasForeColor=   -1  'True
      Columns(1).ForeColor=   -2147483640
      Columns(1).StyleSet=   "Ortho"
      Columns(2).Width=   3200
      Columns(2).Caption=   "Gloss"
      Columns(2).Name =   "Gloss"
      Columns(2).CaptionAlignment=   0
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).Locked=   -1  'True
      Columns(2).StyleSet=   "Gloss"
      Columns(3).Width=   3201
      Columns(3).Caption=   "Dialect"
      Columns(3).Name =   "Dialect"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).StyleSet=   "Dialect"
      Columns(4).Width=   3201
      Columns(4).Caption=   "WavFile"
      Columns(4).Name =   "WavFile"
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3200
      Columns(5).Caption=   "GraphicFile"
      Columns(5).Name =   "GraphicFile"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(5).StyleSet=   "BrowseAudioFile"
      Columns(6).Width=   609
      Columns(6).Name =   "EditWavButton"
      Columns(6).AllowSizing=   0   'False
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).Style=   4
      _ExtentX        =   10530
      _ExtentY        =   2963
      _StockProps     =   79
      Caption         =   "Caption"
      BackColor       =   -2147483636
   End
End
Attribute VB_Name = "frmWordList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
'* frmWordList version info:
'*  See basGlobals (Globals.bas).
'**************************************************

Option Explicit

Private bGridDoubleClicked As Boolean
Private bImportingCategory As Boolean
Private bInEditMode As Boolean
Private iMaxRowHeight As Integer
Private iMinPhonColWidth As Integer
Private iMaxPhonColWidth As Integer
Private iTimesPlayed As Integer
Private sWavFile As String
Private sWordListIndex As String    '* Index of Word List.
Private sWordListCaption As String  '* Name of Word List on main menu
Private sWordListPath As String     '* Path to Word List
Private sWordListSndPath As String  '* Path to Wave Files
Private WLPlayer As clsWordListPlayer
Private xmlWL As clsXMLWordList
Private LastTreeNode As MSComctlLib.node

Private bSplitterDrag As Boolean
Private sngMouseXOffset As Single

Private Const PitchPlotShownEntry = "PitchPlotHasBeenShown"
Private Const TempWLXMLBackup = "~wlXMLBackup.tmp"
Private Const NewCategoryName = "New Category"
Private Const EditModeTBarButtons = "Exit;"
Private Const NoWaveEnalbedTBarButtons = "Test;Exit;"
Private Const TBarButtons = "PlayOnly;PlaySlow;PlaySeparator;Record;StopRec;" & _
                            "PlayRec;PlayRecSpeaker;RecSeparator;" & _
                            "Pitch;PitchSeparator;" & _
                            "Test;TestSeparator;Exit;"

Public Sub AddNewCategory(Index%)

  Dim sNewCategoryName As String
  
  On Error Resume Next
    
  With tvSections
    .SetFocus
    
    sNewCategoryName = GetNewCategoryName()
    
    If (.Nodes.Count = 0) Then
      .Nodes.Add , , , sNewCategoryName
      Set .SelectedItem = .Nodes(1)
    ElseIf (Index = 0) Then
      .Nodes.Add .SelectedItem.Index, tvwPrevious, , sNewCategoryName
      Set .SelectedItem = .SelectedItem.Previous
    Else
      .Nodes.Add .SelectedItem.Index, tvwNext, , sNewCategoryName
      Set .SelectedItem = .SelectedItem.Next
    End If
  
    '*********************************************************
    '* If we just added the tree node to the end of the tree
    '* then add the new word list at the end of the XML file.
    '* Otherwise, specify what list to insert the new list
    '* before.
    '*********************************************************
    If (.SelectedItem.Next Is Nothing) Then
      xmlWL.AddCategory sNewCategoryName
    Else
      xmlWL.AddCategory sNewCategoryName, , .SelectedItem.Next.Text
    End If
    
    '*********************************************************
    '* Save the xml file changes and put the user in the
    '* edit mode to encourage them to give the new word list
    '* section a meaningful name.
    '*********************************************************
    xmlWL.Save
    Call tvSections_NodeClick(.SelectedItem)
    .StartLabelEdit
  End With

End Sub

Private Sub AdjustControlPlacement()

  If Not (bSplitterDrag) Then
    With picSplitter
      .Height = ScaleHeight
      picDragBar.Height = ScaleHeight
    End With
  End If
  
  With picDragBar
    tvSections.Move -30, -30, .Left + (gPixelX * 4), ScaleHeight + (gPixelY * 4)
    ssGrid.Move .Left + .Width + gPixelX, 0, ScaleWidth - (.Left + .Width + gPixelX), ScaleHeight
  End With

  With ssGrid
    .Redraw = False
    Call ResizeGrid
    .Redraw = True
    picTestMode.Move .Left, .Top + (gPixelY * 18), .Width, .Height - (gPixelY * 18)
  End With
    
  With picTestMode
    lblTestMode.Left = (.Width - lblTestMode.Width) \ 2
    txtUserTr.Width = .Width - (gPixelX * 10)
    txtCorrectTr.Width = .Width - (gPixelX * 10)
  End With

  With mdiHelpCharts
    !txtEditModeIndicator.Left = .Width - !txtEditModeIndicator.Width - (gPixelX * 10)
  End With
  
End Sub

Private Sub AdjustGridRowHeightToFitFonts(Optional vForcedHasPitch)

  '****************************************************************
  '* This routine will attempt to adjust the grid's row height to
  '* accomodate the largest font in the grid.
  '****************************************************************

  Dim i, j As Integer
  Dim iHeight As Integer
  
  On Error Resume Next

  iMaxRowHeight = 0

  With ssGrid
    For i = 0 To .Cols - 1
      If (Len(.Columns(i).StyleSet) > 0) Then
        With .StyleSets(.Columns(i).StyleSet).Font
          Font.Name = .Name
          Font.Size = .Size
          Font.Bold = .Bold
          Font.Italic = .Italic
        End With
      Else
        With .Font
          Font.Name = .Name
          Font.Size = .Size
          Font.Bold = .Bold
          Font.Italic = .Italic
        End With
      End If
      
      iHeight = TextHeight("X")
      If (iHeight > iMaxRowHeight) Then iMaxRowHeight = iHeight
    Next
  
    Dim bHasPitch As Boolean
    bHasPitch = DoesGridHavePitchInfo()
    If Not (IsMissing(vForcedHasPitch)) Then bHasPitch = vForcedHasPitch
    .RowHeight = (iMaxRowHeight * IIf(bHasPitch, 2, 1)) + (gPixelY * 3)
  End With

End Sub

Public Sub ApplyPhoneticFontStyle()

  On Error Resume Next
        
  With ssGrid
    .StyleSets("IPAAudio").Font.Name = .StyleSets("IPA").Font.Name
    .StyleSets("IPAAudio").Font.Size = .StyleSets("IPA").Font.Size
    .StyleSets("IPAAudio").Font.Bold = .StyleSets("IPA").Font.Bold
    .StyleSets("IPAAudio").Font.Italic = .StyleSets("IPA").Font.Italic
  End With

End Sub

Public Sub DeleteCategory()

  On Error Resume Next
  
  With tvSections
    If (MsgBox("Are you sure you want to delete " & .SelectedItem.Text & "?", _
        vbYesNo + vbQuestion, App.Title) = vbNo) Then Exit Sub
  
    If (xmlWL.RemoveCategory(.SelectedItem.Text)) Then
      xmlWL.Save
      .SetFocus
      .Nodes.Remove .SelectedItem.Index
      Call tvSections_NodeClick(.SelectedItem)
    End If
  End With
  
End Sub
  
Private Function DoesGridHavePitchInfo() As Boolean

  Dim i As Integer
  
  With ssGrid
    For i = 0 To .Rows - 1
      If (InStr(.Columns("Phonetic").CellText(.AddItemBookmark(i)), vbCrLf) > 0) Then
        DoesGridHavePitchInfo = True
        Exit Function
      End If
    Next
  End With
  
  DoesGridHavePitchInfo = False
  
End Function

Public Sub EditFonts()

  On Error Resume Next
  
  Load frmEditFonts
  
  With frmEditFonts
    Set .WordlistForm = Me
    .Show vbModal
    If Not (.Canceled) Then Call SaveColumnStyles
    Unload frmEditFonts
  End With
  
End Sub

Public Sub EditMode()

  If Not (bInEditMode) Then
    If Not (UserReallyWantsEditMode()) Then Exit Sub
  End If
  
  Call ManageGridColumns(True)
  bInEditMode = Not bInEditMode
    
  Dim i As Integer
  
  With mdiHelpCharts
    !txtEditModeIndicator.Visible = bInEditMode
    .mnuFile.Enabled = Not bInEditMode
    .mnuTest.Enabled = Not bInEditMode
    .mnuWindow.Enabled = Not bInEditMode
    .mnuEditFonts.Enabled = bInEditMode
    .mnuEditSoundPath.Enabled = bInEditMode
    .mnuEditTitle.Enabled = bInEditMode
    .mnuAddNewCategory.Enabled = bInEditMode
    .mnuDeleteCategory.Enabled = bInEditMode
    .mnuEditMode.Caption = "&" & IIf(bInEditMode, "Exit", "Enter") & " Edit Mode"
  End With
  
  With ssGrid
    '******************************************************
    '* First, hide/unhide and lock/unlock the appropriate
    '* columns for going from or to the edit mode.
    '******************************************************
    For i = 0 To .Columns.Count - 1
      .Columns(i).Locked = Not bInEditMode
    Next
    
    .Columns("WavFile").Visible = bInEditMode
    .Columns("WavFile").Locked = True
    .Columns("EditWavButton").Visible = bInEditMode
    .AllowUpdate = bInEditMode
    .AllowAddNew = bInEditMode
    .AllowDelete = bInEditMode
    .SelectTypeRow = IIf(bInEditMode, ssSelectionTypeSingleSelect, ssSelectionTypeMultiSelect)
    Call ManageGridColumns(False)
  End With
  
  tvSections.LabelEdit = IIf(bInEditMode, tvwAutomatic, tvwManual)
    
  If (bInEditMode) Then
    With mdiHelpCharts
      Call .ShowTBarButtons(TBarButtons & EditModeTBarButtons)
      Call .EnableTBarButtons(EditModeTBarButtons)
    End With
  Else
    Call UpdateXMLFileAfterEdits
    Call UserReallyWantsToKeepEdits
    
    With mdiHelpCharts
      Call .ShowTBarButtons(TBarButtons)
      If (Len(ssGrid.Columns("WavFile").Text) = 0) Then
        Call .EnableTBarButtons(NoWaveEnalbedTBarButtons)
      Else
        Call .EnableTBarButtons(TBarButtons)
      End If
    End With
  End If
  
End Sub

Public Sub EditSoundPath()

  '************************************************
  '* Pass gWavPath as default directory.
  '* Show frmFilePath (vbModal means a value must
  '* be returned before execution will continue).
  '* Check to see if a directory was returned.
  '************************************************
  With frmFilePath
    .sDir = xmlWL.SoundPath
    .Show vbModal
    If (Len(.sDir) > 0) Then
      sWordListSndPath = .sDir
      xmlWL.SoundPath = .sDir
      xmlWL.Save
    End If
  End With
  
  Unload frmFilePath

End Sub

Public Sub EditTitle()

  On Error Resume Next
  
  Load frmRenameWLTitle
  
  With frmRenameWLTitle
    .Title = sWordListCaption
    .Show vbModal
    If Not (.Canceled) Then
      sWordListCaption = .Title
      Caption = sWordListCaption & " Word List"
      xmlWL.ID = sWordListCaption
      xmlWL.Save
    End If
  End With
  
  Unload frmRenameWLTitle
    
End Sub

Public Function GetGridData(Column As String) As String

  On Error Resume Next
  GetGridData = ssGrid.Columns(Column).Text
  
End Function

Private Function GetNewCategoryName() As String

  '***********************************************************
  '* This routine will find a new name for a word list by
  '* tacking on numbers to a constant string.
  '***********************************************************
  
  Dim i As Integer
  Dim sName As String
  
  i = 0
  
  Do
    sName = NewCategoryName & IIf(i = 0, "", " " & i)
    i = i + 1
  Loop While (CategoryNameExists(sName, , False))

  GetNewCategoryName = sName
  
End Function

Private Function GetSelectedRowsWavFile(Optional vStartAtTop, Optional vMoveBookmark) As String

  Dim bStartAtTop As Boolean
  Dim bMoveBookmark As Boolean
  Dim i, j As Integer
  Static iStartRow As Integer
  Dim sWavFile As String
  
  On Error Resume Next
  
  GetSelectedRowsWavFile = ""
  sWavFile = ""
  bStartAtTop = False
  bMoveBookmark = False
  If Not (IsMissing(vStartAtTop)) Then bStartAtTop = vStartAtTop
  If Not (IsMissing(vMoveBookmark)) Then bMoveBookmark = vMoveBookmark
  If (bStartAtTop) Then iStartRow = 0

  With ssGrid
    If (.SelBookmarks.Count = 0) Then
      sWavFile = .Columns("WavFile").Text
      GoTo GetSelectedRowsWavFileEnd
    End If
      
    For i = iStartRow To .Rows - 1
      For j = 0 To .SelBookmarks.Count - 1
        If (CStr(.AddItemBookmark(i)) = CStr(.SelBookmarks(j))) Then
          sWavFile = .Columns("WavFile").CellText(.AddItemBookmark(i))
          If (bMoveBookmark) Then .Bookmark = .AddItemBookmark(i)
          GoTo GetSelectedRowsWavFileEnd
        End If
      Next
    Next
  End With
      
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GetSelectedRowsWavFileEnd:
  If (Len(sWavFile) > 0) Then _
    sWavFile = MakeFullPath(sWordListSndPath, sWavFile)
  
  If (Not FileExist(sWavFile)) Then
    iStartRow = 0
    Exit Function
  End If
  
  iStartRow = i + 1
  GetSelectedRowsWavFile = sWavFile

End Function

Private Function GetSelectedRowsWavFileForPitch(Optional vStartAtBottom, Optional vMoveBookmark) As String

  Dim bStartAtBottom As Boolean
  Dim bMoveBookmark As Boolean
  Dim i, j As Integer
  Static iStartRow As Integer
  Dim sWavFile As String
  
  On Error Resume Next
  
  GetSelectedRowsWavFileForPitch = ""
  sWavFile = ""
  bStartAtBottom = False
  bMoveBookmark = False
  If Not (IsMissing(vStartAtBottom)) Then bStartAtBottom = vStartAtBottom
  If Not (IsMissing(vMoveBookmark)) Then bMoveBookmark = vMoveBookmark
  If (bStartAtBottom) Then iStartRow = ssGrid.Rows - 1

  With ssGrid
    If (.SelBookmarks.Count = 0) Then
      sWavFile = .Columns("WavFile").Text
      GoTo GetSelectedRowsWavFileForPitchEnd
    End If
      
    For i = iStartRow To 0 Step -1
      For j = 0 To .SelBookmarks.Count - 1
        If (CStr(.AddItemBookmark(i)) = CStr(.SelBookmarks(j))) Then
          sWavFile = .Columns("WavFile").CellText(.AddItemBookmark(i))
          If (bMoveBookmark) Then .Bookmark = .AddItemBookmark(i)
          GoTo GetSelectedRowsWavFileForPitchEnd
        End If
      Next
    Next
  End With
      
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GetSelectedRowsWavFileForPitchEnd:
  If (Len(sWavFile) > 0) Then _
    sWavFile = MakeFullPath(sWordListSndPath, sWavFile)
  
  If (Not FileExist(sWavFile)) Then
    iStartRow = ssGrid.Rows - 1
    Exit Function
  End If
  
  iStartRow = i - 1
  GetSelectedRowsWavFileForPitch = sWavFile

End Function

Private Sub ImportCategory(sCategory$)
  
  Dim i As Integer
  Dim sListText As String
  Dim vList As Variant
  
  On Error Resume Next
  
  If (Not WLPlayer Is Nothing) Then WLPlayer.Cancel = True
  Call ssGrid_RowColChange(Null, 0)
  vList = xmlWL.WordsInCategory(sCategory)
  
  bImportingCategory = True
  
  '************************************************************
  '* Add each item in the array. An entire row is added in one
  '* call to .AddItem. Column items are separated by vbTab.
  '************************************************************
  With ssGrid
    .Redraw = False
    .RemoveAll
    
    '************************************************************
    '* Add each item in the array. An entire row is added in one
    '* call to .AddItem. Column items are separated by vbTab.
    '************************************************************
    If Not (IsNull(vList)) Then
      For i = 0 To UBound(vList, 1)
        .AddItem IIf(Len(vList(i, 0)) > 0, vList(i, 0) & vbCrLf, "") & _
           vList(i, 1) & vbTab & vList(i, 2) & vbTab & vList(i, 3) & vbTab & _
           vList(i, 4) & vbTab & LCase(vList(i, 5)) & vbTab & vList(i, 6)
      Next
    End If
    
    .Redraw = True
  End With
  
  bImportingCategory = False
  
End Sub

Public Sub Initialize(FormTag As String)

  Dim iIndex As Integer
  
  On Error Resume Next
  
  iIndex = CInt(FormTag)
  sWordListIndex = FormTag
  sWordListCaption = gWordListID(iIndex, 0)
  sWordListPath = gWordListID(iIndex, 1)
  
End Sub

Private Sub InitializeView()

  Dim sCategoryNames() As String
  Dim i As Integer
  
  On Error Resume Next
  
  '**************************************************************
  '* Load xml word list file
  '**************************************************************
  Set xmlWL = New clsXMLWordList
  xmlWL.Load sWordListPath
  
  '**************************************************************
  '* Get the path for the wave files
  '**************************************************************
  sWordListSndPath = xmlWL.SoundPath
  sWordListSndPath = sWordListSndPath & _
                     IIf(Right$(sWordListSndPath, 1) = "\", "", "\")
  
  '**************************************************************
  '* Get the list of word list names and fill the word list
  '* list box.
  '**************************************************************
  sCategoryNames = xmlWL.CategoryNames
  
  '**************************************************************
  '* Fill tree with list names and call the node click method
  '* to fill the grid with the data for the selected tree node.
  '**************************************************************
  With tvSections
    '************************************************************
    '* Clear the tree of all nodes.
    '************************************************************
    For i = 1 To .Nodes.Count
      .Nodes.Remove 1
    Next

    '************************************************************
    '* Add all the list names (i.e. sections names)
    '************************************************************
    For i = 0 To UBound(sCategoryNames)
      .Nodes.Add , , , sCategoryNames(i)
    Next i
    
    '************************************************************
    '* Select the first item in the list call the node click
    '* event to force the filling of the list's grid.
    '************************************************************
    Set .SelectedItem = .Nodes.item(1).FirstSibling
    Set LastTreeNode = Nothing
    Call tvSections_NodeClick(.SelectedItem)
    .SetFocus
  End With
  
  Call SetColumnStyles
  Call AdjustGridRowHeightToFitFonts

End Sub

Private Function KeyHandler(KeyCode%, KeyAscii%, Shift%) As Integer

  On Error Resume Next
  
  If (bInEditMode) Then
    '******************************************************
    '* If the user is in the edit mode then make sure they
    '* aren't allowed to change to a different window
    '******************************************************
    If (KeyCode = vbKeyTab And (Shift And vbCtrlMask)) Then
      KeyCode = 0
      Shift = 0
    ElseIf (KeyCode <> vbKeyReturn) Then
      Exit Function
    End If
      
    With ssGrid
      If (.Col <> 0) Then
         KeyCode = 0
         Exit Function
      End If
        
      If (InStr(.Columns(0).Text, vbCrLf) = 0) Then
        Call AdjustGridRowHeightToFitFonts(True)
        Exit Function
      End If
      
      KeyCode = 0
    End With
    
    Exit Function
  End If
  
  Select Case KeyCode
    Case vbKeyReturn: Call Play(cVocBttn)
    Case Else
  End Select

  Select Case KeyAscii
    Case vbKeyEscape: Unload Me
    Case Else
  End Select

End Function

Private Function CategoryNameExists(sName$, Optional vNode, Optional vShowMsg) As Boolean

  '**********************************************************
  '* This routine determines whether or not a name exists in
  '* the word list. If a node is supplied, that node will be
  '* exempted from the check.
  '**********************************************************
  
  Dim bShowMsg As Boolean
  Dim ExemptNode As MSComctlLib.node
  Dim node As MSComctlLib.node
  
  On Error Resume Next
  
  CategoryNameExists = False
  bShowMsg = True
  If Not (IsMissing(vShowMsg)) Then bShowMsg = vShowMsg
  
  Set ExemptNode = Nothing
  If Not (IsMissing(vNode)) Then Set ExemptNode = vNode

  With tvSections
    For Each node In .Nodes
      If Not (node Is ExemptNode) Then
        If (StrComp(sName, node.Text, vbTextCompare) = 0) Then
          If (bShowMsg) Then _
            MsgBox "'" & sName & "' already exists.", vbOKOnly + vbInformation, App.Title
          CategoryNameExists = True
          Exit Function
        End If
      End If
    Next
  End With

End Function

Private Sub ManageGridColumns(bSave As Boolean)

  Dim i As Integer
  Dim sVal As String
  Dim sEntry As String
  
  sEntry = IIf(bInEditMode, cWLEditModeColsEntry, cWLColsEntry)
  
  For i = 0 To ssGrid.Cols - 1
    If (bSave) Then
      Call WriteINIEntry(cSettingsSect, sEntry & i, ssGrid.Columns(i).Width, gINIPath)
    Else
      sVal = GetINIEntry$(cSettingsSect, sEntry & i, gINIPath)
      If (Val(sVal) > 0) Then ssGrid.Columns(i).Width = Val(sVal)
    End If
  Next
    
End Sub

Public Sub Play(iButton As Integer)

'**************************************************
'* This function allows mdiHelpCharts to play the
'* wave file corresponding to the currently
'* selected symbol, without knowing which form is
'* active.
'**************************************************
  
  Dim i As Integer
  Dim iRptCount As Integer
  Dim sSavBkMrk As String
  Dim sWavFile As String
  
  On Error Resume Next

  If (Len(gMMCtrl.Tag) > 0) Then Exit Sub
  If Not (WLPlayer Is Nothing) Then WLPlayer.Cancel = True
  
  '************************************************************************
  '* This code gets executed when the user clicks on a row in the grid.
  '************************************************************************
  If (iButton = -1) Then
    sWavFile = MakeFullPath(sWordListSndPath, ssGrid.Columns("WavFile").Text)
    If Not (FileExist(sWavFile)) Then Exit Sub
    Set WLPlayer = New clsWordListPlayer
    Call Pause(GetINIEntry$(cSettingsSect, cPlayInitDelayEntry, gINIPath))
    
    iRptCount = GetINIEntry$(cSettingsSect, cRepeatCountEntry, gINIPath)
    gMMCtrl.Wait = True
    gMMCtrl.Tag = cMCIBusy
    For i = 1 To iRptCount
      WLPlayer.Play sWavFile
      If (iRptCount <= 1 Or WLPlayer.Cancel) Then
      Exit For
      End If
      Call Pause(GetINIEntry$(cSettingsSect, cPlayRepeatDelayEntry, gINIPath))
    Next
    gMMCtrl.Tag = ""
  Else
    '**********************************************************************
    '* This code gets executed when the user clicks on the playback button.
    '**********************************************************************
    sSavBkMrk = ssGrid.Bookmark
    sWavFile = GetSelectedRowsWavFile(True, True)
    iRptCount = GetINIEntry$(cSettingsSect, cRepeatCountEntry, gINIPath)
    If (Len(sWavFile) = 0) Then Exit Sub
    Set WLPlayer = New clsWordListPlayer
    Call Pause(GetINIEntry$(cSettingsSect, cPlayInitDelayEntry, gINIPath))
  
    gMMCtrl.Tag = cMCIBusy
    gMMCtrl.Wait = True

'    Do
'      WLPlayer.Play sWavFile
'      sWavFile = GetSelectedRowsWavFile(, True)
'      If (Len(sWavFile) = 0 Or WLPlayer.Cancel Or ssGrid.SelBookmarks.Count = 0) Then Exit Do
'      Call Pause(GetINIEntry$(cSettingsSect, cPlayRepeatDelayEntry, gINIPath))
'    Loop
    For i = 1 To iRptCount
      WLPlayer.Play sWavFile
      sWavFile = GetSelectedRowsWavFile(, True)
      If (Len(sWavFile) = 0 Or WLPlayer.Cancel Or iRptCount <= 1) Then Exit For   ' Or ssGrid.SelBookmarks.Count = 0  used to be in this as well.
      Call Pause(GetINIEntry$(cSettingsSect, cPlayRepeatDelayEntry, gINIPath))
    Next
    
    gMMCtrl.Tag = ""
    ssGrid.Bookmark = sSavBkMrk
  End If
  
  Set WLPlayer = Nothing

End Sub

Public Sub PlaySlow()

  Dim i As Integer
  Dim sSavBkMrk As String
  Dim sWavFile As String
  
  On Error Resume Next
   
  Kill gListFilePath
  
  With ssGrid
    If (Len(.Columns("WavFile").Text) = 0) Then Exit Sub
    .SelBookmarks.RemoveAll
    sWavFile = MakeFullPath(sWordListSndPath, .Columns("WavFile").Text)
    If Not (FileExist(sWavFile)) Then Exit Sub
  End With
  
  Call WriteINIEntry("Settings", "CallingApp", mdiHelpCharts.Caption, gListFilePath)
  Call WriteINIEntry("Settings", "ShowWindow", "Hide", gListFilePath)
  Call WriteINIEntry("Settings", "CallingApp", mdiHelpCharts.Caption, gListFilePath)
  Call WriteINIEntry("AudioFiles", "File0", sWavFile, gListFilePath)
  Call WriteINIEntry("Commands", "command0", "SelectFile(0)", gListFilePath)
  Call WriteINIEntry("Commands", "command1", "Play(" & gSRSpeed & ",50,,)", gListFilePath)
  Call WriteINIEntry("Commands", "command2", "Return(2)", gListFilePath)
  Call mdiHelpCharts.CallSA
  ActiveControl.SetFocus
    
End Sub
  
Private Sub ResizeGrid(Optional vColIndex, Optional vColWidth)

  Dim i As Integer
  Dim iTextWidth As Integer
  Dim iTextHeight As Integer
  Dim iColIndex As Integer
  Dim iRowIndex As Integer
  Dim iGridWidth As Integer
  Dim iPhonColWidth As Integer
  Dim iMaxEticWidth As Integer
  Dim iVisCols
  Dim sCellText As String
  Dim sEticText As String
  Dim sPitchText As String
  Dim iColStart As Integer
  Dim iColEnd As Integer
  Dim iRowHeight As Integer
  Dim iTest As Integer
  
  On Error Resume Next
  
  If (Val(GetINIEntry$(cSettingsSect, cWLColsEntry & "0", gINIPath)) <> 0) Then _
    Exit Sub
  
  With ssGrid
    If Not IsMissing(vColIndex) Then _
      If .Columns(vColIndex).Caption <> "Phonetic" Then Exit Sub
    
    '* column width
    If IsMissing(vColIndex) Then
      '************************************************************
      '* find width from record selector to scrollbar
      '************************************************************
      iGridWidth = .Width - (gPixelX * 40)
      
      '************************************************************
      '* find number of visible columns
      '************************************************************
      iVisCols = 0
      For iColIndex = 0 To .Columns.Count - 1
        If .Columns(iColIndex).Visible Then iVisCols = iVisCols + 1
      Next
      
      '************************************************************
      '* find longest phonetic text
      '************************************************************
      iColIndex = .Columns("Phonetic").Position
      Call SetFormFontToColFont(.Columns("Phonetic").Position)
      
      '************************************************************
      '* set default column widths
      '************************************************************
      iPhonColWidth = iGridWidth / iVisCols
      For iRowIndex = 0 To .Rows - 1
        sPitchText = ""
        sEticText = .Columns("Phonetic").CellText(.AddItemBookmark(iRowIndex))
        i = InStr(sEticText, vbCrLf)
        
        If (i > 0) Then
          sPitchText = Left$(sEticText, i - 1)
          sEticText = Mid$(sEticText, i + Len(vbCrLf))
          iTextWidth = TextWidth(sPitchText)
          If (TextWidth(sEticText) > iTextWidth) Then iTextWidth = TextWidth(sEticText) + (gPixelX * 6)
        Else
          iTextWidth = TextWidth(sEticText) + (gPixelX * 6)
        End If
        
        If (iPhonColWidth < iTextWidth) Then iPhonColWidth = iTextWidth
      Next
      
      '************************************************************
      '* set column widths
      '************************************************************
      iMaxEticWidth = iPhonColWidth
      Select Case iPhonColWidth
        Case Is < iMinPhonColWidth: iPhonColWidth = iMinPhonColWidth
        Case Is > iMaxPhonColWidth: iPhonColWidth = iMaxPhonColWidth
        Case Else
      End Select
      
      .Columns("Phonetic").Width = iPhonColWidth
      
      If (iGridWidth - iPhonColWidth) > iMinPhonColWidth Then
        .Columns("Gloss").Width = (iGridWidth - iPhonColWidth) / 2
      Else
        .Columns("Gloss").Width = iMinPhonColWidth / 2
      End If
      
      .Columns("Dialect").Width = .Columns("Gloss").Width
    End If
  
    Call AdjustGridRowHeightToFitFonts
    
    '**************************************************************
    '* row height
    '**************************************************************
    'iRowHeight = iMaxRowHeight * IIf(DoesGridHavePitchInfo(), 2, 1)
    'If (iMaxEticWidth > iPhonColWidth) Then iRowHeight = iRowHeight + (iMaxRowHeight \ 2)
    '.RowHeight = iRowHeight + (gPixelY * 3)
  
  End With
  
End Sub

Private Sub SaveColumnStyles()

  On Error Resume Next
  
  With ssGrid
    xmlWL.PhoneticFontName = .StyleSets("IPA").Font.Name
    xmlWL.PhoneticFontSize = .StyleSets("IPA").Font.Size
    xmlWL.PhoneticFontBold = .StyleSets("IPA").Font.Bold
    xmlWL.PhoneticFontItalic = .StyleSets("IPA").Font.Italic
    
    xmlWL.OrthoFontName = .StyleSets("Ortho").Font.Name
    xmlWL.OrthoFontSize = .StyleSets("Ortho").Font.Size
    xmlWL.OrthoFontBold = .StyleSets("Ortho").Font.Bold
    xmlWL.OrthoFontItalic = .StyleSets("Ortho").Font.Italic
    
    xmlWL.GlossFontName = .StyleSets("Gloss").Font.Name
    xmlWL.GlossFontSize = .StyleSets("Gloss").Font.Size
    xmlWL.GlossFontBold = .StyleSets("Gloss").Font.Bold
    xmlWL.GlossFontItalic = .StyleSets("Gloss").Font.Italic
    
    xmlWL.DialectFontName = .StyleSets("Dialect").Font.Name
    xmlWL.DialectFontSize = .StyleSets("Dialect").Font.Size
    xmlWL.DialectFontBold = .StyleSets("Dialect").Font.Bold
    xmlWL.DialectFontItalic = .StyleSets("Dialect").Font.Italic
  End With
  
  xmlWL.Save
  
End Sub

Private Sub SetColumnStyles()

  On Error Resume Next
  
  With ssGrid
    If (Len(xmlWL.PhoneticFontName) > 0) Then
      .StyleSets("IPA").Font.Name = xmlWL.PhoneticFontName
      .StyleSets("IPA").Font.Size = xmlWL.PhoneticFontSize
      .StyleSets("IPA").Font.Bold = xmlWL.PhoneticFontBold
      .StyleSets("IPA").Font.Italic = xmlWL.PhoneticFontItalic
    
      .StyleSets("IPAAudio").Font.Name = xmlWL.PhoneticFontName
      .StyleSets("IPAAudio").Font.Size = xmlWL.PhoneticFontSize
      .StyleSets("IPAAudio").Font.Bold = xmlWL.PhoneticFontBold
      .StyleSets("IPAAudio").Font.Italic = xmlWL.PhoneticFontItalic
    End If
    
    If (Len(xmlWL.OrthoFontName) > 0) Then
      .StyleSets("Ortho").Font.Name = xmlWL.OrthoFontName
      .StyleSets("Ortho").Font.Size = xmlWL.OrthoFontSize
      .StyleSets("Ortho").Font.Bold = xmlWL.OrthoFontBold
      .StyleSets("Ortho").Font.Italic = xmlWL.OrthoFontItalic
    End If
    
    If (Len(xmlWL.GlossFontName) > 0) Then
      .StyleSets("Gloss").Font.Name = xmlWL.GlossFontName
      .StyleSets("Gloss").Font.Size = xmlWL.GlossFontSize
      .StyleSets("Gloss").Font.Bold = xmlWL.GlossFontBold
      .StyleSets("Gloss").Font.Italic = xmlWL.GlossFontItalic
    End If
    
    If (Len(xmlWL.DialectFontName) > 0) Then
      .StyleSets("Dialect").Font.Name = xmlWL.DialectFontName
      .StyleSets("Dialect").Font.Size = xmlWL.DialectFontSize
      .StyleSets("Dialect").Font.Bold = xmlWL.DialectFontBold
      .StyleSets("Dialect").Font.Italic = xmlWL.DialectFontItalic
    End If
  End With
  
End Sub

Private Sub SetFormFontToColFont(ColIndex As Integer)

  On Error Resume Next
  
  With ssGrid
    If (Len(.Columns(ColIndex).StyleSet) > 0) Then
      With .StyleSets(.Columns(ColIndex).StyleSet).Font
        Font.Name = .Name
        Font.Size = .Size
        Font.Bold = .Bold
        Font.Italic = .Italic
      End With
    Else
      With .Font
        Font.Name = .Name
        Font.Size = .Size
        Font.Bold = .Bold
        Font.Italic = .Italic
      End With
    End If
  End With

End Sub

Public Sub ShowPitchPlot()

  Dim iCount As Integer
  Dim sPlotWindowCoords As String
  Dim sWaveFile As String
  
  On Error Resume Next
  
  iCount = 0
  Kill gListFilePath
  sWaveFile = GetSelectedRowsWavFile(True)
  If (Len(sWaveFile) = 0) Then Exit Sub
  
  With mdiHelpCharts
    sPlotWindowCoords = (.Left \ gPixelX) + ((.Width * 0.66) \ gPixelX) & "," & _
                        (.Top \ gPixelY) + ((.Height * 0.3) \ gPixelY) & "," & _
                        (.Width * 0.66) \ gPixelX & "," & _
                        (.Height * 0.5) \ gPixelY
  End With
  
  If (Len(GetINIEntry(cSettingsSect, PitchPlotShownEntry, gINIPath)) = 0) Then
    Call WriteINIEntry(cSettingsSect, PitchPlotShownEntry, "1", gINIPath)
    Call WriteINIEntry("Settings", "ShowWindow", "Size(" & sPlotWindowCoords & ")", gListFilePath)
  End If
  
  Call WriteINIEntry("Settings", "CallingApp", mdiHelpCharts.Caption, gListFilePath)
  Call WriteINIEntry("Commands", "Command0", "DisplayPlot(Pitch)", gListFilePath)
  Call WriteINIEntry("Commands", "Command1", "Return(2)", gListFilePath)
  
  Do
    Call WriteINIEntry("AudioFiles", "File" & iCount, sWaveFile, gListFilePath)
    iCount = iCount + 1
    sWaveFile = GetSelectedRowsWavFile()
  Loop Until (Len(sWaveFile) = 0 Or ssGrid.SelBookmarks.Count = 0)
      
  If (gRecordingExists) Then _
    Call WriteINIEntry("AudioFiles", "File" & iCount, gTmpWavPath & gTmpWavName, gListFilePath)
      
  Call mdiHelpCharts.CallSA

End Sub

Public Property Get SoundFile() As String

  On Error Resume Next
  SoundFile = MakeFullPath(sWordListSndPath, ssGrid.Columns("WavFile").Text)
  
End Property

Private Sub StopPlayback()

  On Error Resume Next
  Timer1.Enabled = False
  
End Sub

Public Sub TestPlay()
  
  Dim i As Integer
  Dim iPrevSection As Integer
  Dim iPrevItem As Integer
  Dim lRetryCount As Long
  
  With tvSections
    If (gTestCatChoice = cUser And ValidWavFilesInGrid() = 0) Then
      Call MsgBox("No valid sound files found for this category.", vbOKOnly + vbInformation, App.Title)
      Exit Sub
    End If
    
    iPrevItem = gItemNumber
    If (.SelectedItem Is Nothing) Then
      iPrevSection = 0
    Else
      iPrevSection = .SelectedItem.Index
    End If
      
    '****************************************************************************
    '* If the user is testing random sections, then randomly choose a section.
    '****************************************************************************
    If (gTestCatChoice = cRandom) Then
      Do
        lRetryCount = 0&
        Do
          Randomize
          i = Int(.Nodes.Count * Rnd + 1)
          lRetryCount = lRetryCount + 1&
        Loop While ((i < 1 Or i = iPrevSection) And lRetryCount < 32000)

        Set .SelectedItem = .Nodes(i)
        Call tvSections_NodeClick(.SelectedItem)
        
        '************************************************************************
        '* If the random section chosen has one or more valid wave files in
        '* its word list then exit this loop because we've found what we want.
        '************************************************************************
        If (ValidWavFilesInGrid() > 0) Then
          iPrevSection = i
          Exit Do
        End If
      Loop While True
    End If
  End With
  
  '******************************************************************************
  '* Now randomly select a word within the section's word list. It should be
  '* between the lowest and highest
  '******************************************************************************
  lRetryCount = 0&
  
  With ssGrid
    Do
      Randomize
      i = Int(.Rows * Rnd)
      .Bookmark = .RowBookmark(i)
      gItemNumber = i
      lRetryCount = lRetryCount + 1&
    Loop While ((Not FileExist(MakeFullPath(sWordListSndPath, .Columns("WavFile").Text)) Or _
                i = iPrevItem) And lRetryCount < 32000)
  End With
  
  Call Play(cVocBttn)
  
End Sub

Public Sub UpdateAfterRecordAndPlayback()

  On Error Resume Next
  Call mdiHelpCharts.EnableTBarButtons(TBarButtons)

End Sub

Public Sub UpdateFormAfterTest()

  On Error Resume Next
  
  With tvSections
    .SetFocus
    Set .SelectedItem = .Nodes.item(1).FirstSibling
    Call tvSections_NodeClick(.SelectedItem)
  End With
  
  picTestMode.Visible = False
  Call mdiHelpCharts.EnableTBarButtons(TBarButtons)
  
End Sub

Public Sub UpdateFormForTest()
  
  Dim iBttnTop As Integer
  
  On Error Resume Next
  
  With picTestMode
    '.Move 0, 0 + IIf(gTestCatChoice = cRandom, 0, 270), picGrid.Width, picGrid.Height
    .Visible = True
    .ZOrder vbBringToFront
  End With
  
  With txtUserTr
    With .Font
      .Name = ssGrid.StyleSets("IPA").Font.Name
      .Size = ssGrid.StyleSets("IPA").Font.Size
      .Bold = ssGrid.StyleSets("IPA").Font.Bold
      .Italic = ssGrid.StyleSets("IPA").Font.Italic
    End With
    .Top = .Top                                 'This forces text box to resize after font change above
  End With
  
  With txtCorrectTr
    With .Font
      .Name = ssGrid.StyleSets("IPA").Font.Name
      .Size = ssGrid.StyleSets("IPA").Font.Size
      .Bold = ssGrid.StyleSets("IPA").Font.Bold
      .Italic = ssGrid.StyleSets("IPA").Font.Italic
    End With
  
    .Top = .Top                                 'This forces text box to resize after font change above
    iBttnTop = .Top + .Height + 135
  End With
  
  cmdReplay.Top = iBttnTop
  cmdVerify.Top = iBttnTop
  cmdNext.Top = iBttnTop
  
  Select Case gTestCatChoice
    Case cUser:   gStatLine.SimpleText = "Select category to start test"
    Case cRandom: gStatLine.SimpleText = "Enter transcription when item is pronounced"
  End Select
  
  picTestMode.Refresh
  
End Sub

Private Sub UpdateXMLFileAfterEdits()

  Dim i As Integer
  
  With ssGrid
    .Update
      
    '******************************************************
    '* Save the changed words to the XML file. Do this by
    '* emptying the word list then readding the words.
    '******************************************************
    If Not (xmlWL.EmptyCategory(.Caption)) Then Exit Sub
  
    ReDim sList(0 To .Rows - 1, 0 To 6) As String
  
    For i = 0 To .Rows - 1
      .Bookmark = .RowBookmark(i)
      sList(i, 1) = .Columns(0).Text
      sList(i, 2) = .Columns(1).Text
      sList(i, 3) = .Columns(2).Text
      sList(i, 4) = .Columns(3).Text
      sList(i, 5) = .Columns(4).Text
      sList(i, 6) = .Columns(5).Text
    Next
  
    xmlWL.AddCategory .Caption, sList
    xmlWL.Save
  End With
  
End Sub

Private Function UserReallyWantsEditMode() As Boolean

  '***********************************************************
  '* This routine asks the user whether or not he really
  '* wants to go into the edit mode.
  '***********************************************************

  UserReallyWantsEditMode = False

  If (MsgBox("You are about to enter the edit mode in which you may" & vbCrLf & _
             "make changes to word lists. Are you sure want to proceed?", _
             vbQuestion + vbYesNo, App.Title) = vbNo) Then Exit Function
             
  UserReallyWantsEditMode = True
  
  '***********************************************************
  '* Keep a backup copy of the XML file in case the user
  '* ends up not wanting to keep his changes.
  '***********************************************************
  FileCopy sWordListPath, gINIPath & TempWLXMLBackup
  
End Function

Private Function UserReallyWantsToKeepEdits() As Boolean

  '***********************************************************
  '* This routine asks the user whether or not he really
  '* wants to keep the edit mode changes.
  '***********************************************************

  UserReallyWantsToKeepEdits = True

  '***********************************************************
  '* If the user doesn't want to keep the changes then
  '* restore the backup file.
  '***********************************************************
  If (MsgBox("Do you want to keep your changes?", vbQuestion + vbYesNo, App.Title) = vbNo) Then
    UserReallyWantsToKeepEdits = False
    FileCopy gINIPath & TempWLXMLBackup, sWordListPath
    Call InitializeView
  End If
  
  If (FileExist(sWordListPath)) Then Kill gINIPath & TempWLXMLBackup
  
End Function

Private Function ValidWavFilesInGrid() As Integer

  '*********************************************************
  '* This routine will count how many valid wave files are
  '* in the WavFile column of the grid.
  '*********************************************************

  On Error Resume Next
  
  Dim i As Integer
  Dim iCount As Integer
  
  iCount = 0
  
  With ssGrid
    For i = 0 To .Rows - 1
      If (FileExist(MakeFullPath(sWordListSndPath, .Columns("WavFile").CellText(.RowBookmark(i))))) Then _
        iCount = iCount + 1
    Next
  End With
      
  ValidWavFilesInGrid = iCount
  
End Function

Private Sub cmdNext_Click()

  txtUserTr.Text = ""
  txtCorrectTr.Text = ""
  cmdReplay.Enabled = True
  cmdVerify.Enabled = True
  Call TestPlay
  
End Sub

Private Sub cmdReplay_Click()

  On Error Resume Next
  Call Play(cVocBttn)
  
End Sub

Private Sub cmdVerify_Click()

  On Error Resume Next
  Call StopPlayback
  txtCorrectTr.Text = ssGrid.Columns("Phonetic").Text
  cmdVerify.Enabled = False
  cmdReplay.Enabled = True
  
End Sub

Private Sub Form_Activate()

  Dim i As Integer
  
  On Error Resume Next
  
  If (WindowState <> vbMaximized) Then WindowState = vbMaximized
  Call UpdateTestMenu                          '* Enable test menu.
  
  gStatLine.SimpleText = "Click to select"
  gRecordingExists = False
  tvSections.SetFocus
  
  With mdiHelpCharts
    If Not (!panStatus.Visible) Then !panStatus.Visible = True
    Call .ShowTBarButtons(TBarButtons)
    Call .EnableTBarButtons(TBarButtons)
    !mnuEdit.Visible = gAppPathWriteAccess
    
    For i = 0 To 2
      !mnuPlayback(i).Visible = True
    Next
  End With
  
End Sub

Private Sub Form_Deactivate()

  Dim i As Integer
  
  On Error Resume Next
  
  With mdiHelpCharts
    !mnuEdit.Visible = False
    For i = 0 To 2
      !mnuPlayback(i).Visible = False
    Next
  End With
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  On Error Resume Next
  Call KeyHandler(KeyCode, -1, Shift)
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  
  On Error Resume Next
  Call KeyHandler(-1, KeyAscii, -1)
  
End Sub

Private Sub Form_Load()

  Dim sINIVal As String
  
  On Error Resume Next
  
  With mdiHelpCharts
    If Not (!panStatus.Visible) Then !panStatus.Visible = True
    If Not (!TBar.Visible) Then !TBar.Visible = True
    Call UpdateTestMenu                          '* Enable test menu.
  End With
  
  Call ManageGridColumns(False)
  Set WLPlayer = Nothing
  gStatLine.SimpleText = ""
  Caption = sWordListCaption & " Word List"
  Tag = sWordListIndex
  
  With ssGrid
    .Redraw = False
    .RemoveAll
    .Columns("Ortho").Visible = False
    .Columns("WavFile").Visible = False
    .Columns("GraphicFile").Visible = False
    .Columns("EditWavButton").Visible = False
    .Height = mdiHelpCharts.ScaleHeight
    .Width = mdiHelpCharts.ScaleWidth - .Left
    iMinPhonColWidth = 0.33 * (.Width - gPixelX * 40)
    iMaxPhonColWidth = 0.66 * (.Width - gPixelX * 40)
    .Redraw = True
  End With
  
  picTestMode.BorderStyle = 0
  
  With picDragBar
    .BorderStyle = 0
    .BackColor = vbButtonShadow
  End With
  
  With picSplitter
    .BorderStyle = 0
    .BackColor = vbButtonFace
    
    sINIVal = GetINIEntry("WordList", "SplitterLeft", gINIPath)
    .Move IIf(Val(sINIVal) = 0, (gPixelX * 200), Val(sINIVal)), 0, gPixelX * 8, ScaleHeight
    picDragBar.Move .Left, .Top, .Width, .Height
    picDragBar.ZOrder vbBringToFront
    picDragBar.Visible = False
  End With

  Call InitializeView
  
  With tvSections
    .Height = mdiHelpCharts.ScaleHeight
    .ZOrder vbSendToBack
  End With
  
  Show
  WindowState = vbMaximized
  bGridDoubleClicked = False
  bInEditMode = False
  
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  On Error Resume Next
  gStatLine.SimpleText = ""

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  On Error Resume Next
  Call WriteINIEntry("WordList", "SplitterLeft", picSplitter.Left, gINIPath)
  Call Form_Deactivate
  If (bInEditMode) Then Call EditMode
  Call ManageGridColumns(True)
  Set xmlWL = Nothing
  Set frmWordList = Nothing
  
End Sub

Private Sub Form_Resize()

  On Error Resume Next
  Call AdjustControlPlacement

End Sub

Private Sub ssGrid_BtnClick()

  Load frmGetSoundFile
  
  With frmGetSoundFile
    If (Len(ssGrid.Columns("WavFile").Text) > 0) Then _
      .FileName = ssGrid.Columns("WavFile").Text
    
    .Folder = sWordListSndPath
    .Show vbModal
    
    If Not (.Canceled) Then _
      ssGrid.Columns("WavFile").Text = .FileName
  End With
  
  Unload frmGetSoundFile
  
End Sub

Private Sub ssGrid_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
  
  On Error Resume Next
  Call ResizeGrid(ColIndex, ssGrid.ResizeWidth)
  
End Sub

Private Sub ssGrid_KeyDown(KeyCode As Integer, Shift As Integer)

  On Error Resume Next
  Call KeyHandler(KeyCode, -1, Shift)

End Sub

Private Sub ssGrid_KeyPress(KeyAscii As Integer)

  On Error Resume Next
  Call KeyHandler(-1, KeyAscii, -1)

End Sub

Private Sub ssGrid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  On Error Resume Next
  If (bInEditMode) Then Exit Sub
  If (Shift = 0 And X > 200 And Y > 255) Then ssGrid.SelBookmarks.RemoveAll

End Sub

Private Sub ssGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  '***************************************************************
  '* This will turn on the balloon help for the wave file column
  '* since a wave file's path length will likely exceed the width
  '* of the column in which it's displayed.
  '***************************************************************
  
  Dim iCol As Integer
  
  On Error Resume Next
  
  If Not (bInEditMode) Then Exit Sub
  
  With ssGrid
    iCol = .ColContaining(X)
    If (iCol >= 0) Then
      If (.Columns(iCol).Name = "WavFile") Then
        If Not (.BalloonHelp) Then .BalloonHelp = True
        Exit Sub
      End If
    End If
      
    .BalloonHelp = False
  End With
  
End Sub

Private Sub ssGrid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  On Error Resume Next
  
  If (bInEditMode) Then Exit Sub
  
  With ssGrid
    If (.ColContaining(X) = -1) Then Exit Sub
    Dim iRow As Integer
    iRow = .RowContaining(Y)
    If (iRow >= 0 And iRow <= .Rows And X > 200 And Not .AllowUpdate) Then Call Play(-1)
  End With

End Sub

Private Sub ssGrid_RowColChange(ByVal LastRow As Variant, ByVal LastCol As Integer)

  On Error Resume Next

  If (bImportingCategory Or picTestMode.Visible Or bInEditMode) Or _
     (CStr(LastRow) = CStr(ssGrid.Bookmark)) Then Exit Sub
  
  gRecordingExists = False
  Kill gTmpWavPath & gTmpWavName

  With ssGrid.Columns("WavFile")
    If (Not IsNull(LastRow) And Not FileExist(MakeFullPath(sWordListSndPath, .Text))) Then
      Call mdiHelpCharts.EnableTBarButtons(NoWaveEnalbedTBarButtons)
    Else
      Call mdiHelpCharts.EnableTBarButtons(TBarButtons)
    End If
  End With
  
End Sub

Private Sub ssGrid_RowLoaded(ByVal Bookmark As Variant)

  On Error Resume Next
  With ssGrid
    If (FileExist(MakeFullPath(sWordListSndPath, .Columns("WavFile").Text))) Then
      .Columns("Phonetic").CellStyleSet "IPAAudio", Bookmark
    End If
  End With
  
End Sub

Private Sub ssGrid_SelChange(ByVal SelType As Integer, Cancel As Integer, DispSelRowOverflow As Integer)

  On Error Resume Next
  If (bInEditMode) Then Exit Sub
  If Not (FileExist(MakeFullPath(sWordListSndPath, ssGrid.Columns("WavFile").Text))) Then Cancel = 1

End Sub

Private Sub Timer1_Timer()

  On Error Resume Next
  Call PlayWav(sWavFile)
  iTimesPlayed = iTimesPlayed + 1
  
  With Timer1
    If iTimesPlayed < gWordListRepeat Then
      .Interval = 1500
    Else
      Timer1.Enabled = False
    End If
  End With
  
End Sub

Private Sub picDragBar_Paint()

  With picDragBar
    picDragBar.Line (gPixelX, 0)-(gPixelX, .ScaleHeight), vb3DHighlight
    picDragBar.Line (.ScaleWidth - gPixelX, 0)-(.ScaleWidth - gPixelX, .ScaleHeight), vb3DDKShadow
  End With

End Sub

Private Sub picSplitter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  On Error Resume Next
  
  If (Button <> vbLeftButton) Then Exit Sub
  sngMouseXOffset = X
  bSplitterDrag = True
  picDragBar.Visible = True
  picDragBar.ZOrder vbBringToFront
  picSplitter.Visible = False
  
End Sub

Private Sub picSplitter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  If Not (bSplitterDrag) Then Exit Sub
  picDragBar.Left = picSplitter.Left + X - sngMouseXOffset
  
End Sub

Private Sub picSplitter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  On Error Resume Next
  
  If (bSplitterDrag) Then
    bSplitterDrag = False
  
    With picDragBar
      .Visible = False
      picSplitter.Move .Left, .Top, .Width, .Height
      picSplitter.Visible = True
    End With
    
    Call AdjustControlPlacement
  End If

End Sub

Private Sub picSplitter_Paint()
  
  On Error Resume Next
  
  With picSplitter
    picSplitter.Line (0, 0)-(0, .ScaleHeight), vb3DHighlight
    picSplitter.Line (.ScaleWidth - gPixelX, 0)-(.ScaleWidth - gPixelX, .ScaleHeight), vbButtonShadow
  End With
  
End Sub

Private Sub tvSections_AfterLabelEdit(Cancel As Integer, NewString As String)

  Dim sOldCategoryName As String
  
  With tvSections
    sOldCategoryName = .SelectedItem.Text
    If (CategoryNameExists(NewString, .SelectedItem, True)) Then
      Cancel = True
      .SetFocus
      Exit Sub
    End If
  End With
    
  With xmlWL
    If (.ChangeCategoryName(sOldCategoryName, NewString)) Then .Save
  End With
    
End Sub

Private Sub tvSections_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If (KeyCode = vbKeyF2 And bInEditMode And Shift = 0) Then _
    tvSections.StartLabelEdit

End Sub

Private Sub tvSections_NodeClick(ByVal node As MSComctlLib.node)

  On Error Resume Next
  
  Call StopPlayback
  If (node Is Nothing) Then Exit Sub
  
  If Not (LastTreeNode Is node) Then
    If (bInEditMode) Then Call UpdateXMLFileAfterEdits
    Call ImportCategory(node.Text)
    ssGrid.Caption = node.Text
    Call ResizeGrid
  End If
  
  Set LastTreeNode = node
  If (gTestActive And gTestCatChoice = cUser) Then Call cmdNext_Click

End Sub
