VERSION 5.00
Object = "{BC496AED-9B4E-11CE-A6D5-0000C0BE9395}#2.0#0"; "SSDATB32.OCX"
Begin VB.Form frmEditIPAChars 
   Caption         =   "Edit"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   ControlBox      =   0   'False
   Icon            =   "EditIPAChars.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOKCancel 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   4440
      TabIndex        =   2
      Top             =   4020
      Width           =   915
   End
   Begin VB.CommandButton cmdOKCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   5430
      TabIndex        =   1
      Top             =   4020
      Width           =   915
   End
   Begin SSDataWidgets_B.SSDBGrid ssGrid 
      Align           =   1  'Align Top
      Height          =   1680
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6360
      _Version        =   131078
      DataMode        =   2
      BorderStyle     =   0
      RecordSelectors =   0   'False
      Col.Count       =   8
      stylesets.count =   3
      stylesets(0).Name=   "IPA"
      stylesets(0).ForeColor=   -2147483640
      stylesets(0).BackColor=   -2147483643
      stylesets(0).Picture=   "EditIPAChars.frx":000C
      stylesets(0).AlignmentText=   7
      stylesets(1).Name=   "Number"
      stylesets(1).ForeColor=   -2147483630
      stylesets(1).BackColor=   -2147483633
      stylesets(1).Picture=   "EditIPAChars.frx":0028
      stylesets(1).AlignmentText=   4
      stylesets(2).Name=   "BrowseAudioFile"
      stylesets(2).Picture=   "EditIPAChars.frx":0044
      stylesets(2).AlignmentPicture=   0
      MultiLine       =   0   'False
      AllowRowSizing  =   0   'False
      AllowGroupSizing=   0   'False
      AllowGroupMoving=   0   'False
      AllowColumnMoving=   0
      AllowGroupSwapping=   0   'False
      AllowColumnSwapping=   0
      AllowGroupShrinking=   0   'False
      AllowColumnShrinking=   0   'False
      AllowDragDrop   =   0   'False
      SelectTypeCol   =   0
      SelectTypeRow   =   0
      ForeColorEven   =   -2147483640
      ForeColorOdd    =   -2147483640
      BackColorEven   =   -2147483643
      BackColorOdd    =   -2147483643
      RowHeight       =   423
      ExtraHeight     =   291
      Columns.Count   =   8
      Columns(0).Width=   873
      Columns(0).Name =   "Number"
      Columns(0).DataField=   "Column 0"
      Columns(0).DataType=   2
      Columns(0).FieldLen=   256
      Columns(0).Locked=   -1  'True
      Columns(0).HasForeColor=   -1  'True
      Columns(0).ForeColor=   -2147483640
      Columns(0).StyleSet=   "Number"
      Columns(1).Width=   1614
      Columns(1).Caption=   "Character"
      Columns(1).Name =   "Character"
      Columns(1).CaptionAlignment=   0
      Columns(1).DataField=   "Column 1"
      Columns(1).DataType=   8
      Columns(1).FieldLen=   256
      Columns(1).HasForeColor=   -1  'True
      Columns(1).HasBackColor=   -1  'True
      Columns(1).ForeColor=   -2147483640
      Columns(1).BackColor=   -2147483640
      Columns(1).StyleSet=   "IPA"
      Columns(2).Width=   1588
      Columns(2).Caption=   "Example 1"
      Columns(2).Name =   "Example1"
      Columns(2).DataField=   "Column 2"
      Columns(2).DataType=   8
      Columns(2).FieldLen=   256
      Columns(2).StyleSet=   "IPA"
      Columns(3).Width=   1588
      Columns(3).Caption=   "Example 2"
      Columns(3).Name =   "Example2"
      Columns(3).DataField=   "Column 3"
      Columns(3).DataType=   8
      Columns(3).FieldLen=   256
      Columns(3).StyleSet=   "IPA"
      Columns(4).Width=   3200
      Columns(4).Caption=   "Character Name"
      Columns(4).Name =   "CharName"
      Columns(4).CaptionAlignment=   0
      Columns(4).DataField=   "Column 4"
      Columns(4).DataType=   8
      Columns(4).FieldLen=   256
      Columns(5).Width=   3201
      Columns(5).Caption=   "Character Description"
      Columns(5).Name =   "CharDesc"
      Columns(5).DataField=   "Column 5"
      Columns(5).DataType=   8
      Columns(5).FieldLen=   256
      Columns(6).Width=   3201
      Columns(6).Caption=   "WavFile"
      Columns(6).Name =   "WavFile"
      Columns(6).DataField=   "Column 6"
      Columns(6).DataType=   8
      Columns(6).FieldLen=   256
      Columns(6).Locked=   -1  'True
      Columns(7).Width=   609
      Columns(7).Name =   "EditWavButton"
      Columns(7).AllowSizing=   0   'False
      Columns(7).DataField=   "Column 7"
      Columns(7).DataType=   8
      Columns(7).FieldLen=   256
      Columns(7).Style=   4
      _ExtentX        =   11218
      _ExtentY        =   2963
      _StockProps     =   79
      BackColor       =   -2147483636
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   960
      X2              =   2700
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   1020
      X2              =   2640
      Y1              =   2220
      Y2              =   2220
   End
End
Attribute VB_Name = "frmEditIPAChars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sINISectionCharInfo As String
Private sINISectionSettings As String

Private Function GetCharFromINIStr(sINIStr$) As String

  Dim i As Integer
  Dim j As Integer
  Dim sGridRow As String
  Dim sPiece As String
  
  On Error Resume Next
  
  sGridRow = ""
  sPiece = ""
  
  i = InStr(1, sINIStr, ";")
  j = 1
  
If (i > 0) Then
  sGridRow = sGridRow & Chr$(Mid$(sINIStr, j, i - j)) & vbTab & vbTab & vbTab
  j = i + 1
  i = InStr(j, sINIStr, ";")
End If
  
  While (i > 0)
    sGridRow = sGridRow & Mid$(sINIStr, j, i - j) & vbTab
    j = i + 1
    i = InStr(j, sINIStr, ";")
  Wend
  
  If (j > 0) Then _
    GetCharFromINIStr = sGridRow & Mid$(sINIStr, j, Len(sINIStr) - j + 1)
    
End Function

Public Sub LoadCharInfo()

  Dim i As Integer
  Dim sINIVals() As String
  
  sINIVals = GetAllINISettings(gINIPath, sINISectionCharInfo)
  
  For i = 0 To UBound(sINIVals, 1)
    ssGrid.AddItem sINIVals(i, 0) & vbTab & GetCharFromINIStr(sINIVals(i, 1))
  Next

End Sub

Private Sub SetGridColSizesFromINI()

  '***********************************************************************
  '* Set each grid column's width based on previously saved values.
  '***********************************************************************
  
  Dim i As Integer
  Dim iWid As Integer
  
  With ssGrid
    For i = 0 To .Cols - 1
      iWid = Val(GetINIEntry(sINISectionSettings, "Col" & i, gINIPath))
      If (iWid > 0) Then .Columns(i).Width = iWid
    Next
  End With
  
End Sub

Private Sub SetGridRowHeight()

  Dim iMax As Integer
  
  '***********************************************************************
  '* There are two fonts used in the grid. Find out which one has the
  '* greater height and set the grid's row height to that plus 3 pixels.
  '***********************************************************************
  
  With ssGrid.StyleSets("IPA").Font
    Font.Name = .Name
    Font.Size = .Size
    Font.Bold = .Bold
    Font.Italic = .Italic
    iMax = TextHeight("X")
  End With
  
  With ssGrid.StyleSets("Number").Font
    Font.Name = .Name
    Font.Size = .Size
    Font.Bold = .Bold
    Font.Italic = .Italic
  End With

  ssGrid.RowHeight = IIf(iMax < TextHeight("X"), TextHeight("X"), iMax) + (gPixelY * 3)

End Sub

Private Sub SetIPAFontStyle()

  Dim sINIVal As String
  
  '***********************************************************************
  '* Read the IPA font info. from the ini file and set the grid's IPA
  '* style's font to it.
  '***********************************************************************
  With ssGrid.StyleSets("IPA").Font
    sINIVal = GetINIEntry(sINISectionSettings, cFontNameEntry, gINIPath)
    If (Len(sINIVal) > 0) Then .Name = sINIVal
    sINIVal = GetINIEntry(sINISectionSettings, cFontSizeEntry, gINIPath)
    If (Len(sINIVal) > 0) Then .Size = sINIVal
    sINIVal = GetINIEntry(sINISectionSettings, cFontBoldEntry, gINIPath)
    If (Len(sINIVal) > 0) Then .Bold = sINIVal
    sINIVal = GetINIEntry(sINISectionSettings, cFontItalicEntry, gINIPath)
    If (Len(sINIVal) > 0) Then .Italic = sINIVal
  End With

  Call SetGridRowHeight

End Sub

Public Sub SetINISections(sCharInfo$, sSettings$)

  Dim i As Integer
  
  sINISectionCharInfo = sCharInfo
  sINISectionSettings = sSettings
  Call SetIPAFontStyle
  Call SetGridColSizesFromINI

  '***********************************************************************
  '* Restore this form's position and size to what was previously saved
  '* in the ini file.
  '***********************************************************************
  i = Val(GetINIEntry(sINISectionSettings, cLeftEntry, gINIPath))
  If (i > 0) Then Left = i
  i = Val(GetINIEntry(sINISectionSettings, cTopEntry, gINIPath))
  If (i > 0) Then Top = i
  i = Val(GetINIEntry(sINISectionSettings, cWidthEntry, gINIPath))
  If (i > 0) Then Width = i
  i = Val(GetINIEntry(sINISectionSettings, cHeightEntry, gINIPath))
  If (i > 0) Then Height = i

End Sub

Public Sub ShowExample(iExampleNum%, bShow As Boolean)

  On Error Resume Next
  ssGrid.Columns("Example" & iExampleNum).Visible = bShow
  
End Sub

Private Sub WriteColSizesToINI()

  '***********************************************************************
  '* Save each column's width in the ini file.
  '***********************************************************************
  
  Dim i As Integer
  
  With ssGrid
    For i = 0 To .Cols - 1
      Call WriteINIEntry(sINISectionSettings, "Col" & i, .Columns(i).Width, gINIPath)
    Next
  End With
  
End Sub

Private Sub cmdOKCancel_Click(Index As Integer)

  Unload Me
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  Call WriteColSizesToINI
  
  Call WriteINIEntry(sINISectionSettings, cLeftEntry, Left, gINIPath)
  Call WriteINIEntry(sINISectionSettings, cTopEntry, Top, gINIPath)
  Call WriteINIEntry(sINISectionSettings, cWidthEntry, Width, gINIPath)
  Call WriteINIEntry(sINISectionSettings, cHeightEntry, Height, gINIPath)
  
End Sub

Private Sub Form_Resize()

  With cmdOKCancel(0)
    .Left = ScaleWidth - (.Width * 2) - (gPixelX * 7)
    .Top = ScaleHeight - .Height - (gPixelX * 3)
  End With

  With cmdOKCancel(1)
    .Left = ScaleWidth - .Width - (gPixelX * 3)
    .Top = cmdOKCancel(0).Top
  End With

  ssGrid.Height = ScaleHeight - cmdOKCancel(0).Height - (gPixelY * 9)
  
  With Line1
    .X1 = 0
    .X2 = Width
    .Y1 = ssGrid.Height
    .Y2 = .Y1
  End With
  
  With Line2
    .X1 = 0
    .X2 = Width
    .Y1 = ssGrid.Height + gPixelY
    .Y2 = .Y1
  End With
    
End Sub
