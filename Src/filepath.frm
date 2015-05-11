VERSION 5.00
Begin VB.Form frmFilePath 
   Caption         =   "Set Sound File Path"
   ClientHeight    =   3780
   ClientLeft      =   9105
   ClientTop       =   3555
   ClientWidth     =   4995
   ControlBox      =   0   'False
   Icon            =   "FILEPATH.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3780
   ScaleWidth      =   4995
   Begin VB.DirListBox dirSndDir 
      Height          =   2340
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   3540
   End
   Begin VB.TextBox txtDir 
      Height          =   300
      Left            =   60
      TabIndex        =   1
      Top             =   300
      Width           =   3540
   End
   Begin VB.DriveListBox drvSndDrive 
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   3360
      Width           =   3540
   End
   Begin VB.CommandButton cmdOKCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   355
      Index           =   1
      Left            =   3840
      TabIndex        =   6
      Top             =   780
      Width           =   1020
   End
   Begin VB.CommandButton cmdOKCancel 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   355
      Index           =   0
      Left            =   3840
      TabIndex        =   5
      Top             =   300
      Width           =   1020
   End
   Begin VB.Label Label4 
      Caption         =   "Dri&ves:"
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1785
   End
   Begin VB.Label lblFolder 
      Caption         =   "&Folder:"
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   1785
   End
End
Attribute VB_Name = "frmFilePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************
'* frmFilePath version info:
'*  See basGlobals (Globals.bas).
'**************************************************

Option Explicit

Public sDir As String                               '* Directory Passed between caller & callee.

Private Function InvalDir() As Integer

'**************************************************
'* This function checks the validity of the
'* selected directory. It returns true if the
'* directory is invalid, false if it is valid.
'**************************************************

  Dim i As Integer
  Dim tmpDir As String
  Dim sMsg As String

  tmpDir = LCase$(txtDir.Text)
  If Mid$(tmpDir, 2, 1) = ":" Then                                   'Does input have drive specification?
    For i = 0 To drvSndDrive.ListCount - 1                                'Search drive list to determine whether
      If Left$(tmpDir, 1) = Left$(drvSndDrive.List(i), 1) Then Exit For   '  or not specified drive exists.
    Next i
  
    '**********************************************
    '* If you get to the end of the drive list and
    '* don't find the specified drive ...
    '**********************************************
    If i = drvSndDrive.ListCount Then                                     'Was specified drive found?
      sMsg = "Invalid drive specification."
      GoTo BadDirSpec
    Else
      tmpDir = Mid$(tmpDir, 3)                                       'Save input without drive spec.
    End If
  End If

  '************************************************
  '* Disallow invalid characters.
  '************************************************
  If InStr(tmpDir, ":") _
    Or InStr(tmpDir, "*") _
    Or InStr(tmpDir, "?") _
    Or InStr(tmpDir, Chr$(34)) _
    Or InStr(tmpDir, "<") _
    Or InStr(tmpDir, ">") _
    Or InStr(tmpDir, "|") Then                                      'If directory name contains :, *, ?, ", <, >, or |
    sMsg = "Invalid character in " & _
             "directory specification."
    GoTo BadDirSpec
  End If

  InvalDir = False
  Exit Function

BadDirSpec:
  MsgBox sMsg, vbExclamation, App.Title
  InvalDir = True
 
End Function

Private Sub cmdOKCancel_Click(Index As Integer)

  Dim savDir As String
  Dim sMsg As String
  
  savDir = CurDir$                                            'Save current directory.
  
  '************************************************
  '* If OK pressed, check if the path is valid.
  '* Otherwise, just exit.
  '************************************************
  If cmdOKCancel(0).Value Then                                'Did user press OK button?
    sDir = Trim$(txtDir.Text)                          'Trim any leading & trailing spaces.
    
    If sDir = "" Or InvalDir() Then                    'Does input contain valid characters?
      txtDir.SetFocus
      Exit Sub
    End If
    
    sDir = txtDir.Text & IIf(Right$(txtDir.Text, 1) = "\", "", "\")            'Make sure path has trailing backslash.
    
    On Error GoTo DirNotExist                                 'Prepare not to find drive or directory.
    If Mid$(sDir, 2, 1) = ":" Then ChDrive Left$(sDir, 2)     'If input has drive in it then access it.
    ChDir sDir                                                'Try to access specified directory.
    On Error GoTo 0                                           'If we're here then input directory exists.
    
    ChDrive Left$(savDir, 2)                                  'Set drive and directory back to what
    ChDir savDir                                              '  it was before dir exists test.
    Call SaveSetting(gINIPath, cPathsSect, cSoundsEntry, sDir)
  Else                                                'OK Not pressed means Cancel was pressed
    sDir = ""                                         '  so return nothing.
  End If
  
  frmFilePath.Visible = False
  
  Exit Sub

DirNotExist:                                                  'At this point the user specified dir doesn't exist.
  On Error GoTo 0
  ChDrive Left$(savDir, 2)                                    'Set drive and directory back to what
  ChDir savDir                                                '  it was before dir exisits test.
  sMsg = "Directory does not exist. " & "Would you like to create it?"
  
  If MsgBox(sMsg, vbQuestion + vbYesNo, App.Title) = vbYes Then
    MkDir txtDir.Text                                         'User wants to make new directory.
  Else
    txtDir.SetFocus                                           'Re-highlight directory text-box
    txtDir.SelStart = 0
    txtDir.SelLength = Len(txtDir.Text)
  End If
  
  Exit Sub

End Sub

Private Sub dirSndDir_Change()

  On Error Resume Next
  txtDir.Text = dirSndDir.Path

End Sub

Private Sub drvSndDrive_Change()

  On Error GoTo DrvSndDriveChangeErr
  dirSndDir.Path = Left(drvSndDrive.Drive, 2)
  txtDir.Text = dirSndDir.Path                      'Update text box to coincide with specified directory.
  
  Exit Sub
  
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
DrvSndDriveChangeErr:
  MsgBox Err.Description, vbInformation, App.Title
  drvSndDrive.Drive = Left(dirSndDir.Path, 2)
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  '************************************************
  '* Enter = OK.
  '* Escape = Cancel.
  '************************************************
  Select Case KeyCode
    Case vbKeyReturn
      cmdOKCancel(0).Value = True
    Case vbKeyEscape
      cmdOKCancel(1).Value = True
  End Select
    
End Sub

Private Sub Form_Load()

  On Error Resume Next

  KeyPreview = True
  Call CenterForm(Me)                             ' Center Form on MDI Form
  lblFolder.Caption = "&Folder:"
  
  If gWavPath = "" Then                           ' If there's no initial path specified
    txtDir.Text = dirSndDir.Path                  ' Use whatever the DirListBox defaults to
  Else
    txtDir.Text = sDir
    dirSndDir.Path = sDir                         ' Otherwise use the initial path
    drvSndDrive.Drive = Left(sDir, 2)
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  sDir = ""                                         'In case form closed by event other
                                                    'than cmdOKCancel()
End Sub

Private Sub txtDir_GotFocus()

  '************************************************
  '* Select all the path text.
  '************************************************
  txtDir.SelStart = 0
  txtDir.SelLength = Len(txtDir.Text)

End Sub
