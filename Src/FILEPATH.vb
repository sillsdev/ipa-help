Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmFilePath
	Inherits System.Windows.Forms.Form
	'**************************************************
	'* frmFilePath version info:
	'*  See basGlobals (Globals.bas).
	'**************************************************
	
	
	Public sDir As String '* Directory Passed between caller & callee.
	
	Private Function InvalDir() As Short
		
		'**************************************************
		'* This function checks the validity of the
		'* selected directory. It returns true if the
		'* directory is invalid, false if it is valid.
		'**************************************************
		
		Dim i As Short
		Dim tmpDir As String
		Dim sMsg As String
		
		tmpDir = LCase(txtDir.Text)
		If Mid(tmpDir, 2, 1) = ":" Then 'Does input have drive specification?
			For i = 0 To drvSndDrive.Items.Count - 1 'Search drive list to determine whether
				If VB.Left(tmpDir, 1) = VB.Left(drvSndDrive.Items(i), 1) Then Exit For '  or not specified drive exists.
			Next i
			
			'**********************************************
			'* If you get to the end of the drive list and
			'* don't find the specified drive ...
			'**********************************************
			If i = drvSndDrive.Items.Count Then 'Was specified drive found?
				sMsg = "Invalid drive specification."
				GoTo BadDirSpec
			Else
				tmpDir = Mid(tmpDir, 3) 'Save input without drive spec.
			End If
		End If
		
		'************************************************
		'* Disallow invalid characters.
		'************************************************
		If InStr(tmpDir, ":") Or InStr(tmpDir, "*") Or InStr(tmpDir, "?") Or InStr(tmpDir, Chr(34)) Or InStr(tmpDir, "<") Or InStr(tmpDir, ">") Or InStr(tmpDir, "|") Then 'If directory name contains :, *, ?, ", <, >, or |
			sMsg = "Invalid character in " & "directory specification."
			GoTo BadDirSpec
		End If
		
		InvalDir = False
		Exit Function
		
BadDirSpec: 
		MsgBox(sMsg, MsgBoxStyle.Exclamation, My.Application.Info.Title)
		InvalDir = True
		
	End Function
	
	Private Sub cmdOKCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOKCancel.Click
		Dim Index As Short = cmdOKCancel.GetIndex(eventSender)
		
		Dim savDir As String
		Dim sMsg As String
		
		savDir = CurDir() 'Save current directory.
		
		'************************************************
		'* If OK pressed, check if the path is valid.
		'* Otherwise, just exit.
		'************************************************
        If cmdOKCancel(0).Text Then 'Did user press OK button?
            sDir = Trim(txtDir.Text) 'Trim any leading & trailing spaces.

            If sDir = "" Or InvalDir() Then 'Does input contain valid characters?
                txtDir.Focus()
                Exit Sub
            End If

            sDir = txtDir.Text & IIf(VB.Right(txtDir.Text, 1) = "\", "", "\") 'Make sure path has trailing backslash.

            On Error GoTo DirNotExist 'Prepare not to find drive or directory.
            If Mid(sDir, 2, 1) = ":" Then ChDrive(VB.Left(sDir, 2)) 'If input has drive in it then access it.
            ChDir(sDir) 'Try to access specified directory.
            On Error GoTo 0 'If we're here then input directory exists.

            ChDrive(VB.Left(savDir, 2)) 'Set drive and directory back to what
            ChDir(savDir) '  it was before dir exists test.
            Call SaveSetting(gINIPath, cPathsSect, cSoundsEntry, sDir)
        Else 'OK Not pressed means Cancel was pressed
            sDir = "" '  so return nothing.
        End If
		
		Me.Visible = False
		
		Exit Sub
		
DirNotExist: 'At this point the user specified dir doesn't exist.
		On Error GoTo 0
		ChDrive(VB.Left(savDir, 2)) 'Set drive and directory back to what
		ChDir(savDir) '  it was before dir exisits test.
		sMsg = "Directory does not exist. " & "Would you like to create it?"
		
		If MsgBox(sMsg, MsgBoxStyle.Question + MsgBoxStyle.YesNo, My.Application.Info.Title) = MsgBoxResult.Yes Then
			MkDir(txtDir.Text) 'User wants to make new directory.
		Else
			txtDir.Focus() 'Re-highlight directory text-box
			txtDir.SelectionStart = 0
			txtDir.SelectionLength = Len(txtDir.Text)
		End If
		
		Exit Sub
		
	End Sub
	
	Private Sub dirSndDir_Change(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles dirSndDir.Change
		
		On Error Resume Next
		txtDir.Text = dirSndDir.Path
		
	End Sub
	
	Private Sub drvSndDrive_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles drvSndDrive.SelectedIndexChanged
		
		On Error GoTo DrvSndDriveChangeErr
		dirSndDir.Path = VB.Left(drvSndDrive.Drive, 2)
		txtDir.Text = dirSndDir.Path 'Update text box to coincide with specified directory.
		
		Exit Sub
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
DrvSndDriveChangeErr: 
		MsgBox(Err.Description, MsgBoxStyle.Information, My.Application.Info.Title)
		drvSndDrive.Drive = VB.Left(dirSndDir.Path, 2)
		
	End Sub
	
	Private Sub frmFilePath_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		'************************************************
		'* Enter = OK.
		'* Escape = Cancel.
		'************************************************
		Select Case KeyCode
			Case System.Windows.Forms.Keys.Return
				cmdOKCancel(0).PerformClick()
			Case System.Windows.Forms.Keys.Escape
				cmdOKCancel(1).PerformClick()
		End Select
		
	End Sub
	
	Private Sub frmFilePath_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		On Error Resume Next
		
		KeyPreview = True
		Call CenterForm(Me) ' Center Form on MDI Form
		lblFolder.Text = "&Folder:"
		
		If gWavPath = "" Then ' If there's no initial path specified
			txtDir.Text = dirSndDir.Path ' Use whatever the DirListBox defaults to
		Else
			txtDir.Text = sDir
			dirSndDir.Path = sDir ' Otherwise use the initial path
			drvSndDrive.Drive = VB.Left(sDir, 2)
		End If
		
	End Sub
	
	Private Sub frmFilePath_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		sDir = "" 'In case form closed by event other
		'than cmdOKCancel()
		eventArgs.Cancel = Cancel
	End Sub
	
	Private Sub txtDir_Enter(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDir.Enter
		
		'************************************************
		'* Select all the path text.
		'************************************************
		txtDir.SelectionStart = 0
		txtDir.SelectionLength = Len(txtDir.Text)
		
	End Sub
End Class