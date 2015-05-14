<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmFilePath
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents dirSndDir As Microsoft.VisualBasic.Compatibility.VB6.DirListBox
	Public WithEvents txtDir As System.Windows.Forms.TextBox
	Public WithEvents drvSndDrive As Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
	Public WithEvents _cmdOKCancel_1 As System.Windows.Forms.Button
	Public WithEvents _cmdOKCancel_0 As System.Windows.Forms.Button
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents lblFolder As System.Windows.Forms.Label
	Public WithEvents cmdOKCancel As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmFilePath))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.dirSndDir = New Microsoft.VisualBasic.Compatibility.VB6.DirListBox
		Me.txtDir = New System.Windows.Forms.TextBox
		Me.drvSndDrive = New Microsoft.VisualBasic.Compatibility.VB6.DriveListBox
		Me._cmdOKCancel_1 = New System.Windows.Forms.Button
		Me._cmdOKCancel_0 = New System.Windows.Forms.Button
		Me.Label4 = New System.Windows.Forms.Label
		Me.lblFolder = New System.Windows.Forms.Label
		Me.cmdOKCancel = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.cmdOKCancel, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
		Me.Text = "Set Sound File Path"
		Me.ClientSize = New System.Drawing.Size(333, 252)
		Me.Location = New System.Drawing.Point(607, 237)
		Me.ControlBox = False
		Me.Icon = CType(resources.GetObject("frmFilePath.Icon"), System.Drawing.Icon)
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.Enabled = True
		Me.KeyPreview = False
		Me.MaximizeBox = True
		Me.MinimizeBox = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmFilePath"
		Me.dirSndDir.Size = New System.Drawing.Size(236, 156)
		Me.dirSndDir.Location = New System.Drawing.Point(4, 44)
		Me.dirSndDir.TabIndex = 2
		Me.dirSndDir.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.dirSndDir.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.dirSndDir.BackColor = System.Drawing.SystemColors.Window
		Me.dirSndDir.CausesValidation = True
		Me.dirSndDir.Enabled = True
		Me.dirSndDir.ForeColor = System.Drawing.SystemColors.WindowText
		Me.dirSndDir.Cursor = System.Windows.Forms.Cursors.Default
		Me.dirSndDir.TabStop = True
		Me.dirSndDir.Visible = True
		Me.dirSndDir.Name = "dirSndDir"
		Me.txtDir.AutoSize = False
		Me.txtDir.Size = New System.Drawing.Size(236, 20)
		Me.txtDir.Location = New System.Drawing.Point(4, 20)
		Me.txtDir.TabIndex = 1
		Me.txtDir.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtDir.AcceptsReturn = True
		Me.txtDir.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtDir.BackColor = System.Drawing.SystemColors.Window
		Me.txtDir.CausesValidation = True
		Me.txtDir.Enabled = True
		Me.txtDir.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtDir.HideSelection = True
		Me.txtDir.ReadOnly = False
		Me.txtDir.Maxlength = 0
		Me.txtDir.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtDir.MultiLine = False
		Me.txtDir.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtDir.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtDir.TabStop = True
		Me.txtDir.Visible = True
		Me.txtDir.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtDir.Name = "txtDir"
		Me.drvSndDrive.Size = New System.Drawing.Size(236, 21)
		Me.drvSndDrive.Location = New System.Drawing.Point(4, 224)
		Me.drvSndDrive.TabIndex = 4
		Me.drvSndDrive.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.drvSndDrive.BackColor = System.Drawing.SystemColors.Window
		Me.drvSndDrive.CausesValidation = True
		Me.drvSndDrive.Enabled = True
		Me.drvSndDrive.ForeColor = System.Drawing.SystemColors.WindowText
		Me.drvSndDrive.Cursor = System.Windows.Forms.Cursors.Default
		Me.drvSndDrive.TabStop = True
		Me.drvSndDrive.Visible = True
		Me.drvSndDrive.Name = "drvSndDrive"
		Me._cmdOKCancel_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me._cmdOKCancel_1
		Me._cmdOKCancel_1.Text = "Cancel"
		Me._cmdOKCancel_1.Size = New System.Drawing.Size(68, 24)
		Me._cmdOKCancel_1.Location = New System.Drawing.Point(256, 52)
		Me._cmdOKCancel_1.TabIndex = 6
		Me._cmdOKCancel_1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdOKCancel_1.BackColor = System.Drawing.SystemColors.Control
		Me._cmdOKCancel_1.CausesValidation = True
		Me._cmdOKCancel_1.Enabled = True
		Me._cmdOKCancel_1.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdOKCancel_1.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdOKCancel_1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdOKCancel_1.TabStop = True
		Me._cmdOKCancel_1.Name = "_cmdOKCancel_1"
		Me._cmdOKCancel_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me._cmdOKCancel_0.Text = "OK"
		Me.AcceptButton = Me._cmdOKCancel_0
		Me._cmdOKCancel_0.Size = New System.Drawing.Size(68, 24)
		Me._cmdOKCancel_0.Location = New System.Drawing.Point(256, 20)
		Me._cmdOKCancel_0.TabIndex = 5
		Me._cmdOKCancel_0.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me._cmdOKCancel_0.BackColor = System.Drawing.SystemColors.Control
		Me._cmdOKCancel_0.CausesValidation = True
		Me._cmdOKCancel_0.Enabled = True
		Me._cmdOKCancel_0.ForeColor = System.Drawing.SystemColors.ControlText
		Me._cmdOKCancel_0.Cursor = System.Windows.Forms.Cursors.Default
		Me._cmdOKCancel_0.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me._cmdOKCancel_0.TabStop = True
		Me._cmdOKCancel_0.Name = "_cmdOKCancel_0"
		Me.Label4.Text = "Dri&ves:"
		Me.Label4.Size = New System.Drawing.Size(119, 16)
		Me.Label4.Location = New System.Drawing.Point(8, 208)
		Me.Label4.TabIndex = 3
		Me.Label4.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label4.BackColor = System.Drawing.SystemColors.Control
		Me.Label4.Enabled = True
		Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label4.UseMnemonic = True
		Me.Label4.Visible = True
		Me.Label4.AutoSize = False
		Me.Label4.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label4.Name = "Label4"
		Me.lblFolder.Text = "&Folder:"
		Me.lblFolder.Size = New System.Drawing.Size(119, 16)
		Me.lblFolder.Location = New System.Drawing.Point(8, 4)
		Me.lblFolder.TabIndex = 0
		Me.lblFolder.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblFolder.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblFolder.BackColor = System.Drawing.SystemColors.Control
		Me.lblFolder.Enabled = True
		Me.lblFolder.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFolder.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFolder.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFolder.UseMnemonic = True
		Me.lblFolder.Visible = True
		Me.lblFolder.AutoSize = False
		Me.lblFolder.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblFolder.Name = "lblFolder"
		Me.Controls.Add(dirSndDir)
		Me.Controls.Add(txtDir)
		Me.Controls.Add(drvSndDrive)
		Me.Controls.Add(_cmdOKCancel_1)
		Me.Controls.Add(_cmdOKCancel_0)
		Me.Controls.Add(Label4)
		Me.Controls.Add(lblFolder)
		Me.cmdOKCancel.SetIndex(_cmdOKCancel_1, CType(1, Short))
		Me.cmdOKCancel.SetIndex(_cmdOKCancel_0, CType(0, Short))
		CType(Me.cmdOKCancel, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class