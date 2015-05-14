<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmGetSoundFile
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
	Public WithEvents cmdOK As System.Windows.Forms.Button
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cboFileTypes As System.Windows.Forms.ComboBox
	Public WithEvents File1 As Microsoft.VisualBasic.Compatibility.VB6.FileListBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmGetSoundFile))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdOK = New System.Windows.Forms.Button
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cboFileTypes = New System.Windows.Forms.ComboBox
		Me.File1 = New Microsoft.VisualBasic.Compatibility.VB6.FileListBox
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Choose a Sound File"
		Me.ClientSize = New System.Drawing.Size(270, 186)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.ControlBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
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
		Me.Name = "frmGetSoundFile"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(65, 25)
		Me.cmdOK.Location = New System.Drawing.Point(132, 156)
		Me.cmdOK.TabIndex = 2
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(200, 156)
		Me.cmdCancel.TabIndex = 3
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cboFileTypes.Size = New System.Drawing.Size(122, 21)
		Me.cboFileTypes.Location = New System.Drawing.Point(4, 159)
		Me.cboFileTypes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
		Me.cboFileTypes.TabIndex = 1
		Me.cboFileTypes.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cboFileTypes.BackColor = System.Drawing.SystemColors.Window
		Me.cboFileTypes.CausesValidation = True
		Me.cboFileTypes.Enabled = True
		Me.cboFileTypes.ForeColor = System.Drawing.SystemColors.WindowText
		Me.cboFileTypes.IntegralHeight = True
		Me.cboFileTypes.Cursor = System.Windows.Forms.Cursors.Default
		Me.cboFileTypes.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cboFileTypes.Sorted = False
		Me.cboFileTypes.TabStop = True
		Me.cboFileTypes.Visible = True
		Me.cboFileTypes.Name = "cboFileTypes"
		Me.File1.Size = New System.Drawing.Size(261, 149)
		Me.File1.Location = New System.Drawing.Point(4, 4)
		Me.File1.TabIndex = 0
		Me.File1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.File1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.File1.Archive = True
		Me.File1.BackColor = System.Drawing.SystemColors.Window
		Me.File1.CausesValidation = True
		Me.File1.Enabled = True
		Me.File1.ForeColor = System.Drawing.SystemColors.WindowText
		Me.File1.Hidden = False
		Me.File1.Cursor = System.Windows.Forms.Cursors.Default
		Me.File1.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.File1.Normal = True
		Me.File1.Pattern = "*.*"
		Me.File1.ReadOnly = True
		Me.File1.System = False
		Me.File1.TabStop = True
		Me.File1.TopIndex = 0
		Me.File1.Visible = True
		Me.File1.Name = "File1"
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cboFileTypes)
		Me.Controls.Add(File1)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class