<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmWordList
#Region "Windows Form Designer generated code "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'This call is required by the Windows Form Designer.
		InitializeComponent()
		'This form is an MDI child.
		'This code simulates the VB6 
		' functionality of automatically
		' loading and showing an MDI
		' child's parent.
		Me.MDIParent = IPAHelp.mdiHelpCharts
		IPAHelp.mdiHelpCharts.Show
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
	Public WithEvents tvSections As System.Windows.Forms.TreeView
	Public dlgOpen As System.Windows.Forms.OpenFileDialog
	Public dlgSave As System.Windows.Forms.SaveFileDialog
	Public dlgFont As System.Windows.Forms.FontDialog
	Public dlgColor As System.Windows.Forms.ColorDialog
	Public dlgPrint As System.Windows.Forms.PrintDialog
	Public WithEvents picDragBar As System.Windows.Forms.PictureBox
	Public WithEvents picSplitter As System.Windows.Forms.PictureBox
	Public WithEvents txtUserTr As System.Windows.Forms.TextBox
	Public WithEvents cmdVerify As System.Windows.Forms.Button
	Public WithEvents txtCorrectTr As System.Windows.Forms.TextBox
	Public WithEvents cmdReplay As System.Windows.Forms.Button
	Public WithEvents cmdNext As System.Windows.Forms.Button
	Public WithEvents lblTestMode As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents picTestMode As System.Windows.Forms.Panel
	Public WithEvents Timer1 As System.Windows.Forms.Timer
    Public WithEvents ssGrid As SheridanNotSSGrid
    'KG TODO Public WithEvents ssGrid As SSDataWidgets_B.SSDBGrid
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmWordList))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.tvSections = New System.Windows.Forms.TreeView
		Me.dlgOpen = New System.Windows.Forms.OpenFileDialog
		Me.dlgSave = New System.Windows.Forms.SaveFileDialog
		Me.dlgFont = New System.Windows.Forms.FontDialog
		Me.dlgColor = New System.Windows.Forms.ColorDialog
		Me.dlgPrint = New System.Windows.Forms.PrintDialog
		Me.picDragBar = New System.Windows.Forms.PictureBox
		Me.picSplitter = New System.Windows.Forms.PictureBox
		Me.picTestMode = New System.Windows.Forms.Panel
		Me.txtUserTr = New System.Windows.Forms.TextBox
		Me.cmdVerify = New System.Windows.Forms.Button
		Me.txtCorrectTr = New System.Windows.Forms.TextBox
		Me.cmdReplay = New System.Windows.Forms.Button
		Me.cmdNext = New System.Windows.Forms.Button
		Me.lblTestMode = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Label2 = New System.Windows.Forms.Label
		Me.Timer1 = New System.Windows.Forms.Timer(components)
        Me.ssGrid = New SheridanNotSSGrid
        'KG TODO Me.ssGrid = New SSDataWidgets_B.SSDBGrid
		Me.picTestMode.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		CType(Me.ssGrid, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.Text = "Word List"
		Me.ClientSize = New System.Drawing.Size(609, 399)
		Me.Location = New System.Drawing.Point(225, 336)
		Me.ControlBox = False
		Me.KeyPreview = True
		Me.MaximizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultBounds
		Me.MinimizeBox = False
		Me.ShowInTaskbar = False
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.Enabled = True
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmWordList"
		Me.tvSections.CausesValidation = True
		Me.tvSections.Size = New System.Drawing.Size(195, 119)
		Me.tvSections.Location = New System.Drawing.Point(0, 0)
		Me.tvSections.TabIndex = 12
		Me.tvSections.HideSelection = False
		Me.tvSections.LabelEdit = False
		Me.tvSections.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.tvSections.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.tvSections.Name = "tvSections"
		Me.picDragBar.BackColor = System.Drawing.SystemColors.Window
		Me.picDragBar.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picDragBar.Size = New System.Drawing.Size(12, 70)
		Me.picDragBar.Location = New System.Drawing.Point(184, 168)
		Me.picDragBar.TabIndex = 11
		Me.picDragBar.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picDragBar.Dock = System.Windows.Forms.DockStyle.None
		Me.picDragBar.CausesValidation = True
		Me.picDragBar.Enabled = True
		Me.picDragBar.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picDragBar.TabStop = True
		Me.picDragBar.Visible = True
		Me.picDragBar.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.picDragBar.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picDragBar.Name = "picDragBar"
		Me.picSplitter.BackColor = System.Drawing.SystemColors.Window
		Me.picSplitter.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picSplitter.Size = New System.Drawing.Size(11, 67)
		Me.picSplitter.Location = New System.Drawing.Point(184, 248)
		Me.picSplitter.TabIndex = 10
		Me.picSplitter.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picSplitter.Dock = System.Windows.Forms.DockStyle.None
		Me.picSplitter.CausesValidation = True
		Me.picSplitter.Enabled = True
		Me.picSplitter.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picSplitter.TabStop = True
		Me.picSplitter.Visible = True
		Me.picSplitter.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Normal
		Me.picSplitter.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.picSplitter.Name = "picSplitter"
		Me.picTestMode.ForeColor = System.Drawing.SystemColors.WindowText
		Me.picTestMode.Size = New System.Drawing.Size(389, 243)
		Me.picTestMode.Location = New System.Drawing.Point(200, 125)
		Me.picTestMode.TabIndex = 0
		Me.picTestMode.Visible = False
		Me.picTestMode.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.picTestMode.Dock = System.Windows.Forms.DockStyle.None
		Me.picTestMode.BackColor = System.Drawing.SystemColors.Control
		Me.picTestMode.CausesValidation = True
		Me.picTestMode.Enabled = True
		Me.picTestMode.Cursor = System.Windows.Forms.Cursors.Default
		Me.picTestMode.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.picTestMode.TabStop = True
		Me.picTestMode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
		Me.picTestMode.Name = "picTestMode"
		Me.txtUserTr.AutoSize = False
		Me.txtUserTr.Size = New System.Drawing.Size(374, 26)
		Me.txtUserTr.Location = New System.Drawing.Point(5, 87)
		Me.txtUserTr.TabIndex = 5
		Me.txtUserTr.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtUserTr.AcceptsReturn = True
		Me.txtUserTr.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtUserTr.BackColor = System.Drawing.SystemColors.Window
		Me.txtUserTr.CausesValidation = True
		Me.txtUserTr.Enabled = True
		Me.txtUserTr.ForeColor = System.Drawing.SystemColors.WindowText
		Me.txtUserTr.HideSelection = True
		Me.txtUserTr.ReadOnly = False
		Me.txtUserTr.Maxlength = 0
		Me.txtUserTr.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtUserTr.MultiLine = False
		Me.txtUserTr.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtUserTr.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtUserTr.TabStop = True
		Me.txtUserTr.Visible = True
		Me.txtUserTr.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtUserTr.Name = "txtUserTr"
		Me.cmdVerify.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdVerify.Text = "&Verify"
		Me.cmdVerify.Size = New System.Drawing.Size(80, 24)
		Me.cmdVerify.Location = New System.Drawing.Point(143, 194)
		Me.cmdVerify.TabIndex = 4
		Me.cmdVerify.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdVerify.BackColor = System.Drawing.SystemColors.Control
		Me.cmdVerify.CausesValidation = True
		Me.cmdVerify.Enabled = True
		Me.cmdVerify.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdVerify.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdVerify.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdVerify.TabStop = True
		Me.cmdVerify.Name = "cmdVerify"
		Me.txtCorrectTr.AutoSize = False
		Me.txtCorrectTr.BackColor = System.Drawing.SystemColors.Control
		Me.txtCorrectTr.ForeColor = System.Drawing.SystemColors.ControlText
		Me.txtCorrectTr.Size = New System.Drawing.Size(374, 26)
		Me.txtCorrectTr.Location = New System.Drawing.Point(5, 148)
		Me.txtCorrectTr.ReadOnly = True
		Me.txtCorrectTr.TabIndex = 3
		Me.txtCorrectTr.TabStop = False
		Me.txtCorrectTr.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.txtCorrectTr.AcceptsReturn = True
		Me.txtCorrectTr.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
		Me.txtCorrectTr.CausesValidation = True
		Me.txtCorrectTr.Enabled = True
		Me.txtCorrectTr.HideSelection = True
		Me.txtCorrectTr.Maxlength = 0
		Me.txtCorrectTr.Cursor = System.Windows.Forms.Cursors.IBeam
		Me.txtCorrectTr.MultiLine = False
		Me.txtCorrectTr.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.txtCorrectTr.ScrollBars = System.Windows.Forms.ScrollBars.None
		Me.txtCorrectTr.Visible = True
		Me.txtCorrectTr.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.txtCorrectTr.Name = "txtCorrectTr"
		Me.cmdReplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdReplay.Text = "&Replay"
		Me.AcceptButton = Me.cmdReplay
		Me.cmdReplay.Size = New System.Drawing.Size(80, 24)
		Me.cmdReplay.Location = New System.Drawing.Point(54, 194)
		Me.cmdReplay.TabIndex = 2
		Me.cmdReplay.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdReplay.BackColor = System.Drawing.SystemColors.Control
		Me.cmdReplay.CausesValidation = True
		Me.cmdReplay.Enabled = True
		Me.cmdReplay.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdReplay.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdReplay.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdReplay.TabStop = True
		Me.cmdReplay.Name = "cmdReplay"
		Me.cmdNext.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdNext.Text = "&Next Word >"
		Me.cmdNext.Size = New System.Drawing.Size(80, 24)
		Me.cmdNext.Location = New System.Drawing.Point(231, 194)
		Me.cmdNext.TabIndex = 1
		Me.cmdNext.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdNext.BackColor = System.Drawing.SystemColors.Control
		Me.cmdNext.CausesValidation = True
		Me.cmdNext.Enabled = True
		Me.cmdNext.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdNext.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdNext.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdNext.TabStop = True
		Me.cmdNext.Name = "cmdNext"
		Me.lblTestMode.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblTestMode.BackColor = System.Drawing.SystemColors.Highlight
		Me.lblTestMode.Text = "TEST MODE"
		Me.lblTestMode.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblTestMode.ForeColor = System.Drawing.SystemColors.highlightText
		Me.lblTestMode.Size = New System.Drawing.Size(81, 13)
		Me.lblTestMode.Location = New System.Drawing.Point(140, 24)
		Me.lblTestMode.TabIndex = 8
		Me.lblTestMode.Enabled = True
		Me.lblTestMode.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblTestMode.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblTestMode.UseMnemonic = True
		Me.lblTestMode.Visible = True
		Me.lblTestMode.AutoSize = False
		Me.lblTestMode.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblTestMode.Name = "lblTestMode"
		Me.Label1.Text = "&Enter Transcription Here:"
		Me.Label1.Size = New System.Drawing.Size(118, 13)
		Me.Label1.Location = New System.Drawing.Point(8, 72)
		Me.Label1.TabIndex = 7
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.Color.Transparent
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = True
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Label2.BackColor = System.Drawing.Color.Transparent
		Me.Label2.Text = "&Correct Transcription:"
		Me.Label2.ForeColor = System.Drawing.SystemColors.WindowText
		Me.Label2.Size = New System.Drawing.Size(101, 13)
		Me.Label2.Location = New System.Drawing.Point(8, 133)
		Me.Label2.TabIndex = 6
		Me.Label2.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label2.Enabled = True
		Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label2.UseMnemonic = True
		Me.Label2.Visible = True
		Me.Label2.AutoSize = True
		Me.Label2.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label2.Name = "Label2"
		Me.Timer1.Enabled = False
		Me.Timer1.Interval = 1
		ssGrid.OcxState = CType(resources.GetObject("ssGrid.OcxState"), System.Windows.Forms.AxHost.State)
		Me.ssGrid.Size = New System.Drawing.Size(398, 112)
		Me.ssGrid.Location = New System.Drawing.Point(202, 4)
		Me.ssGrid.TabIndex = 9
		Me.ssGrid.Name = "ssGrid"
		CType(Me.ssGrid, System.ComponentModel.ISupportInitialize).EndInit()
		Me.Controls.Add(tvSections)
		Me.Controls.Add(picDragBar)
		Me.Controls.Add(picSplitter)
		Me.Controls.Add(picTestMode)
		Me.Controls.Add(ssGrid)
		Me.picTestMode.Controls.Add(txtUserTr)
		Me.picTestMode.Controls.Add(cmdVerify)
		Me.picTestMode.Controls.Add(txtCorrectTr)
		Me.picTestMode.Controls.Add(cmdReplay)
		Me.picTestMode.Controls.Add(cmdNext)
		Me.picTestMode.Controls.Add(lblTestMode)
		Me.picTestMode.Controls.Add(Label1)
		Me.picTestMode.Controls.Add(Label2)
		Me.picTestMode.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class