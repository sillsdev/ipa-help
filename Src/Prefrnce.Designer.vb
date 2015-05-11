<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPreferences
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
	Public WithEvents _cmdOKCancel_1 As System.Windows.Forms.Button
	Public WithEvents _cmdOKCancel_0 As System.Windows.Forms.Button
	Public WithEvents _cmdBrowse_1 As System.Windows.Forms.Button
	Public WithEvents txtSASLoc As System.Windows.Forms.TextBox
	Public WithEvents _cmdBrowse_0 As System.Windows.Forms.Button
	Public WithEvents txtSndsLoc As System.Windows.Forms.TextBox
	Public WithEvents _Label1_6 As System.Windows.Forms.Label
	Public WithEvents _Label1_5 As System.Windows.Forms.Label
	Public WithEvents _Frame1_1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdAdd As System.Windows.Forms.Button
	Public WithEvents cmdRemove As System.Windows.Forms.Button
	Public WithEvents cmdBrowseWL As System.Windows.Forms.Button
	Public WithEvents _lvWLLocations_ColumnHeader_1 As System.Windows.Forms.ColumnHeader
	Public WithEvents _lvWLLocations_ColumnHeader_2 As System.Windows.Forms.ColumnHeader
	Public WithEvents lvWLLocations As System.Windows.Forms.ListView
	Public WithEvents _Frame1_0 As System.Windows.Forms.GroupBox
	Public WithEvents updnSRSpeed As System.Windows.Forms.Label
	Public WithEvents _updnDelay_1 As System.Windows.Forms.Label
	Public WithEvents updnPlaybackTimes As System.Windows.Forms.Label
	Public WithEvents txtSRSpeed As System.Windows.Forms.TextBox
	Public WithEvents txtRptDelay As System.Windows.Forms.TextBox
	Public WithEvents txtRptCount As System.Windows.Forms.TextBox
	Public WithEvents txtInitDelay As System.Windows.Forms.TextBox
	Public WithEvents _updnDelay_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_7 As System.Windows.Forms.Label
	Public WithEvents _Label1_8 As System.Windows.Forms.Label
	Public WithEvents _Label1_4 As System.Windows.Forms.Label
	Public WithEvents _Label1_0 As System.Windows.Forms.Label
	Public WithEvents _Label1_3 As System.Windows.Forms.Label
	Public WithEvents _Label1_2 As System.Windows.Forms.Label
	Public WithEvents _Label1_1 As System.Windows.Forms.Label
	Public WithEvents _Frame1_2 As System.Windows.Forms.GroupBox
	Public dlgOpen As System.Windows.Forms.OpenFileDialog
	Public WithEvents Frame1 As Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray
	Public WithEvents Label1 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
	Public WithEvents cmdBrowse As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	Public WithEvents cmdOKCancel As Microsoft.VisualBasic.Compatibility.VB6.ButtonArray
	Public WithEvents updnDelay As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmPreferences))
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
        Me._cmdOKCancel_1 = New System.Windows.Forms.Button
        Me._cmdOKCancel_0 = New System.Windows.Forms.Button
        Me._Frame1_1 = New System.Windows.Forms.GroupBox
        Me._cmdBrowse_1 = New System.Windows.Forms.Button
        Me.txtSASLoc = New System.Windows.Forms.TextBox
        Me._cmdBrowse_0 = New System.Windows.Forms.Button
        Me.txtSndsLoc = New System.Windows.Forms.TextBox
        Me._Label1_6 = New System.Windows.Forms.Label
        Me._Label1_5 = New System.Windows.Forms.Label
        Me._Frame1_0 = New System.Windows.Forms.GroupBox
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.cmdRemove = New System.Windows.Forms.Button
        Me.cmdBrowseWL = New System.Windows.Forms.Button
        Me.lvWLLocations = New System.Windows.Forms.ListView
        Me._lvWLLocations_ColumnHeader_1 = New System.Windows.Forms.ColumnHeader
        Me._lvWLLocations_ColumnHeader_2 = New System.Windows.Forms.ColumnHeader
        Me._Frame1_2 = New System.Windows.Forms.GroupBox
        Me.updnSRSpeed = New System.Windows.Forms.Label
        Me._updnDelay_1 = New System.Windows.Forms.Label
        Me.updnPlaybackTimes = New System.Windows.Forms.Label
        Me.txtSRSpeed = New System.Windows.Forms.TextBox
        Me.txtRptDelay = New System.Windows.Forms.TextBox
        Me.txtRptCount = New System.Windows.Forms.TextBox
        Me.txtInitDelay = New System.Windows.Forms.TextBox
        Me._updnDelay_0 = New System.Windows.Forms.Label
        Me._Label1_7 = New System.Windows.Forms.Label
        Me._Label1_8 = New System.Windows.Forms.Label
        Me._Label1_4 = New System.Windows.Forms.Label
        Me._Label1_0 = New System.Windows.Forms.Label
        Me._Label1_3 = New System.Windows.Forms.Label
        Me._Label1_2 = New System.Windows.Forms.Label
        Me._Label1_1 = New System.Windows.Forms.Label
        Me.dlgOpen = New System.Windows.Forms.OpenFileDialog
        Me.Frame1 = New Microsoft.VisualBasic.Compatibility.VB6.GroupBoxArray(components)
        Me.Label1 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
        Me.cmdBrowse = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
        Me.cmdOKCancel = New Microsoft.VisualBasic.Compatibility.VB6.ButtonArray(components)
        Me.updnDelay = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
        Me._Frame1_1.SuspendLayout()
        Me._Frame1_0.SuspendLayout()
        Me.lvWLLocations.SuspendLayout()
        Me._Frame1_2.SuspendLayout()
        Me.SuspendLayout()
        Me.ToolTip1.Active = True
        CType(Me.Frame1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdBrowse, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdOKCancel, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.updnDelay, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "IPA Help Preferences"
        Me.ClientSize = New System.Drawing.Size(577, 240)
        Me.Location = New System.Drawing.Point(98, 415)
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
        Me.ControlBox = True
        Me.Enabled = True
        Me.KeyPreview = False
        Me.MaximizeBox = True
        Me.MinimizeBox = True
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = True
        Me.HelpButton = False
        Me.WindowState = System.Windows.Forms.FormWindowState.Normal
        Me.Name = "frmPreferences"
        Me._cmdOKCancel_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.CancelButton = Me._cmdOKCancel_1
        Me._cmdOKCancel_1.Text = "Cancel"
        Me._cmdOKCancel_1.Size = New System.Drawing.Size(61, 25)
        Me._cmdOKCancel_1.Location = New System.Drawing.Point(513, 207)
        Me._cmdOKCancel_1.TabIndex = 20
        Me._cmdOKCancel_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
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
        Me._cmdOKCancel_0.Size = New System.Drawing.Size(61, 25)
        Me._cmdOKCancel_0.Location = New System.Drawing.Point(447, 207)
        Me._cmdOKCancel_0.TabIndex = 19
        Me._cmdOKCancel_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cmdOKCancel_0.BackColor = System.Drawing.SystemColors.Control
        Me._cmdOKCancel_0.CausesValidation = True
        Me._cmdOKCancel_0.Enabled = True
        Me._cmdOKCancel_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._cmdOKCancel_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmdOKCancel_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmdOKCancel_0.TabStop = True
        Me._cmdOKCancel_0.Name = "_cmdOKCancel_0"
        Me._Frame1_1.Text = "Other Locations"
        Me._Frame1_1.Size = New System.Drawing.Size(366, 108)
        Me._Frame1_1.Location = New System.Drawing.Point(2, 128)
        Me._Frame1_1.TabIndex = 24
        Me._Frame1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Frame1_1.BackColor = System.Drawing.SystemColors.Control
        Me._Frame1_1.Enabled = True
        Me._Frame1_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Frame1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame1_1.Visible = True
        Me._Frame1_1.Padding = New System.Windows.Forms.Padding(0)
        Me._Frame1_1.Name = "_Frame1_1"
        Me._cmdBrowse_1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me._cmdBrowse_1.Text = "Browse..."
        Me._cmdBrowse_1.Size = New System.Drawing.Size(59, 22)
        Me._cmdBrowse_1.Location = New System.Drawing.Point(300, 77)
        Me._cmdBrowse_1.TabIndex = 10
        Me._cmdBrowse_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cmdBrowse_1.BackColor = System.Drawing.SystemColors.Control
        Me._cmdBrowse_1.CausesValidation = True
        Me._cmdBrowse_1.Enabled = True
        Me._cmdBrowse_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me._cmdBrowse_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmdBrowse_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmdBrowse_1.TabStop = True
        Me._cmdBrowse_1.Name = "_cmdBrowse_1"
        Me.txtSASLoc.AutoSize = False
        Me.txtSASLoc.Size = New System.Drawing.Size(287, 21)
        Me.txtSASLoc.Location = New System.Drawing.Point(9, 77)
        Me.txtSASLoc.ReadOnly = True
        Me.txtSASLoc.TabIndex = 9
        Me.txtSASLoc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSASLoc.AcceptsReturn = True
        Me.txtSASLoc.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtSASLoc.BackColor = System.Drawing.SystemColors.Window
        Me.txtSASLoc.CausesValidation = True
        Me.txtSASLoc.Enabled = True
        Me.txtSASLoc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSASLoc.HideSelection = True
        Me.txtSASLoc.MaxLength = 0
        Me.txtSASLoc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSASLoc.Multiline = False
        Me.txtSASLoc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSASLoc.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtSASLoc.TabStop = True
        Me.txtSASLoc.Visible = True
        Me.txtSASLoc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtSASLoc.Name = "txtSASLoc"
        Me._cmdBrowse_0.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me._cmdBrowse_0.Text = "Browse..."
        Me._cmdBrowse_0.Size = New System.Drawing.Size(59, 22)
        Me._cmdBrowse_0.Location = New System.Drawing.Point(300, 35)
        Me._cmdBrowse_0.TabIndex = 7
        Me._cmdBrowse_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._cmdBrowse_0.BackColor = System.Drawing.SystemColors.Control
        Me._cmdBrowse_0.CausesValidation = True
        Me._cmdBrowse_0.Enabled = True
        Me._cmdBrowse_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._cmdBrowse_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._cmdBrowse_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._cmdBrowse_0.TabStop = True
        Me._cmdBrowse_0.Name = "_cmdBrowse_0"
        Me.txtSndsLoc.AutoSize = False
        Me.txtSndsLoc.Size = New System.Drawing.Size(287, 21)
        Me.txtSndsLoc.Location = New System.Drawing.Point(9, 35)
        Me.txtSndsLoc.ReadOnly = True
        Me.txtSndsLoc.TabIndex = 6
        Me.txtSndsLoc.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSndsLoc.AcceptsReturn = True
        Me.txtSndsLoc.TextAlign = System.Windows.Forms.HorizontalAlignment.Left
        Me.txtSndsLoc.BackColor = System.Drawing.SystemColors.Window
        Me.txtSndsLoc.CausesValidation = True
        Me.txtSndsLoc.Enabled = True
        Me.txtSndsLoc.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSndsLoc.HideSelection = True
        Me.txtSndsLoc.MaxLength = 0
        Me.txtSndsLoc.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSndsLoc.Multiline = False
        Me.txtSndsLoc.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSndsLoc.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtSndsLoc.TabStop = True
        Me.txtSndsLoc.Visible = True
        Me.txtSndsLoc.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtSndsLoc.Name = "txtSndsLoc"
        Me._Label1_6.BackColor = System.Drawing.Color.Transparent
        Me._Label1_6.Text = "Speech Analy&zer Server Location"
        Me._Label1_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me._Label1_6.Size = New System.Drawing.Size(158, 13)
        Me._Label1_6.Location = New System.Drawing.Point(11, 61)
        Me._Label1_6.TabIndex = 8
        Me._Label1_6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_6.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._Label1_6.Enabled = True
        Me._Label1_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_6.UseMnemonic = True
        Me._Label1_6.Visible = True
        Me._Label1_6.AutoSize = True
        Me._Label1_6.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._Label1_6.Name = "_Label1_6"
        Me._Label1_5.BackColor = System.Drawing.Color.Transparent
        Me._Label1_5.Text = "&Phonetic Sound Files Location"
        Me._Label1_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me._Label1_5.Size = New System.Drawing.Size(144, 13)
        Me._Label1_5.Location = New System.Drawing.Point(11, 19)
        Me._Label1_5.TabIndex = 5
        Me._Label1_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_5.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._Label1_5.Enabled = True
        Me._Label1_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_5.UseMnemonic = True
        Me._Label1_5.Visible = True
        Me._Label1_5.AutoSize = True
        Me._Label1_5.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._Label1_5.Name = "_Label1_5"
        Me._Frame1_0.Text = "&Word List Locations"
        Me._Frame1_0.Size = New System.Drawing.Size(365, 118)
        Me._Frame1_0.Location = New System.Drawing.Point(3, 4)
        Me._Frame1_0.TabIndex = 0
        Me._Frame1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Frame1_0.BackColor = System.Drawing.SystemColors.Control
        Me._Frame1_0.Enabled = True
        Me._Frame1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Frame1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame1_0.Visible = True
        Me._Frame1_0.Padding = New System.Windows.Forms.Padding(0)
        Me._Frame1_0.Name = "_Frame1_0"
        Me.cmdAdd.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.cmdAdd.Text = "&Add"
        Me.cmdAdd.Size = New System.Drawing.Size(59, 25)
        Me.cmdAdd.Location = New System.Drawing.Point(297, 51)
        Me.cmdAdd.TabIndex = 3
        Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAdd.CausesValidation = True
        Me.cmdAdd.Enabled = True
        Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAdd.TabStop = True
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdRemove.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.cmdRemove.Text = "&Remove"
        Me.cmdRemove.Size = New System.Drawing.Size(59, 25)
        Me.cmdRemove.Location = New System.Drawing.Point(297, 82)
        Me.cmdRemove.TabIndex = 4
        Me.cmdRemove.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRemove.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRemove.CausesValidation = True
        Me.cmdRemove.Enabled = True
        Me.cmdRemove.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRemove.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRemove.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRemove.TabStop = True
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdBrowseWL.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.cmdBrowseWL.Text = "&Browse..."
        Me.cmdBrowseWL.Size = New System.Drawing.Size(59, 25)
        Me.cmdBrowseWL.Location = New System.Drawing.Point(297, 21)
        Me.cmdBrowseWL.TabIndex = 2
        Me.cmdBrowseWL.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowseWL.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBrowseWL.CausesValidation = True
        Me.cmdBrowseWL.Enabled = True
        Me.cmdBrowseWL.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowseWL.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBrowseWL.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBrowseWL.TabStop = True
        Me.cmdBrowseWL.Name = "cmdBrowseWL"
        Me.lvWLLocations.Size = New System.Drawing.Size(283, 94)
        Me.lvWLLocations.Location = New System.Drawing.Point(6, 18)
        Me.lvWLLocations.TabIndex = 1
        Me.lvWLLocations.View = System.Windows.Forms.View.Details
        Me.lvWLLocations.LabelWrap = False
        Me.lvWLLocations.HideSelection = False
        Me.lvWLLocations.FullRowSelect = True
        Me.lvWLLocations.ForeColor = System.Drawing.SystemColors.WindowText
        Me.lvWLLocations.BackColor = System.Drawing.SystemColors.Window
        Me.lvWLLocations.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lvWLLocations.LabelEdit = True
        Me.lvWLLocations.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.lvWLLocations.Name = "lvWLLocations"
        Me._lvWLLocations_ColumnHeader_1.Text = "Word List Title"
        Me._lvWLLocations_ColumnHeader_1.Width = 212
        Me._lvWLLocations_ColumnHeader_2.Text = "File"
        Me._lvWLLocations_ColumnHeader_2.Width = 170
        Me._Frame1_2.Text = "Word List Audio Playback Settings"
        Me._Frame1_2.Size = New System.Drawing.Size(203, 195)
        Me._Frame1_2.Location = New System.Drawing.Point(373, 4)
        Me._Frame1_2.TabIndex = 21
        Me._Frame1_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Frame1_2.BackColor = System.Drawing.SystemColors.Control
        Me._Frame1_2.Enabled = True
        Me._Frame1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Frame1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Frame1_2.Visible = True
        Me._Frame1_2.Padding = New System.Windows.Forms.Padding(0)
        Me._Frame1_2.Name = "_Frame1_2"
        Me.updnSRSpeed.Size = New System.Drawing.Size(16, 20)
        Me.updnSRSpeed.Location = New System.Drawing.Point(154, 156)
        Me.updnSRSpeed.TabIndex = 28
        Me.updnSRSpeed.Text = "updnSRSpeed"
        Me.updnSRSpeed.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.updnSRSpeed.BackColor = System.Drawing.Color.Red
        Me.updnSRSpeed.Name = "updnSRSpeed"
        Me._updnDelay_1.Size = New System.Drawing.Size(16, 20)
        Me._updnDelay_1.Location = New System.Drawing.Point(154, 72)
        Me._updnDelay_1.TabIndex = 27
        Me._updnDelay_1.Text = "_updnDelay_1"
        Me._updnDelay_1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._updnDelay_1.BackColor = System.Drawing.Color.Red
        Me._updnDelay_1.Name = "_updnDelay_1"
        Me.updnPlaybackTimes.Size = New System.Drawing.Size(16, 20)
        Me.updnPlaybackTimes.Location = New System.Drawing.Point(154, 114)
        Me.updnPlaybackTimes.TabIndex = 26
        Me.updnPlaybackTimes.Text = "updnPlaybackTimes"
        Me.updnPlaybackTimes.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.updnPlaybackTimes.BackColor = System.Drawing.Color.Red
        Me.updnPlaybackTimes.Name = "updnPlaybackTimes"
        Me.txtSRSpeed.AutoSize = False
        Me.txtSRSpeed.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtSRSpeed.Size = New System.Drawing.Size(37, 19)
        Me.txtSRSpeed.Location = New System.Drawing.Point(117, 157)
        Me.txtSRSpeed.TabIndex = 18
        Me.txtSRSpeed.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtSRSpeed.AcceptsReturn = True
        Me.txtSRSpeed.BackColor = System.Drawing.SystemColors.Window
        Me.txtSRSpeed.CausesValidation = True
        Me.txtSRSpeed.Enabled = True
        Me.txtSRSpeed.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtSRSpeed.HideSelection = True
        Me.txtSRSpeed.ReadOnly = False
        Me.txtSRSpeed.MaxLength = 0
        Me.txtSRSpeed.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtSRSpeed.Multiline = False
        Me.txtSRSpeed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtSRSpeed.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtSRSpeed.TabStop = True
        Me.txtSRSpeed.Visible = True
        Me.txtSRSpeed.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtSRSpeed.Name = "txtSRSpeed"
        Me.txtRptDelay.AutoSize = False
        Me.txtRptDelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtRptDelay.Size = New System.Drawing.Size(37, 19)
        Me.txtRptDelay.Location = New System.Drawing.Point(117, 73)
        Me.txtRptDelay.TabIndex = 14
        Me.txtRptDelay.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRptDelay.AcceptsReturn = True
        Me.txtRptDelay.BackColor = System.Drawing.SystemColors.Window
        Me.txtRptDelay.CausesValidation = True
        Me.txtRptDelay.Enabled = True
        Me.txtRptDelay.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRptDelay.HideSelection = True
        Me.txtRptDelay.ReadOnly = False
        Me.txtRptDelay.MaxLength = 0
        Me.txtRptDelay.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRptDelay.Multiline = False
        Me.txtRptDelay.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRptDelay.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtRptDelay.TabStop = True
        Me.txtRptDelay.Visible = True
        Me.txtRptDelay.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtRptDelay.Name = "txtRptDelay"
        Me.txtRptCount.AutoSize = False
        Me.txtRptCount.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtRptCount.Size = New System.Drawing.Size(37, 19)
        Me.txtRptCount.Location = New System.Drawing.Point(117, 115)
        Me.txtRptCount.TabIndex = 16
        Me.txtRptCount.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRptCount.AcceptsReturn = True
        Me.txtRptCount.BackColor = System.Drawing.SystemColors.Window
        Me.txtRptCount.CausesValidation = True
        Me.txtRptCount.Enabled = True
        Me.txtRptCount.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRptCount.HideSelection = True
        Me.txtRptCount.ReadOnly = False
        Me.txtRptCount.MaxLength = 0
        Me.txtRptCount.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRptCount.Multiline = False
        Me.txtRptCount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRptCount.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtRptCount.TabStop = True
        Me.txtRptCount.Visible = True
        Me.txtRptCount.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtRptCount.Name = "txtRptCount"
        Me.txtInitDelay.AutoSize = False
        Me.txtInitDelay.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtInitDelay.Size = New System.Drawing.Size(37, 19)
        Me.txtInitDelay.Location = New System.Drawing.Point(117, 31)
        Me.txtInitDelay.TabIndex = 12
        Me.txtInitDelay.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInitDelay.AcceptsReturn = True
        Me.txtInitDelay.BackColor = System.Drawing.SystemColors.Window
        Me.txtInitDelay.CausesValidation = True
        Me.txtInitDelay.Enabled = True
        Me.txtInitDelay.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInitDelay.HideSelection = True
        Me.txtInitDelay.ReadOnly = False
        Me.txtInitDelay.MaxLength = 0
        Me.txtInitDelay.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInitDelay.Multiline = False
        Me.txtInitDelay.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInitDelay.ScrollBars = System.Windows.Forms.ScrollBars.None
        Me.txtInitDelay.TabStop = True
        Me.txtInitDelay.Visible = True
        Me.txtInitDelay.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.txtInitDelay.Name = "txtInitDelay"
        Me._updnDelay_0.Size = New System.Drawing.Size(16, 20)
        Me._updnDelay_0.Location = New System.Drawing.Point(154, 30)
        Me._updnDelay_0.TabIndex = 29
        Me._updnDelay_0.Text = "_updnDelay_0"
        Me._updnDelay_0.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me._updnDelay_0.BackColor = System.Drawing.Color.Red
        Me._updnDelay_0.Name = "_updnDelay_0"
        Me._Label1_7.Text = "%"
        Me._Label1_7.Size = New System.Drawing.Size(30, 13)
        Me._Label1_7.Location = New System.Drawing.Point(175, 160)
        Me._Label1_7.TabIndex = 25
        Me._Label1_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_7.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._Label1_7.BackColor = System.Drawing.Color.Transparent
        Me._Label1_7.Enabled = True
        Me._Label1_7.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_7.UseMnemonic = True
        Me._Label1_7.Visible = True
        Me._Label1_7.AutoSize = True
        Me._Label1_7.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._Label1_7.Name = "_Label1_7"
        Me._Label1_8.Text = "&Slowed replay speed (as % of normal):"
        Me._Label1_8.Size = New System.Drawing.Size(98, 27)
        Me._Label1_8.Location = New System.Drawing.Point(12, 149)
        Me._Label1_8.TabIndex = 17
        Me._Label1_8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_8.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._Label1_8.BackColor = System.Drawing.Color.Transparent
        Me._Label1_8.Enabled = True
        Me._Label1_8.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_8.UseMnemonic = True
        Me._Label1_8.Visible = True
        Me._Label1_8.AutoSize = False
        Me._Label1_8.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._Label1_8.Name = "_Label1_8"
        Me._Label1_4.Text = "Sec."
        Me._Label1_4.Size = New System.Drawing.Size(26, 13)
        Me._Label1_4.Location = New System.Drawing.Point(175, 76)
        Me._Label1_4.TabIndex = 23
        Me._Label1_4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_4.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._Label1_4.BackColor = System.Drawing.Color.Transparent
        Me._Label1_4.Enabled = True
        Me._Label1_4.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_4.UseMnemonic = True
        Me._Label1_4.Visible = True
        Me._Label1_4.AutoSize = True
        Me._Label1_4.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._Label1_4.Name = "_Label1_4"
        Me._Label1_0.Text = "Sec."
        Me._Label1_0.Size = New System.Drawing.Size(22, 13)
        Me._Label1_0.Location = New System.Drawing.Point(175, 34)
        Me._Label1_0.TabIndex = 22
        Me._Label1_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_0.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._Label1_0.BackColor = System.Drawing.Color.Transparent
        Me._Label1_0.Enabled = True
        Me._Label1_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_0.UseMnemonic = True
        Me._Label1_0.Visible = True
        Me._Label1_0.AutoSize = True
        Me._Label1_0.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._Label1_0.Name = "_Label1_0"
        Me._Label1_3.Text = "Delay &between repeated playbacks:"
        Me._Label1_3.Size = New System.Drawing.Size(97, 26)
        Me._Label1_3.Location = New System.Drawing.Point(12, 65)
        Me._Label1_3.TabIndex = 13
        Me._Label1_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_3.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._Label1_3.BackColor = System.Drawing.Color.Transparent
        Me._Label1_3.Enabled = True
        Me._Label1_3.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_3.UseMnemonic = True
        Me._Label1_3.Visible = True
        Me._Label1_3.AutoSize = False
        Me._Label1_3.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._Label1_3.Name = "_Label1_3"
        Me._Label1_2.Text = "&Times to playback selected word(s):"
        Me._Label1_2.Size = New System.Drawing.Size(95, 26)
        Me._Label1_2.Location = New System.Drawing.Point(12, 107)
        Me._Label1_2.TabIndex = 15
        Me._Label1_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_2.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._Label1_2.BackColor = System.Drawing.Color.Transparent
        Me._Label1_2.Enabled = True
        Me._Label1_2.ForeColor = System.Drawing.SystemColors.ControlText
        Me._Label1_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_2.UseMnemonic = True
        Me._Label1_2.Visible = True
        Me._Label1_2.AutoSize = False
        Me._Label1_2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._Label1_2.Name = "_Label1_2"
        Me._Label1_1.BackColor = System.Drawing.Color.Transparent
        Me._Label1_1.Text = "&Delay before playback begins:"
        Me._Label1_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._Label1_1.Size = New System.Drawing.Size(84, 26)
        Me._Label1_1.Location = New System.Drawing.Point(12, 23)
        Me._Label1_1.TabIndex = 11
        Me._Label1_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._Label1_1.TextAlign = System.Drawing.ContentAlignment.TopLeft
        Me._Label1_1.Enabled = True
        Me._Label1_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._Label1_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._Label1_1.UseMnemonic = True
        Me._Label1_1.Visible = True
        Me._Label1_1.AutoSize = False
        Me._Label1_1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me._Label1_1.Name = "_Label1_1"
        Me.Controls.Add(_cmdOKCancel_1)
        Me.Controls.Add(_cmdOKCancel_0)
        Me.Controls.Add(_Frame1_1)
        Me.Controls.Add(_Frame1_0)
        Me.Controls.Add(_Frame1_2)
        Me._Frame1_1.Controls.Add(_cmdBrowse_1)
        Me._Frame1_1.Controls.Add(txtSASLoc)
        Me._Frame1_1.Controls.Add(_cmdBrowse_0)
        Me._Frame1_1.Controls.Add(txtSndsLoc)
        Me._Frame1_1.Controls.Add(_Label1_6)
        Me._Frame1_1.Controls.Add(_Label1_5)
        Me._Frame1_0.Controls.Add(cmdAdd)
        Me._Frame1_0.Controls.Add(cmdRemove)
        Me._Frame1_0.Controls.Add(cmdBrowseWL)
        Me._Frame1_0.Controls.Add(lvWLLocations)
        Me.lvWLLocations.Columns.Add(_lvWLLocations_ColumnHeader_1)
        Me.lvWLLocations.Columns.Add(_lvWLLocations_ColumnHeader_2)
        Me._Frame1_2.Controls.Add(updnSRSpeed)
        Me._Frame1_2.Controls.Add(_updnDelay_1)
        Me._Frame1_2.Controls.Add(updnPlaybackTimes)
        Me._Frame1_2.Controls.Add(txtSRSpeed)
        Me._Frame1_2.Controls.Add(txtRptDelay)
        Me._Frame1_2.Controls.Add(txtRptCount)
        Me._Frame1_2.Controls.Add(txtInitDelay)
        Me._Frame1_2.Controls.Add(_updnDelay_0)
        Me._Frame1_2.Controls.Add(_Label1_7)
        Me._Frame1_2.Controls.Add(_Label1_8)
        Me._Frame1_2.Controls.Add(_Label1_4)
        Me._Frame1_2.Controls.Add(_Label1_0)
        Me._Frame1_2.Controls.Add(_Label1_3)
        Me._Frame1_2.Controls.Add(_Label1_2)
        Me._Frame1_2.Controls.Add(_Label1_1)
        Me.Frame1.SetIndex(_Frame1_1, CType(1, Short))
        Me.Frame1.SetIndex(_Frame1_0, CType(0, Short))
        Me.Frame1.SetIndex(_Frame1_2, CType(2, Short))
        Me.Label1.SetIndex(_Label1_6, CType(6, Short))
        Me.Label1.SetIndex(_Label1_5, CType(5, Short))
        Me.Label1.SetIndex(_Label1_7, CType(7, Short))
        Me.Label1.SetIndex(_Label1_8, CType(8, Short))
        Me.Label1.SetIndex(_Label1_4, CType(4, Short))
        Me.Label1.SetIndex(_Label1_0, CType(0, Short))
        Me.Label1.SetIndex(_Label1_3, CType(3, Short))
        Me.Label1.SetIndex(_Label1_2, CType(2, Short))
        Me.Label1.SetIndex(_Label1_1, CType(1, Short))
        Me.cmdBrowse.SetIndex(_cmdBrowse_1, CType(1, Short))
        Me.cmdBrowse.SetIndex(_cmdBrowse_0, CType(0, Short))
        Me.cmdOKCancel.SetIndex(_cmdOKCancel_1, CType(1, Short))
        Me.cmdOKCancel.SetIndex(_cmdOKCancel_0, CType(0, Short))
        Me.updnDelay.SetIndex(_updnDelay_1, CType(1, Short))
        Me.updnDelay.SetIndex(_updnDelay_0, CType(0, Short))
        CType(Me.updnDelay, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdOKCancel, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdBrowse, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Frame1, System.ComponentModel.ISupportInitialize).EndInit()
        Me._Frame1_1.ResumeLayout(False)
        Me._Frame1_0.ResumeLayout(False)
        Me.lvWLLocations.ResumeLayout(False)
        Me._Frame1_2.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
#End Region 
End Class