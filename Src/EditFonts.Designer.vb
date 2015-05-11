<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmEditFonts
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
	Public dlgFont As System.Windows.Forms.FontDialog
	Public WithEvents cmdApply As System.Windows.Forms.Button
	Public WithEvents lblSample As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents cmdCancel As System.Windows.Forms.Button
	Public WithEvents cmdChange As System.Windows.Forms.Button
	Public WithEvents lstCols As System.Windows.Forms.ListBox
	Public WithEvents lblFontSpec As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEditFonts))
		Me.components = New System.ComponentModel.Container()
		Me.ToolTip1 = New System.Windows.Forms.ToolTip(components)
		Me.cmdOK = New System.Windows.Forms.Button
		Me.dlgFont = New System.Windows.Forms.FontDialog
		Me.cmdApply = New System.Windows.Forms.Button
		Me.Frame1 = New System.Windows.Forms.GroupBox
		Me.lblSample = New System.Windows.Forms.Label
		Me.cmdCancel = New System.Windows.Forms.Button
		Me.cmdChange = New System.Windows.Forms.Button
		Me.lstCols = New System.Windows.Forms.ListBox
		Me.lblFontSpec = New System.Windows.Forms.Label
		Me.Label1 = New System.Windows.Forms.Label
		Me.Frame1.SuspendLayout()
		Me.SuspendLayout()
		Me.ToolTip1.Active = True
		Me.Text = "Edit Fonts"
		Me.ClientSize = New System.Drawing.Size(315, 134)
		Me.Location = New System.Drawing.Point(4, 23)
		Me.ControlBox = False
		Me.MaximizeBox = False
		Me.MinimizeBox = False
		Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
		Me.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.BackColor = System.Drawing.SystemColors.Control
		Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
		Me.Enabled = True
		Me.KeyPreview = False
		Me.Cursor = System.Windows.Forms.Cursors.Default
		Me.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.ShowInTaskbar = True
		Me.HelpButton = False
		Me.WindowState = System.Windows.Forms.FormWindowState.Normal
		Me.Name = "frmEditFonts"
		Me.cmdOK.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdOK.Text = "OK"
		Me.AcceptButton = Me.cmdOK
		Me.cmdOK.Size = New System.Drawing.Size(65, 25)
		Me.cmdOK.Location = New System.Drawing.Point(108, 104)
		Me.cmdOK.TabIndex = 3
		Me.cmdOK.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdOK.BackColor = System.Drawing.SystemColors.Control
		Me.cmdOK.CausesValidation = True
		Me.cmdOK.Enabled = True
		Me.cmdOK.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdOK.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdOK.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdOK.TabStop = True
		Me.cmdOK.Name = "cmdOK"
		Me.cmdApply.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdApply.Text = "&Apply"
		Me.cmdApply.Enabled = False
		Me.cmdApply.Size = New System.Drawing.Size(65, 25)
		Me.cmdApply.Location = New System.Drawing.Point(244, 104)
		Me.cmdApply.TabIndex = 5
		Me.cmdApply.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdApply.BackColor = System.Drawing.SystemColors.Control
		Me.cmdApply.CausesValidation = True
		Me.cmdApply.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdApply.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdApply.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdApply.TabStop = True
		Me.cmdApply.Name = "cmdApply"
		Me.Frame1.Size = New System.Drawing.Size(201, 61)
		Me.Frame1.Location = New System.Drawing.Point(108, 16)
		Me.Frame1.TabIndex = 7
		Me.Frame1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Frame1.BackColor = System.Drawing.SystemColors.Control
		Me.Frame1.Enabled = True
		Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Frame1.Visible = True
		Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
		Me.Frame1.Name = "Frame1"
		Me.lblSample.TextAlign = System.Drawing.ContentAlignment.TopCenter
		Me.lblSample.Text = "Sample"
		Me.lblSample.Font = New System.Drawing.Font("Arial", 24!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSample.Size = New System.Drawing.Size(193, 45)
		Me.lblSample.Location = New System.Drawing.Point(4, 12)
		Me.lblSample.TabIndex = 8
		Me.lblSample.BackColor = System.Drawing.Color.Transparent
		Me.lblSample.Enabled = True
		Me.lblSample.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblSample.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblSample.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblSample.UseMnemonic = True
		Me.lblSample.Visible = True
		Me.lblSample.AutoSize = False
		Me.lblSample.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblSample.Name = "lblSample"
		Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.CancelButton = Me.cmdCancel
		Me.cmdCancel.Text = "Cancel"
		Me.cmdCancel.Size = New System.Drawing.Size(65, 25)
		Me.cmdCancel.Location = New System.Drawing.Point(176, 104)
		Me.cmdCancel.TabIndex = 4
		Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
		Me.cmdCancel.CausesValidation = True
		Me.cmdCancel.Enabled = True
		Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdCancel.TabStop = True
		Me.cmdCancel.Name = "cmdCancel"
		Me.cmdChange.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		Me.cmdChange.Text = "C&hange"
		Me.cmdChange.Size = New System.Drawing.Size(93, 25)
		Me.cmdChange.Location = New System.Drawing.Point(4, 104)
		Me.cmdChange.TabIndex = 2
		Me.cmdChange.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.cmdChange.BackColor = System.Drawing.SystemColors.Control
		Me.cmdChange.CausesValidation = True
		Me.cmdChange.Enabled = True
		Me.cmdChange.ForeColor = System.Drawing.SystemColors.ControlText
		Me.cmdChange.Cursor = System.Windows.Forms.Cursors.Default
		Me.cmdChange.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.cmdChange.TabStop = True
		Me.cmdChange.Name = "cmdChange"
		Me.lstCols.Size = New System.Drawing.Size(93, 75)
		Me.lstCols.IntegralHeight = False
		Me.lstCols.Location = New System.Drawing.Point(4, 24)
		Me.lstCols.TabIndex = 1
		Me.lstCols.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lstCols.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
		Me.lstCols.BackColor = System.Drawing.SystemColors.Window
		Me.lstCols.CausesValidation = True
		Me.lstCols.Enabled = True
		Me.lstCols.ForeColor = System.Drawing.SystemColors.WindowText
		Me.lstCols.Cursor = System.Windows.Forms.Cursors.Default
		Me.lstCols.SelectionMode = System.Windows.Forms.SelectionMode.One
		Me.lstCols.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lstCols.Sorted = False
		Me.lstCols.TabStop = True
		Me.lstCols.Visible = True
		Me.lstCols.MultiColumn = False
		Me.lstCols.Name = "lstCols"
		Me.lblFontSpec.Text = "Label3"
		Me.lblFontSpec.Size = New System.Drawing.Size(197, 21)
		Me.lblFontSpec.Location = New System.Drawing.Point(112, 80)
		Me.lblFontSpec.TabIndex = 6
		Me.lblFontSpec.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblFontSpec.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.lblFontSpec.BackColor = System.Drawing.Color.Transparent
		Me.lblFontSpec.Enabled = True
		Me.lblFontSpec.ForeColor = System.Drawing.SystemColors.ControlText
		Me.lblFontSpec.Cursor = System.Windows.Forms.Cursors.Default
		Me.lblFontSpec.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.lblFontSpec.UseMnemonic = True
		Me.lblFontSpec.Visible = True
		Me.lblFontSpec.AutoSize = False
		Me.lblFontSpec.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.lblFontSpec.Name = "lblFontSpec"
		Me.Label1.Text = "Word List &Column"
		Me.Label1.Size = New System.Drawing.Size(117, 17)
		Me.Label1.Location = New System.Drawing.Point(8, 8)
		Me.Label1.TabIndex = 0
		Me.Label1.Font = New System.Drawing.Font("Arial", 8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopLeft
		Me.Label1.BackColor = System.Drawing.Color.Transparent
		Me.Label1.Enabled = True
		Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
		Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
		Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
		Me.Label1.UseMnemonic = True
		Me.Label1.Visible = True
		Me.Label1.AutoSize = False
		Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.Label1.Name = "Label1"
		Me.Controls.Add(cmdOK)
		Me.Controls.Add(cmdApply)
		Me.Controls.Add(Frame1)
		Me.Controls.Add(cmdCancel)
		Me.Controls.Add(cmdChange)
		Me.Controls.Add(lstCols)
		Me.Controls.Add(lblFontSpec)
		Me.Controls.Add(Label1)
		Me.Frame1.Controls.Add(lblSample)
		Me.Frame1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()
	End Sub
#End Region 
End Class