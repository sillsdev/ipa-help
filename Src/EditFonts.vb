Option Strict Off
Option Explicit On
Friend Class frmEditFonts
	Inherits System.Windows.Forms.Form
	
	Private Structure ColFontInfoStruct
		Dim StyleNum As Short
		'UPGRADE_NOTE: Name was upgraded to Name_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Name_Renamed As String
		'UPGRADE_NOTE: Size was upgraded to Size_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
		Dim Size_Renamed As Short
		Dim Bold As Boolean
		Dim Italic As Boolean
	End Structure
	
	Dim bCanceled As Boolean
	Dim ColFontInfo() As ColFontInfoStruct
    Dim frmWL As frmWordList
	
	Public ReadOnly Property Canceled() As Boolean
		Get
            On Error Resume Next
			Canceled = bCanceled
        End Get
	End Property
	
	Public WriteOnly Property WordlistForm() As System.Windows.Forms.Form
		Set(ByVal Value As System.Windows.Forms.Form)
            On Error Resume Next
			frmWL = Value
        End Set
	End Property
	
    Private Function StyleNumber(ByRef sStyleName As String) As Short

        Dim i As Short

        On Error Resume Next

        With CType(frmWL.Controls("ssGrid"), Object)
            'UPGRADE_WARNING: Couldn't resolve default property of object frmWL!ssGrid.StyleSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            For i = 0 To .StyleSets.Count - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object frmWL!ssGrid.StyleSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If (.StyleSets(i).Name = sStyleName) Then
                    StyleNumber = i
                    Exit Function
                End If
            Next
        End With

        'UPGRADE_WARNING: Couldn't resolve default property of object StyleNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        StyleNumber = -1

    End Function
	
	Private Sub cmdApply_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdApply.Click
		
		Dim i As Short
        On Error Resume Next
		
        If (lstCols.SelectedIndex < 0) Then Exit Sub
        i = lstCols.SelectedIndex

        With CType(frmWL.Controls("ssGrid"), Object)
            For i = 0 To UBound(ColFontInfo)
                'UPGRADE_WARNING: Couldn't resolve default property of object frmWL!ssGrid.StyleSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                With .StyleSets(ColFontInfo(i).StyleNum).Font
                    'UPGRADE_WARNING: Couldn't resolve default property of object frmWL!ssGrid.StyleSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Name = ColFontInfo(i).Name_Renamed
                    'UPGRADE_WARNING: Couldn't resolve default property of object frmWL!ssGrid.StyleSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Size = ColFontInfo(i).Size_Renamed
                    'UPGRADE_WARNING: Couldn't resolve default property of object frmWL!ssGrid.StyleSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Bold = ColFontInfo(i).Bold
                    'UPGRADE_WARNING: Couldn't resolve default property of object frmWL!ssGrid.StyleSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Italic = ColFontInfo(i).Italic
                End With

                frmWL.ApplyPhoneticFontStyle()
                .Refresh()
                cmdOK.Focus()
                cmdApply.Enabled = False
            Next
        End With

	End Sub
	
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		
		On Error Resume Next
		bCanceled = True
		Visible = False
		
	End Sub
	
	Private Sub cmdChange_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdChange.Click
		
		Dim i As Short
        Dim dlg As New FontDialog()

        On Error Resume Next

        If (lstCols.SelectedIndex < 0) Then Exit Sub
        'dlg.Title = "Set " & VB6.GetItemString(lstCols, lstCols.SelectedIndex) & " Font"
        i = lstCols.SelectedIndex

        dlg.Font = VB6.FontChangeName(dlg.Font, ColFontInfo(i).Name_Renamed)
        dlg.Font = VB6.FontChangeSize(dlg.Font, ColFontInfo(i).Size_Renamed)
        dlg.Font = VB6.FontChangeBold(dlg.Font, ColFontInfo(i).Bold)
        dlg.Font = VB6.FontChangeItalic(dlg.Font, ColFontInfo(i).Italic)
        'UPGRADE_ISSUE: Constant cdlCFScreenFonts was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="55B59875-9A95-4B71-9D6A-7C294BF7139D"'
        'UPGRADE_ISSUE: MSComDlg.CommonDialog property dlg.Flags was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'dlgOpen.Flags = MSComDlg.FontsConstants.cdlCFScreenFonts
        On Error GoTo cmdChange_Canceled

        Dim result As System.Windows.Forms.DialogResult
        result = dlg.ShowDialog()
        If (result = System.Windows.Forms.DialogResult.Cancel) Then
            GoTo cmdChange_Canceled
        End If

        ColFontInfo(i).Name_Renamed = dlg.Font.Name
        ColFontInfo(i).Size_Renamed = dlg.Font.Size
        ColFontInfo(i).Bold = dlg.Font.Bold
        ColFontInfo(i).Italic = dlg.Font.Italic
        Call lstCols_SelectedIndexChanged(lstCols, New System.EventArgs())
        cmdApply.Enabled = True


cmdChange_Canceled:

    End Sub
	
	Private Sub cmdOK_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOK.Click
		
		On Error Resume Next
		If (cmdApply.Enabled) Then Call cmdApply_Click(cmdApply, New System.EventArgs())
		bCanceled = False
		Visible = False
		
	End Sub
	
	'UPGRADE_WARNING: Form event frmEditFonts.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmEditFonts_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		Dim i As Short
		Dim j As Short
		
		On Error Resume Next
		
		With CType(frmWL.Controls("ssGrid"), Object)
			'UPGRADE_WARNING: Couldn't resolve default property of object frmWL!ssGrid.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For i = 0 To .Columns.Count - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object frmWL!ssGrid.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If ((.Columns(i).Visible) And Not (.Columns(i).Caption = "WavFile")) Then
                    j = lstCols.Items.Add(.Columns(i).Caption)
                    ReDim Preserve ColFontInfo(j)
					'UPGRADE_WARNING: Couldn't resolve default property of object frmWL!ssGrid.Columns. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object StyleNumber(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ColFontInfo(j).StyleNum = StyleNumber(.Columns(i).StyleSet)
					
					'UPGRADE_WARNING: Couldn't resolve default property of object frmWL!ssGrid.StyleSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					With .StyleSets(ColFontInfo(j).StyleNum).Font
						'UPGRADE_WARNING: Couldn't resolve default property of object frmWL!ssGrid.StyleSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ColFontInfo(j).Name_Renamed = .Name
						'UPGRADE_WARNING: Couldn't resolve default property of object frmWL!ssGrid.StyleSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ColFontInfo(j).Size_Renamed = .Size
						'UPGRADE_WARNING: Couldn't resolve default property of object frmWL!ssGrid.StyleSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ColFontInfo(j).Bold = .Bold
						'UPGRADE_WARNING: Couldn't resolve default property of object frmWL!ssGrid.StyleSets. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						ColFontInfo(j).Italic = .Italic
					End With
				End If
			Next 
		End With
		
		lstCols.SelectedIndex = 0
		
	End Sub
	
	Private Sub frmEditFonts_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		On Error Resume Next
		Call CenterForm(Me, True)
		
	End Sub
	
	Private Sub frmEditFonts_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel

		On Error Resume Next
        If eventArgs.CloseReason <> CloseReason.UserClosing Then Exit Sub
        Cancel = True
        Call cmdOK_Click(cmdOK, New System.EventArgs())

        eventArgs.Cancel = Cancel
	End Sub
	
	'UPGRADE_WARNING: Event lstCols.SelectedIndexChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub lstCols_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstCols.SelectedIndexChanged
		
		Dim i As Short
		
		On Error Resume Next
		
		With lstCols
			If (.SelectedIndex < 0) Then Exit Sub
			i = .SelectedIndex
		End With
		
		With lblSample
			.Font = VB6.FontChangeName(.Font, ColFontInfo(i).Name_Renamed)
			.Font = VB6.FontChangeBold(.Font, ColFontInfo(i).Bold)
			.Font = VB6.FontChangeItalic(.Font, ColFontInfo(i).Italic)
		End With
		
		With ColFontInfo(i)
			lblFontSpec.Text = .Name_Renamed & ", " & System.Math.Round(.Size_Renamed) & " pt." & IIf(.Bold, ", bold", "") & IIf(.Italic, ", italic", "")
		End With
		
	End Sub
End Class