Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmWordList
	Inherits System.Windows.Forms.Form
	'**************************************************
	'* frmWordList version info:
	'*  See basGlobals (Globals.bas).
	'**************************************************
	
	
	Private bGridDoubleClicked As Boolean
	Private bImportingCategory As Boolean
	Private bInEditMode As Boolean
	Private iMaxRowHeight As Short
	Private iMinPhonColWidth As Short
	Private iMaxPhonColWidth As Short
	Private iTimesPlayed As Short
	Private sWavFile As String
	Private sWordListIndex As String '* Index of Word List.
	Private sWordListCaption As String '* Name of Word List on main menu
	Private sWordListPath As String '* Path to Word List
	Private sWordListSndPath As String '* Path to Wave Files
	Private WLPlayer As clsWordListPlayer
	Private xmlWL As clsXMLWordList
	Private LastTreeNode As System.Windows.Forms.TreeNode
	
	Private bSplitterDrag As Boolean
	Private sngMouseXOffset As Single
	
	Private Const PitchPlotShownEntry As String = "PitchPlotHasBeenShown"
	Private Const TempWLXMLBackup As String = "~wlXMLBackup.tmp"
	Private Const NewCategoryName As String = "New Category"
	Private Const EditModeTBarButtons As String = "Exit;"
	Private Const NoWaveEnalbedTBarButtons As String = "Test;Exit;"
	Private Const TBarButtons As String = "PlayOnly;PlaySlow;PlaySeparator;Record;StopRec;" & "PlayRec;PlayRecSpeaker;RecSeparator;" & "Pitch;PitchSeparator;" & "Test;TestSeparator;Exit;"
	
	Public Sub AddNewCategory(ByRef Index As Short)
		
		Dim sNewCategoryName As String
		
		On Error Resume Next
		
		With tvSections
			.Focus()
			
			sNewCategoryName = GetNewCategoryName()
			
			If (.Nodes.Count = 0) Then
				.Nodes.Add(sNewCategoryName)
				'UPGRADE_WARNING: Lower bound of collection tvSections.Nodes has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
				.SelectedNode = .Nodes.Item(1)
			ElseIf (Index = 0) Then 
				'UPGRADE_WARNING: Cannot determine Node location Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="196D987F-2118-46D0-80D2-92FB2909C206"'
				.Nodes.Insert(.SelectedNode.Index, sNewCategoryName)
				.SelectedNode = .SelectedNode.PrevNode
			Else
				'UPGRADE_WARNING: Cannot determine Node location Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="196D987F-2118-46D0-80D2-92FB2909C206"'
				.Nodes.Insert(.SelectedNode.Index, sNewCategoryName)
				.SelectedNode = .SelectedNode.NextNode
			End If
			
			'*********************************************************
			'* If we just added the tree node to the end of the tree
			'* then add the new word list at the end of the XML file.
			'* Otherwise, specify what list to insert the new list
			'* before.
			'*********************************************************
			If (.SelectedNode.NextNode Is Nothing) Then
				xmlWL.AddCategory(sNewCategoryName)
			Else
				xmlWL.AddCategory(sNewCategoryName,  , .SelectedNode.NextNode.Text)
			End If
			
			'*********************************************************
			'* Save the xml file changes and put the user in the
			'* edit mode to encourage them to give the new word list
			'* section a meaningful name.
			'*********************************************************
			xmlWL.Save()
			Call tvSections_NodeClick(tvSections, New System.Windows.Forms.TreeNodeMouseClickEventArgs(.SelectedNode, System.Windows.Forms.MouseButtons.None, 0, 0, 0))
			.SelectedNode.BeginEdit()
		End With
		
	End Sub
	
	Private Sub AdjustControlPlacement()
		
		If Not (bSplitterDrag) Then
			With picSplitter
				.Height = ClientRectangle.Height
				picDragBar.Height = ClientRectangle.Height
			End With
		End If
		
		With picDragBar
			tvSections.SetBounds(VB6.TwipsToPixelsX(-30), VB6.TwipsToPixelsY(-30), VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(.Left) + (gPixelX * 4)), VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(ClientRectangle.Height) + (gPixelY * 4)))
			ssGrid.SetBounds(VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(.Left) + VB6.PixelsToTwipsX(.Width) + gPixelX), 0, VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(ClientRectangle.Width) - (VB6.PixelsToTwipsX(.Left) + VB6.PixelsToTwipsX(.Width) + gPixelX)), ClientRectangle.Height)
		End With
		
		With ssGrid
			.Redraw = False
			Call ResizeGrid()
			.Redraw = True
			picTestMode.SetBounds(.Left, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(.Top) + (gPixelY * 18)), .Width, VB6.TwipsToPixelsY(VB6.PixelsToTwipsY(.Height) - (gPixelY * 18)))
		End With
		
		With picTestMode
			lblTestMode.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(.Width) - VB6.PixelsToTwipsX(lblTestMode.Width)) \ 2)
			txtUserTr.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(.Width) - (gPixelX * 10))
			txtCorrectTr.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(.Width) - (gPixelX * 10))
		End With
		
		With mdiHelpCharts
			CType(.Controls("txtEditModeIndicator"), Object).Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(.Width) - VB6.PixelsToTwipsX(CType(.Controls("txtEditModeIndicator"), Object).Width) - (gPixelX * 10))
		End With
		
	End Sub
	
	Private Sub AdjustGridRowHeightToFitFonts(Optional ByRef vForcedHasPitch As Object = Nothing)
		
		'****************************************************************
		'* This routine will attempt to adjust the grid's row height to
		'* accomodate the largest font in the grid.
		'****************************************************************
        Dim i As Object
		Dim j As Short
		Dim iHeight As Short
		
		On Error Resume Next
		
		iMaxRowHeight = 0
		
		Dim bHasPitch As Boolean
        For i = 0 To ssGrid.Cols - 1
            If (Len(ssGrid.Columns(i).StyleSet) > 0) Then
                Font = VB6.FontChangeName(Font, ssGrid.StyleSets(ssGrid.Columns(i).StyleSet).Font.Name)
                Font = VB6.FontChangeSize(Font, ssGrid.StyleSets(ssGrid.Columns(i).StyleSet).Font.Size)
                Font = VB6.FontChangeBold(Font, ssGrid.StyleSets(ssGrid.Columns(i).StyleSet).Font.Bold)
                Font = VB6.FontChangeItalic(Font, ssGrid.StyleSets(ssGrid.Columns(i).StyleSet).Font.Italic)
            Else
                Font = VB6.FontChangeName(Font, ssGrid.Font.Name)
                Font = VB6.FontChangeSize(Font, ssGrid.Font.SizeInPoints)
                Font = VB6.FontChangeBold(Font, ssGrid.Font.Bold)
                Font = VB6.FontChangeItalic(Font, ssGrid.Font.Italic)
            End If

            Dim graphics As Graphics
            graphics = Me.CreateGraphics
            iHeight = graphics.MeasureString("X", Font).Height

            If (iHeight > iMaxRowHeight) Then iMaxRowHeight = iHeight
        Next

        bHasPitch = DoesGridHavePitchInfo()
        If Not (IsNothing(vForcedHasPitch)) Then bHasPitch = vForcedHasPitch
        ssGrid.RowHeight = (iMaxRowHeight * IIf(bHasPitch, 2, 1)) + (gPixelY * 3)

	End Sub
	
	Public Sub ApplyPhoneticFontStyle()
		
		On Error Resume Next
		
		With ssGrid
			'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets(IPAAudio).Font.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.StyleSets("IPAAudio").Font.Name = .StyleSets("IPA").Font.Name
			'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets(IPAAudio).Font.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.StyleSets("IPAAudio").Font.Size = .StyleSets("IPA").Font.Size
			'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets(IPAAudio).Font.Bold. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Bold. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.StyleSets("IPAAudio").Font.Bold = .StyleSets("IPA").Font.Bold
			'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets(IPAAudio).Font.Italic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Italic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			.StyleSets("IPAAudio").Font.Italic = .StyleSets("IPA").Font.Italic
		End With
		
	End Sub
	
	Public Sub DeleteCategory()
		
		On Error Resume Next
		
		With tvSections
			If (MsgBox("Are you sure you want to delete " & .SelectedNode.Text & "?", MsgBoxStyle.YesNo + MsgBoxStyle.Question, My.Application.Info.Title) = MsgBoxResult.No) Then Exit Sub
			
			If (xmlWL.RemoveCategory(.SelectedNode.Text)) Then
				xmlWL.Save()
				.Focus()
				'UPGRADE_WARNING: MSComctlLib.Nodes method tvSections.Nodes.Remove has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
				.Nodes.RemoveAt(.SelectedNode.Index)
				Call tvSections_NodeClick(tvSections, New System.Windows.Forms.TreeNodeMouseClickEventArgs(.SelectedNode, System.Windows.Forms.MouseButtons.None, 0, 0, 0))
			End If
		End With
		
	End Sub
	
	Private Function DoesGridHavePitchInfo() As Boolean
		
		Dim i As Short
		
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

        Dim frm As New frmEditFonts
        frm.WordlistForm = Me
        frm.ShowDialog()
        If Not (frm.Canceled) Then Call SaveColumnStyles()
        frm.Close()

	End Sub
	
	Public Sub EditMode()
		
		If Not (bInEditMode) Then
			If Not (UserReallyWantsEditMode()) Then Exit Sub
		End If
		
		Call ManageGridColumns(True)
		bInEditMode = Not bInEditMode
		
		Dim i As Short
		
		With mdiHelpCharts
			CType(.Controls("txtEditModeIndicator"), Object).Visible = bInEditMode
			.mnuFile.Enabled = Not bInEditMode
			.mnuTest.Enabled = Not bInEditMode
			.mnuWindow.Enabled = Not bInEditMode
			.mnuEditFonts.Enabled = bInEditMode
			.mnuEditSoundPath.Enabled = bInEditMode
			.mnuEditTitle.Enabled = bInEditMode
			.mnuAddNewCategory.Enabled = bInEditMode
			.mnuDeleteCategory.Enabled = bInEditMode
			.mnuEditMode.Text = "&" & IIf(bInEditMode, "Exit", "Enter") & " Edit Mode"
		End With
		
        '******************************************************
        '* First, hide/unhide and lock/unlock the appropriate
        '* columns for going from or to the edit mode.
        '******************************************************
        For i = 0 To ssGrid.Columns.Count - 1
            ssGrid.Columns(i).Locked = Not bInEditMode
        Next

        ssGrid.Columns("WavFile").Visible = bInEditMode
        ssGrid.Columns("WavFile").Locked = True
        ssGrid.Columns("EditWavButton").Visible = bInEditMode
        ssGrid.AllowUpdate = bInEditMode
        ssGrid.AllowAddNew = bInEditMode
        ssGrid.AllowDelete = bInEditMode
        If (bInEditMode) Then
            'KG-TODO ssGrid.SelectTypeRow = SSDataWidgets_B.Constants_SelectionType.ssSelectionTypeSingleSelect
        Else
            'KG-TODO ssGrid.SelectTypeRow = SSDataWidgets_B.Constants_SelectionType.ssSelectionTypeMultiSelect
        End If
        Call ManageGridColumns(False)

        tvSections.LabelEdit = IIf(bInEditMode, True, False)

        If (bInEditMode) Then
            With mdiHelpCharts
                Call .ShowTBarButtons(TBarButtons & EditModeTBarButtons)
                Call .EnableTBarButtons(EditModeTBarButtons)
            End With
        Else
            Call UpdateXMLFileAfterEdits()
            Call UserReallyWantsToKeepEdits()

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
			.ShowDialog()
			If (Len(.sDir) > 0) Then
				sWordListSndPath = .sDir
				xmlWL.SoundPath = .sDir
				xmlWL.Save()
			End If
		End With
		
		frmFilePath.Close()
		
	End Sub
	
	Public Sub EditTitle()
		
		On Error Resume Next

        Dim frm As New frmRenameWLTitle
        frmRenameWLTitle.Title = sWordListCaption
        frmRenameWLTitle.ShowDialog()
        If Not (frmRenameWLTitle.Canceled) Then
            sWordListCaption = frmRenameWLTitle.Title
            Text = sWordListCaption & " Word List"
            xmlWL.ID = sWordListCaption
            xmlWL.Save()
        End If
        frm.Close()
		
	End Sub
	
	Public Function GetGridData(ByRef Column As String) As String
		
		On Error Resume Next
		GetGridData = ssGrid.Columns(Column).Text
		
	End Function
	
	Private Function GetNewCategoryName() As String
		
		'***********************************************************
		'* This routine will find a new name for a word list by
		'* tacking on numbers to a constant string.
		'***********************************************************
		
		Dim i As Short
		Dim sName As String
		
		i = 0
		
		Do 
			sName = NewCategoryName & IIf(i = 0, "", " " & i)
			i = i + 1
		Loop While (CategoryNameExists(sName,  , False))
		
		GetNewCategoryName = sName
		
	End Function
	
	Private Function GetSelectedRowsWavFile(Optional ByRef vStartAtTop As Object = Nothing, Optional ByRef vMoveBookmark As Object = Nothing) As String
		
		Dim bStartAtTop As Boolean
		Dim bMoveBookmark As Boolean
		Dim i As Object
		Dim j As Short
		Static iStartRow As Short
		Dim sWavFile As String
		
		On Error Resume Next
		
		GetSelectedRowsWavFile = ""
		sWavFile = ""
		bStartAtTop = False
		bMoveBookmark = False
        If Not (IsNothing(vStartAtTop)) Then bStartAtTop = vStartAtTop
        If Not (IsNothing(vMoveBookmark)) Then bMoveBookmark = vMoveBookmark
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
		'UPGRADE_WARNING: Couldn't resolve default property of object MakeFullPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Len(sWavFile) > 0) Then sWavFile = MakeFullPath(sWordListSndPath, sWavFile)
		
		If (Not FileExist(sWavFile)) Then
			iStartRow = 0
			Exit Function
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iStartRow = i + 1
		GetSelectedRowsWavFile = sWavFile
		
	End Function
	
	Private Function GetSelectedRowsWavFileForPitch(Optional ByRef vStartAtBottom As Object = Nothing, Optional ByRef vMoveBookmark As Object = Nothing) As String
		
		Dim bStartAtBottom As Boolean
		Dim bMoveBookmark As Boolean
		Dim i As Object
		Dim j As Short
		Static iStartRow As Short
		Dim sWavFile As String
		
		On Error Resume Next
		
		GetSelectedRowsWavFileForPitch = ""
		sWavFile = ""
		bStartAtBottom = False
		bMoveBookmark = False
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		'UPGRADE_WARNING: Couldn't resolve default property of object vStartAtBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Not (IsNothing(vStartAtBottom)) Then bStartAtBottom = vStartAtBottom
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		'UPGRADE_WARNING: Couldn't resolve default property of object vMoveBookmark. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Not (IsNothing(vMoveBookmark)) Then bMoveBookmark = vMoveBookmark
		If (bStartAtBottom) Then iStartRow = ssGrid.Rows - 1
		
		With ssGrid
			If (.SelBookmarks.Count = 0) Then
				sWavFile = .Columns("WavFile").Text
				GoTo GetSelectedRowsWavFileForPitchEnd
			End If
			
			For i = iStartRow To 0 Step -1
				For j = 0 To .SelBookmarks.Count - 1
					'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.SelBookmarks(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.AddItemBookmark(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					If (CStr(.AddItemBookmark(i)) = CStr(.SelBookmarks(j))) Then
						'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						sWavFile = .Columns("WavFile").CellText(.AddItemBookmark(i))
						'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.AddItemBookmark(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
						If (bMoveBookmark) Then .Bookmark = .AddItemBookmark(i)
						GoTo GetSelectedRowsWavFileForPitchEnd
					End If
				Next 
			Next 
		End With
		
		'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
GetSelectedRowsWavFileForPitchEnd: 
		'UPGRADE_WARNING: Couldn't resolve default property of object MakeFullPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If (Len(sWavFile) > 0) Then sWavFile = MakeFullPath(sWordListSndPath, sWavFile)
		
		If (Not FileExist(sWavFile)) Then
			iStartRow = ssGrid.Rows - 1
			Exit Function
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iStartRow = i - 1
		GetSelectedRowsWavFileForPitch = sWavFile
		
	End Function
	
	Private Sub ImportCategory(ByRef sCategory As String)
		
		Dim i As Short
		Dim sListText As String
		Dim vList As Object
		
		On Error Resume Next
		
		If (Not WLPlayer Is Nothing) Then WLPlayer.Cancel = True
        'KG TODO Call ssGrid_RowColChange(ssGrid, New SSDataWidgets_B._DSSDBGridEvents_RowColChangeEvent(System.DBNull.Value, 0))
        vList = xmlWL.WordsInCategory(sCategory)
		
		bImportingCategory = True
		
		'************************************************************
		'* Add each item in the array. An entire row is added in one
		'* call to .AddItem. Column items are separated by vbTab.
		'************************************************************
		With ssGrid
			.Redraw = False
			.RemoveAll()
			
			'************************************************************
			'* Add each item in the array. An entire row is added in one
			'* call to .AddItem. Column items are separated by vbTab.
			'************************************************************
			'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
			If Not (IsDbNull(vList)) Then
				For i = 0 To UBound(vList, 1)
					'UPGRADE_WARNING: Couldn't resolve default property of object vList(i, 6). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object vList(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object vList(i, 4). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object vList(i, 3). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object vList(i, 2). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object vList(i, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object vList(i, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					.AddItem(IIf(Len(vList(i, 0)) > 0, vList(i, 0) & vbCrLf, "") & vList(i, 1) & vbTab & vList(i, 2) & vbTab & vList(i, 3) & vbTab & vList(i, 4) & vbTab & LCase(vList(i, 5)) & vbTab & vList(i, 6))
				Next 
			End If
			
			.Redraw = True
		End With
		
		bImportingCategory = False
		
	End Sub
	
	Public Sub Initialize(ByRef FormTag As String)
		
		Dim iIndex As Short
		
		On Error Resume Next
		
		iIndex = CShort(FormTag)
		sWordListIndex = FormTag
		'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(iIndex, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sWordListCaption = gWordListID(iIndex, 0)
		'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(iIndex, 1). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		sWordListPath = gWordListID(iIndex, 1)
		
	End Sub
	
	Private Sub InitializeView()
		
		Dim sCategoryNames() As String
		Dim i As Short
		
		On Error Resume Next
		
		'**************************************************************
		'* Load xml word list file
		'**************************************************************
		xmlWL = New clsXMLWordList
		xmlWL.Load(sWordListPath)
		
		'**************************************************************
		'* Get the path for the wave files
		'**************************************************************
		sWordListSndPath = xmlWL.SoundPath
		sWordListSndPath = sWordListSndPath & IIf(VB.Right(sWordListSndPath, 1) = "\", "", "\")
		
		'**************************************************************
		'* Get the list of word list names and fill the word list
		'* list box.
		'**************************************************************
		sCategoryNames = VB6.CopyArray(xmlWL.CategoryNames)
		
		'**************************************************************
		'* Fill tree with list names and call the node click method
		'* to fill the grid with the data for the selected tree node.
		'**************************************************************
        '************************************************************
        '* Clear the tree of all nodes.
        '************************************************************
        For i = 1 To tvSections.Nodes.Count
            'UPGRADE_WARNING: MSComctlLib.Nodes method tvSections.Nodes.Remove has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
            tvSections.Nodes.RemoveAt(1)
        Next

        '************************************************************
        '* Add all the list names (i.e. sections names)
        '************************************************************
        For i = 0 To UBound(sCategoryNames)
            tvSections.Nodes.Add(sCategoryNames(i))
        Next i

        '************************************************************
        '* Select the first item in the list call the node click
        '* event to force the filling of the list's grid.
        '************************************************************
        tvSections.SelectedNode = tvSections.Nodes.Item(0).Parent.FirstNode
        LastTreeNode = Nothing
        Call tvSections_NodeClick(tvSections, New System.Windows.Forms.TreeNodeMouseClickEventArgs(tvSections.SelectedNode, System.Windows.Forms.MouseButtons.None, 0, 0, 0))
        tvSections.Focus()

        Call SetColumnStyles()
        Call AdjustGridRowHeightToFitFonts()

	End Sub
	
	Private Function KeyHandler(ByRef KeyCode As Short, ByRef KeyAscii As Short, ByRef Shift As Short) As Short
		
		On Error Resume Next
		
		If (bInEditMode) Then
			'******************************************************
			'* If the user is in the edit mode then make sure they
			'* aren't allowed to change to a different window
			'******************************************************
			If (KeyCode = System.Windows.Forms.Keys.Tab And (Shift And VB6.ShiftConstants.CtrlMask)) Then
				KeyCode = 0
				Shift = 0
			ElseIf (KeyCode <> System.Windows.Forms.Keys.Return) Then 
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
			Case System.Windows.Forms.Keys.Return : Call Play(cVocBttn)
			Case Else
		End Select
		
		Select Case KeyAscii
			Case System.Windows.Forms.Keys.Escape : Me.Close()
			Case Else
		End Select
		
	End Function
	
	Private Function CategoryNameExists(ByRef sName As String, Optional ByRef vNode As Object = Nothing, Optional ByRef vShowMsg As Object = Nothing) As Boolean
		
		'**********************************************************
		'* This routine determines whether or not a name exists in
		'* the word list. If a node is supplied, that node will be
		'* exempted from the check.
		'**********************************************************
		
		Dim bShowMsg As Boolean
		Dim ExemptNode As System.Windows.Forms.TreeNode
		Dim node As System.Windows.Forms.TreeNode
		
		On Error Resume Next
		
		CategoryNameExists = False
		bShowMsg = True
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		'UPGRADE_WARNING: Couldn't resolve default property of object vShowMsg. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Not (IsNothing(vShowMsg)) Then bShowMsg = vShowMsg
		
		'UPGRADE_NOTE: Object ExemptNode may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		ExemptNode = Nothing
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		If Not (IsNothing(vNode)) Then ExemptNode = vNode
		
		With tvSections
			For	Each node In .Nodes
				If Not (node Is ExemptNode) Then
					If (StrComp(sName, node.Text, CompareMethod.Text) = 0) Then
						If (bShowMsg) Then MsgBox("'" & sName & "' already exists.", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, My.Application.Info.Title)
						CategoryNameExists = True
						Exit Function
					End If
				End If
			Next node
		End With
		
	End Function
	
	Private Sub ManageGridColumns(ByRef bSave As Boolean)
		
		Dim i As Short
		Dim sVal As String
		Dim sEntry As String
		
		sEntry = IIf(bInEditMode, cWLEditModeColsEntry, cWLColsEntry)
		
		For i = 0 To ssGrid.Cols - 1
			If (bSave) Then
				Call WriteINIEntry(cSettingsSect, sEntry & i, CStr(ssGrid.Columns(i).Width), gINIPath)
			Else
				sVal = GetINIEntry(cSettingsSect, sEntry & i, gINIPath)
				If (Val(sVal) > 0) Then ssGrid.Columns(i).Width = Val(sVal)
			End If
		Next 
		
	End Sub
	
	Public Sub Play(ByRef iButton As Short)
		
		'**************************************************
		'* This function allows mdiHelpCharts to play the
		'* wave file corresponding to the currently
		'* selected symbol, without knowing which form is
		'* active.
		'**************************************************
		
		Dim i As Short
		Dim iRptCount As Short
		Dim sSavBkMrk As String
		Dim sWavFile As String
		
		On Error Resume Next
		
		If (Len(gMMCtrl.Tag) > 0) Then Exit Sub
		If Not (WLPlayer Is Nothing) Then WLPlayer.Cancel = True
		
		'************************************************************************
		'* This code gets executed when the user clicks on a row in the grid.
		'************************************************************************
		If (iButton = -1) Then
			'UPGRADE_WARNING: Couldn't resolve default property of object MakeFullPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sWavFile = MakeFullPath(sWordListSndPath, (ssGrid.Columns("WavFile").Text))
			If Not (FileExist(sWavFile)) Then Exit Sub
			WLPlayer = New clsWordListPlayer
			Call Pause(CDbl(GetINIEntry(cSettingsSect, cPlayInitDelayEntry, gINIPath)))
			
			iRptCount = CShort(GetINIEntry(cSettingsSect, cRepeatCountEntry, gINIPath))
			gMMCtrl.Wait = True
			gMMCtrl.Tag = cMCIBusy
			For i = 1 To iRptCount
				WLPlayer.Play(sWavFile)
				If (iRptCount <= 1 Or WLPlayer.Cancel) Then
					Exit For
				End If
				Call Pause(CDbl(GetINIEntry(cSettingsSect, cPlayRepeatDelayEntry, gINIPath)))
			Next 
			gMMCtrl.Tag = ""
		Else
			'**********************************************************************
			'* This code gets executed when the user clicks on the playback button.
			'**********************************************************************
			'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.Bookmark. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sSavBkMrk = ssGrid.Bookmark
			sWavFile = GetSelectedRowsWavFile(True, True)
			iRptCount = CShort(GetINIEntry(cSettingsSect, cRepeatCountEntry, gINIPath))
			If (Len(sWavFile) = 0) Then Exit Sub
			WLPlayer = New clsWordListPlayer
			Call Pause(CDbl(GetINIEntry(cSettingsSect, cPlayInitDelayEntry, gINIPath)))
			
			gMMCtrl.Tag = cMCIBusy
			gMMCtrl.Wait = True
			
			'    Do
			'      WLPlayer.Play sWavFile
			'      sWavFile = GetSelectedRowsWavFile(, True)
			'      If (Len(sWavFile) = 0 Or WLPlayer.Cancel Or ssGrid.SelBookmarks.Count = 0) Then Exit Do
			'      Call Pause(GetINIEntry$(cSettingsSect, cPlayRepeatDelayEntry, gINIPath))
			'    Loop
			For i = 1 To iRptCount
				WLPlayer.Play(sWavFile)
				sWavFile = GetSelectedRowsWavFile( , True)
				If (Len(sWavFile) = 0 Or WLPlayer.Cancel Or iRptCount <= 1) Then Exit For ' Or ssGrid.SelBookmarks.Count = 0  used to be in this as well.
				Call Pause(CDbl(GetINIEntry(cSettingsSect, cPlayRepeatDelayEntry, gINIPath)))
			Next 
			
			gMMCtrl.Tag = ""
			ssGrid.Bookmark = sSavBkMrk
		End If
		
		'UPGRADE_NOTE: Object WLPlayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		WLPlayer = Nothing
		
	End Sub
	
	Public Sub PlaySlow()
		
		Dim i As Short
		Dim sSavBkMrk As String
		Dim sWavFile As String
		
		On Error Resume Next
		
		Kill(gListFilePath)
		
		With ssGrid
			If (Len(.Columns("WavFile").Text) = 0) Then Exit Sub
			.SelBookmarks.RemoveAll()
			'UPGRADE_WARNING: Couldn't resolve default property of object MakeFullPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			sWavFile = MakeFullPath(sWordListSndPath, (.Columns("WavFile").Text))
			If Not (FileExist(sWavFile)) Then Exit Sub
		End With
		
		Call WriteINIEntry("Settings", "CallingApp", mdiHelpCharts.Text, gListFilePath)
		Call WriteINIEntry("Settings", "ShowWindow", "Hide", gListFilePath)
		Call WriteINIEntry("Settings", "CallingApp", mdiHelpCharts.Text, gListFilePath)
		Call WriteINIEntry("AudioFiles", "File0", sWavFile, gListFilePath)
		Call WriteINIEntry("Commands", "command0", "SelectFile(0)", gListFilePath)
		Call WriteINIEntry("Commands", "command1", "Play(" & gSRSpeed & ",50,,)", gListFilePath)
		Call WriteINIEntry("Commands", "command2", "Return(2)", gListFilePath)
		Call mdiHelpCharts.CallSA()
		ActiveControl.Focus()
		
	End Sub
	
    Private Sub ResizeGrid(Optional ByRef vColIndex As Object = Nothing, Optional ByRef vColWidth As Object = Nothing)

        Dim i As Short
        Dim iTextWidth As Short
        Dim iTextHeight As Short
        Dim iColIndex As Short
        Dim iRowIndex As Short
        Dim iGridWidth As Short
        Dim iPhonColWidth As Short
        Dim iMaxEticWidth As Short
        Dim iVisCols As Object
        Dim sCellText As String
        Dim sEticText As String
        Dim sPitchText As String
        Dim iColStart As Short
        Dim iColEnd As Short
        Dim iRowHeight As Short
        Dim iTest As Short

        On Error Resume Next

        If (Val(GetINIEntry(cSettingsSect, cWLColsEntry & "0", gINIPath)) <> 0) Then Exit Sub

        With ssGrid
            'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
            If Not IsNothing(vColIndex) Then If .Columns(vColIndex).Caption <> "Phonetic" Then Exit Sub

            '* column width
            'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
            If IsNothing(vColIndex) Then
                '************************************************************
                '* find width from record selector to scrollbar
                '************************************************************
                iGridWidth = VB6.PixelsToTwipsX(.Width) - (gPixelX * 40)

                '************************************************************
                '* find number of visible columns
                '************************************************************
                'UPGRADE_WARNING: Couldn't resolve default property of object iVisCols. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                iVisCols = 0
                For iColIndex = 0 To .Columns.Count - 1
                    'UPGRADE_WARNING: Couldn't resolve default property of object iVisCols. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    If .Columns(iColIndex).Visible Then iVisCols = iVisCols + 1
                Next

                '************************************************************
                '* find longest phonetic text
                '************************************************************
                iColIndex = .Columns("Phonetic").Position
                Call SetFormFontToColFont((.Columns("Phonetic").Position))

                'KG-TODO-specify font being used in display
                Dim font As New Font("Arial", 16)

                '************************************************************
                '* set default column widths
                '************************************************************
                'UPGRADE_WARNING: Couldn't resolve default property of object iVisCols. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                iPhonColWidth = iGridWidth / iVisCols
                For iRowIndex = 0 To .Rows - 1
                    sPitchText = ""
                    sEticText = .Columns("Phonetic").CellText(.AddItemBookmark(iRowIndex))
                    i = InStr(sEticText, vbCrLf)

                    Dim graphics As Graphics
                    graphics = Me.CreateGraphics

                    If (i > 0) Then
                        sPitchText = VB.Left(sEticText, i - 1)
                        sEticText = Mid(sEticText, i + Len(vbCrLf))
                        iTextWidth = graphics.MeasureString(sPitchText, font).Width
                        If (graphics.MeasureString(sEticText, font).Width > iTextWidth) Then iTextWidth = graphics.MeasureString(sEticText, font).Width + (gPixelX * 6)
                    Else
                        iTextWidth = graphics.MeasureString(sEticText, font).Width + (gPixelX * 6)
                    End If

                    If (iPhonColWidth < iTextWidth) Then iPhonColWidth = iTextWidth
                Next

                '************************************************************
                '* set column widths
                '************************************************************
                iMaxEticWidth = iPhonColWidth
                Select Case iPhonColWidth
                    Case Is < iMinPhonColWidth : iPhonColWidth = iMinPhonColWidth
                    Case Is > iMaxPhonColWidth : iPhonColWidth = iMaxPhonColWidth
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

            Call AdjustGridRowHeightToFitFonts()

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
            xmlWL.PhoneticFontName = ssGrid.StyleSets("IPA").Font.Name
            xmlWL.PhoneticFontSize = ssGrid.StyleSets("IPA").Font.Size
            xmlWL.PhoneticFontBold = ssGrid.StyleSets("IPA").Font.Bold
            xmlWL.PhoneticFontItalic = ssGrid.StyleSets("IPA").Font.Italic
			
            xmlWL.OrthoFontName = ssGrid.StyleSets("Ortho").Font.Name
            xmlWL.OrthoFontSize = ssGrid.StyleSets("Ortho").Font.Size
            xmlWL.OrthoFontBold = ssGrid.StyleSets("Ortho").Font.Bold
            xmlWL.OrthoFontItalic = ssGrid.StyleSets("Ortho").Font.Italic
			
            xmlWL.GlossFontName = ssGrid.StyleSets("Gloss").Font.Name
            xmlWL.GlossFontSize = ssGrid.StyleSets("Gloss").Font.Size
            xmlWL.GlossFontBold = ssGrid.StyleSets("Gloss").Font.Bold
            xmlWL.GlossFontItalic = ssGrid.StyleSets("Gloss").Font.Italic
			
            xmlWL.DialectFontName = ssGrid.StyleSets("Dialect").Font.Name
            xmlWL.DialectFontSize = ssGrid.StyleSets("Dialect").Font.Size
            xmlWL.DialectFontBold = ssGrid.StyleSets("Dialect").Font.Bold
            xmlWL.DialectFontItalic = ssGrid.StyleSets("Dialect").Font.Italic
		End With
		
		xmlWL.Save()
		
	End Sub
	
	Private Sub SetColumnStyles()
		
		On Error Resume Next
		
		With ssGrid
			If (Len(xmlWL.PhoneticFontName) > 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.StyleSets("IPA").Font.Name = xmlWL.PhoneticFontName
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.StyleSets("IPA").Font.Size = xmlWL.PhoneticFontSize
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Bold. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.StyleSets("IPA").Font.Bold = xmlWL.PhoneticFontBold
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Italic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.StyleSets("IPA").Font.Italic = xmlWL.PhoneticFontItalic
				
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.StyleSets("IPAAudio").Font.Name = xmlWL.PhoneticFontName
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.StyleSets("IPAAudio").Font.Size = xmlWL.PhoneticFontSize
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Bold. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.StyleSets("IPAAudio").Font.Bold = xmlWL.PhoneticFontBold
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Italic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.StyleSets("IPAAudio").Font.Italic = xmlWL.PhoneticFontItalic
			End If
			
			If (Len(xmlWL.OrthoFontName) > 0) Then
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.StyleSets("Ortho").Font.Name = xmlWL.OrthoFontName
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.StyleSets("Ortho").Font.Size = xmlWL.OrthoFontSize
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Bold. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.StyleSets("Ortho").Font.Bold = xmlWL.OrthoFontBold
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Italic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
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
	
	Private Sub SetFormFontToColFont(ByRef ColIndex As Short)
		
		On Error Resume Next
		
		With ssGrid
			If (Len(.Columns(ColIndex).StyleSet) > 0) Then
				With .StyleSets(.Columns(ColIndex).StyleSet).Font
					'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Font = VB6.FontChangeName(Font, .Name)
					'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Font = VB6.FontChangeSize(Font, .Size)
					'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Bold. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Font = VB6.FontChangeBold(Font, .Bold)
					'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Italic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					Font = VB6.FontChangeItalic(Font, .Italic)
				End With
			Else
				With .Font
					Font = VB6.FontChangeName(Font, .Name)
					Font = VB6.FontChangeSize(Font, .SizeInPoints)
					Font = VB6.FontChangeBold(Font, .Bold)
					Font = VB6.FontChangeItalic(Font, .Italic)
				End With
			End If
		End With
		
	End Sub
	
	Public Sub ShowPitchPlot()
		
		Dim iCount As Short
		Dim sPlotWindowCoords As String
		Dim sWaveFile As String
		
		On Error Resume Next
		
		iCount = 0
		Kill(gListFilePath)
		sWaveFile = GetSelectedRowsWavFile(True)
		If (Len(sWaveFile) = 0) Then Exit Sub
		
		With mdiHelpCharts
			sPlotWindowCoords = (VB6.PixelsToTwipsX(.Left) \ gPixelX) + ((VB6.PixelsToTwipsX(.Width) * 0.66) \ gPixelX) & "," & (VB6.PixelsToTwipsY(.Top) \ gPixelY) + ((VB6.PixelsToTwipsY(.Height) * 0.3) \ gPixelY) & "," & (VB6.PixelsToTwipsX(.Width) * 0.66) \ gPixelX & "," & (VB6.PixelsToTwipsY(.Height) * 0.5) \ gPixelY
		End With
		
		If (Len(GetINIEntry(cSettingsSect, PitchPlotShownEntry, gINIPath)) = 0) Then
			Call WriteINIEntry(cSettingsSect, PitchPlotShownEntry, "1", gINIPath)
			Call WriteINIEntry("Settings", "ShowWindow", "Size(" & sPlotWindowCoords & ")", gListFilePath)
		End If
		
		Call WriteINIEntry("Settings", "CallingApp", mdiHelpCharts.Text, gListFilePath)
		Call WriteINIEntry("Commands", "Command0", "DisplayPlot(Pitch)", gListFilePath)
		Call WriteINIEntry("Commands", "Command1", "Return(2)", gListFilePath)
		
		Do 
			Call WriteINIEntry("AudioFiles", "File" & iCount, sWaveFile, gListFilePath)
			iCount = iCount + 1
			sWaveFile = GetSelectedRowsWavFile()
		Loop Until (Len(sWaveFile) = 0 Or ssGrid.SelBookmarks.Count = 0)
		
		If (gRecordingExists) Then Call WriteINIEntry("AudioFiles", "File" & iCount, gTmpWavPath & gTmpWavName, gListFilePath)
		
		Call mdiHelpCharts.CallSA()
		
	End Sub
	
	Public ReadOnly Property SoundFile() As String
		Get
			
			On Error Resume Next
			'UPGRADE_WARNING: Couldn't resolve default property of object MakeFullPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			SoundFile = MakeFullPath(sWordListSndPath, (ssGrid.Columns("WavFile").Text))
			
		End Get
	End Property
	
	Private Sub StopPlayback()
		
		On Error Resume Next
		Timer1.Enabled = False
		
	End Sub
	
	Public Sub TestPlay()
		
		Dim i As Short
		Dim iPrevSection As Short
		Dim iPrevItem As Short
		Dim lRetryCount As Integer
		
		With tvSections
			If (gTestCatChoice = cUser And ValidWavFilesInGrid() = 0) Then
				Call MsgBox("No valid sound files found for this category.", MsgBoxStyle.OKOnly + MsgBoxStyle.Information, My.Application.Info.Title)
				Exit Sub
			End If
			
			'UPGRADE_WARNING: Couldn't resolve default property of object gItemNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			iPrevItem = gItemNumber
			If (.SelectedNode Is Nothing) Then
				iPrevSection = 0
			Else
				iPrevSection = .SelectedNode.Index
			End If
			
			'****************************************************************************
			'* If the user is testing random sections, then randomly choose a section.
			'****************************************************************************
			If (gTestCatChoice = cRandom) Then
				Do 
					lRetryCount = 0
					Do 
						Randomize()
						i = Int(.Nodes.Count * Rnd() + 1)
						lRetryCount = lRetryCount + 1
					Loop While ((i < 1 Or i = iPrevSection) And lRetryCount < 32000)
					
					'UPGRADE_WARNING: Lower bound of collection tvSections.Nodes has changed from 1 to 0. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
					.SelectedNode = .Nodes.Item(i)
					Call tvSections_NodeClick(tvSections, New System.Windows.Forms.TreeNodeMouseClickEventArgs(.SelectedNode, System.Windows.Forms.MouseButtons.None, 0, 0, 0))
					
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
		lRetryCount = 0
		
		With ssGrid
			Do 
				Randomize()
				i = Int(.Rows * Rnd())
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.RowBookmark(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Bookmark = .RowBookmark(i)
				'UPGRADE_WARNING: Couldn't resolve default property of object gItemNumber. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				gItemNumber = i
				lRetryCount = lRetryCount + 1
				'UPGRADE_WARNING: Couldn't resolve default property of object MakeFullPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			Loop While ((Not FileExist(MakeFullPath(sWordListSndPath, (.Columns("WavFile").Text))) Or i = iPrevItem) And lRetryCount < 32000)
		End With
		
		Call Play(cVocBttn)
		
	End Sub
	
	Public Sub UpdateAfterRecordAndPlayback()
		
		On Error Resume Next
		Call mdiHelpCharts.EnableTBarButtons(TBarButtons)
		
	End Sub
	
	Public Sub UpdateFormAfterTest()
		
		On Error Resume Next
		
        tvSections.Focus()
        tvSections.SelectedNode = tvSections.Nodes.Item(0).Parent.FirstNode
        Call tvSections_NodeClick(tvSections, New System.Windows.Forms.TreeNodeMouseClickEventArgs(tvSections.SelectedNode, System.Windows.Forms.MouseButtons.None, 0, 0, 0))

        picTestMode.Visible = False
        Call mdiHelpCharts.EnableTBarButtons(TBarButtons)

	End Sub
	
	Public Sub UpdateFormForTest()
		
		Dim iBttnTop As Short
		
		On Error Resume Next
		
		With picTestMode
			'.Move 0, 0 + IIf(gTestCatChoice = cRandom, 0, 270), picGrid.Width, picGrid.Height
			.Visible = True
			.BringToFront()
		End With
		
		With txtUserTr
			With .Font
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				txtUserTr.Font = VB6.FontChangeName(txtUserTr.Font, ssGrid.StyleSets("IPA").Font.Name)
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				txtUserTr.Font = VB6.FontChangeSize(txtUserTr.Font, ssGrid.StyleSets("IPA").Font.Size)
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Bold. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				txtUserTr.Font = VB6.FontChangeBold(txtUserTr.Font, ssGrid.StyleSets("IPA").Font.Bold)
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Italic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				txtUserTr.Font = VB6.FontChangeItalic(txtUserTr.Font, ssGrid.StyleSets("IPA").Font.Italic)
			End With
			.Top = .Top 'This forces text box to resize after font change above
		End With
		
		With txtCorrectTr
			With .Font
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Name. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				txtCorrectTr.Font = VB6.FontChangeName(txtCorrectTr.Font, ssGrid.StyleSets("IPA").Font.Name)
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Size. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				txtCorrectTr.Font = VB6.FontChangeSize(txtCorrectTr.Font, ssGrid.StyleSets("IPA").Font.Size)
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Bold. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				txtCorrectTr.Font = VB6.FontChangeBold(txtCorrectTr.Font, ssGrid.StyleSets("IPA").Font.Bold)
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.StyleSets().Font.Italic. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				txtCorrectTr.Font = VB6.FontChangeItalic(txtCorrectTr.Font, ssGrid.StyleSets("IPA").Font.Italic)
			End With
			
			.Top = .Top 'This forces text box to resize after font change above
			iBttnTop = VB6.PixelsToTwipsY(.Top) + VB6.PixelsToTwipsY(.Height) + 135
		End With
		
		cmdReplay.Top = VB6.TwipsToPixelsY(iBttnTop)
		cmdVerify.Top = VB6.TwipsToPixelsY(iBttnTop)
		cmdNext.Top = VB6.TwipsToPixelsY(iBttnTop)
		
		Select Case gTestCatChoice
			Case cUser : gStatLine.Text = "Select category to start test"
			Case cRandom : gStatLine.Text = "Enter transcription when item is pronounced"
		End Select
		
		picTestMode.Refresh()
		
	End Sub
	
	Private Sub UpdateXMLFileAfterEdits()
		
		Dim i As Short
		
		With ssGrid
            .CtlUpdate()
			
			'******************************************************
			'* Save the changed words to the XML file. Do this by
			'* emptying the word list then readding the words.
			'******************************************************
			If Not (xmlWL.EmptyCategory(.Text)) Then Exit Sub
			
			Dim sList(.Rows - 1, 6) As String
			
			For i = 0 To .Rows - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object ssGrid.RowBookmark(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				.Bookmark = .RowBookmark(i)
				sList(i, 1) = .Columns(0).Text
				sList(i, 2) = .Columns(1).Text
				sList(i, 3) = .Columns(2).Text
				sList(i, 4) = .Columns(3).Text
				sList(i, 5) = .Columns(4).Text
				sList(i, 6) = .Columns(5).Text
			Next 
			
			xmlWL.AddCategory(.Text, sList)
			xmlWL.Save()
		End With
		
	End Sub
	
	Private Function UserReallyWantsEditMode() As Boolean
		
		'***********************************************************
		'* This routine asks the user whether or not he really
		'* wants to go into the edit mode.
		'***********************************************************
		
		UserReallyWantsEditMode = False
		
		If (MsgBox("You are about to enter the edit mode in which you may" & vbCrLf & "make changes to word lists. Are you sure want to proceed?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, My.Application.Info.Title) = MsgBoxResult.No) Then Exit Function
		
		UserReallyWantsEditMode = True
		
		'***********************************************************
		'* Keep a backup copy of the XML file in case the user
		'* ends up not wanting to keep his changes.
		'***********************************************************
		FileCopy(sWordListPath, gINIPath & TempWLXMLBackup)
		
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
		If (MsgBox("Do you want to keep your changes?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, My.Application.Info.Title) = MsgBoxResult.No) Then
			UserReallyWantsToKeepEdits = False
			FileCopy(gINIPath & TempWLXMLBackup, sWordListPath)
			Call InitializeView()
		End If
		
		If (FileExist(sWordListPath)) Then Kill(gINIPath & TempWLXMLBackup)
		
	End Function
	
	Private Function ValidWavFilesInGrid() As Short
		
		'*********************************************************
		'* This routine will count how many valid wave files are
		'* in the WavFile column of the grid.
		'*********************************************************
		
		On Error Resume Next
		
		Dim i As Short
		Dim iCount As Short
		
		iCount = 0
		
		With ssGrid
			For i = 0 To .Rows - 1
				'UPGRADE_WARNING: Couldn't resolve default property of object MakeFullPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If (FileExist(MakeFullPath(sWordListSndPath, .Columns("WavFile").CellText(.RowBookmark(i))))) Then iCount = iCount + 1
			Next 
		End With
		
		ValidWavFilesInGrid = iCount
		
	End Function
	
	Private Sub cmdNext_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdNext.Click
		
		txtUserTr.Text = ""
		txtCorrectTr.Text = ""
		cmdReplay.Enabled = True
		cmdVerify.Enabled = True
		Call TestPlay()
		
	End Sub
	
	Private Sub cmdReplay_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReplay.Click
		
		On Error Resume Next
		Call Play(cVocBttn)
		
	End Sub
	
	Private Sub cmdVerify_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdVerify.Click
		
		On Error Resume Next
		Call StopPlayback()
		txtCorrectTr.Text = ssGrid.Columns("Phonetic").Text
		cmdVerify.Enabled = False
		cmdReplay.Enabled = True
		
	End Sub
	
	'UPGRADE_WARNING: Form event frmWordList.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmWordList_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
		
		Dim i As Short
		
		On Error Resume Next
		
		If (WindowState <> System.Windows.Forms.FormWindowState.Maximized) Then WindowState = System.Windows.Forms.FormWindowState.Maximized
		Call UpdateTestMenu() '* Enable test menu.
		
		gStatLine.Text = "Click to select"
		gRecordingExists = False
		tvSections.Focus()
		
		With mdiHelpCharts
			If Not (CType(.Controls("panStatus"), Object).Visible) Then CType(.Controls("panStatus"), Object).Visible = True
			Call .ShowTBarButtons(TBarButtons)
			Call .EnableTBarButtons(TBarButtons)
			CType(.Controls("mnuEdit"), Object).Visible = gAppPathWriteAccess
			
			For i = 0 To 2
				CType(.Controls("mnuPlayback"), Object)(i).Visible = True
			Next 
		End With
		
	End Sub
	
	'UPGRADE_WARNING: Form event frmWordList.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
	Private Sub frmWordList_Deactivate(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Deactivate
		
		Dim i As Short
		
		On Error Resume Next
		
		With mdiHelpCharts
			CType(.Controls("mnuEdit"), Object).Visible = False
			For i = 0 To 2
				CType(.Controls("mnuPlayback"), Object)(i).Visible = False
			Next 
		End With
		
	End Sub
	
	Private Sub frmWordList_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		On Error Resume Next
		Call KeyHandler(KeyCode, -1, Shift)
		
	End Sub
	
	Private Sub frmWordList_KeyPress(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
		Dim KeyAscii As Short = Asc(eventArgs.KeyChar)
		
		On Error Resume Next
		Call KeyHandler(-1, KeyAscii, -1)
		
		eventArgs.KeyChar = Chr(KeyAscii)
		If KeyAscii = 0 Then
			eventArgs.Handled = True
		End If
	End Sub
	
	Private Sub frmWordList_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		
		Dim sINIVal As String
		
		On Error Resume Next
		
		With mdiHelpCharts
			If Not (CType(.Controls("panStatus"), Object).Visible) Then CType(.Controls("panStatus"), Object).Visible = True
			If Not (CType(.Controls("TBar"), Object).Visible) Then CType(.Controls("TBar"), Object).Visible = True
			Call UpdateTestMenu() '* Enable test menu.
		End With
		
		Call ManageGridColumns(False)
		'UPGRADE_NOTE: Object WLPlayer may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		WLPlayer = Nothing
		gStatLine.Text = ""
		Text = sWordListCaption & " Word List"
		Tag = sWordListIndex
		
		With ssGrid
			.Redraw = False
			.RemoveAll()
			.Columns("Ortho").Visible = False
			.Columns("WavFile").Visible = False
			.Columns("GraphicFile").Visible = False
			.Columns("EditWavButton").Visible = False
			.Height = mdiHelpCharts.ClientRectangle.Height
			.Width = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(mdiHelpCharts.ClientRectangle.Width) - VB6.PixelsToTwipsX(.Left))
			iMinPhonColWidth = 0.33 * (VB6.PixelsToTwipsX(.Width) - gPixelX * 40)
			iMaxPhonColWidth = 0.66 * (VB6.PixelsToTwipsX(.Width) - gPixelX * 40)
			.Redraw = True
		End With
		
		picTestMode.BorderStyle = System.Windows.Forms.BorderStyle.None
		
		With picDragBar
			.BorderStyle = System.Windows.Forms.BorderStyle.None
			.BackColor = System.Drawing.SystemColors.ControlDark
		End With
		
		With picSplitter
			.BorderStyle = System.Windows.Forms.BorderStyle.None
			.BackColor = System.Drawing.SystemColors.Control
			
			sINIVal = GetINIEntry("WordList", "SplitterLeft", gINIPath)
			.SetBounds(VB6.TwipsToPixelsX(IIf(Val(sINIVal) = 0, (gPixelX * 200), Val(sINIVal))), 0, VB6.TwipsToPixelsX(gPixelX * 8), ClientRectangle.Height)
			picDragBar.SetBounds(.Left, .Top, .Width, .Height)
			picDragBar.BringToFront()
			picDragBar.Visible = False
		End With
		
		Call InitializeView()
		
		With tvSections
			.Height = mdiHelpCharts.ClientRectangle.Height
			.SendToBack()
		End With
		
		Show()
		WindowState = System.Windows.Forms.FormWindowState.Maximized
		bGridDoubleClicked = False
		bInEditMode = False
		
	End Sub
	
	Private Sub frmWordList_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles MyBase.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		On Error Resume Next
		gStatLine.Text = ""
		
	End Sub
	
	Private Sub frmWordList_FormClosing(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		Dim Cancel As Boolean = eventArgs.Cancel
		Dim UnloadMode As System.Windows.Forms.CloseReason = eventArgs.CloseReason
		
		On Error Resume Next
		Call WriteINIEntry("WordList", "SplitterLeft", CStr(VB6.PixelsToTwipsX(picSplitter.Left)), gINIPath)
		'UPGRADE_WARNING: Form event frmWordList.Deactivate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
		frmWordList_Deactivate(Me, New System.EventArgs())
		If (bInEditMode) Then Call EditMode()
		Call ManageGridColumns(True)
		'UPGRADE_NOTE: Object xmlWL may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
		xmlWL = Nothing
		'UPGRADE_NOTE: Object frmWordList may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        Me.Dispose()
		
		eventArgs.Cancel = Cancel
	End Sub
	
    Private Sub frmWordList_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
        On Error Resume Next
        Call AdjustControlPlacement()
    End Sub
#If OMIT Then
	Private Sub ssGrid_BtnClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ssGrid.BtnClick
		
		'UPGRADE_ISSUE: Load statement is not supported. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B530EFF2-3132-48F8-B8BC-D88AF543D321"'
		Load(frmGetSoundFile)
		
		With frmGetSoundFile
			If (Len(ssGrid.Columns("WavFile").Text) > 0) Then .FileName = ssGrid.Columns("WavFile").Text
			
			.Folder = sWordListSndPath
			.ShowDialog()
			
			If Not (.Canceled) Then ssGrid.Columns("WavFile").Text = .FileName
		End With
		
		frmGetSoundFile.Close()
		
	End Sub
	
    Private Sub ssGrid_ColResize(ByVal eventSender As System.Object, ByVal eventArgs As SSDataWidgets_B._DSSDBGridEvents) Handles ssGrid.ColResize

        On Error Resume Next
        Call ResizeGrid(eventArgs.ColIndex, ssGrid.ResizeWidth)

    End Sub
	
    Private Sub ssGrid_KeyDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As SSDataWidgets_B._DSSDBGridEvents_KeyDownEvent) Handles ssGrid.KeyDownEvent

        On Error Resume Next
        Call KeyHandler(eventArgs.KeyCode, -1, eventArgs.Shift)

    End Sub
	
    Private Sub ssGrid_KeyPressEvent(ByVal eventSender As System.Object, ByVal eventArgs As SSDataWidgets_B._DSSDBGridEvents) Handles ssGrid.KeyPressEvent

        On Error Resume Next
        Call KeyHandler(-1, eventArgs.KeyAscii, -1)

    End Sub
	
    Private Sub ssGrid_MouseDownEvent(ByVal eventSender As System.Object, ByVal eventArgs As SSDataWidgets_B._DSSDBGridEvents) Handles ssGrid.MouseDownEvent

        On Error Resume Next
        If (bInEditMode) Then Exit Sub
        If (eventArgs.Shift = 0 And eventArgs.X > 200 And eventArgs.Y > 255) Then ssGrid.SelBookmarks.RemoveAll()

    End Sub
	
    Private Sub ssGrid_MouseMoveEvent(ByVal eventSender As System.Object, ByVal eventArgs As SSDataWidgets_B._DSSDBGridEvents) Handles ssGrid.MouseMoveEvent

        '***************************************************************
        '* This will turn on the balloon help for the wave file column
        '* since a wave file's path length will likely exceed the width
        '* of the column in which it's displayed.
        '***************************************************************

        Dim iCol As Short

        On Error Resume Next

        If Not (bInEditMode) Then Exit Sub

        With ssGrid
            iCol = .ColContaining(eventArgs.X)
            If (iCol >= 0) Then
                If (.Columns(iCol).Name = "WavFile") Then
                    If Not (.BalloonHelp) Then .BalloonHelp = True
                    Exit Sub
                End If
            End If

            .BalloonHelp = False
        End With

    End Sub
	
    Private Sub ssGrid_MouseUpEvent(ByVal eventSender As System.Object, ByVal eventArgs As SSDataWidgets_B._DSSDBGridEvents) Handles ssGrid.MouseUpEvent()

        On Error Resume Next

        If (bInEditMode) Then Exit Sub

        Dim iRow As Short
        With ssGrid
            If (.ColContaining(eventArgs.X) = -1) Then Exit Sub
            iRow = .RowContaining(eventArgs.Y)
            If (iRow >= 0 And iRow <= .Rows And eventArgs.X > 200 And Not .AllowUpdate) Then Call Play(-1)
        End With

    End Sub
	
    Private Sub ssGrid_RowColChange(ByVal eventSender As System.Object, ByVal eventArgs As SSDataWidgets_B._DSSDBGridEvents_RowColChangeEvent) Handles ssGrid.RowColChange

        On Error Resume Next

        'UPGRADE_WARNING: Couldn't resolve default property of object LastRow. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (bImportingCategory Or picTestMode.Visible Or bInEditMode) Or (CStr(eventArgs.LastRow) = CStr(ssGrid.Bookmark)) Then Exit Sub

        gRecordingExists = False
        Kill(gTmpWavPath & gTmpWavName)

        With ssGrid.Columns("WavFile")
            'UPGRADE_WARNING: Couldn't resolve default property of object MakeFullPath(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Use of Null/IsNull() detected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="2EED02CB-5C0E-4DC1-AE94-4FAA3A30F51A"'
            If (Not IsDBNull(eventArgs.LastRow) And Not FileExist(MakeFullPath(sWordListSndPath, .Text))) Then
                Call mdiHelpCharts.EnableTBarButtons(NoWaveEnalbedTBarButtons)
            Else
                Call mdiHelpCharts.EnableTBarButtons(TBarButtons)
            End If
        End With

    End Sub
	
    Private Sub ssGrid_RowLoaded(ByVal eventSender As System.Object, ByVal eventArgs As SSDataWidgets_B._DSSDBGridEvents_RowLoadedEvent) Handles ssGrid.RowLoaded

        On Error Resume Next
        If (FileExist(MakeFullPath(sWordListSndPath, (ssGrid.Columns("WavFile").Text)))) Then
            ssGrid.Columns("Phonetic").CellStyleSet("IPAAudio", eventArgs.Bookmark)
        End If
        
    End Sub
	
    Private Sub ssGrid_SelChange(ByVal eventSender As System.Object, ByVal eventArgs As SSDataWidgets_B._DSSDBGridEvents_SelChangeEvent) Handles ssGrid.SelChange

        On Error Resume Next
        If (bInEditMode) Then Exit Sub
        If Not (FileExist(MakeFullPath(sWordListSndPath, (ssGrid.Columns("WavFile").Text)))) Then eventArgs.Cancel = 1

    End Sub
#End If
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
		
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
	
	Private Sub picDragBar_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles picDragBar.Paint

        Dim graphics As Graphics
        graphics = picDragBar.CreateGraphics
        Dim pen As New Pen(System.Drawing.SystemColors.ControlLightLight, 1)

        graphics.DrawLine(pen, New Point(gPixelX, 0), New Point(gPixelX, VB6.PixelsToTwipsY(Me.ClientRectangle.Height)))

        pen.Color = System.Drawing.SystemColors.ControlDarkDark
        graphics.DrawLine(pen, New Point(VB6.PixelsToTwipsX(picDragBar.ClientRectangle.Width) - gPixelX, 0), New Point(VB6.PixelsToTwipsX(picDragBar.ClientRectangle.Width) - gPixelX, VB6.PixelsToTwipsY(picDragBar.ClientRectangle.Height)))

	End Sub
	
	Private Sub picSplitter_MouseDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picSplitter.MouseDown
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		On Error Resume Next
		
		If (Button <> VB6.MouseButtonConstants.LeftButton) Then Exit Sub
		sngMouseXOffset = X
		bSplitterDrag = True
		picDragBar.Visible = True
		picDragBar.BringToFront()
		picSplitter.Visible = False
		
	End Sub
	
	Private Sub picSplitter_MouseMove(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picSplitter.MouseMove
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		If Not (bSplitterDrag) Then Exit Sub
		picDragBar.Left = VB6.TwipsToPixelsX(VB6.PixelsToTwipsX(picSplitter.Left) + X - sngMouseXOffset)
		
	End Sub
	
	Private Sub picSplitter_MouseUp(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.MouseEventArgs) Handles picSplitter.MouseUp
		Dim Button As Short = eventArgs.Button \ &H100000
		Dim Shift As Short = System.Windows.Forms.Control.ModifierKeys \ &H10000
		Dim X As Single = VB6.PixelsToTwipsX(eventArgs.X)
		Dim Y As Single = VB6.PixelsToTwipsY(eventArgs.Y)
		
		On Error Resume Next
		
		If (bSplitterDrag) Then
			bSplitterDrag = False
			
			With picDragBar
				.Visible = False
				picSplitter.SetBounds(.Left, .Top, .Width, .Height)
				picSplitter.Visible = True
			End With
			
			Call AdjustControlPlacement()
		End If
		
	End Sub
	
	Private Sub picSplitter_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles picSplitter.Paint
		
		On Error Resume Next

        Dim graphics As Graphics
        Dim pen As New Pen(System.Drawing.SystemColors.ControlLightLight, 1)

        graphics = picSplitter.CreateGraphics

        graphics.DrawLine(pen, New Point(0, 0), New Point(0, VB6.PixelsToTwipsY(picSplitter.ClientRectangle.Height)))
        pen.Color = System.Drawing.SystemColors.ControlDark
        graphics.DrawLine(pen, New Point(VB6.PixelsToTwipsX(picSplitter.ClientRectangle.Width) - gPixelX, 0), New Point(VB6.PixelsToTwipsX(picSplitter.ClientRectangle.Width) - gPixelX, VB6.PixelsToTwipsY(picSplitter.ClientRectangle.Height)))

	End Sub
	
	Private Sub tvSections_AfterLabelEdit(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.NodeLabelEditEventArgs) Handles tvSections.AfterLabelEdit
		Dim Cancel As Boolean = eventArgs.CancelEdit
		Dim NewString As String = eventArgs.Label
		
		Dim sOldCategoryName As String
		
		With tvSections
			sOldCategoryName = .SelectedNode.Text
			If (CategoryNameExists(NewString, .SelectedNode, True)) Then
				Cancel = True
				.Focus()
				Exit Sub
			End If
		End With
		
		With xmlWL
			If (.ChangeCategoryName(sOldCategoryName, NewString)) Then .Save()
		End With
		
	End Sub
	
	Private Sub tvSections_KeyDown(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.KeyEventArgs) Handles tvSections.KeyDown
		Dim KeyCode As Short = eventArgs.KeyCode
		Dim Shift As Short = eventArgs.KeyData \ &H10000
		
		If (KeyCode = System.Windows.Forms.Keys.F2 And bInEditMode And Shift = 0) Then tvSections.SelectedNode.BeginEdit()
		
	End Sub
	
	Private Sub tvSections_NodeClick(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.TreeNodeMouseClickEventArgs) Handles tvSections.NodeMouseClick
		Dim node As System.Windows.Forms.TreeNode = eventArgs.Node
		
		On Error Resume Next
		
		Call StopPlayback()
		If (node Is Nothing) Then Exit Sub
		
		If Not (LastTreeNode Is node) Then
			If (bInEditMode) Then Call UpdateXMLFileAfterEdits()
			Call ImportCategory((node.Text))
			ssGrid.Text = node.Text
			Call ResizeGrid()
		End If
		
		LastTreeNode = node
		If (gTestActive And gTestCatChoice = cUser) Then Call cmdNext_Click(cmdNext, New System.EventArgs())
		
	End Sub
End Class