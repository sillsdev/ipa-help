Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks

Friend Class frmMenu

    Inherits System.Windows.Forms.Form

    '**************************************************
    '* frmMenu version info:
    '*  See basGlobals (Globals.bas).
    '**************************************************

    Private Const MaxWordListsVisible As Short = 5

    Private iTopWordListName As Short

    Public Sub SetupWordListOptions()

        Dim i As Short
        Dim iIndex As Short

        On Error GoTo SetupWordListOptionsErr

        '* Setup word list tooltips and display

        iIndex = iTopWordListName

        '**********************************************
        '* Setup word list tooltips and display
        '**********************************************
        For i = 4 To 8
            With lnkChart(i)
                If (iIndex > WordListArraySize()) Then
                    .Visible = False
                ElseIf (Len(gWordListID(iIndex, 0)) > 0) Then
                    'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(iIndex, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Text = gWordListID(iIndex, 0)
                    'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(iIndex, 0). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    ToolTip1.SetToolTip(lnkChart(i), "Display " & gWordListID(iIndex, 0) & " Word List")
                    'UPGRADE_WARNING: Couldn't resolve default property of object gWordListID(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    .Enabled = FileExist(gWordListID(iIndex, 1))
                    .Visible = True
                Else
                    .Visible = False
                    .Text = ""
                End If
            End With
            iIndex = iIndex + 1
        Next i

        lnkUpDown(0).Visible = (WordListArraySize() >= MaxWordListsVisible)
        lnkUpDown(1).Visible = (WordListArraySize() >= MaxWordListsVisible)

        ' Line1(1).Y2 = lblChart(iLastList).Top + lblChart(iLastList).Height

        Exit Sub

        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
SetupWordListOptionsErr:
        MsgBox(Err.Description & " - (SetupWordListOptions)", MsgBoxStyle.Information, My.Application.Info.Title)

    End Sub

    'UPGRADE_WARNING: Form event frmMenu.Activate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub frmMenu_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated

        With mdiHelpCharts
            If (CType(.Controls("TBar"), Object).Visible) Then CType(.Controls("TBar"), Object).Visible = False
            If (CType(.Controls("panStatus"), Object).Visible) Then CType(.Controls("panStatus"), Object).Visible = False
            CType(.Controls("mnuTest"), Object).Enabled = True
        End With

        Call UpdateTestMenu()

        'UPGRADE_ISSUE: Unable to determine which constant to upgrade vbNormal to. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="B3B44E51-B5F1-4FD7-AA29-CAD31B71F487"'
        If WindowState = vbNormal Then
            Top = VB6.TwipsToPixelsY(-VB6.PixelsToTwipsY(Height))
            Show()
            WindowState = System.Windows.Forms.FormWindowState.Maximized
        End If

    End Sub

    Private Sub frmMenu_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        Dim i As Short

        On Error Resume Next
        mdiHelpCharts.panStatus.Visible = False

        Randomize() ' Set a new seed value
        Top = VB6.TwipsToPixelsY(-VB6.PixelsToTwipsY(Height))
        Show()
        WindowState = System.Windows.Forms.FormWindowState.Maximized
        Call SetupWordListOptions()

        iTopWordListName = 0

        With imgBkg
            For i = 0 To 14
                lnkChart(i).Picture = .Image
            Next

            lnkUpDown(0).Picture = .Image
            lnkUpDown(1).Picture = .Image
        End With

    End Sub

    Private Sub frmMenu_Paint(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint

        Dim i As Short
        Dim j As Short

        On Error Resume Next
        For i = 0 To VB6.PixelsToTwipsX(ClientRectangle.Width) Step VB6.PixelsToTwipsX(imgBkg.Width) - gPixelX
            For j = 0 To VB6.PixelsToTwipsY(ClientRectangle.Height) Step VB6.PixelsToTwipsY(imgBkg.Height)
                Dim x As Integer
                Dim y As Integer
                Dim w As Integer
                Dim h As Integer
                x = i
                y = j
                w = VB6.PixelsToTwipsX(imgBkg.Width)
                h = VB6.PixelsToTwipsY(imgBkg.Height)
                eventArgs.Graphics.DrawImage(imgBkg.Image, x, y, w, h)
            Next j
        Next i

    End Sub

    Private Sub lnkChart_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lnkChart.ClickEvent

        Dim Index As Short = lnkChart.GetIndex(eventSender)

        Dim frm As System.Windows.Forms.Form

        On Error Resume Next
        'UPGRADE_WARNING: Screen property Screen.MousePointer has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        Select Case Index
            Case 0 : frmDispCon.Show()
            Case 1 : frmDispVow.Show()
            Case 2 : frmDispDia.Show()
            Case 3 : frmDispSSeg.Show()
            Case 9 : frmDiagArtrs.Show()
            Case 10 : frmDiagPtArtn.Show()
            Case 13 : frmDispAmerDia2.Show()
            Case 14 : frmDispAmerOther2.Show()

            Case 4 To 8
                For Each frm In My.Application.OpenForms
                    If frm.Tag = Str(Index - 4) Then
                        frm.Show()
                        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                        Exit Sub
                    End If
                Next frm

                CType(frm, frmWordList).Initialize(Index - 4)
                frm.Show()

            Case 11
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                mdiHelpCharts.Controls("mnuSILCons").Show()

            Case 12
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                mdiHelpCharts.Controls("mnuSILVows").Show()
        End Select

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

    End Sub

    Private Sub lnkUpDown_ClickEvent(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lnkUpDown.ClickEvent
        Dim Index As Short = lnkUpDown.GetIndex(eventSender)

        On Error Resume Next

        If (Index = 0) Then
            If (iTopWordListName = 0) Then Exit Sub
            iTopWordListName = iTopWordListName - 1
        Else
            If ((iTopWordListName + MaxWordListsVisible) > WordListArraySize()) Then Exit Sub
            iTopWordListName = iTopWordListName + 1
        End If

        Call SetupWordListOptions()

    End Sub
End Class