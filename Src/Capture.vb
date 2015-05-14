Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks.Printing.Compatibility.VB6
Friend Class clsCapture
	
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'       Visual Basic 4.0 16 Capture Routines
	'
	' This module contains several routines for capturing windows into a
	' picture.  All the routines work on 16 bit Windows
	' platforms.
	' The routines also have palette support.
	'
	' CreateBitmapPicture - Creates a picture object from a bitmap and
	' palette.
	' CaptureWindow - Captures any window given a window handle.
	' CaptureActiveWindow - Captures the active window on the desktop.
	' CaptureForm - Captures the entire form.
	' CaptureClient - Captures the client area of a form.
	' CaptureScreen - Captures the entire screen.
	' PrintPictureToFitPage - prints any picture as big as possible on
	' the page.
	'
	' NOTES
	'    - No error trapping is included in these routines.
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	Private Const NormFont As String = "Arial"
	Private Const NormSize As Short = 9
	Private Const RASTERCAPS As Integer = 38
	Private Const RC_PALETTE As Integer = &H100
	Private Const SIZEPALETTE As Integer = 104
	Private Const COLOR_BTNFACE As Short = 15
	
	'**************************************************************************
	' CreateBitmapPicture - Creates a bitmap type Picture object from a
	'                       bitmap and palette.
	' hBmp - Handle to a bitmap.
	'
	' hPal - Handle to a Palette. (Can be null if the bitmap doesn't
	'        use a palette.)
	'
	' Returns - A Picture object containing the bitmap.
	'**************************************************************************
	Public Function CreateBitmapPicture(ByVal hBmp As Integer, ByVal hPal As Integer) As System.Drawing.Image
		
		Dim r As Integer
		Dim pic As PicBmp
        Dim IPic As System.Drawing.Bitmap
		'UPGRADE_WARNING: Arrays in structure IID_IDispatch may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim IID_IDispatch As GUID
		
		'**************************************************************************
		'* Fill in with IDispatch Interface ID.
        '**************************************************************************
        IID_IDispatch.Data1 = &H20400
        IID_IDispatch.Data4(0) = &HC0
        IID_IDispatch.Data4(7) = &H46
		
        '**************************************************************************
        '* Fill Pic with necessary parts.
        '**************************************************************************
        pic.Size = Len(pic) ' Length of structure.
        pic.hBmp = hBmp ' Handle to bitmap.
        pic.hPal = hPal ' Handle to palette (may be null).

        '**************************************************************************
        '* Create Picture object and return it.
        '**************************************************************************
        r = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
        CreateBitmapPicture = IPic

	End Function
	
	'**************************************************************************
	' CaptureWindow - Captures any portion of a window.
	'
	' hWndSrc - Handle to the window to be captured.
	'
	' Client - If True CaptureWindow captures from the client area of the
	'          window. If False CaptureWindow captures from the entire window.
	'
	' iLeft, iTop, iWidth, iHeight - Specify the portion of the window
	'                                        to capture. Dimensions need to be
	'                                        specified in pixels.
	'
	' Returns - A Picture object containing a bitmap of the specified
	'           portion of the window that was captured.
	'**************************************************************************
	Public Function CaptureWindow(ByVal hWndSrc As Integer, ByVal Client As Boolean, ByVal iLeft As Integer, ByVal iTop As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, Optional ByRef vMakeSemiMono As Object = Nothing) As System.Drawing.Image
		
		Dim bMakeSemiMono As Boolean
		Dim hDCMemory As Integer
		Dim hBmp As Integer
		Dim hBmpPrev As Integer
		Dim r As Integer
		Dim hDCSrc As Integer
		Dim hPal As Integer
		Dim hPalPrev As Integer
		Dim RasterCapsScrn As Integer
		Dim HasPaletteScrn As Integer
		Dim PaletteSizeScrn As Integer
		'UPGRADE_WARNING: Arrays in structure LogPal may need to be initialized before they can be used. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="814DF224-76BD-4BB4-BFFB-EA359CB9FC48"'
		Dim LogPal As LOGPALETTE
		
		On Error Resume Next
		
		bMakeSemiMono = False
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		'UPGRADE_WARNING: Couldn't resolve default property of object vMakeSemiMono. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Not (IsNothing(vMakeSemiMono)) Then bMakeSemiMono = vMakeSemiMono
		
		'**************************************************************************
		'* Depending on the value of Client get the proper device context.
		'**************************************************************************
		hDCSrc = IIf(Client, GetDC(hWndSrc), GetWindowDC(hWndSrc))
		
		'**************************************************************************
		'* Create a memory device context for the copy process. Then create a
		'* bitmap and place it in the memory DC.
		'**************************************************************************
		hDCMemory = CreateCompatibleDC(hDCSrc)
		hBmp = CreateCompatibleBitmap(hDCSrc, iWidth, iHeight)
		hBmpPrev = SelectObject(hDCMemory, hBmp)
		
		'**************************************************************************
		'* Set all the pixels in the memory DC to white. (This is incase we call
		'* TransBmp below.)
		'**************************************************************************
		r = BitBlt(hDCMemory, 0, 0, iWidth, iHeight, hDCSrc, iLeft, iTop, &HFF0062)
		
		'**************************************************************************
		'* Get screen properties.
		'**************************************************************************
		RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster capabilities.
		HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette support.
		PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of palette.
		
		'**************************************************************************
		'* If the screen has a palette make a copy and realize it.
		'**************************************************************************
		If (HasPaletteScrn And (PaletteSizeScrn = 256)) Then
			LogPal.palVersion = &H300
			LogPal.palNumEntries = 256
			r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
			hPal = CreatePalette(LogPal)
			
			'************************************************************************
			'* Select the new palette into the memory DC and realize it.
			'************************************************************************
			hPalPrev = SelectPalette(hDCMemory, hPal, 0)
			r = RealizePalette(hDCMemory)
		End If
		
		'**************************************************************************
		'* If the button face color should be converted to white then create the
		'* bitmap in the memory DC via the TransBmp function rather than BitBlt.
		'**************************************************************************
		If (bMakeSemiMono) Then
			Call TransBmp(hDCMemory, hDCSrc, iLeft, iTop, iWidth, iHeight, 0, 0, GetSysColor(COLOR_BTNFACE))
		Else
			'************************************************************************
			'* Copy the on-screen image into the memory DC. Then remove the new copy
			'* of the on-screen image.
			'************************************************************************
            Dim vbSrcCopy As Integer = &HCC0020
            r = BitBlt(hDCMemory, 0, 0, iWidth, iHeight, hDCSrc, iLeft, iTop, vbSrcCopy)
		End If
		
		hBmp = SelectObject(hDCMemory, hBmpPrev)
		
		'**************************************************************************
		'* If screen has a palette, get back the palette selected previously.
		'**************************************************************************
		If (HasPaletteScrn And (PaletteSizeScrn = 256)) Then hPal = SelectPalette(hDCMemory, hPalPrev, 0)
		
		'**************************************************************************
		'* Release the device context resources back to the system.
		'**************************************************************************
		r = DeleteDC(hDCMemory)
		r = ReleaseDC(hWndSrc, hDCSrc)
		
		'**************************************************************************
		'* Call CreateBitmapPicture to create a picture object from the bitmap and
		'* palette handles. Then return the resulting picture object.
		'**************************************************************************
		CaptureWindow = CreateBitmapPicture(hBmp, hPal)
		
	End Function
	
	'**************************************************************************
	' CaptureWindowArea - Captures an area of a form.
	'
	' frmSrc - The Form object to capture.
	'
	' Returns - A Picture object containing a bitmap of the entire form.
	'**************************************************************************
	Public Function CaptureWindowArea(ByRef frmSrc As System.Windows.Forms.Form, ByRef iLeft As Integer, ByRef iTop As Integer, ByRef iWidth As Integer, ByRef iHeight As Integer, Optional ByRef vMakeSemiMono As Object = Nothing) As System.Drawing.Image
		
		Dim bMakeSemiMono As Boolean
		
		On Error Resume Next
		
		bMakeSemiMono = False
		'UPGRADE_NOTE: IsMissing() was changed to IsNothing(). Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="8AE1CB93-37AB-439A-A4FF-BE3B6760BB23"'
		'UPGRADE_WARNING: Couldn't resolve default property of object vMakeSemiMono. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If Not (IsNothing(vMakeSemiMono)) Then bMakeSemiMono = vMakeSemiMono
		
		'**************************************************************************
		'* Call CaptureWindow to capture the entire form given its window handle
		'* and then return the resulting Picture object.
		'**************************************************************************
		With frmSrc
			CaptureWindowArea = CaptureWindow(.Handle.ToInt32, True, iLeft, iTop, iWidth, iHeight, bMakeSemiMono)
		End With
		
	End Function
	
	'**************************************************************************
	' CaptureScreen - Captures the entire screen.
	'
	' Returns - A Picture object containing a bitmap of the screen.
	'**************************************************************************
	Public Function CaptureScreen() As System.Drawing.Image
		
		'  Dim hWndScreen As Long
		
		'  On Error Resume Next
		
		'**************************************************************************
		'* Get a handle to the desktop window.
		'**************************************************************************
		'  hWndScreen = GetDesktopWindow()
		
		'**************************************************************************
		'* Call CaptureWindow to capture the entire desktop give the handle and
		'* and return the resulting Picture object.
		'**************************************************************************
		'  Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, _
		'Screen.Width \ Screen.TwipsPerPixelX, _
		'Screen.Height \ Screen.TwipsPerPixelY)
		
	End Function
	
	'**************************************************************************
	' CaptureForm - Captures an entire form including title bar and border.
	'
	' frmSrc - The Form object to capture.
	'
	' Returns - A Picture object containing a bitmap of the entire form.
	'**************************************************************************
	Public Function CaptureForm(ByRef frmSrc As System.Windows.Forms.Form) As System.Drawing.Image
		
		'  On Error Resume Next
		
		'**************************************************************************
		'* Call CaptureWindow to capture the entire form given its window handle
		'* and then return the resulting Picture object.
		'**************************************************************************
		'  With frmSrc
		'    Set CaptureForm = CaptureWindow(.hWnd, False, 0, 0, _
		'.ScaleX(.Width, vbTwips, vbPixels), _
		'.ScaleY(.Height, vbTwips, vbPixels))
		'  End With
		
	End Function
	
	'**************************************************************************
	' CaptureClient - Captures the client area of a form.
	'
	' frmSrc - The Form object to capture.
	'
	' Returns - A Picture object containing a bitmap of the form's client area.
	'**************************************************************************
	Public Function CaptureClient(ByRef frmSrc As System.Windows.Forms.Form) As System.Drawing.Image
		
		'  On Error Resume Next
		
		'**************************************************************************
		'* Call CaptureWindow to capture the client area of the form given its
		'* window handle and return the resulting Picture object.
		'**************************************************************************
		'  With frmSrc
		'    Set CaptureClient = CaptureWindow(.hWnd, True, 0, 0, _
		'.ScaleX(.ScaleWidth, .ScaleMode, vbPixels), _
		'.ScaleY(.ScaleHeight, .ScaleMode, vbPixels))
		'  End With
		
	End Function
	
	'**************************************************************************
	' CaptureActiveWindow - Captures the currently active window on the screen.
	'
	' Returns - A Picture object containing a bitmap of the active window.
	'**************************************************************************
	Public Function CaptureActiveWindow() As System.Drawing.Image
		
		'  Dim hWndActive As Integer
		'  Dim r As Integer
		'  Dim RectActive As RECT
		
		'  On Error Resume Next
		
		'**************************************************************************
		'* Get a handle to the active/foreground window and the dimensions of that
		'* window.
		'**************************************************************************
		'  hWndActive = GetForegroundWindow()
		'  r = GetWindowRect(hWndActive, RectActive)
		
		'**************************************************************************
		'* Call CaptureWindow to capture the active window given its handle and
		'* return the Resulting Picture object.
		'**************************************************************************
		'  With RectActive
		'    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, _
		'.Right - .Left, .Bottom - .Top)
		'  End With
		
	End Function
	
	Public Sub PrintChart(ByRef pic As System.Drawing.Image, ByRef sDescription As String)
		Dim Printer As New Printer
		
        On Error Resume Next
		
        '*******************************************************
        '* Don't change frmMenu to 'Printer' or this won't
        '* work. It really should be 'Printer' but doing
        '* a Printer.ScaleX method has the effect of beginning
        '* a new page after which setting the orientation won't
        '* work. ARRRGGGGHHH! frmMenu.ScaleX and Printer.ScaleX
        '* appeared to return the same values. I hope that will
        '* always be the case.
        '*******************************************************
        'UPGRADE_ISSUE: Picture property pic.Width was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'UPGRADE_ISSUE: Form method frmMenu.ScaleX was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'If (frmMenu.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, 8) < pic.Width) Then
        'Printer.Orientation = PrinterObjectConstants.vbPRORLandscape
        'Else
        Printer.Orientation = PrinterObjectConstants.vbPRORPortrait
        'End If

        Printer.Font = VB6.FontChangeName(Printer.Font, "")
        Printer.Font = VB6.FontChangeName(Printer.Font, "Arial")
        Printer.Font = VB6.FontChangeSize(Printer.Font, 10)

        Printer.Print()
        Printer.Print(sDescription)
        Printer.Print("Date/Time: " & vbTab & VB6.Format(Now, "Medium Date") & " " & VB6.Format(Now, "Medium Time"))
        Printer.PaintPicture(pic, 0, (Printer.TextHeight("X") * 5.5))
        Printer.EndDoc()


	End Sub
	
	'**************************************************************************
	' PrintPicture - Prints a Picture object.
	'
	' Prn - Destination Printer object.
	'
	' Pic - Source Picture object.
	'**************************************************************************
	Public Sub PrintPicture(ByRef prn As Printer, ByRef pic As System.Drawing.Image)
		
		'  Const vbHiMetric As Integer = 8
		'  Dim PicRatio As Double
		'  Dim PrnWidth As Double
		'  Dim PrnHeight As Double
		'  Dim PrnRatio As Double
		'  Dim PrnPicWidth As Double
		'  Dim PrnPicHeight As Double
		
		'  On Error Resume Next
		
		'  With Prn
		'**************************************************************************
		'* Determine if picture should be printed in landscape or portrait and set
		'* the orientation. Then calculate device independent Width-to-Height
		'* ratio for picture.
		'**************************************************************************
		'    .Orientation = IIf(pic.Height >= pic.Width, vbPRORPortrait, vbPRORLandscape)
		'    PicRatio = pic.Width / pic.Height
		
		'**************************************************************************
		'* Calculate the dimentions of the printable area in HiMetric.
		'**************************************************************************
		'    PrnWidth = .ScaleX(.ScaleWidth, .ScaleMode, vbHiMetric)
		'    PrnHeight = .ScaleY(.ScaleHeight, .ScaleMode, vbHiMetric)
		
		'**************************************************************************
		'* Calculate device independent Width to Height ratio for printer.
		'**************************************************************************
		'    PrnRatio = PrnWidth / PrnHeight
		
		'**************************************************************************
		'* Scale the output to the printable area.
		'**************************************************************************
		'    If (PicRatio >= PrnRatio) Then
		'      '************************************************************************
		'      '* Scale picture to fit full width of printable area.
		'      '************************************************************************
		'      PrnPicWidth = .ScaleX(PrnWidth, vbHiMetric, .ScaleMode)
		'      PrnPicHeight = .ScaleY(PrnWidth / PicRatio, vbHiMetric, .ScaleMode)
		'    Else
		'      '************************************************************************
		'      '* Scale picture to fit full height of printable area.
		'      '************************************************************************
		'      PrnPicHeight = .ScaleY(PrnHeight, vbHiMetric, .ScaleMode)
		'      PrnPicWidth = .ScaleX(PrnHeight * PicRatio, vbHiMetric, .ScaleMode)
		'    End If
		
		'**************************************************************************
		'* Print the picture using the PaintPicture method.
		'**************************************************************************
		'    .PaintPicture pic, 0, 0, pic.Width, pic.Height
		'  End With
		
	End Sub
	
	'**************************************************************************
	' PrintPictureToFitPage - Prints a Picture object as big as possible.
	'
	' Prn - Destination Printer object.
	'
	' Pic - Source Picture object.
	'**************************************************************************
	Public Sub PrintPictureToFitPage(ByRef prn As Printer, ByRef pic As System.Drawing.Image)
		
		'  Const vbHiMetric As Integer = 8
		'  Dim PicRatio As Double
		'  Dim PrnWidth As Double
		'  Dim PrnHeight As Double
		'  Dim PrnRatio As Double
		'  Dim PrnPicWidth As Double
		'  Dim PrnPicHeight As Double
		
		'  On Error Resume Next
		
		'  With Prn
		'**************************************************************************
		'* Determine if picture should be printed in landscape or portrait and set
		'* the orientation. Then calculate device independent Width-to-Height
		'* ratio for picture.
		'**************************************************************************
		'    .Orientation = IIf(pic.Height >= pic.Width, vbPRORPortrait, vbPRORLandscape)
		'    PicRatio = pic.Width / pic.Height
		
		'**************************************************************************
		'* Calculate the dimentions of the printable area in HiMetric.
		'**************************************************************************
		'    PrnWidth = .ScaleX(.ScaleWidth, .ScaleMode, vbHiMetric)
		'    PrnHeight = .ScaleY(.ScaleHeight, .ScaleMode, vbHiMetric)
		
		'**************************************************************************
		'* Calculate device independent Width to Height ratio for printer.
		'**************************************************************************
		'    PrnRatio = PrnWidth / PrnHeight
		
		'**************************************************************************
		'* Scale the output to the printable area.
		'**************************************************************************
		'    If (PicRatio >= PrnRatio) Then
		'************************************************************************
		'* Scale picture to fit full width of printable area.
		'************************************************************************
		'      PrnPicWidth = .ScaleX(PrnWidth, vbHiMetric, .ScaleMode)
		'      PrnPicHeight = .ScaleY(PrnWidth / PicRatio, vbHiMetric, .ScaleMode)
		'    Else
		'************************************************************************
		'* Scale picture to fit full height of printable area.
		'************************************************************************
		'      PrnPicHeight = .ScaleY(PrnHeight, vbHiMetric, .ScaleMode)
		'      PrnPicWidth = .ScaleX(PrnHeight * PicRatio, vbHiMetric, .ScaleMode)
		'    End If
		
		'**************************************************************************
		'* Print the picture using the PaintPicture method.
		'**************************************************************************
		'    .PaintPicture pic, 0, 0, PrnPicWidth, PrnPicHeight
		'  End With
		
	End Sub
	
	Public Sub TransBmp(ByRef OutDstDC As Integer, ByRef SrcDC As Integer, ByVal iLeft As Integer, ByVal iTop As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, ByRef DstX As Integer, ByRef DstY As Integer, ByRef TransColor As Integer)
		
		'DstDC- Device context into which image must be
		'drawn transparently
		
		'OutDstDC- Device context into image is actually drawn,
		'even though it is made transparent in terms of DstDC
		
		'Src- Device context of source to be made transparent
		'in color TransColor
		
		'SrcRect- Rectangular region within SrcDC to be made
		'transparent in terms of DstDC, and drawn to OutDstDC
		
		'DstX, DstY - Coordinates in OutDstDC (and DstDC)
		'where the transparent bitmap must go. In most
		'cases, OutDstDC and DstDC will be the same
		
		Dim DstDC As Integer
		Dim nRet As Integer
		Dim MonoMaskDC, hMonoMask As Integer
		Dim MonoInvDC, hMonoInv As Integer
		Dim ResultDstDC, hResultDst As Integer
		Dim ResultSrcDC, hResultSrc As Integer
		Dim hPrevMask, hPrevInv As Integer
		Dim hPrevSrc, hPrevDst As Integer
		
		DstDC = OutDstDC
		
		'create monochrome mask and inverse masks
		MonoMaskDC = CreateCompatibleDC(DstDC)
		MonoInvDC = CreateCompatibleDC(DstDC)
		hMonoMask = CreateBitmap(iWidth, iHeight, 1, 1, 0)
		hMonoInv = CreateBitmap(iWidth, iHeight, 1, 1, 0)
		hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
		hPrevInv = SelectObject(MonoInvDC, hMonoInv)
		
		'create keeper DCs and bitmaps
		ResultDstDC = CreateCompatibleDC(DstDC)
		ResultSrcDC = CreateCompatibleDC(DstDC)
		hResultDst = CreateCompatibleBitmap(DstDC, iWidth, iHeight)
		hResultSrc = CreateCompatibleBitmap(DstDC, iWidth, iHeight)
		hPrevDst = SelectObject(ResultDstDC, hResultDst)
		hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)

        Dim vbSrcCopy As Integer = &HCC0020
        Dim vbSrcAnd As Integer = &H8800C6
        Dim vbSrcInvert As Integer = &H660046
        Dim vbNotSrcCopy As Integer = &H330008

		'copy src to monochrome mask
		Dim OldBC As Integer
		OldBC = SetBkColor(SrcDC, TransColor)
        nRet = BitBlt(MonoMaskDC, 0, 0, iWidth, iHeight, SrcDC, iLeft, iTop, vbSrcCopy)
		TransColor = SetBkColor(SrcDC, OldBC)
		
		'create inverse of mask
        nRet = BitBlt(MonoInvDC, 0, 0, iWidth, iHeight, MonoMaskDC, 0, 0, vbNotSrcCopy)
		
		'get background
        nRet = BitBlt(ResultDstDC, 0, 0, iWidth, iHeight, DstDC, DstX, DstY, vbSrcCopy)
		
		'AND with Monochrome mask
        nRet = BitBlt(ResultDstDC, 0, 0, iWidth, iHeight, MonoMaskDC, 0, 0, vbSrcAnd)
		
		'get overlapper
        nRet = BitBlt(ResultSrcDC, 0, 0, iWidth, iHeight, SrcDC, iLeft, iTop, vbSrcCopy)
		
		'AND with inverse monochrome mask
        nRet = BitBlt(ResultSrcDC, 0, 0, iWidth, iHeight, MonoInvDC, 0, 0, vbSrcAnd)
		
		'XOR these two
        nRet = BitBlt(ResultDstDC, 0, 0, iWidth, iHeight, ResultSrcDC, 0, 0, vbSrcInvert)
		
		'output results
        nRet = BitBlt(OutDstDC, DstX, DstY, iWidth, iHeight, ResultDstDC, 0, 0, vbSrcCopy)
		
		'clean up
		hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
		DeleteObject(hMonoMask)
		
		hMonoInv = SelectObject(MonoInvDC, hPrevInv)
		DeleteObject(hMonoInv)
		
		hResultDst = SelectObject(ResultDstDC, hPrevDst)
		DeleteObject(hResultDst)
		
		hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
		DeleteObject(hResultSrc)
		
		DeleteDC(MonoMaskDC)
		DeleteDC(MonoInvDC)
		DeleteDC(ResultDstDC)
		DeleteDC(ResultSrcDC)
		
	End Sub
End Class