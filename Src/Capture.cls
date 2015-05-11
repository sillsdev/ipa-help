VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

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
      
Private Const NormFont = "Arial"
Private Const NormSize = 9
Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104
Private Const COLOR_BTNFACE = 15
        
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
Public Function CreateBitmapPicture(ByVal hBmp As Long, _
                                    ByVal hPal As Long) As Picture
            
  Dim r As Long
  Dim pic As PicBmp
  Dim IPic As IPicture ' IPicture requires a reference to "Standard OLE Types."
  Dim IID_IDispatch As GUID
         
  '**************************************************************************
  '* Fill in with IDispatch Interface ID.
  '**************************************************************************
  With IID_IDispatch
    .Data1 = &H20400
    .Data4(0) = &HC0
    .Data4(7) = &H46
  End With
         
  '**************************************************************************
  '* Fill Pic with necessary parts.
  '**************************************************************************
  With pic
    .Size = Len(pic)          ' Length of structure.
    .Type = vbPicTypeBitmap   ' Type of Picture (bitmap).
    .hBmp = hBmp              ' Handle to bitmap.
    .hPal = hPal              ' Handle to palette (may be null).
  End With
         
  '**************************************************************************
  '* Create Picture object and return it.
  '**************************************************************************
  r = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
  Set CreateBitmapPicture = IPic
    
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
Public Function CaptureWindow(ByVal hWndSrc As Long, _
            ByVal Client As Boolean, ByVal iLeft As Long, _
            ByVal iTop As Long, ByVal iWidth As Long, _
            ByVal iHeight As Long, Optional vMakeSemiMono) As Picture
            
  Dim bMakeSemiMono As Boolean
  Dim hDCMemory As Long
  Dim hBmp As Long
  Dim hBmpPrev As Long
  Dim r As Long
  Dim hDCSrc As Long
  Dim hPal As Long
  Dim hPalPrev As Long
  Dim RasterCapsScrn As Long
  Dim HasPaletteScrn As Long
  Dim PaletteSizeScrn As Long
  Dim LogPal As LOGPALETTE
         
  On Error Resume Next
  
  bMakeSemiMono = False
  If Not (IsMissing(vMakeSemiMono)) Then bMakeSemiMono = vMakeSemiMono
  
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
  RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)   ' Raster capabilities.
  HasPaletteScrn = RasterCapsScrn And RC_PALETTE       ' Palette support.
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
    Call TransBmp(hDCMemory, hDCSrc, iLeft, iTop, _
                  iWidth, iHeight, 0, 0, GetSysColor(COLOR_BTNFACE))
  Else
    '************************************************************************
    '* Copy the on-screen image into the memory DC. Then remove the new copy
    '* of the on-screen image.
    '************************************************************************
    r = BitBlt(hDCMemory, 0, 0, iWidth, iHeight, hDCSrc, iLeft, iTop, vbSrcCopy)
  End If
  
  hBmp = SelectObject(hDCMemory, hBmpPrev)
         
  '**************************************************************************
  '* If screen has a palette, get back the palette selected previously.
  '**************************************************************************
  If (HasPaletteScrn And (PaletteSizeScrn = 256)) Then _
    hPal = SelectPalette(hDCMemory, hPalPrev, 0)
         
  '**************************************************************************
  '* Release the device context resources back to the system.
  '**************************************************************************
  r = DeleteDC(hDCMemory)
  r = ReleaseDC(hWndSrc, hDCSrc)
         
  '**************************************************************************
  '* Call CreateBitmapPicture to create a picture object from the bitmap and
  '* palette handles. Then return the resulting picture object.
  '**************************************************************************
  Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
         
End Function
      
'**************************************************************************
' CaptureWindowArea - Captures an area of a form.
'
' frmSrc - The Form object to capture.
'
' Returns - A Picture object containing a bitmap of the entire form.
'**************************************************************************
Public Function CaptureWindowArea(frmSrc As Form, iLeft&, iTop&, _
                                  iWidth&, iHeight&, Optional vMakeSemiMono) As Picture
         
  Dim bMakeSemiMono As Boolean
  
  On Error Resume Next
  
  bMakeSemiMono = False
  If Not (IsMissing(vMakeSemiMono)) Then bMakeSemiMono = vMakeSemiMono
  
  '**************************************************************************
  '* Call CaptureWindow to capture the entire form given its window handle
  '* and then return the resulting Picture object.
  '**************************************************************************
  With frmSrc
    Set CaptureWindowArea = CaptureWindow(.hWnd, True, iLeft, iTop, iWidth, _
                                          iHeight, bMakeSemiMono)
  End With
  
End Function

'**************************************************************************
' CaptureScreen - Captures the entire screen.
'
' Returns - A Picture object containing a bitmap of the screen.
'**************************************************************************
Public Function CaptureScreen() As Picture
      
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
                                    Screen.Width \ Screen.TwipsPerPixelX, _
                                    Screen.Height \ Screen.TwipsPerPixelY)
            
End Function
            
'**************************************************************************
' CaptureForm - Captures an entire form including title bar and border.
'
' frmSrc - The Form object to capture.
'
' Returns - A Picture object containing a bitmap of the entire form.
'**************************************************************************
Public Function CaptureForm(frmSrc As Form) As Picture
         
'  On Error Resume Next
  
  '**************************************************************************
  '* Call CaptureWindow to capture the entire form given its window handle
  '* and then return the resulting Picture object.
  '**************************************************************************
'  With frmSrc
'    Set CaptureForm = CaptureWindow(.hWnd, False, 0, 0, _
                                    .ScaleX(.Width, vbTwips, vbPixels), _
                                    .ScaleY(.Height, vbTwips, vbPixels))
'  End With
  
End Function
      
'**************************************************************************
' CaptureClient - Captures the client area of a form.
'
' frmSrc - The Form object to capture.
'
' Returns - A Picture object containing a bitmap of the form's client area.
'**************************************************************************
Public Function CaptureClient(frmSrc As Form) As Picture
         
'  On Error Resume Next
  
  '**************************************************************************
  '* Call CaptureWindow to capture the client area of the form given its
  '* window handle and return the resulting Picture object.
  '**************************************************************************
'  With frmSrc
'    Set CaptureClient = CaptureWindow(.hWnd, True, 0, 0, _
                                      .ScaleX(.ScaleWidth, .ScaleMode, vbPixels), _
                                      .ScaleY(.ScaleHeight, .ScaleMode, vbPixels))
'  End With
  
End Function
      
'**************************************************************************
' CaptureActiveWindow - Captures the currently active window on the screen.
'
' Returns - A Picture object containing a bitmap of the active window.
'**************************************************************************
Public Function CaptureActiveWindow() As Picture
      
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
                                            .Right - .Left, .Bottom - .Top)
'  End With

End Function

Public Sub PrintChart(pic As Picture, sDescription$)

  '***********************************************************
  '***********************************************************

  On Error Resume Next
    
  With Printer
    '*******************************************************
    '* Don't change frmMenu to 'Printer' or this won't
    '* work. It really should be 'Printer' but doing
    '* a Printer.ScaleX method has the effect of beginning
    '* a new page after which setting the orientation won't
    '* work. ARRRGGGGHHH! frmMenu.ScaleX and Printer.ScaleX
    '* appeared to return the same values. I hope that will
    '* always be the case.
    '*******************************************************
    If (frmMenu.ScaleX(.ScaleWidth, .ScaleMode, 8) < pic.Width) Then
      .Orientation = vbPRORLandscape
    Else
      .Orientation = vbPRORPortrait
    End If
    
    .Font = ""
    .Font.Name = "Arial"
    .Font.Size = 10
   
    Printer.Print
    Printer.Print sDescription
    Printer.Print "Date/Time: " & vbTab & Format$(Now, "Medium Date") & " " & _
                                          Format$(Now, "Medium Time")
    .PaintPicture pic, 0, (.TextHeight("X") * 5.5)
    .EndDoc
  End With
  
End Sub

'**************************************************************************
' PrintPicture - Prints a Picture object.
'
' Prn - Destination Printer object.
'
' Pic - Source Picture object.
'**************************************************************************
Public Sub PrintPicture(prn As Printer, pic As Picture)
         
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
Public Sub PrintPictureToFitPage(prn As Printer, pic As Picture)
         
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

Public Sub TransBmp(OutDstDC&, SrcDC&, ByVal iLeft&, ByVal iTop&, ByVal iWidth&, _
                    ByVal iHeight&, DstX&, DstY&, TransColor&)
   
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
   
  Dim DstDC As Long
  Dim nRet As Long
  Dim MonoMaskDC As Long, hMonoMask As Long
  Dim MonoInvDC As Long, hMonoInv As Long
  Dim ResultDstDC As Long, hResultDst As Long
  Dim ResultSrcDC As Long, hResultSrc As Long
  Dim hPrevMask As Long, hPrevInv As Long
  Dim hPrevSrc As Long, hPrevDst As Long

  DstDC = OutDstDC
   
 'create monochrome mask and inverse masks
  MonoMaskDC = CreateCompatibleDC(DstDC)
  MonoInvDC = CreateCompatibleDC(DstDC)
  hMonoMask = CreateBitmap(iWidth, iHeight, 1&, 1&, ByVal 0&)
  hMonoInv = CreateBitmap(iWidth, iHeight, 1&, 1&, ByVal 0&)
  hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
  hPrevInv = SelectObject(MonoInvDC, hMonoInv)
   
 'create keeper DCs and bitmaps
  ResultDstDC = CreateCompatibleDC(DstDC)
  ResultSrcDC = CreateCompatibleDC(DstDC)
  hResultDst = CreateCompatibleBitmap(DstDC, iWidth, iHeight)
  hResultSrc = CreateCompatibleBitmap(DstDC, iWidth, iHeight)
  hPrevDst = SelectObject(ResultDstDC, hResultDst)
  hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)
   
'copy src to monochrome mask
  Dim OldBC As Long
  OldBC = SetBkColor(SrcDC, TransColor)
  nRet = BitBlt(MonoMaskDC, 0&, 0&, iWidth, iHeight, SrcDC, iLeft, iTop, vbSrcCopy)
  TransColor = SetBkColor(SrcDC, OldBC)
   
 'create inverse of mask
  nRet = BitBlt(MonoInvDC, 0&, 0&, iWidth, iHeight, MonoMaskDC, 0&, 0&, vbNotSrcCopy)
   
 'get background
  nRet = BitBlt(ResultDstDC, 0&, 0&, iWidth, iHeight, DstDC, DstX, DstY, vbSrcCopy)
   
 'AND with Monochrome mask
  nRet = BitBlt(ResultDstDC, 0&, 0&, iWidth, iHeight, MonoMaskDC, 0&, 0&, vbSrcAnd)
   
 'get overlapper
  nRet = BitBlt(ResultSrcDC, 0&, 0&, iWidth, iHeight, SrcDC, iLeft, iTop, vbSrcCopy)
   
 'AND with inverse monochrome mask
  nRet = BitBlt(ResultSrcDC, 0&, 0&, iWidth, iHeight, MonoInvDC, 0&, 0&, vbSrcAnd)
   
 'XOR these two
  nRet = BitBlt(ResultDstDC, 0&, 0&, iWidth, iHeight, ResultSrcDC, 0&, 0&, vbSrcInvert)
   
 'output results
  nRet = BitBlt(OutDstDC, DstX, DstY, iWidth, iHeight, ResultDstDC, 0&, 0&, vbSrcCopy)
   
 'clean up
  hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
  DeleteObject hMonoMask

  hMonoInv = SelectObject(MonoInvDC, hPrevInv)
  DeleteObject hMonoInv

  hResultDst = SelectObject(ResultDstDC, hPrevDst)
  DeleteObject hResultDst

  hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
  DeleteObject hResultSrc

  DeleteDC MonoMaskDC
  DeleteDC MonoInvDC
  DeleteDC ResultDstDC
  DeleteDC ResultSrcDC

End Sub
