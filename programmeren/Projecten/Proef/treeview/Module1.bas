Attribute VB_Name = "Module1"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias _
    "RtlMoveMemory" (lpDest As Any, lpSrc As Any, _
    ByVal dwLen As Long)

Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" _
    Alias "SetWindowLongA" (ByVal hWnd As Long, _
    ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_WNDPROC = (-4)
Private Declare Function CallWindowProc Lib "user32" _
    Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, ByVal Msg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetProp Lib "user32" Alias _
    "GetPropA" (ByVal hWnd As Long, _
    ByVal lpString As String) As Long
Private Declare Function SetProp Lib "user32" Alias _
    "SetPropA" (ByVal hWnd As Long, _
    ByVal lpString As String, ByVal hData As Long) As Long
Private Declare Function RemoveProp Lib "user32" Alias _
    "RemovePropA" (ByVal hWnd As Long, _
    ByVal lpString As String) As Long

Private m_hpalHalftone As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function CreateSolidBrush Lib "gdi32" _
    (ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" _
    (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal dwRop As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" _
    (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" _
    (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" _
    (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" _
    (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" _
    (ByVal hDC As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function SelectObject Lib "gdi32" _
    (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" _
    (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" _
    (ByVal hWnd As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" _
    (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" _
    (ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal nPlanes As Long, ByVal nBitCount As Long, _
    lpBits As Any) As Long
Private Declare Function GetBkColor Lib "gdi32" _
    (ByVal hDC As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" _
    (ByVal hDC As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" _
    (ByVal hDC As Long, ByVal hPalette As Long, _
    ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" _
    (ByVal hDC As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" _
    (ByVal lOleColor As Long, ByVal lHPalette As Long, _
    lColorRef As Long) As Long
Private Declare Function DrawIconEx Lib "user32" _
    (ByVal hDC As Long, ByVal xLeft As Long, _
    ByVal yTop As Long, ByVal hIcon As Long, _
    ByVal cxWidth As Long, ByVal cyHeight As Long, _
    ByVal istepIfAniCur As Long, _
    ByVal hbrFlickerFreeDraw As Long, _
    ByVal diFlags As Long) As Long
Private Declare Function FillRect Lib "user32" _
    (ByVal hDC As Long, lpRect As RECT, _
    ByVal hBrush As Long) As Long

'DrawIconEx Flags
Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = &H3
Private Const DI_COMPAT = &H4
Private Const DI_DEFAULTSIZE = &H8

'Raster Operation Codes
Private Const DSna = &H220326 '0x00220326

'VB Errors
Private Const giINVALID_PICTURE As Integer = 481


Public Function TranslateColor(inCol As Long) As Long

'A simple wrapper for OleTranslateColor
Dim retCol As Long
OleTranslateColor inCol, 0&, retCol
TranslateColor = retCol

End Function

Public Sub PaintNormalStdPic(ByVal hdcDest As Long, _
    ByVal xDest As Long, _
    ByVal yDest As Long, _
    ByVal width As Long, _
    ByVal Height As Long, _
    ByVal picSource As Picture, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    Optional ByVal hPal As Long = 0)

    Dim hdcTemp As Long
    Dim hPalOld As Long
    Dim hbmMemSrcOld As Long
    Dim hdcScreen As Long
    Dim hbmMemSrc As Long
    'Validate that a bitmap was passed in
    If picSource Is Nothing Then GoTo PaintNormalStdPic_InvalidParam
    Select Case picSource.Type
        Case vbPicTypeBitmap
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            hdcScreen = GetDC(0&)
            'Create a DC to select bitmap into
            hdcTemp = CreateCompatibleDC(hdcScreen)
            hPalOld = SelectPalette(hdcTemp, hPal, True)
            RealizePalette hdcTemp
            'Select bitmap into DC
            hbmMemSrcOld = SelectObject(hdcTemp, picSource.Handle)
            'Copy to destination DC
            BitBlt hdcDest, xDest, yDest, width, Height, hdcTemp, xSrc, ySrc, vbSrcCopy
            'Cleanup
            SelectObject hdcTemp, hbmMemSrcOld
            SelectPalette hdcTemp, hPalOld, True
            RealizePalette hdcTemp
            DeleteDC hdcTemp
            ReleaseDC 0&, hdcScreen
        Case vbPicTypeIcon
            'Create a bitmap and select it into an DC
            'Draw Icon onto DC
            DrawIconEx hdcDest, xDest, yDest, picSource.Handle, 0, 0, 0&, 0&, DI_NORMAL
        Case Else
            GoTo PaintNormalStdPic_InvalidParam
    End Select
    Exit Sub
PaintNormalStdPic_InvalidParam:
    Err.Raise giINVALID_PICTURE

End Sub

Public Sub PaintTransparentDC(ByVal hdcDest As Long, _
    ByVal xDest As Long, _
    ByVal yDest As Long, _
    ByVal width As Long, _
    ByVal Height As Long, _
    ByVal hdcSrc As Long, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal clrMask As OLE_COLOR, _
    Optional ByVal hPal As Long = 0)

    Dim hdcMask As Long        'HDC of the created mask image
    Dim hdcColor As Long       'HDC of the created color image
    Dim hbmMask As Long        'Bitmap handle to the mask image
    Dim hbmColor As Long       'Bitmap handle to the color image
    Dim hbmColorOld As Long
    Dim hbmMaskOld As Long
    Dim hPalOld As Long
    Dim hdcScreen As Long
    Dim hdcScnBuffer As Long   'Buffer to do all work on
    Dim hbmScnBuffer As Long
    Dim hbmScnBufferOld As Long
    Dim hPalBufferOld As Long
    Dim lMaskColor As Long
    
    hdcScreen = GetDC(0&)
    'Validate palette
    If hPal = 0 Then
        hPal = m_hpalHalftone
    End If
    OleTranslateColor clrMask, hPal, lMaskColor
    
    'Create a color bitmap to server as a copy of the destination
    'Do all work on this bitmap and then copy it back over the destination
    'when it's done.
    hbmScnBuffer = CreateCompatibleBitmap(hdcScreen, width, Height)
    'Create DC for screen buffer
    hdcScnBuffer = CreateCompatibleDC(hdcScreen)
    hbmScnBufferOld = SelectObject(hdcScnBuffer, hbmScnBuffer)
    hPalBufferOld = SelectPalette(hdcScnBuffer, hPal, True)
    RealizePalette hdcScnBuffer
    'Copy the destination to the screen buffer
    BitBlt hdcScnBuffer, 0, 0, width, Height, hdcDest, xDest, yDest, vbSrcCopy
    
    'Create a (color) bitmap for the cover (can't use CompatibleBitmap with
    'hdcSrc, because this will create a DIB section if the original bitmap
    'is a DIB section)
    hbmColor = CreateCompatibleBitmap(hdcScreen, width, Height)
    'Now create a monochrome bitmap for the mask
    hbmMask = CreateBitmap(width, Height, 1, 1, ByVal 0&)
    'First, blt the source bitmap onto the cover.  We do this first
    'and then use it instead of the source bitmap
    'because the source bitmap may be
    'a DIB section, which behaves differently than a bitmap.
    '(Specifically, copying from a DIB section to a monochrome bitmap
    'does a nearest-color selection rather than painting based on the
    'backcolor and forecolor.
    hdcColor = CreateCompatibleDC(hdcScreen)
    hbmColorOld = SelectObject(hdcColor, hbmColor)
    hPalOld = SelectPalette(hdcColor, hPal, True)
    RealizePalette hdcColor
    'In case hdcSrc contains a monochrome bitmap, we must set the destination
    'foreground/background colors according to those currently set in hdcSrc
    '(because Windows will associate these colors with the two monochrome colors)
    SetBkColor hdcColor, GetBkColor(hdcSrc)
    SetTextColor hdcColor, GetTextColor(hdcSrc)
    BitBlt hdcColor, 0, 0, width, Height, hdcSrc, xSrc, ySrc, vbSrcCopy
    'Paint the mask.  What we want is white at the transparent color
    'from the source, and black everywhere else.
    hdcMask = CreateCompatibleDC(hdcScreen)
    hbmMaskOld = SelectObject(hdcMask, hbmMask)

    'When bitblt'ing from color to monochrome, Windows sets to 1
    'all pixels that match the background color of the source DC.  All
    'other bits are set to 0.
    SetBkColor hdcColor, lMaskColor
    SetTextColor hdcColor, vbWhite
    BitBlt hdcMask, 0, 0, width, Height, hdcColor, 0, 0, vbSrcCopy
    'Paint the rest of the cover bitmap.
    '
    'What we want here is black at the transparent color, and
    'the original colors everywhere else.  To do this, we first
    'paint the original onto the cover (which we already did), then we
    'AND the inverse of the mask onto that using the DSna ternary raster
    'operation (0x00220326 - see Win32 SDK reference, Appendix, "Raster
    'Operation Codes", "Ternary Raster Operations", or search in MSDN
    'for 00220326).  DSna [reverse polish] means "(not SRC) and DEST".
    '
    'When bitblt'ing from monochrome to color, Windows transforms all white
    'bits (1) to the background color of the destination hdc.  All black (0)
    'bits are transformed to the foreground color.
    SetTextColor hdcColor, vbBlack
    SetBkColor hdcColor, vbWhite
    BitBlt hdcColor, 0, 0, width, Height, hdcMask, 0, 0, DSna
    'Paint the Mask to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, width, Height, hdcMask, 0, 0, vbSrcAnd
    'Paint the Color to the Screen buffer
    BitBlt hdcScnBuffer, 0, 0, width, Height, hdcColor, 0, 0, vbSrcPaint
    'Copy the screen buffer to the screen
    BitBlt hdcDest, xDest, yDest, width, Height, hdcScnBuffer, 0, 0, vbSrcCopy
    'All done!
    DeleteObject SelectObject(hdcColor, hbmColorOld)
    SelectPalette hdcColor, hPalOld, True
    RealizePalette hdcColor
    DeleteDC hdcColor
    DeleteObject SelectObject(hdcScnBuffer, hbmScnBufferOld)
    SelectPalette hdcScnBuffer, hPalBufferOld, True
    RealizePalette hdcScnBuffer
    DeleteDC hdcScnBuffer

    DeleteObject SelectObject(hdcMask, hbmMaskOld)
    DeleteDC hdcMask
    ReleaseDC 0&, hdcScreen
End Sub

Public Sub PaintTransparentStdPic(ByVal hdcDest As Long, _
    ByVal xDest As Long, _
    ByVal yDest As Long, _
    ByVal width As Long, _
    ByVal Height As Long, _
    ByVal picSource As Picture, _
    ByVal xSrc As Long, _
    ByVal ySrc As Long, _
    ByVal clrMask As OLE_COLOR, _
    Optional ByVal hPal As Long = 0)

    Dim hdcSrc As Long         'HDC that the source bitmap is selected into
    Dim hbmMemSrcOld As Long
    Dim hbmMemSrc As Long
    Dim udtRect As RECT
    Dim hbrMask As Long
    Dim lMaskColor As Long
    Dim hdcScreen As Long
    Dim hPalOld As Long
    'Verify that the passed picture is a Bitmap
    If picSource Is Nothing Then
        GoTo PaintTransparentStdPic_InvalidParam
    End If

    Select Case picSource.Type
        Case vbPicTypeBitmap
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            'Select passed picture into an HDC
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrcOld = SelectObject(hdcSrc, picSource.Handle)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw the bitmap
            PaintTransparentDC hdcDest, xDest, yDest, width, Height, hdcSrc, xSrc, ySrc, clrMask, hPal

            SelectObject hdcSrc, hbmMemSrcOld
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
        Case vbPicTypeIcon
            'Create a bitmap and select it into an DC
            hdcScreen = GetDC(0&)
            'Validate palette
            If hPal = 0 Then
                hPal = m_hpalHalftone
            End If
            hdcSrc = CreateCompatibleDC(hdcScreen)
            hbmMemSrc = CreateCompatibleBitmap(hdcScreen, width, Height)
            hbmMemSrcOld = SelectObject(hdcSrc, hbmMemSrc)
            hPalOld = SelectPalette(hdcSrc, hPal, True)
            RealizePalette hdcSrc
            'Draw Icon onto DC
            udtRect.Bottom = Height
            udtRect.Right = width
            OleTranslateColor clrMask, 0&, lMaskColor
            hbrMask = CreateSolidBrush(lMaskColor)
            FillRect hdcSrc, udtRect, hbrMask
            DeleteObject hbrMask
            DrawIconEx hdcSrc, 0, 0, picSource.Handle, 0, 0, 0, 0, DI_NORMAL
            'Draw Transparent image
            PaintTransparentDC hdcDest, xDest, yDest, width, Height, hdcSrc, 0, 0, lMaskColor, hPal
            'Clean up
            DeleteObject SelectObject(hdcSrc, hbmMemSrcOld)
            SelectPalette hdcSrc, hPalOld, True
            RealizePalette hdcSrc
            DeleteDC hdcSrc
            ReleaseDC 0&, hdcScreen
        Case Else
            GoTo PaintTransparentStdPic_InvalidParam
    End Select
    Exit Sub
PaintTransparentStdPic_InvalidParam:
    'Err.Raise giINVALID_PICTURE
    Exit Sub
End Sub

Public Sub Subclass(frm As Form, tv As Treeview)

'Subclass the TreeView and store an object
'pointer to the form.

Dim lProc As Long

If GetProp(tv.hWnd, "VBTWndProc") <> 0 Then
    Exit Sub
End If

lProc = GetWindowLong(tv.hWnd, GWL_WNDPROC)
SetProp tv.hWnd, "VBTWndProc", lProc
SetProp tv.hWnd, "VBTWndPtr", ObjPtr(frm)

SetWindowLong tv.hWnd, GWL_WNDPROC, _
    AddressOf WndProcTV

End Sub

Public Sub UnSubclass(tv As Treeview)

Dim lProc As Long

lProc = GetProp(tv.hWnd, "VBTWndProc")
If lProc = 0 Then
    Exit Sub
End If

SetWindowLong tv.hWnd, GWL_WNDPROC, lProc
RemoveProp tv.hWnd, "VBTWndProc"
RemoveProp tv.hWnd, "VBTWndPtr"

End Sub

Public Function WndProcTV(ByVal hWnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long

On Error Resume Next

Dim lProc As Long
Dim lPtr As Long
Dim tmpForm As Form
Dim bUseRetVal As Boolean
Dim lRetVal As Long

bUseRetVal = False
lProc = GetProp(hWnd, "VBTWndProc")
lPtr = GetProp(hWnd, "VBTWndPtr")
'Copy the form's object pointer into an
'object variable and call the message handler.
CopyMemory tmpForm, lPtr, 4
tmpForm.TreeViewMessage hWnd, wMsg, wParam, lParam, _
    lRetVal, bUseRetVal
CopyMemory tmpForm, 0&, 4

If bUseRetVal = True Then
    'Use the return value from the form's
    'handler
    WndProcTV = lRetVal
Else
    'Pass on to original wndproc
    WndProcTV = CallWindowProc(lProc, hWnd, wMsg, _
        wParam, lParam)
End If

End Function


