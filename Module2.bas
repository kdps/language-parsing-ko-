Attribute VB_Name = "Reflection"
Option Explicit

' Declares...
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    ' Note no palette entry here, not needed
End Type

Private Const BI_RGB = 0&
Private Const BI_RLE4 = 2&
Private Const BI_RLE8 = 1&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCurrentObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Private Const OBJ_BITMAP As Long = 7

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Declare Function UpdateLayeredWindow Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hdcDst As Long, _
    pptDst As Any, _
    psize As Any, _
    ByVal hdcSrc As Long, _
    pptSrc As Any, _
    ByVal crKey As Long, _
    pblend As BLENDFUNCTION, _
    ByVal dwFlags As Long) As Long

' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hdc As Long, _
    pBitmapInfo As BITMAPINFO, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long
    
Private Type SIZEAPI
   cx As Long
   cy As Long
End Type

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type BLENDFUNCTION
   BlendOp As Byte
   BlendFlags As Byte
   SourceConstantAlpha As Byte
   AlphaFormat As Byte
End Type

Private Const AC_SRC_OVER As Long = &H0&
Private Const ULW_COLORKEY As Long = &H1&
Private Const ULW_ALPHA As Long = &H2&
Private Const ULW_OPAQUE As Long = &H4&
Private Const AC_SRC_ALPHA = &H1


Private Const WS_EX_TOPMOST As Long = &H8&
Private Const WS_EX_TRANSPARENT  As Long = &H20&
Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_POPUP = &H80000000
Private Const WS_VISIBLE = &H10000000
Private Const WS_DISABLED As Long = &H8000000

Private Const WM_DESTROY = &H2
Private Const WM_SIZE = &H5
Private Const WM_SIZING = &H214
Private Const WM_MOVING = &H216&
Private Const WM_ENTERSIZEMOVE = &H231&
Private Const WM_EXITSIZEMOVE = &H232&
Private Const WM_MOVE As Long = &H3

Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const SWP_ASYNCWINDOWPOS As Long = &H4000
Private Const SWP_DEFERERASE As Long = &H2000
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_DRAWFRAME As Long = SWP_FRAMECHANGED
Private Const SWP_HIDEWINDOW As Long = &H80
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOCOPYBITS As Long = &H100
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOREDRAW As Long = &H8
Private Const SWP_NOREPOSITION As Long = SWP_NOOWNERZORDER
Private Const SWP_NOSENDCHANGING As Long = &H400
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_SHOWWINDOW As Long = &H40

Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_EXSTYLE As Long = -20
Private Const GWL_WNDPROC As Long = -4

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function IsWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Sub OutputDebugString Lib "kernel32.dll" Alias "OutputDebugStringA" (ByVal lpOutputString As String)
Private Declare Function Ellipse Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private hWndAttach As Long
Private hWndReflection As Long
Private pFnOldWindowProc As Long
Private iTimerID As Long
Private bDying As Boolean

Const WAVE_SIZE As Long = 20 ' Amplitude
Const WAVE_GAP As Long = 5 ' Extra space to try and stop truncation
Const WAVE_FREQ As Long = 4
Private OffsetLookup(0 To 256) As Long

' Graphic objects
Private hDib As Long
Private hBmpOld As Long
Private hdc As Long
Private hDD As Long
Private lPtr As Long
Private bi As BITMAPINFO

Public Sub Attach(ByVal hwnd As Long)
    If hWndAttach <> 0 Then
        Err.Raise vbObjectError, "Reflection::Attach()", "Only one window supported in this version"
        Exit Sub
    End If
    If IsWindow(hwnd) = 0 Then
        Err.Raise vbObjectError, "Reflection::Attach()", "Not a valid window"
        Exit Sub
    End If
    hWndAttach = hwnd
    
    ' Build lookup tables
    Dim i As Long
    Dim two_pi As Double
    two_pi = Atn(1) * 8#
    For i = 0 To 256
        OffsetLookup(i) = Round((0 * WAVE_SIZE - WAVE_GAP) * Sin(CDbl(WAVE_FREQ) * two_pi * CDbl(i) / 2#))
    Next
    
    
    ' Get window info
    
    Dim rc As RECT, cy As Long
    GetWindowRect hWndAttach, rc
    cy = rc.Bottom - rc.Top
    rc.Top = rc.Top + cy
    rc.Bottom = rc.Top + cy
    
    ' Create reflection window
    
    hWndReflection = CreateWindowEx(WS_EX_LAYERED Or WS_EX_TRANSPARENT Or WS_EX_TOOLWINDOW, _
                                    "STATIC", "Reflection", WS_POPUP Or WS_VISIBLE Or WS_DISABLED, _
                                    rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, _
                                    0, 0, App.hInstance, ByVal 0&)
                                    
    If hWndReflection = 0 Then
        hWndAttach = 0
        Err.Raise vbObjectError, "Reflection::Create()", "Could not create window"
        Exit Sub
    End If
    
    ' Subclass the parent window
    pFnOldWindowProc = SetWindowLong(hWndAttach, GWL_WNDPROC, AddressOf AttachedWindow_WindowProc)
    
    ' Set a tiemr
    iTimerID = SetTimer(0, 0, 50, AddressOf TimerProc)
    
    bDying = False
    
End Sub

Public Sub Detach()
    bDying = True
    
    ODS "Detach()..."
    If hWndAttach = 0 Then Exit Sub
    
    ' Kill timer
    KillTimer 0, iTimerID
    
    ' Unsubclass
    ODS "Unsubclassing..."
    SetWindowLong hWndAttach, GWL_WNDPROC, pFnOldWindowProc
    
    ' Destroy our reflection window
    ODS "Destroying our window..."
    If hWndReflection <> 0 Then
        If IsWindow(hWndReflection) Then
            SetWindowLong hWndReflection, GWL_EXSTYLE, GetWindowLong(hWndReflection, GWL_EXSTYLE) And Not (WS_EX_LAYERED)
            DestroyWindow hWndReflection
        End If
        hWndReflection = 0
    End If
    
    ' Delete graphic objects
    ODS "Deleting graphics..."
    ClearUp
    
    hWndAttach = 0
    ODS "Done"
End Sub

Private Function CreateDIB( _
        ByVal hDCRef As Long, _
        ByVal w As Long, _
        ByVal h As Long, _
        ByRef hDib As Long _
    ) As Boolean
    With bi.bmiHeader
        .biSize = Len(bi.bmiHeader)
        .biWidth = w
        .biHeight = h
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        '.biSizeImage = BytesPerScanLine * .biHeight
        .biSizeImage = 4& * .biHeight * .biHeight
    End With
    hDib = CreateDIBSection( _
            hDCRef, _
            bi, _
            DIB_RGB_COLORS, _
            lPtr, _
            0, 0)
    CreateDIB = (hDib <> 0)
End Function

Private Function Create( _
        ByVal w As Long, _
        ByVal h As Long _
    ) As Boolean
    
   ' Don't bother creating if it's the same size as what we already have
   ' This could be further optimized to keep larger bitmaps and not
   ' re-create smaller ones.
   If w = bi.bmiHeader.biWidth And h = bi.bmiHeader.biHeight Then
        Create = True
        Exit Function
   End If
    
   ClearUp
   hdc = CreateCompatibleDC(0)
   If (hdc <> 0) Then
       If (CreateDIB(hdc, w, h, hDib)) Then
           hBmpOld = SelectObject(hdc, hDib)
           Create = True
       Else
           DeleteObject hdc
           hdc = 0
       End If
   End If
End Function
Private Sub ClearUp()
    If (hdc <> 0) Then
        If (hDib <> 0) Then
            SelectObject hdc, hBmpOld
            DeleteObject hDib
        End If
        DeleteObject hdc
    End If
    hdc = 0: hDib = 0: hBmpOld = 0: lPtr = 0
End Sub

Private Function AttachedWindow_WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If bDying = False Then
    Static bInSizeMove As Boolean
    Dim rc As RECT
    Select Case uMsg
    Case WM_MOVING
        ' Retrieve the rectangle
        'CopyMemory rc, ByVal lParam, Len(rc)
        GetWindowRect hwnd, rc
        Reposition rc
    Case WM_SIZING
        ' Retrieve the rectangle
        'CopyMemory rc, ByVal lParam, Len(rc)
        GetWindowRect hwnd, rc
        Reposition rc
    Case WM_SIZE, WM_MOVE
        If Not bInSizeMove Then
            GetWindowRect hwnd, rc
            Reposition rc
        End If
    Case WM_ENTERSIZEMOVE
        bInSizeMove = True
    Case WM_EXITSIZEMOVE
        bInSizeMove = False
    Case WM_DESTROY
        rc.Left = pFnOldWindowProc
        Detach
        AttachedWindow_WindowProc = CallWindowProc(rc.Left, hwnd, uMsg, wParam, lParam)
        Exit Function
    End Select
    End If
    AttachedWindow_WindowProc = CallWindowProc(pFnOldWindowProc, hwnd, uMsg, wParam, lParam)
    
End Function

Private Function Reposition(rc As RECT)
    If bDying Then Exit Function
    
    Dim cy As Long
    cy = rc.Bottom - rc.Top
    rc.Top = rc.Top + cy
    
    If cy > 128 * 2 Then
        rc.Bottom = rc.Top + 128
    Else
        rc.Bottom = rc.Top + cy \ 2
    End If
            
    rc.Left = rc.Left - WAVE_SIZE
    rc.Right = rc.Right + WAVE_SIZE
    
    SetWindowPos hWndReflection, 0, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, SWP_NOOWNERZORDER Or SWP_NOACTIVATE
    
    Create rc.Right - rc.Left, rc.Bottom - rc.Top
    Redraw
End Function

Public Function Redraw()
    If bDying Then Exit Function
    
    Dim rc As RECT
    Dim si As SIZEAPI
    GetWindowRect hWndReflection, rc
    si.cx = rc.Right - rc.Left
    si.cy = rc.Bottom - rc.Top
    Dim bf As BLENDFUNCTION
    bf.BlendOp = AC_SRC_OVER
    bf.BlendFlags = 0
    bf.AlphaFormat = AC_SRC_ALPHA
    bf.SourceConstantAlpha = 192    ' Not fully opaque at any point
    Dim pt As POINTAPI
    pt.x = 0
    pt.y = 0
    
    Create rc.Right - rc.Left, rc.Bottom - rc.Top
    Render
    
    UpdateLayeredWindow hWndReflection, ByVal 0&, ByVal 0&, si, hdc, pt, 0, bf, ULW_ALPHA
End Function

Private Function Render()
    Static LastRenderTime As Long
    Dim ThisRenderTime As Long
    
    ThisRenderTime = timeGetTime
    
    If bDying Then Exit Function
    
    Dim SrcBits() As Byte
    Dim DstBits() As Byte
    Dim x As Long, y As Long
    Dim tSA As SAFEARRAY2D

    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = bi.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = bi.bmiHeader.biWidth * 4 ' Bytes per scanline
        .pvData = lPtr
    End With
    CopyMemory ByVal VarPtrArray(DstBits()), VarPtr(tSA), 4
    
    
    GetWindowBits hWndAttach, SrcBits()
    'Dim rc As RECT
    'GetWindowRect hWndAttach, rc
    'ReDim WinBits(0 To ((rc.Right - rc.Left) * 4) - 1, 0 To (rc.Bottom - rc.Top) - 1)
    
    
    Dim SrcX As Long, SrcY As Long
    'SrcY = UBound(SrcBits, 2)
    SrcY = 2
    
    Dim alpha As Long, alpha_delta As Long
    alpha_delta = Int(255# / CDbl(bi.bmiHeader.biHeight) + 0.5)
    alpha = 0
    
    Dim phase As Double
    phase = CDbl(ThisRenderTime Mod 1000) * 0.001
    phase = phase * 3.1415927 * 2#
    
    Dim phase2 As Long
    phase2 = (ThisRenderTime Mod 1000)
    phase2 = phase2 \ 16
    
    Dim pos As Double
    Dim pos2 As Long

    For y = bi.bmiHeader.biHeight - 1 To 0 Step -1
        pos2 = y * 255 \ bi.bmiHeader.biHeight
        alpha = pos2
        For x = 0 To bi.bmiHeader.biWidth * 4 - 1 Step 4
            'alpha = 127
            
            SrcX = x \ 4
            SrcX = SrcX - WAVE_SIZE + ((255 - pos2) * OffsetLookup(((y + phase2)) Mod 256)) \ 255
            SrcX = SrcX * 4
            
            If SrcX < 0 Or SrcX > UBound(SrcBits, 1) Then
                DstBits(x + 3, y) = 0   ' Alpha
                DstBits(x + 2, y) = 0
                DstBits(x + 1, y) = 0
                DstBits(x + 0, y) = 0
            Else
                DstBits(x + 3, y) = alpha   ' Alpha
                DstBits(x + 2, y) = (SrcBits(SrcX + 2, SrcY) * alpha) \ 255
                DstBits(x + 1, y) = (SrcBits(SrcX + 1, SrcY) * alpha) \ 255
                DstBits(x + 0, y) = (SrcBits(SrcX + 0, SrcY) * alpha) \ 255
            End If
        Next
        SrcY = SrcY + 1
        alpha = alpha + alpha_delta
    Next

    CopyMemory ByVal VarPtrArray(DstBits()), 0&, 4
    
    LastRenderTime = ThisRenderTime
End Function

Private Function GetWindowBits(ByVal hwnd As Long, ByRef WinBits() As Byte)
    Dim rc As RECT
    Dim hWinDC As Long
    Dim hWinBmp As Long
    Dim hWinOldBmp As Long
    Dim tSA As SAFEARRAY2D
    Dim biWin As BITMAPINFO
        
    ' Get dimensions
    GetWindowRect hwnd, rc
    
    ' Fill in bitmap structure
    With biWin.bmiHeader
        .biSize = Len(bi.bmiHeader)
        .biWidth = rc.Right - rc.Left
        .biHeight = rc.Bottom - rc.Top
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = BI_RGB
        .biSizeImage = 4& * .biHeight * .biHeight
    End With
    
    ReDim WinBits(0 To ((rc.Right - rc.Left) * 4) - 1, 0 To (rc.Bottom - rc.Top) - 1)

Dim ret As Long
Dim hTempDC As Long
    
    hWinDC = GetWindowDC(hwnd)
    hTempDC = CreateCompatibleDC(0)
    hWinBmp = CreateCompatibleBitmap(hWinDC, biWin.bmiHeader.biWidth, biWin.bmiHeader.biHeight)
    hWinOldBmp = SelectObject(hTempDC, hWinBmp)
    BitBlt hTempDC, 0, 0, biWin.bmiHeader.biWidth, biWin.bmiHeader.biHeight, hWinDC, 0, 0, vbSrcCopy
    SelectObject hTempDC, hWinOldBmp
    '
    ret = GetDIBits(hWinDC, hWinBmp, 0, rc.Bottom - rc.Top, WinBits(0, 0), biWin, DIB_RGB_COLORS)
    '
    ReleaseDC hwnd, hWinDC
    DeleteDC hTempDC
    DeleteObject hWinBmp
    
End Function

Public Function ODS(ParamArray s())
    OutputDebugString Join(s, ", ") & vbCrLf
End Function

Private Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    If bDying Then Exit Sub
    Dim rc As RECT
    GetWindowRect hWndAttach, rc
    Reposition rc
End Sub



