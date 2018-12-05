Option Strict On
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Drawing.Imaging

Namespace Screen
    Public Class ScreenCapture

        Public Sub New()
            CheckKey()
        End Sub



        ''' <summary>
        ''' Creates an Image object containing a screen shot of the entire desktop
        ''' </summary>
        ''' <returns></returns>
        Public Function CaptureScreen() As Image
            Return CaptureWindow(User32.GetDesktopWindow())
        End Function
        ''' <summary>
        ''' Creates an Image object containing a screen shot of the active window
        ''' </summary>
        ''' <returns></returns>
        Public Function CaptureActiveWindow() As Image
            Return CaptureWindow(User32.GetForegroundWindow())
        End Function
        ''' <summary>
        ''' Creates an Image object containing a screen shot of a specific window
        ''' </summary>
        ''' <param name="handle">The handle to the window. (In windows forms, this is obtained by the Handle property)</param>
        ''' <returns></returns>
        Public Function CaptureWindow(ByVal handle As IntPtr) As Image
            '' get te hDC of the target window
            Dim hdcSrc As IntPtr = User32.GetWindowDC(handle)
            '' get the size
            Dim windowRect As User32.RECT = New User32.RECT()
            User32.GetWindowRect(handle, windowRect)
            Dim width As Integer = windowRect.right - windowRect.left
            Dim height As Integer = windowRect.bottom - windowRect.top
            '' create a device context we can copy to
            Dim hdcDest As IntPtr = GDI32.CreateCompatibleDC(hdcSrc)
            '' create a bitmap we can copy it to,
            '' using GetDeviceCaps to get the width/height
            Dim hBitmap As IntPtr = GDI32.CreateCompatibleBitmap(hdcSrc, width, height)
            '' select the bitmap object
            Dim hOld As IntPtr = GDI32.SelectObject(hdcDest, hBitmap)
            '' bitblt over
            GDI32.BitBlt(hdcDest, 0, 0, width, height, hdcSrc, 0, 0, CopyPixelOperation.SourceCopy)
            '' restore selection
            GDI32.SelectObject(hdcDest, hOld)
            '' clean up 
            GDI32.DeleteDC(hdcDest)
            User32.ReleaseDC(handle, hdcSrc)
            '' get a .NET image object for it
            Dim img As Image = Image.FromHbitmap(hBitmap)
            '' free up the Bitmap object
            GDI32.DeleteObject(hBitmap)
            Return img
        End Function

        Public Function CaptureWindow(ByVal left As Integer, ByVal top As Integer, ByVal width As Integer, ByVal height As Integer, ByVal handle As IntPtr) As Image

            '' get te hDC of the target window
            Dim hdcSrc As IntPtr = User32.GetWindowDC(Nothing)
            ' '' get the size
            'Dim windowRect As User32.RECT = New User32.RECT()
            'User32.GetWindowRect(handle, windowRect)
            'Dim width As Integer = CInt(windowRect.right - windowRect.left)
            'Dim height As Integer = CInt(windowRect.bottom - windowRect.top)
            '' create a device context we can copy to
            Dim hdcDest As IntPtr = GDI32.CreateCompatibleDC(hdcSrc)
            '' create a bitmap we can copy it to,
            '' using GetDeviceCaps to get the width/height
            Dim hBitmap As IntPtr = GDI32.CreateCompatibleBitmap(hdcSrc, width, height)
            '' select the bitmap object
            Dim hOld As IntPtr = GDI32.SelectObject(hdcDest, hBitmap)
            '' bitblt over
            GDI32.BitBlt(hdcDest, 0, 0, width, height, hdcSrc, left, top, (CopyPixelOperation.SourceCopy Or CopyPixelOperation.CaptureBlt))
            '' restore selection
            GDI32.SelectObject(hdcDest, hOld)
            '' clean up 
            GDI32.DeleteDC(hdcDest)
            User32.ReleaseDC(handle, hdcSrc)
            '' get a .NET image object for it
            Dim img As Image = Image.FromHbitmap(hBitmap)

            '' free up the Bitmap object
            GDI32.DeleteObject(hBitmap)
            Return img
        End Function

        ''' <summary>
        ''' Captures a screen shot of a specific window, and saves it to a file
        ''' </summary>
        ''' <param name="handle"></param>
        ''' <param name="filename"></param>
        ''' <param name="format"></param>
        Public Sub CaptureWindowToFile(ByVal handle As IntPtr, ByVal filename As String, ByVal format As ImageFormat)

            Dim img As Image = CaptureWindow(handle)
            img.Save(filename, format)
        End Sub
        ''' <summary>
        ''' Captures a screen shot of the entire desktop, and saves it to a file
        ''' </summary>
        ''' <param name="filename"></param>
        ''' <param name="format"></param>
        Public Sub CaptureScreenToFile(ByVal filename As String, ByVal format As ImageFormat)

            Dim img As Image = CaptureScreen()
            img.Save(filename, format)
        End Sub

        ''' <summary>
        ''' Helper class containing Gdi32 API functions
        ''' </summary>
        Private Class GDI32

            Public Enum TernaryRasterOperations
                SRCCOPY = &HCC0020 ' dest = source
                SRCPAINT = &HEE0086 ' dest = source OR dest
                SRCAND = &H8800C6 ' dest = source AND dest
                SRCINVERT = &H660046 ' dest = source XOR dest
                SRCERASE = &H440328 ' dest = source AND (NOT dest )
                NOTSRCCOPY = &H330008 ' dest = (NOT source)
                NOTSRCERASE = &H1100A6 ' dest = (NOT src) AND (NOT dest) 
                MERGECOPY = &HC000CA ' dest = (source AND pattern)
                MERGEPAINT = &HBB0226 ' dest = (NOT source) OR dest
                PATCOPY = &HF00021 ' dest = pattern
                PATPAINT = &HFB0A09 ' dest = DPSnoo
                PATINVERT = &H5A0049 ' dest = pattern XOR dest
                DSTINVERT = &H550009 ' dest = (NOT dest)
                BLACKNESS = &H42 ' dest = BLACK
                WHITENESS = &HFF0062 ' dest = WHITE
            End Enum

            Public Const SRCCOPY As Integer = &HCC0020 '' BitBlt dwRop parameter
            ''' <summary>
            '''    Performs a bit-block transfer of the color data corresponding to a
            '''    rectangle of pixels from the specified source device context into
            '''    a destination device context.
            ''' </summary>
            ''' <param name="hdc">Handle to the destination device context.</param>
            ''' <param name="nXDest">The leftmost x-coordinate of the destination rectangle (in pixels).</param>
            ''' <param name="nYDest">The topmost y-coordinate of the destination rectangle (in pixels).</param>
            ''' <param name="nWidth">The width of the source and destination rectangles (in pixels).</param>
            ''' <param name="nHeight">The height of the source and the destination rectangles (in pixels).</param>
            ''' <param name="hdcSrc">Handle to the source device context.</param>
            ''' <param name="nXSrc">The leftmost x-coordinate of the source rectangle (in pixels).</param>
            ''' <param name="nYSrc">The topmost y-coordinate of the source rectangle (in pixels).</param>
            ''' <param name="dwRop">A raster-operation code.</param>
            ''' <returns>
            '''    <c>true</c> if the operation succeeded, <c>false</c> otherwise.
            ''' </returns>
            Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hdc As IntPtr, ByVal nXDest As Integer, ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As CopyPixelOperation) As Boolean

            Public Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hdc As IntPtr, ByVal nWidth As Integer, ByVal nHeight As Integer) As IntPtr

            Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As IntPtr) As IntPtr

            Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As IntPtr) As Boolean

            Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As IntPtr) As Boolean

            Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As IntPtr, ByVal hgdiobj As IntPtr) As IntPtr
        End Class

        ''' <summary>
        ''' Helper class containing User32 API functions
        ''' </summary>
        Private Class User32

            <StructLayout(LayoutKind.Sequential)> Public Structure RECT
                Public left As Integer
                Public top As Integer
                Public right As Integer
                Public bottom As Integer
            End Structure

            Public Declare Auto Function GetDesktopWindow Lib "user32.dll" () As IntPtr

            Public Declare Function GetWindowDC Lib "user32" (ByVal hwnd As IntPtr) As IntPtr

            Public Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal hDC As IntPtr) As IntPtr

            Public Overloads Declare Function GetWindowRect Lib "User32" Alias "GetWindowRect" (ByVal hWnd As IntPtr, ByRef lpRect As RECT) As Integer

            Public Declare Function GetForegroundWindow Lib "user32" () As IntPtr


        End Class

    End Class


End Namespace
