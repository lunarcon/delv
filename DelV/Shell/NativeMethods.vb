Imports System.Runtime.InteropServices
Imports System.Text
Friend Class NativeMethods
    <DllImport("user32.dll")>
    Public Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll")>
    Public Shared Function GetForegroundWindow() As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function FindWindowEx(ByVal hWndParent As IntPtr, ByVal hWndChildAfter As IntPtr, ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetWindowRect(ByVal hwnd As IntPtr, <Out> ByRef lpRect As RECT) As Boolean
    End Function
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)>
    Public Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure
    Public Shared Function GetActiveWindowTitle() As IntPtr
        Const nChars As Integer = 256
        Dim Buff As StringBuilder = New StringBuilder(nChars)
        Dim handle As IntPtr = GetForegroundWindow()
        If GetWindowText(handle, Buff, nChars) > 0 Then
            Return handle
        End If
        Return Nothing
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function PrintWindow(ByVal hwnd As IntPtr, ByVal hDC As IntPtr, ByVal nFlags As UInteger) As Boolean
    End Function
    <DllImport("user32.dll")>
    Public Shared Function WindowFromPoint(ByVal p As Point) As IntPtr
    End Function
    <DllImport("user32.dll")>
    Public Shared Function SetCapture(ByVal hWnd As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll")>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll")>
    Public Shared Function LoadIcon(ByVal hInstance As IntPtr, ByVal lpIconName As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll", EntryPoint:="GetClassLong")>
    Public Shared Function GetClassLongPtr32(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As UInteger
    End Function
    <DllImport("user32.dll", EntryPoint:="GetClassLongPtr")>
    Public Shared Function GetClassLongPtr64(ByVal hWnd As IntPtr, ByVal nIndex As Integer) As IntPtr
    End Function
    Public Shared WM_GETICON As UInteger = &H7F
    Public Shared WM_CLOSE As UInteger = &H10
    Public Shared ICON_SMALL2 As IntPtr = New IntPtr(2)
    Public Shared SHGFI_LARGEICON As Integer = &H0
    Public Shared IDI_APPLICATION As IntPtr = New IntPtr(&H7F00)
    Public Shared GCL_HICON As Integer = -14
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function GetWindowTextLength(ByVal hWnd As IntPtr) As Integer
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function GetWindowText(ByVal hWnd As IntPtr, ByVal lpString As StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function
    <DllImport("user32.dll")>
    Public Shared Function IsWindow(ByVal hWnd As IntPtr) As Boolean
    End Function
    <DllImport("user32")>
    Public Shared Function GetClassName(ByVal hWnd As IntPtr, ByVal lpString As StringBuilder, ByVal nMaxCount As Integer) As Integer
    End Function
    <DllImport("user32.dll")>
    Public Shared Function EnumWindows(ByVal lpEnumFunc As EnumWindowsProc, ByVal lParam As IntPtr) As Integer
    End Function
    Public Delegate Function EnumWindowsProc(ByVal hWnd As IntPtr, ByVal lParam As IntPtr) As Boolean
    <DllImport("USER32.DLL")>
    Public Shared Function IsWindowVisible(ByVal IntPtr As IntPtr) As Boolean
    End Function
    Public Shared Function GetLnkTarget(lnkPath As String) As String
        Dim shl = New Shell32.Shell()
        ' Move this to class scope
        lnkPath = System.IO.Path.GetFullPath(lnkPath)
        Dim dir = shl.[NameSpace](System.IO.Path.GetDirectoryName(lnkPath))
        Dim itm = dir.Items().Item(System.IO.Path.GetFileName(lnkPath))
        Dim lnk = DirectCast(itm.GetLink, Shell32.ShellLinkObject)
        Return lnk.Target.Path
    End Function
End Class

