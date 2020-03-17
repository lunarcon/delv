Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.IO
Imports DelV.UX
Imports System.Text
Imports System.Management

Public Class acrylic_panel
    Dim timelinestr As String = ""
    Dim pinnedapps As String = CStr("C:\Users\" & GetUserName() & "\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\")
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Const CS_DROPSHADOW As Integer = &H20000
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H80
            cp.ClassStyle = cp.ClassStyle Or CS_DROPSHADOW
            Return cp
        End Get
    End Property
    <DllImport("shell32", EntryPoint:="#261", CharSet:=CharSet.Unicode, PreserveSig:=False)>
    Public Shared Sub GetUserTilePath(username As String, whatever As UInt32, picpath As StringBuilder, maxLength As Integer)
    End Sub
    Declare Function GetUserName Lib "advapi32.dll" Alias _
       "GetUserNameA" (ByVal lpBuffer As String,
       ByRef nSize As Integer) As Integer
    Public Function GetUserName() As String
        Dim iReturn As Integer
        Dim userName As String
        userName = New String(CChar(" "), 50)
        iReturn = GetUserName(userName, 50)
        GetUserName = userName.Substring(0, userName.IndexOf(Chr(0)))
    End Function
    Public Function GetUserTilePath(username As String) As String
        Dim sb As StringBuilder
        sb = New StringBuilder(1000)
        GetUserTilePath(username, 2147483648, sb, sb.Capacity)
        Return sb.ToString()
    End Function
    Public Function GetUserTile(username As String) As Image
        Return Image.FromFile(GetUserTilePath(username))
    End Function
    Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVallpClassName As String, ByVal lpWindowName As String) As Integer
    Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByValhWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    <DllImport("user32.dll", EntryPoint:="SendMessage", SetLastError:=True)>
    Private Shared Function SendMessage(hWnd As IntPtr, Msg As Int32, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function
    Public Const SWP_HIDEWINDOW = &H80
    Public Const SWP_SHOWWINDOW = &H40
    Public Sub ShowDesktop()
        keybd_event(VK_LWIN, 0, 0, 0)
        keybd_event(77, 0, 0, 0)
        keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
    End Sub
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    Private Const KEYEVENTF_KEYUP = &H2
    Private Const VK_LWIN = &H5B
    Private Function GetMainModuleFilepath(ByVal processId As Integer) As String
        Dim wmiQueryString As String = "SELECT ProcessId, ExecutablePath FROM Win32_Process WHERE ProcessId = " & processId

        Using searcher = New ManagementObjectSearcher(wmiQueryString)

            Using results = searcher.[Get]()
                Dim mo As ManagementObject = results.Cast(Of ManagementObject)().FirstOrDefault()

                If mo IsNot Nothing Then
                    Return CStr(mo("ExecutablePath"))
                End If
            End Using
        End Using

        Return Nothing
    End Function
    Private Sub loadprocesses()
        Dim processlist = Process.GetProcesses
        Dim idx As Integer = 0
        For Each c As Control In AppsPanel.Controls
            Dim cont As Boolean = False
            For Each process In processlist
                If Not String.IsNullOrEmpty(process.MainWindowTitle) Then
                    If process.Id = Val(c.Tag) Then
                        cont = True
                    End If
                End If
            Next
            If cont = False Then
                AppsPanel.Controls.Remove(c)
            End If
        Next
        For Each process In processlist
            If Not String.IsNullOrEmpty(process.MainWindowTitle) And Not (process.MainWindowTitle.Contains("input") Or process.MainWindowTitle.Contains("Input")) Then
                Try
                    Dim btx As New taskbar_button
                    btx.Tag = process.Id
                    btx.Button1.Tag = process.MainWindowTitle.ToString
                    Dim contains = False
                    For Each c As Control In AppsPanel.Controls
                        If c.Tag = btx.Tag Then
                            contains = True
                        End If
                    Next
                    If contains = False Then
                        Dim xath = GetMainModuleFilepath(process.Id)
                        Dim appicon = AddIcon(xath, 1)
                        btx.Label1.Visible = True
                        btx.BackColor = Color.Transparent
                        btx.Button1.BackColor = Color.FromArgb(5, 0, 0, 0)
                        btx.Size = New Size(50, 42)
                        btx.Name = process.ProcessName
                        btx.Dock = DockStyle.Left
                        btx.Button1.Image = appicon
                        Rescale(0.45, appicon, btx)
                        btx.Margin = New Padding(10, 0, 0, 0)
                        AppsPanel.Controls.Add(btx)
                        idx += 1
                    End If
                Catch
                End Try
            End If
        Next
    End Sub
    Private Sub Rescale(scale_factor As Single, bm_source As Bitmap, source As taskbar_button)

        Dim bm_dest As New Bitmap(
        CInt(bm_source.Width * scale_factor),
        CInt(bm_source.Height * scale_factor))
        Dim gr_dest As Graphics = Graphics.FromImage(bm_dest)

        gr_dest.DrawImage(bm_source, 0, 0,
        bm_dest.Width + 1,
        bm_dest.Height + 1)

        source.Button1.Image = bm_dest
    End Sub
    Private Function AddIcon(fname As String, mode As Integer)
        Dim size
        If mode = 1 Then
            size = CType(ExtractLargeIconFromFile.ShellEx.IconSizeEnum.LargeIcon48, ExtractLargeIconFromFile.ShellEx.IconSizeEnum)
        ElseIf mode = 2 Then
            size = CType(ExtractLargeIconFromFile.ShellEx.IconSizeEnum.MediumIcon32, ExtractLargeIconFromFile.ShellEx.IconSizeEnum)
        End If
#Disable Warning BC42104
        Return ExtractLargeIconFromFile.ShellEx.GetBitmapFromFilePath(fname, size)
#Enable Warning BC42104
    End Function
    Dim l1 As New ListBox
    Private Sub get_pinned_apps(filepath As String)
        If Not filepath Is Nothing AndAlso Directory.Exists(filepath) Then
            For Each file As String In Directory.GetFiles(filepath)
                If file.Contains(".lnk") Then
                    Dim btx As New taskbar_button
                    Dim appicon = AddIcon(file, 1)
                    btx.Tag = file
                    l1.Items.Add(NativeMethods.GetLnkTarget(file.ToString))
                    btx.BackColor = Color.Transparent
                    'btx.Button1.FlatStyle = FlatStyle.Flat
                    btx.Button1.BackColor = Color.FromArgb(5, 0, 0, 0)
                    btx.Size = New Size(50, 42)
                    btx.Name = Path.GetFileNameWithoutExtension(file)
                    btx.Dock = DockStyle.Left
                    Rescale(0.45, appicon, btx)
                    btx.Margin = New Padding(10, 0, 0, 0)
                    pnned_panel.Controls.Add(btx)
                End If
            Next
        End If
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        BringToFront()
        Width = My.Computer.Screen.WorkingArea.Width
        Location = New Point(0, My.Computer.Screen.WorkingArea.Height)
        Dim intReturn As Integer = FindWindow("Shell_traywnd", "")
        SetWindowPos(intReturn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
        EnableBlur()
        get_pinned_apps(pinnedapps)
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        SafeClose()
    End Sub

    Public Sub SafeClose()
        Dim intReturn As Integer = FindWindow("Shell_traywnd", "")
        SetWindowPos(intReturn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    End Sub

    Friend Enum AccentState
        ACCENT_DISABLED = 0
        ACCENT_ENABLE_GRADIENT = 1
        ACCENT_ENABLE_TRANSPARENTGRADIENT = 2
        ACCENT_ENABLE_BLURBEHIND = 3
        ACCENT_ENABLE_ACRYLICBLURBEHIND = 4
        ACCENT_INVALID_STATE = 5
    End Enum

    <StructLayout(LayoutKind.Sequential)>
    Friend Structure AccentPolicy
        Public AccentState As AccentState
        Public AccentFlags As Integer
        Public GradientColor As Integer
        Public AnimationId As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Friend Structure WindowCompositionAttributeData
        Public Attribute As WindowCompositionAttribute
        Public Data As IntPtr
        Public SizeOfData As Integer
    End Structure

    Friend Enum WindowCompositionAttribute
        ' ...
        WCA_ACCENT_POLICY = 19
        ' ...
    End Enum

    <DllImport("user32.dll")>
    Friend Shared Function SetWindowCompositionAttribute(ByVal hwnd As IntPtr, ByRef data As WindowCompositionAttributeData) As Integer
    End Function

    Private _blurOpacity As UInteger = 50

    Public Property BlurOpacity As Double
        Get
            Return _blurOpacity
        End Get
        Set(ByVal value As Double)
            _blurOpacity = CUInt(value)
            EnableBlur()
        End Set
    End Property
    Private _blurBackgroundColor As UInteger = &HFFFFFF

    Friend Sub EnableBlur()
        Dim windowHelper = Me
        Dim accent = New AccentPolicy()
        accent.AccentState = AccentState.ACCENT_ENABLE_ACRYLICBLURBEHIND
        accent.GradientColor = _blurBackgroundColor
        accent.AccentFlags = &H20 Or &H40 Or &H80 Or &H100
        Dim accentStructSize = Marshal.SizeOf(accent)
        Dim accentPtr = Marshal.AllocHGlobal(accentStructSize)
        Marshal.StructureToPtr(accent, accentPtr, False)
        Dim data = New WindowCompositionAttributeData()
        data.Attribute = WindowCompositionAttribute.WCA_ACCENT_POLICY
        data.SizeOfData = accentStructSize
        data.Data = accentPtr
        SetWindowCompositionAttribute(windowHelper.Handle, data)
        Marshal.FreeHGlobal(accentPtr)
    End Sub


    Public Function GetWindowIcon(ByVal windowHandle As IntPtr) As Image
        Try
            Dim hIcon As IntPtr = Nothing
            hIcon = NativeMethods.SendMessage(windowHandle, NativeMethods.WM_GETICON, NativeMethods.SHGFI_LARGEICON, IntPtr.Zero)
            If hIcon = IntPtr.Zero Then hIcon = GetClassLongPtr(windowHandle, NativeMethods.GCL_HICON)
            If hIcon = IntPtr.Zero Then hIcon = NativeMethods.LoadIcon(IntPtr.Zero, CType(&H7F00, IntPtr))
            If hIcon <> IntPtr.Zero Then
                Return New Bitmap(Icon.FromHandle(hIcon).ToBitmap(), 23, 23)
            Else
                Return Nothing
            End If

        Catch ola_amigo_si_dora As Exception
            Return Nothing
        End Try
    End Function
    Private Function GetCOnnectionStrength()
        Dim proc As ProcessStartInfo = New ProcessStartInfo("cmd.exe")
        Dim pr As Process
        proc.CreateNoWindow = True
        proc.UseShellExecute = False
        proc.RedirectStandardInput = True
        proc.RedirectStandardOutput = True
        pr = Process.Start(proc)
        pr.StandardInput.WriteLine("netsh wlan show interfaces")
        pr.StandardInput.Close()
        Dim a As String = ""
        a = pr.StandardOutput.ReadToEnd
        a = a.Remove(0, a.LastIndexOf("Name") - 1)
        a.Remove(a.LastIndexOf("Hosted"), a.Length - a.LastIndexOf("Hosted"))
        pr.StandardOutput.Close()
        Dim parts As String() = a.ToString.Split(vbNewLine)
        Dim syslist As New ListBox
        For Each part As String In parts
            syslist.Items.Add(part)
        Next
        For i = 0 To syslist.Items.Count
            Try
                syslist.Items(i) = syslist.Items(i).ToString.Remove(0, syslist.Items(i).ToString.LastIndexOf(" : ") + 2)
            Catch ec As Exception
            End Try
        Next
        Return syslist.Items.Item(16)
#Disable Warning BC42105
    End Function
#Enable Warning BC42105
    Private Function GetClassLongPtr(ByVal windowHandle As IntPtr, ByVal nIndex As Integer) As IntPtr
        If IntPtr.Size = 4 Then
            Return New IntPtr(CLng(NativeMethods.GetClassLongPtr32(windowHandle, nIndex)))
        Else
            Return NativeMethods.GetClassLongPtr64(windowHandle, nIndex)
        End If
    End Function

    Private Sub AppsPanel_Paint(sender As Object, e As PaintEventArgs) Handles AppsPanel.Paint

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        loadprocesses()
        systime.Text = Now.ToShortTimeString + vbNewLine + Now.Date.ToShortDateString
    End Sub

    Private Sub Fore_panel_Paint(sender As Object, e As PaintEventArgs) Handles fore_panel.Paint

    End Sub
    Private Sub Panel2_Click(sender As Object, e As EventArgs) Handles Panel2.Click
        SafeClose()
        Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        keybd_event(VK_LWIN, 0, 0, 0)
        keybd_event(Keys.A, 0, 0, 0)
        keybd_event(Keys.A, 0, KEYEVENTF_KEYUP, 0)
        keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        keybd_event(VK_LWIN, 0, 0, 0)
        keybd_event(Keys.Tab, 0, 0, 0)
        keybd_event(Keys.Tab, 0, KEYEVENTF_KEYUP, 0)
        keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        keybd_event(VK_LWIN, 0, 0, 0)
        keybd_event(Keys.S, 0, 0, 0)
        keybd_event(Keys.S, 0, KEYEVENTF_KEYUP, 0)
        keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        keybd_event(VK_LWIN, 0, 0, 0)
        keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
    End Sub

    Private Sub Peek_Click(sender As Object, e As EventArgs) Handles Peek.Click
        keybd_event(VK_LWIN, 0, 0, 0)
        keybd_event(Keys.Oemcomma, 0, 0, 0)
        keybd_event(Keys.Oemcomma, 0, KEYEVENTF_KEYUP, 0)
        keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
    End Sub

    Private Sub Systime_Click(sender As Object, e As EventArgs) Handles systime.Click
        Shell("explorer.exe outlookcal:")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        keybd_event(VK_LWIN, 0, 0, 0)
        keybd_event(Keys.C, 0, 0, 0)
        keybd_event(Keys.C, 0, KEYEVENTF_KEYUP, 0)
        keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Shell("explorer.exe ms-availablenetworks:")
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Shell("explorer.exe ms-whiteboard-cmd:")
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Dim st As String = GetCOnnectionStrength()
        Dim s As Integer = Val(st.Replace("%", ""))
        If s <= 25 And s >= 0 Then
            Button5.Text = ""
        ElseIf s > 25 And s <= 50 Then
            Button5.Text = ""
        ElseIf s > 50 And s <= 75 Then
            Button5.Text = ""
        ElseIf s > 70 And s <= 100 Then
            Button5.Text = ""
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        SafeClose()
        Close()
    End Sub
End Class
