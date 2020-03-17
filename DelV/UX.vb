Imports System.Runtime.InteropServices

Public Class UX
    Public Shared colorSystemAccent As UInteger = GetImmersiveColorFromColorSetEx(GetImmersiveUserColorSetPreference(False, False), GetImmersiveColorTypeFromName(Marshal.StringToHGlobalUni("ImmersiveSystemAccent")), False, 0)
    Public Shared colorAccent As System.Drawing.Color = System.Drawing.Color.FromArgb((&HFF000000 And colorSystemAccent) >> 24, &HFF And colorSystemAccent, (&HFF00 And colorSystemAccent) >> 8, (&HFF0000 And colorSystemAccent) >> 16)
    Shared strPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase)
    <DllImport("Uxtheme.dll", SetLastError:=True, CharSet:=CharSet.Auto, EntryPoint:="#95")>
    Public Shared Function GetImmersiveColorFromColorSetEx(ByVal dwImmersiveColorSet As UInteger, ByVal dwImmersiveColorType As UInteger, ByVal bIgnoreHighContrast As Boolean, ByVal dwHighContrastCacheMode As UInteger) As UInteger
    End Function
    <DllImport("Uxtheme.dll", SetLastError:=True, CharSet:=CharSet.Auto, EntryPoint:="#96")>
    Public Shared Function GetImmersiveColorTypeFromName(ByVal pName As IntPtr) As UInteger
    End Function
    <DllImport("Uxtheme.dll", SetLastError:=True, CharSet:=CharSet.Auto, EntryPoint:="#98")>
    Public Shared Function GetImmersiveUserColorSetPreference(ByVal bForceCheckRegistry As Boolean, ByVal bSkipCheckOnFail As Boolean) As UInteger
    End Function
    Public Shared inactivecolor As Color = ColorTranslator.FromHtml("#" + Hex(My.Computer.Registry.GetValue("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\DWM", "AccentColorInactive", Nothing)).ToString)
    Public Shared darkness As Double = 1 - ((0.299 * Int(colorAccent.R)) + (0.587 * Int(colorAccent.G)) + (0.114 * Int(colorAccent.B)))
    Public Shared Sub roundobject(oj As Object, rad As Integer)
        Dim p As New Drawing2D.GraphicsPath()
        p.StartFigure()
        p.AddArc(New Rectangle(0, 0, rad, rad), 180, 90)
        p.AddLine(rad, 0, oj.Width - rad, 0)
        p.AddArc(New Rectangle(oj.Width - rad, 0, rad, rad), -90, 90)
        p.AddLine(oj.Width, rad, oj.Width, oj.Height - rad)
        p.AddArc(New Rectangle(oj.Width - rad, oj.Height - rad, rad, rad), 0, 90)
        p.AddLine(oj.Width - rad, oj.Height, rad, oj.Height)
        p.AddArc(New Rectangle(0, oj.Height - rad, rad, rad), 90, 90)
        p.CloseFigure()
        oj.Region = New Region(p)
    End Sub
    Public Shared Sub circularobject(oj As Object)
        Dim p As New Drawing2D.GraphicsPath()
        p.StartFigure()
        p.AddEllipse(0, 0, oj.width, oj.width)
        p.CloseFigure()
        oj.Region = New Region(p)
    End Sub
End Class
