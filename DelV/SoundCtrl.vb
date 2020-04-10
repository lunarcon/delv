Imports System.Runtime.InteropServices

Public Class SoundCtrl
    Private Sub SoundCtrl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        EnableBlur()
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

    Friend Declare Function SetWindowCompositionAttribute Lib "user32.dll" (hwnd As IntPtr, ByRef data As WindowCompositionAttributeData) As Integer

    Private _blurOpacity As UInteger = 60

    Public Property BlurOpacity As Double
        Get
            Return _blurOpacity
        End Get
        Set(value As Double)
            _blurOpacity = CUInt(Fix(value))
            EnableBlur()
        End Set
    End Property

    Private _blurBackgroundColor As UInteger = &H0

    Friend Sub EnableBlur()
        Dim windowHelper = Me
        Dim accent As New AccentPolicy With {
            .AccentState = AccentState.ACCENT_ENABLE_ACRYLICBLURBEHIND,
            .AccentFlags = 2,
            .GradientColor = (_blurOpacity << 24) Or (_blurBackgroundColor And &HFFFFFFUI)
        }
        If taskbar.IsWindows10() = False Then
            accent.AccentState = AccentState.ACCENT_ENABLE_BLURBEHIND
        End If
        Dim hGc = GCHandle.Alloc(accent, GCHandleType.Pinned)
        Dim data As New WindowCompositionAttributeData With {
            .Attribute = WindowCompositionAttribute.WCA_ACCENT_POLICY,
            .SizeOfData = Marshal.SizeOf(accent),
            .Data = hGc.AddrOfPinnedObject
        }
        SetWindowCompositionAttribute(windowHelper.Handle, data)
        hGc.Free()
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub TrackBar1_Scroll(sender As Object, e As EventArgs)

    End Sub

End Class