Imports System.Runtime.InteropServices
Module TranspUX
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bDefaut As Byte, ByVal dwFlags As Long) As Long
    Private Const GWL_EXSTYLE As Long = (-20)
    Private Const LWA_COLORKEY As Long = &H1
    Private Const LWA_Defaut As Long = &H2
    Private Const WS_EX_LAYERED As Long = &H80000
    Public Sub Transparency(ByVal hWnd As Long, Optional ByVal Col As Long = 0, Optional ByVal PcTransp As Byte = 255, Optional ByVal TrMode As Boolean = True)
        Dim DisplayStyle As Long
        Try
            If DisplayStyle <> (DisplayStyle Or WS_EX_LAYERED) Then
                DisplayStyle = (DisplayStyle Or WS_EX_LAYERED)
                Call SetWindowLong(hWnd, GWL_EXSTYLE, DisplayStyle)
            End If
            SetLayeredWindowAttributes(hWnd, Col, PcTransp, IIf(TrMode, LWA_COLORKEY Or LWA_Defaut, LWA_COLORKEY))
        Catch

        End Try
    End Sub

    Public Sub ActiveTransparency(M As Form, d As Boolean, F As Boolean, T_Transparency As Integer)
        If d And F Then

        ElseIf d Then
            Transparency(M.Handle, 0, T_Transparency, True)
        Else
            Transparency(M.Handle,, 255, True)
        End If
    End Sub
End Module
