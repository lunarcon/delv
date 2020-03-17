Imports DelV.UX
Imports System.Runtime.InteropServices
Public Class RoundedModularPanel
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Const CS_DROPSHADOW As Integer = &H20000
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ClassStyle = cp.ClassStyle Or CS_DROPSHADOW
            Return cp
        End Get
    End Property
    Private Sub RoundedModularPanel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        roundobject(Me, 10)

    End Sub

    Private Sub RoundedModularPanel_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        roundobject(Me, 10)
    End Sub

End Class
