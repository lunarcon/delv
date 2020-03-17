Imports System.Runtime.InteropServices

Namespace ExtractLargeIconFromFile.Shell

    <StructLayout(LayoutKind.Sequential)>
    Public Structure POINT

        Public X As Integer

        Public Y As Integer

        Public Sub New(ByVal x As Integer, ByVal y As Integer)
            '   MyBase.New
            Me.X = x
            Me.Y = y
        End Sub

        Public Sub New(ByVal pt As System.Drawing.Point)
            Me.New(pt.X, pt.Y)

        End Sub

        Public Overloads Shared Function implicitOperator(ByVal p As POINT) As System.Drawing.Point
            Return New System.Drawing.Point(p.X, p.Y)
        End Function

        Public Overloads Shared Function implicitOperator(ByVal p As System.Drawing.Point) As POINT
            Return New POINT(p.X, p.Y)
        End Function
    End Structure
End Namespace