Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Tasks

Namespace ExtractLargeIconFromFile.Shell
    <StructLayout(LayoutKind.Sequential)>
    Public Structure IMAGEINFO
        Public hbmImage As IntPtr
        Public hbmMask As IntPtr
        Public Unused1 As Integer
        Public Unused2 As Integer
        Public rcImage As Shell.RECT
    End Structure
End Namespace
