Imports System
Imports System.Runtime.InteropServices

Namespace ExtractLargeIconFromFile.Shell
    Public Structure SHFILEINFO
        Public hIcon As IntPtr
        Public iIcon As Integer
        Public dwAttributes As UInteger
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=254)>
        Public szDisplayName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public szTypeName As String
    End Structure
End Namespace
