Imports System

Namespace ExtractLargeIconFromFile.Shell
    Public Structure IMAGELISTDRAWPARAMS
        Public cbSize As Integer
        Public himl As IntPtr
        Public i As Integer
        Public hdcDst As IntPtr
        Public x As Integer
        Public y As Integer
        Public cx As Integer
        Public cy As Integer
        Public xBitmap As Integer
        Public yBitmap As Integer
        Public rgbBk As Integer
        Public rgbFg As Integer
        Public fStyle As Integer
        Public dwRop As Integer
        Public fState As Integer
        Public Frame As Integer
        Public crEffect As Integer
    End Structure
End Namespace
