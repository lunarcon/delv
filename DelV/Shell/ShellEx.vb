Imports System
Imports System.IO
Imports System.Runtime.InteropServices

Namespace ExtractLargeIconFromFile
    Public Class ShellEx
        Private Const SHGFI_SMALLICON As Integer = &H1
        Private Const SHGFI_LARGEICON As Integer = &H0
        Private Const SHIL_JUMBO As Integer = &H4
        Private Const SHIL_EXTRALARGE As Integer = &H2
        Private Const WM_CLOSE As Integer = &H10

        Public Enum IconSizeEnum
            SmallIcon16 = SHGFI_SMALLICON
            MediumIcon32 = SHGFI_LARGEICON
            LargeIcon48 = SHIL_EXTRALARGE
            ExtraLargeIcon = SHIL_JUMBO
        End Enum

        <DllImport("user32")>
        Private Shared Function SendMessage(ByVal handle As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr

        End Function
        <DllImport("shell32.dll")>
        Private Shared Function SHGetImageList(ByVal iImageList As Integer, ByRef riid As Guid, <Out> ByRef ppv As Shell.IImageList) As Integer

        End Function
        <DllImport("Shell32.dll")>
        Public Shared Function SHGetFileInfo(ByVal pszPath As String, ByVal dwFileAttributes As Integer, ByRef psfi As Shell.SHFILEINFO, ByVal cbFileInfo As Integer, ByVal uFlags As UInteger) As Integer
        End Function
        <DllImport("user32")>
        Public Shared Function DestroyIcon(ByVal hIcon As IntPtr) As Integer

        End Function

        Public Shared Function GetBitmapFromFolderPath(ByVal filepath As String, ByVal iconsize As IconSizeEnum) As Bitmap
            Dim hIcon As IntPtr = GetIconHandleFromFolderPath(filepath, iconsize)
            Return getBitmapFromIconHandle(hIcon)
        End Function

        Public Shared Function GetBitmapFromFilePath(ByVal filepath As String, ByVal iconsize As IconSizeEnum) As System.Drawing.Bitmap
            Dim hIcon As IntPtr = GetIconHandleFromFilePath(filepath, iconsize)
            Return getBitmapFromIconHandle(hIcon)
        End Function

        Public Shared Function GetBitmapFromPath(ByVal filepath As String, ByVal iconsize As IconSizeEnum) As System.Drawing.Bitmap
            Dim hIcon As IntPtr = IntPtr.Zero

            If Directory.Exists(filepath) Then
                hIcon = GetIconHandleFromFolderPath(filepath, iconsize)
            Else

                If File.Exists(filepath) Then
                    hIcon = GetIconHandleFromFilePath(filepath, iconsize)
                End If
            End If

            Return getBitmapFromIconHandle(hIcon)
        End Function

        Private Shared Function getBitmapFromIconHandle(ByVal hIcon As IntPtr) As System.Drawing.Bitmap
            If hIcon = IntPtr.Zero Then Throw New System.IO.FileNotFoundException()
            Dim myIcon = System.Drawing.Icon.FromHandle(hIcon)
            Dim bitmap = myIcon.ToBitmap()
            myIcon.Dispose()
            DestroyIcon(hIcon)
            SendMessage(hIcon, WM_CLOSE, IntPtr.Zero, IntPtr.Zero)
            Return bitmap
        End Function

        Private Shared Function GetIconHandleFromFilePath(ByVal filepath As String, ByVal iconsize As IconSizeEnum) As IntPtr
            Dim shinfo = New Shell.SHFILEINFO()
            Const SHGFI_SYSICONINDEX As UInteger = &H4000
            Const FILE_ATTRIBUTE_NORMAL As Integer = &H80
            Dim flags As UInteger = SHGFI_SYSICONINDEX
            Return getIconHandleFromFilePathWithFlags(filepath, iconsize, shinfo, FILE_ATTRIBUTE_NORMAL, flags)
        End Function

        Private Shared Function GetIconHandleFromFolderPath(ByVal folderpath As String, ByVal iconsize As IconSizeEnum) As IntPtr
            Dim shinfo = New Shell.SHFILEINFO()
            Const SHGFI_ICON As UInteger = &H100
            Const SHGFI_USEFILEATTRIBUTES As UInteger = &H10
            Const FILE_ATTRIBUTE_DIRECTORY As Integer = &H10
            Dim flags As UInteger = SHGFI_ICON Or SHGFI_USEFILEATTRIBUTES
            Return getIconHandleFromFilePathWithFlags(folderpath, iconsize, shinfo, FILE_ATTRIBUTE_DIRECTORY, flags)
        End Function

        Private Shared Function getIconHandleFromFilePathWithFlags(ByVal filepath As String, ByVal iconsize As IconSizeEnum, ByRef shinfo As Shell.SHFILEINFO, ByVal fileAttributeFlag As Integer, ByVal flags As UInteger) As IntPtr
            Const ILD_TRANSPARENT As Integer = 1
            Dim retval = SHGetFileInfo(filepath, fileAttributeFlag, shinfo, Marshal.SizeOf(shinfo), flags)
            If retval = 0 Then Throw (New System.IO.FileNotFoundException())
            Dim iconIndex = shinfo.iIcon
            Dim iImageListGuid = New Guid("46EB5926-582E-4017-9FDF-E8998DAA0950")
            Dim iml As Shell.IImageList
            Dim hres = SHGetImageList(CInt(iconsize), iImageListGuid, iml)
            Dim hIcon = IntPtr.Zero
            hres = iml.GetIcon(iconIndex, ILD_TRANSPARENT, hIcon)
            Return hIcon
        End Function
    End Class
End Namespace
