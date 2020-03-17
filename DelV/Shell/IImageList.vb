Imports System
Imports System.Runtime.InteropServices

Namespace ExtractLargeIconFromFile.Shell

    <ComImportAttribute(),
     GuidAttribute("46EB5926-582E-4017-9FDF-E8998DAA0950"),
     InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IImageList

        <PreserveSig()>
        Function Add(ByVal hbmImage As IntPtr, ByVal hbmMask As IntPtr, ByRef pi As Integer) As Integer

        <PreserveSig()>
        Function ReplaceIcon(ByVal i As Integer, ByVal hicon As IntPtr, ByRef pi As Integer) As Integer

        <PreserveSig()>
        Function SetOverlayImage(ByVal iImage As Integer, ByVal iOverlay As Integer) As Integer

        <PreserveSig()>
        Function Replace(ByVal i As Integer, ByVal hbmImage As IntPtr, ByVal hbmMask As IntPtr) As Integer

        <PreserveSig()>
        Function AddMasked(ByVal hbmImage As IntPtr, ByVal crMask As Integer, ByRef pi As Integer) As Integer

        <PreserveSig()>
        Function Draw(ByRef pimldp As Shell.IMAGELISTDRAWPARAMS) As Integer

        <PreserveSig()>
        Function Remove(ByVal i As Integer) As Integer

        <PreserveSig()>
        Function GetIcon(ByVal i As Integer, ByVal flags As Integer, ByRef picon As IntPtr) As Integer
    End Interface
End Namespace