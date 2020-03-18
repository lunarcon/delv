Imports System.Runtime.InteropServices

Module VolumeUtilities
    Function GetMasterVolume() As Single
        Dim deviceEnumerator As IMMDeviceEnumerator = CType((New MMDeviceEnumerator()), IMMDeviceEnumerator)
        Dim speakers As IMMDevice
        Const eRender As Integer = 0
        Const eMultimedia As Integer = 1
        deviceEnumerator.GetDefaultAudioEndpoint(eRender, eMultimedia, speakers)
        Dim o As Object
        speakers.Activate(GetType(IAudioEndpointVolume).GUID, 0, IntPtr.Zero, o)
        Dim aepv As IAudioEndpointVolume = CType(o, IAudioEndpointVolume)
        Dim volume As Single = aepv.GetMasterVolumeLevelScalar()
        Marshal.ReleaseComObject(aepv)
        Marshal.ReleaseComObject(speakers)
        Marshal.ReleaseComObject(deviceEnumerator)
        Return volume
    End Function

    <ComImport>
    <Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")>
    Private Class MMDeviceEnumerator
    End Class

    <Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IAudioEndpointVolume
        Sub _VtblGap1_6()
        Function GetMasterVolumeLevelScalar() As Single
    End Interface

    <Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IMMDeviceEnumerator
        Sub _VtblGap1_1()
        <PreserveSig>
        Function GetDefaultAudioEndpoint(ByVal dataFlow As Integer, ByVal role As Integer, <Out> ByRef ppDevice As IMMDevice) As Integer
    End Interface

    <Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Private Interface IMMDevice
        <PreserveSig>
        Function Activate(
        <MarshalAs(UnmanagedType.LPStruct)> ByVal iid As Guid, ByVal dwClsCtx As Integer, ByVal pActivationParams As IntPtr, <Out>
        <MarshalAs(UnmanagedType.IUnknown)> ByRef ppInterface As Object) As Integer
    End Interface
End Module
