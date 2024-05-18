
# Define the necessary COM interfaces and enums
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class AudioHelper
{
    public static float GetPeakValue()
    {
        IMMDeviceEnumerator enumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
        IMMDevice speakers = enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
        IAudioMeterInformation meter = (IAudioMeterInformation)speakers.Activate(typeof(IAudioMeterInformation).GUID, 0, IntPtr.Zero);
        return meter.GetPeakValue();
    }

    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    private class MMDeviceEnumerator
    {
    }

    private enum EDataFlow
    {
        eRender,
        eCapture,
        eAll,
    }

    private enum ERole
    {
        eConsole,
        eMultimedia,
        eCommunications,
    }

    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    private interface IMMDeviceEnumerator
    {
        void NotNeeded();
        IMMDevice GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role);
    }

    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    private interface IMMDevice
    {
        [return: MarshalAs(UnmanagedType.IUnknown)]
        object Activate([MarshalAs(UnmanagedType.LPStruct)] Guid iid, int dwClsCtx, IntPtr pActivationParams);
    }

    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("C02216F6-8C67-4B5B-9D00-D008E73E0064")]
    private interface IAudioMeterInformation
    {
        float GetPeakValue();
    }
}
"@

# Function to check if Windows is playing sound
function Is-WindowsPlayingSound {
    try {
        $peakValue = [AudioHelper]::GetPeakValue()
        return $peakValue -gt 1E-08
    } catch {
        Write-Error "An error occurred: $_"
        return $false
    }
}

# Example usage
$result = Is-WindowsPlayingSound
Write-Output "Is Windows Playing Sound: $result"
