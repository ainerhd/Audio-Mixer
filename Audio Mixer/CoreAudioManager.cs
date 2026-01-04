using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Audio_Mixer
{
    public sealed class CoreAudioManager
    {
        private readonly Dictionary<string, IAudioEndpointVolume> volumeCache = new();

        public IReadOnlyList<AudioDevice> GetOutputDevices()
        {
            var devices = new List<AudioDevice>();

            var enumerator = CreateDeviceEnumerator();
            int hr = enumerator.EnumAudioEndpoints(EDataFlow.eRender, (uint)DeviceState.Active, out var collection);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);

            hr = collection.GetCount(out var count);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);

            for (uint i = 0; i < count; i++)
            {
                hr = collection.Item(i, out var device);
                if (hr != 0) Marshal.ThrowExceptionForHR(hr);

                hr = device.GetId(out var id);
                if (hr != 0) Marshal.ThrowExceptionForHR(hr);

                var name = GetDeviceName(device);
                devices.Add(new AudioDevice(id, name));
            }

            return devices;
        }

        public void SetDeviceVolume(string deviceId, float volumeScalar)
        {
            if (volumeScalar < 0f) volumeScalar = 0f;
            if (volumeScalar > 1f) volumeScalar = 1f;

            if (!volumeCache.TryGetValue(deviceId, out var endpoint))
            {
                var enumerator = CreateDeviceEnumerator();

                int hr = enumerator.GetDevice(deviceId, out var device);
                if (hr != 0) Marshal.ThrowExceptionForHR(hr);

                endpoint = GetEndpointVolume(device);
                volumeCache[deviceId] = endpoint;
            }

            int hr2 = endpoint.SetMasterVolumeLevelScalar(volumeScalar, Guid.Empty);
            if (hr2 != 0) Marshal.ThrowExceptionForHR(hr2);
        }

        private static IMMDeviceEnumerator CreateDeviceEnumerator()
        {
            // Deutlich robuster als selbst CoCreateInstance aufzurufen
            return (IMMDeviceEnumerator)new MMDeviceEnumeratorComObject();
        }

        private static string GetDeviceName(IMMDevice device)
        {
            int hr = device.OpenPropertyStore(StgmAccess.Read, out var propertyStore);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);

            var friendlyNameKey = PropertyKeys.PKEY_Device_FriendlyName;
            hr = propertyStore.GetValue(ref friendlyNameKey, out var propVariant);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);

            string? name = propVariant.GetValue();
            propVariant.Clear();

            return name ?? string.Empty;
        }

        private static IAudioEndpointVolume GetEndpointVolume(IMMDevice device)
        {
            var iid = typeof(IAudioEndpointVolume).GUID;
            int hr = device.Activate(ref iid, ClsCtx.InprocServer, IntPtr.Zero, out var obj);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);

            return (IAudioEndpointVolume)obj;
        }
    }

    public sealed record AudioDevice(string Id, string Name);

    // COM coclass
    [ComImport]
    [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    internal class MMDeviceEnumeratorComObject
    {
    }

    internal enum EDataFlow
    {
        eRender = 0,
        eCapture = 1,
        eAll = 2,
    }

    [Flags]
    internal enum DeviceState : uint
    {
        Active = 0x00000001,
        Disabled = 0x00000002,
        NotPresent = 0x00000004,
        Unplugged = 0x00000008,
        All = 0x0000000F,
    }

    [Flags]
    internal enum ClsCtx
    {
        InprocServer = 0x1,
        InprocHandler = 0x2,
        LocalServer = 0x4,
        InprocServer16 = 0x8,
        RemoteServer = 0x10,
        InprocHandler16 = 0x20,
        Reserved1 = 0x40,
        Reserved2 = 0x80,
        Reserved3 = 0x100,
        Reserved4 = 0x200,
        NoCodeDownload = 0x400,
        Reserved5 = 0x800,
        NoCustomMarshal = 0x1000,
        EnableCodeDownload = 0x2000,
        NoFailureLog = 0x4000,
        DisableActivateAsActivator = 0x8000,
        EnableActivateAsActivator = 0x10000,
        FromDefaultContext = 0x20000,
        Inproc = InprocServer | InprocHandler,
        Server = InprocServer | LocalServer | RemoteServer,
        All = Server | InprocHandler,
    }

    internal enum StgmAccess
    {
        Read = 0x00000000,
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct PropertyKey
    {
        public Guid FormatId;
        public int PropertyId;

        public PropertyKey(Guid formatId, int propertyId)
        {
            FormatId = formatId;
            PropertyId = propertyId;
        }
    }

    internal static class PropertyKeys
    {
        public static readonly PropertyKey PKEY_Device_FriendlyName =
            new(new Guid("A45C254E-DF1C-4EFD-8020-67D146A850E0"), 14);
    }

    [StructLayout(LayoutKind.Explicit)]
    internal struct PropVariant
    {
        [FieldOffset(0)]
        private ushort vt;

        [FieldOffset(8)]
        private IntPtr pointerValue;

        public string? GetValue()
        {
            // VT_LPWSTR = 31
            return vt == 31 ? Marshal.PtrToStringUni(pointerValue) : null;
        }

        public void Clear()
        {
            PropVariantClear(ref this);
        }

        [DllImport("ole32.dll")]
        private static extern int PropVariantClear(ref PropVariant pvar);
    }

    [ComImport]
    [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceEnumerator
    {
        // WICHTIG: dwStateMask muss uint sein + MarshalAs Interface beim Out
        int EnumAudioEndpoints(
            EDataFlow dataFlow,
            uint dwStateMask,
            [MarshalAs(UnmanagedType.Interface)] out IMMDeviceCollection ppDevices);

        int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppEndpoint);

        int GetDevice([MarshalAs(UnmanagedType.LPWStr)] string pwstrId, out IMMDevice ppDevice);

        int RegisterEndpointNotificationCallback(IntPtr pClient);
        int UnregisterEndpointNotificationCallback(IntPtr pClient);
    }

    internal enum ERole
    {
        eConsole = 0,
        eMultimedia = 1,
        eCommunications = 2,
    }

    [ComImport]
    [Guid("0BD7A1BE-7A1A-44DB-8397-C0F6EEDCC8DD")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDeviceCollection
    {
        int GetCount(out uint pcDevices);
        int Item(uint nDevice, out IMMDevice ppDevice);
    }

    [ComImport]
    [Guid("D666063F-1587-4E43-81F1-B948E807363F")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IMMDevice
    {
        int Activate(ref Guid iid, ClsCtx dwClsCtx, IntPtr pActivationParams,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);

        int OpenPropertyStore(StgmAccess stgmAccess, out IPropertyStore ppProperties);

        int GetId([MarshalAs(UnmanagedType.LPWStr)] out string ppstrId);

        int GetState(out DeviceState pdwState);
    }

    [ComImport]
    [Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPropertyStore
    {
        int GetCount(out uint cProps);
        int GetAt(uint iProp, out PropertyKey pkey);
        int GetValue(ref PropertyKey key, out PropVariant pv);
        int SetValue(ref PropertyKey key, ref PropVariant propvar);
        int Commit();
    }

    [ComImport]
    [Guid("5CDF2C82-841E-4546-9722-0CF74078229A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IAudioEndpointVolume
    {
        int RegisterControlChangeNotify(IntPtr pNotify);
        int UnregisterControlChangeNotify(IntPtr pNotify);
        int GetChannelCount(out uint pnChannelCount);
        int SetMasterVolumeLevel(float fLevelDB, Guid pguidEventContext);
        int SetMasterVolumeLevelScalar(float fLevel, Guid pguidEventContext);
        int GetMasterVolumeLevel(out float pfLevelDB);
        int GetMasterVolumeLevelScalar(out float pfLevel);
        int SetChannelVolumeLevel(uint nChannel, float fLevelDB, Guid pguidEventContext);
        int SetChannelVolumeLevelScalar(uint nChannel, float fLevel, Guid pguidEventContext);
        int GetChannelVolumeLevel(uint nChannel, out float pfLevelDB);
        int GetChannelVolumeLevelScalar(uint nChannel, out float pfLevel);
        int SetMute(bool bMute, Guid pguidEventContext);
        int GetMute(out bool pbMute);
        int GetVolumeStepInfo(out uint pnStep, out uint pnStepCount);
        int VolumeStepUp(Guid pguidEventContext);
        int VolumeStepDown(Guid pguidEventContext);
        int QueryHardwareSupport(out uint pdwHardwareSupportMask);
        int GetVolumeRange(out float pflVolumeMindB, out float pflVolumeMaxdB, out float pflVolumeIncrementdB);
    }
}
