using System;
using System.Collections.Generic;
using System.Linq;
using NAudio.CoreAudioApi;

namespace Audio_Mixer.Core
{
    public sealed class CoreAudioManager : IDisposable
    {
        private readonly MMDeviceEnumerator enumerator = new();

        public IReadOnlyList<AudioDevice> GetOutputDevices()
        {
            var devices = enumerator.EnumerateAudioEndPoints(DataFlow.Render, DeviceState.Active);
            return devices.Select(d => new AudioDevice(d.ID, d.FriendlyName)).ToList();
        }

        public void SetDeviceVolume(string deviceId, float volumeScalar)
        {
            if (deviceId is null) throw new ArgumentNullException(nameof(deviceId));

            if (volumeScalar < 0f) volumeScalar = 0f;
            if (volumeScalar > 1f) volumeScalar = 1f;

            using var device = enumerator.GetDevice(deviceId);
            device.AudioEndpointVolume.MasterVolumeLevelScalar = volumeScalar;
        }

        public float GetDeviceVolume(string deviceId)
        {
            if (deviceId is null) throw new ArgumentNullException(nameof(deviceId));

            using var device = enumerator.GetDevice(deviceId);
            return device.AudioEndpointVolume.MasterVolumeLevelScalar;
        }

        public void Dispose()
        {
            enumerator.Dispose();
        }
    }

    public sealed record AudioDevice(string Id, string Name);
}
