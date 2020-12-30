using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace AUdioTesting
{
    class Program
    {
        // Exposes Win32 APIs to C Sharp for use
        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        static void Main(string[] args)
        {
            switch (args[0]) {
                case "appvolume":
                    float volumeAmount = float.Parse(args[1]);
                    ChangeAppVolume(volumeAmount);
                    break;
                case "appmutetoggle":
                    ToggleMute();
                    break;
            }
            
        }

        private static void ChangeAppVolume(float volumeAmount)
        {
            IntPtr foregroundWindow = GetForegroundWindow();
            int windowProcessId;
            GetWindowThreadProcessId(foregroundWindow, out windowProcessId);

            var windowProcess = Process.GetProcessById(windowProcessId);
            Process[] appProcessCollection = Process.GetProcessesByName(windowProcess.ProcessName);
            var appProcessIdCollection = getProcessIds(appProcessCollection).ToHashSet();

            var defaultAudioDevice = VolumeMixer.GetOutputDevice();
            var sessionManager = VolumeMixer.GetAudioSessionManager2(defaultAudioDevice);
            var sessions = VolumeMixer.GetAudioSessionEnumerator(sessionManager);
            var audioControls = VolumeMixer.GetAudioContols(sessions);
            var audioProcessIdCollection = audioControls.Keys.ToHashSet();

            var commonProcessIdCollection = appProcessIdCollection.Intersect(audioProcessIdCollection);
            foreach (int processId in commonProcessIdCollection)
            {
                var volumeControl = audioControls[processId] as ISimpleAudioVolume;
                var newVolumeLevel = VolumeMixer.GetApplicationVolume(volumeControl) + volumeAmount;
                VolumeMixer.SetApplicationVolume(volumeControl, Math.Min(100, Math.Max(0, newVolumeLevel ?? 30f)));
                Marshal.ReleaseComObject(volumeControl);
            }

            Marshal.ReleaseComObject(defaultAudioDevice);
            Marshal.ReleaseComObject(sessionManager);
            Marshal.ReleaseComObject(sessions);
        }

        private static void ToggleMute()
        {
            IntPtr foregroundWindow = GetForegroundWindow();
            int windowProcessId;
            GetWindowThreadProcessId(foregroundWindow, out windowProcessId);

            var windowProcess = Process.GetProcessById(windowProcessId);
            Process[] appProcessCollection = Process.GetProcessesByName(windowProcess.ProcessName);
            var appProcessIdCollection = getProcessIds(appProcessCollection).ToHashSet();

            var defaultAudioDevice = VolumeMixer.GetOutputDevice();
            var sessionManager = VolumeMixer.GetAudioSessionManager2(defaultAudioDevice);
            var sessions = VolumeMixer.GetAudioSessionEnumerator(sessionManager);
            var audioControls = VolumeMixer.GetAudioContols(sessions);
            var audioProcessIdCollection = audioControls.Keys.ToHashSet();

            var commonProcessIdCollection = appProcessIdCollection.Intersect(audioProcessIdCollection);
            foreach (int processId in commonProcessIdCollection)
            {
                var volumeControl = audioControls[processId] as ISimpleAudioVolume;
                var currentMute = VolumeMixer.GetApplicationMute(volumeControl);
                VolumeMixer.SetApplicationMute(volumeControl, !currentMute ?? false);
                Marshal.ReleaseComObject(volumeControl);
            }

            Marshal.ReleaseComObject(defaultAudioDevice);
            Marshal.ReleaseComObject(sessionManager);
            Marshal.ReleaseComObject(sessions);
        }

        private static IEnumerable<int> getProcessIds(Process[] processes)
        {
            foreach (Process p in processes)
            {
                yield return p.Id;
            }
        }

        public class VolumeMixer
        {
            public static float? GetApplicationVolume(ISimpleAudioVolume volume)
            {
                if (volume == null)
                    return null;

                float level;
                volume.GetMasterVolume(out level);
                return level * 100;
            }

            public static void SetApplicationVolume(ISimpleAudioVolume volume, float level)
            {
                if (volume == null)
                    return;

                Guid guid = Guid.Empty;
                volume.SetMasterVolume(level / 100, ref guid);
                Marshal.ReleaseComObject(volume);
            }

            public static bool? GetApplicationMute(ISimpleAudioVolume volume)
            {
                if (volume == null)
                    return null;

                bool mute;
                volume.GetMute(out mute);
                return mute;
            }

            public static void SetApplicationMute(ISimpleAudioVolume volume, bool mute)
            {
                if (volume == null)
                    return;

                Guid guid = Guid.Empty;
                volume.SetMute(mute, ref guid);
            }

            public static IMMDevice GetOutputDevice()
            {
                IMMDeviceEnumerator deviceEnumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
                IMMDevice speakers;
                deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out speakers);
                return speakers;
            }

            public static IAudioSessionManager2 GetAudioSessionManager2(IMMDevice device)
            {
                Guid IID_IAudioSessionManager2 = typeof(IAudioSessionManager2).GUID;
                object o;
                device.Activate(ref IID_IAudioSessionManager2, 0, IntPtr.Zero, out o);
                IAudioSessionManager2 mgr = (IAudioSessionManager2)o;
                return mgr;
            }

            public static IAudioSessionEnumerator GetAudioSessionEnumerator(IAudioSessionManager2 manager)
            {
                IAudioSessionEnumerator sessionEnumerator;
                manager.GetSessionEnumerator(out sessionEnumerator);
                return sessionEnumerator;
            }

            public static IDictionary<int, IAudioSessionControl2> GetAudioContols(IAudioSessionEnumerator sessions)
            {
                int count;
                sessions.GetCount(out count);
                IDictionary<int, IAudioSessionControl2> controls = new Dictionary<int, IAudioSessionControl2>();
                for (int i = 0; i < count; i++)
                {
                    IAudioSessionControl2 ctl;
                    sessions.GetSession(i, out ctl);
                    int cpid;
                    ctl.GetProcessId(out cpid);
                    controls[cpid] = ctl;
                }
                return controls;
            }
        }
    }
}