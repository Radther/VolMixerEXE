using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace VolMixerEXE
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
            // Switch on commands
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
        // Change the focused apps volume by the amount specified
        private static void ChangeAppVolume(float volumeAmount)
        {
            // Get the focused apps process Id
            IntPtr foregroundWindow = GetForegroundWindow();
            int windowProcessId;
            GetWindowThreadProcessId(foregroundWindow, out windowProcessId);

            // Find all process with the same Id
            var windowProcess = Process.GetProcessById(windowProcessId);
            Process[] appProcessCollection = Process.GetProcessesByName(windowProcess.ProcessName);
            var appProcessIdCollection = getProcessIds(appProcessCollection).ToHashSet();

            // Get all the processes that are an audio session
            var defaultAudioDevice = VolumeMixer.GetOutputDevice();
            var sessionManager = VolumeMixer.GetAudioSessionManager2(defaultAudioDevice);
            var sessions = VolumeMixer.GetAudioSessionEnumerator(sessionManager);
            var audioControls = VolumeMixer.GetAudioContols(sessions);
            var audioProcessIdCollection = audioControls.Keys.ToHashSet();

            // Get all the processes that are audio sessions of the focused application
            var commonProcessIdCollection = appProcessIdCollection.Intersect(audioProcessIdCollection);

            // Change the volume of all the audio processes of the focused application
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
            // Get the focused apps process Id
            IntPtr foregroundWindow = GetForegroundWindow();
            int windowProcessId;
            GetWindowThreadProcessId(foregroundWindow, out windowProcessId);

            // Find all process with the same Id
            var windowProcess = Process.GetProcessById(windowProcessId);
            Process[] appProcessCollection = Process.GetProcessesByName(windowProcess.ProcessName);
            var appProcessIdCollection = getProcessIds(appProcessCollection).ToHashSet();

            // Get all the processes that are an audio session
            var defaultAudioDevice = VolumeMixer.GetOutputDevice();
            var sessionManager = VolumeMixer.GetAudioSessionManager2(defaultAudioDevice);
            var sessions = VolumeMixer.GetAudioSessionEnumerator(sessionManager);
            var audioControls = VolumeMixer.GetAudioContols(sessions);
            var audioProcessIdCollection = audioControls.Keys.ToHashSet();

            // Get all the processes that are audio sessions of the focused application
            var commonProcessIdCollection = appProcessIdCollection.Intersect(audioProcessIdCollection);

            // Change the volume of all the audio processes of the focused application
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

        // Get the process Ids from process objects
        private static IEnumerable<int> getProcessIds(Process[] processes)
        {
            foreach (Process p in processes)
            {
                yield return p.Id;
            }
        }

        // Helper Class to manage Volume and Mute
        // Based on code by Anders Carstensen: https://stackoverflow.com/a/25584074
        public class VolumeMixer
        {
            // Get the volume
            public static float? GetApplicationVolume(ISimpleAudioVolume volume)
            {
                if (volume == null)
                    return null;

                float level;
                volume.GetMasterVolume(out level);
                return level * 100;
            }

            // Set the volume to level
            public static void SetApplicationVolume(ISimpleAudioVolume volume, float level)
            {
                if (volume == null)
                    return;

                Guid guid = Guid.Empty;
                volume.SetMasterVolume(level / 100, ref guid);
                Marshal.ReleaseComObject(volume);
            }

            // Get the mute state
            public static bool? GetApplicationMute(ISimpleAudioVolume volume)
            {
                if (volume == null)
                    return null;

                bool mute;
                volume.GetMute(out mute);
                return mute;
            }

            // Set the mute state
            public static void SetApplicationMute(ISimpleAudioVolume volume, bool mute)
            {
                if (volume == null)
                    return;

                Guid guid = Guid.Empty;
                volume.SetMute(mute, ref guid);
            }

            // Get the current default audio device
            public static IMMDevice GetOutputDevice()
            {
                IMMDeviceEnumerator deviceEnumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
                IMMDevice speakers;
                deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out speakers);
                return speakers;
            }

            // Get the audio session manager
            public static IAudioSessionManager2 GetAudioSessionManager2(IMMDevice device)
            {
                Guid IID_IAudioSessionManager2 = typeof(IAudioSessionManager2).GUID;
                object o;
                device.Activate(ref IID_IAudioSessionManager2, 0, IntPtr.Zero, out o);
                IAudioSessionManager2 mgr = (IAudioSessionManager2)o;
                return mgr;
            }

            // Get an enumerator of all the current audio sessions from a session manager
            public static IAudioSessionEnumerator GetAudioSessionEnumerator(IAudioSessionManager2 manager)
            {
                IAudioSessionEnumerator sessionEnumerator;
                manager.GetSessionEnumerator(out sessionEnumerator);
                return sessionEnumerator;
            }

            // Get audio controls and process Ids for audio sessions from enumerator
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