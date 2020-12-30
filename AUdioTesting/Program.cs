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
        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string strClassName, string strWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int processId);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        static void Main(string[] args)
        {
            Stopwatch watch = new Stopwatch();
            //ListProcesses();
            // MAIN
            IntPtr hwnd = GetForegroundWindow();
            int pid;
            GetWindowThreadProcessId(hwnd, out pid);
            
            var process = Process.GetProcessById(pid);
            Process[] pcollection = Process.GetProcessesByName(process.ProcessName);
            //foreach (Process p in pcollection)
            //{
            //    VolumeMixer.SetApplicationVolume(p.Id, 100f);

            //}

            //using (System.IO.StreamWriter file =
            //new System.IO.StreamWriter(@".\WriteLines2.txt"))
            //{

            //}
            watch.Start();
            //VolumeMixer.SetApplicationVolume(pid, 100f);
            foreach (Process p in pcollection)
            {
                //VolumeMixer.performanceTest(p.Id);
                //VolumeMixer.SetApplicationVolume(p.Id, 100f);
            }
            watch.Stop();


            System.IO.File.WriteAllText(@".\WriteText.txt", watch.ElapsedMilliseconds.ToString());

            //var pID = 0;
            //foreach (var process in Process.GetProcesses())
            //{
            //    pID = process.Id;
            //    if (pID != 0)
            //        VolumeMixer.SetApplicationVolume(pID, 100f);

            //}

            //foreach (int name in EnumerateApplications())
            //{
            //    Console.WriteLine("name:" + name);

            //    IntPtr hwnd = GetForegroundWindow();
            //    uint pid;
            //    GetWindowThreadProcessId(hwnd, out pid);
            //    Console.WriteLine("currentProcess:" + pid);
            //    if (pid == name)
            //    {
            //        //display mute state & volume level(% of master)
            //        //Console.WriteLine("Mute:" + GetApplicationMute(app));
            //        //Console.WriteLine("Volume:" + GetApplicationVolume(app));

            //        //mute the application
            //        //SetApplicationMute(app, true);

            //        //set the volume to half of master volume(50 %)
            //        //SetApplicationVolume(app, 50);
            //    }
            //}
            //Console.Read();
        }

        private static void ListProcesses()
        {
            Process[] processCollection = Process.GetProcesses();
            foreach (Process p in processCollection)
            {
                Console.WriteLine(p.ProcessName);
            }
        }

        public class VolumeMixer
        {
            public static float? GetApplicationVolume(int pid)
            {
                ISimpleAudioVolume volume = GetVolumeObject(pid);
                if (volume == null)
                    return null;

                float level;
                volume.GetMasterVolume(out level);
                Marshal.ReleaseComObject(volume);
                return level * 100;
            }

            public static bool? GetApplicationMute(int pid)
            {
                ISimpleAudioVolume volume = GetVolumeObject(pid);
                if (volume == null)
                    return null;

                bool mute;
                volume.GetMute(out mute);
                Marshal.ReleaseComObject(volume);
                return mute;
            }

            public static void SetApplicationVolume(int pid, float level)
            {
                ISimpleAudioVolume volume = GetVolumeObject(pid);
                if (volume == null)
                    return;

                Guid guid = Guid.Empty;
                volume.SetMasterVolume(level / 100, ref guid);
                Marshal.ReleaseComObject(volume);
            }

            public static void SetApplicationMute(int pid, bool mute)
            {
                ISimpleAudioVolume volume = GetVolumeObject(pid);
                if (volume == null)
                    return;

                Guid guid = Guid.Empty;
                volume.SetMute(mute, ref guid);
                Marshal.ReleaseComObject(volume);
            }

            private static ISimpleAudioVolume GetVolumeObject(int pid)
            {
                // get the speakers (1st render + multimedia) device
                IMMDeviceEnumerator deviceEnumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
                IMMDevice speakers;
                deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out speakers);

                // activate the session manager. we need the enumerator
                Guid IID_IAudioSessionManager2 = typeof(IAudioSessionManager2).GUID;
                object o;
                speakers.Activate(ref IID_IAudioSessionManager2, 0, IntPtr.Zero, out o);
                IAudioSessionManager2 mgr = (IAudioSessionManager2)o;

                // enumerate sessions for on this device
                IAudioSessionEnumerator sessionEnumerator;
                mgr.GetSessionEnumerator(out sessionEnumerator);
                int count;
                sessionEnumerator.GetCount(out count);

                // search for an audio session with the required name
                // NOTE: we could also use the process id instead of the app name (with IAudioSessionControl2)
                ISimpleAudioVolume volumeControl = null;
                for (int i = 0; i < count; i++)
                {
                    IAudioSessionControl2 ctl;
                    sessionEnumerator.GetSession(i, out ctl);
                    int cpid;
                    ctl.GetProcessId(out cpid);

                    if (cpid == pid)
                    {
                        volumeControl = ctl as ISimpleAudioVolume;
                        break;
                    }
                    Marshal.ReleaseComObject(ctl);
                }
                Marshal.ReleaseComObject(sessionEnumerator);
                Marshal.ReleaseComObject(mgr);
                Marshal.ReleaseComObject(speakers);
                Marshal.ReleaseComObject(deviceEnumerator);
                return volumeControl;
            }

            public static void performanceTest(int pid)
            {
                // get the speakers (1st render + multimedia) device
                IMMDeviceEnumerator deviceEnumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
                IMMDevice speakers;
                deviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia, out speakers);

                // activate the session manager. we need the enumerator
                Guid IID_IAudioSessionManager2 = typeof(IAudioSessionManager2).GUID;
                object o;
                speakers.Activate(ref IID_IAudioSessionManager2, 0, IntPtr.Zero, out o);
                IAudioSessionManager2 mgr = (IAudioSessionManager2)o;

                //// enumerate sessions for on this device
                //IAudioSessionEnumerator sessionEnumerator;
                //mgr.GetSessionEnumerator(out sessionEnumerator);
                //int count;
                //sessionEnumerator.GetCount(out count);

                //// search for an audio session with the required name
                //// NOTE: we could also use the process id instead of the app name (with IAudioSessionControl2)
                //ISimpleAudioVolume volumeControl = null;
                //for (int i = 0; i < count; i++)
                //{
                //    IAudioSessionControl2 ctl;
                //    sessionEnumerator.GetSession(i, out ctl);
                //    int cpid;
                //    ctl.GetProcessId(out cpid);

                //    if (cpid == pid)
                //    {
                //        volumeControl = ctl as ISimpleAudioVolume;
                //        break;
                //    }
                //    Marshal.ReleaseComObject(ctl);
                //}
                //Marshal.ReleaseComObject(sessionEnumerator);
                //Marshal.ReleaseComObject(mgr);
                //Marshal.ReleaseComObject(speakers);
                //Marshal.ReleaseComObject(deviceEnumerator);
                //return volumeControl;
            }
        }

        [ComImport]
        [Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
        internal class MMDeviceEnumerator
        {
        }

        internal enum EDataFlow
        {
            eRender,
            eCapture,
            eAll,
            EDataFlow_enum_count
        }

        internal enum ERole
        {
            eConsole,
            eMultimedia,
            eCommunications,
            ERole_enum_count
        }

        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IMMDeviceEnumerator
        {
            int NotImpl1();

            [PreserveSig]
            int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice ppDevice);

            // the rest is not implemented
        }

        [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IMMDevice
        {
            [PreserveSig]
            int Activate(ref Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);

            // the rest is not implemented
        }

        [Guid("77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IAudioSessionManager2
        {
            int NotImpl1();
            int NotImpl2();

            [PreserveSig]
            int GetSessionEnumerator(out IAudioSessionEnumerator SessionEnum);

            // the rest is not implemented
        }

        [Guid("E2F5BB11-0570-40CA-ACDD-3AA01277DEE8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IAudioSessionEnumerator
        {
            [PreserveSig]
            int GetCount(out int SessionCount);

            [PreserveSig]
            int GetSession(int SessionCount, out IAudioSessionControl2 Session);
        }

        [Guid("87CE5498-68D6-44E5-9215-6DA47EF883D8"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface ISimpleAudioVolume
        {
            [PreserveSig]
            int SetMasterVolume(float fLevel, ref Guid EventContext);

            [PreserveSig]
            int GetMasterVolume(out float pfLevel);

            [PreserveSig]
            int SetMute(bool bMute, ref Guid EventContext);

            [PreserveSig]
            int GetMute(out bool pbMute);
        }

        [Guid("bfb7ff88-7239-4fc9-8fa2-07c950be9c6d"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        internal interface IAudioSessionControl2
        {
            // IAudioSessionControl
            [PreserveSig]
            int NotImpl0();

            [PreserveSig]
            int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

            [PreserveSig]
            int SetDisplayName([MarshalAs(UnmanagedType.LPWStr)] string Value, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

            [PreserveSig]
            int GetIconPath([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

            [PreserveSig]
            int SetIconPath([MarshalAs(UnmanagedType.LPWStr)] string Value, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

            [PreserveSig]
            int GetGroupingParam(out Guid pRetVal);

            [PreserveSig]
            int SetGroupingParam([MarshalAs(UnmanagedType.LPStruct)] Guid Override, [MarshalAs(UnmanagedType.LPStruct)] Guid EventContext);

            [PreserveSig]
            int NotImpl1();

            [PreserveSig]
            int NotImpl2();

            // IAudioSessionControl2
            [PreserveSig]
            int GetSessionIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

            [PreserveSig]
            int GetSessionInstanceIdentifier([MarshalAs(UnmanagedType.LPWStr)] out string pRetVal);

            [PreserveSig]
            int GetProcessId(out int pRetVal);

            [PreserveSig]
            int IsSystemSoundsSession();

            [PreserveSig]
            int SetDuckingPreference(bool optOut);
        }

    }
}