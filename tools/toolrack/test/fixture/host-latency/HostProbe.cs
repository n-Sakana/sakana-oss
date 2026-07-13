using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ToolRackProbe
{
    public sealed class ProbeResult
    {
        public bool AddType;
        public bool HookInstalled;
        public bool HotkeyAvailable;
        public bool MessageLoop;
        public bool Cleanup;
        public bool DwmAvailable;
        public bool AlreadyRunning;
        public int ProcessId;
    }

    public static class HostProbe
    {
        private const int WhMouseLl = 14;
        private const uint ModAlt = 0x0001;
        private const uint ModControl = 0x0002;
        private const uint ModShift = 0x0004;
        private const uint ModNoRepeat = 0x4000;
        private const uint VkF13 = 0x7C;
        private const uint VkF24 = 0x87;

        private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

        private static HookProc hookProc;
        private static IntPtr hookHandle;
        private static System.Windows.Forms.Timer timer;
        private static Stopwatch stopwatch;
        private static string stateRoot;
        private static int autoStopMilliseconds;
        private static ProbeResult result;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int hookId, HookProc callback, IntPtr module, uint threadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hook);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint key);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnregisterHotKey(IntPtr window, int id);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string moduleName);

        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmIsCompositionEnabled(out bool enabled);

        public static ProbeResult Run(string path, int stopAfterMilliseconds)
        {
            if (String.IsNullOrWhiteSpace(path))
            {
                throw new ArgumentException("state root is required", "path");
            }

            stateRoot = Path.GetFullPath(path);
            Directory.CreateDirectory(stateRoot);
            autoStopMilliseconds = stopAfterMilliseconds;
            result = new ProbeResult();
            result.AddType = true;
            result.ProcessId = Process.GetCurrentProcess().Id;
            Mutex instanceMutex = new Mutex(false, GetMutexName(stateRoot));
            bool ownsMutex = false;
            try
            {
                try
                {
                    ownsMutex = instanceMutex.WaitOne(0, false);
                }
                catch (AbandonedMutexException)
                {
                    ownsMutex = true;
                }
                if (!ownsMutex)
                {
                    result.AlreadyRunning = true;
                    result.Cleanup = true;
                    return result;
                }

                result.DwmAvailable = TestDwm();
                result.HotkeyAvailable = TestHotkey();
                hookProc = HookCallback;
                hookHandle = InstallHook();
                result.HookInstalled = hookHandle != IntPtr.Zero;
                stopwatch = Stopwatch.StartNew();

                timer = new System.Windows.Forms.Timer();
                timer.Interval = 50;
                timer.Tick += OnTick;
                timer.Start();

                try
                {
                    Application.Run(new ApplicationContext());
                }
                finally
                {
                    if (timer != null)
                    {
                        timer.Stop();
                        timer.Dispose();
                        timer = null;
                    }

                    bool cleanup = true;
                    if (hookHandle != IntPtr.Zero)
                    {
                        cleanup = UnhookWindowsHookEx(hookHandle);
                        hookHandle = IntPtr.Zero;
                    }
                    result.Cleanup = cleanup;
                    hookProc = null;
                    WriteHeartbeat(false);
                }
                return result;
            }
            finally
            {
                if (ownsMutex)
                {
                    instanceMutex.ReleaseMutex();
                }
                instanceMutex.Dispose();
            }
        }

        private static string GetMutexName(string path)
        {
            uint hash = 2166136261;
            string normalized = path.ToUpperInvariant();
            for (int index = 0; index < normalized.Length; index++)
            {
                hash ^= normalized[index];
                hash *= 16777619;
            }
            return "Local\\ToolRackHostProbe-" + hash.ToString("X8", CultureInfo.InvariantCulture);
        }

        private static IntPtr InstallHook()
        {
            using (Process process = Process.GetCurrentProcess())
            using (ProcessModule module = process.MainModule)
            {
                IntPtr moduleHandle = GetModuleHandle(module.ModuleName);
                return SetWindowsHookEx(WhMouseLl, hookProc, moduleHandle, 0);
            }
        }

        private static IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
        {
            return CallNextHookEx(hookHandle, code, wParam, lParam);
        }

        private static bool TestHotkey()
        {
            int id = 0x6F31;
            uint modifiers = ModControl | ModAlt | ModShift | ModNoRepeat;
            for (uint key = VkF24; key >= VkF13; key--)
            {
                if (RegisterHotKey(IntPtr.Zero, id, modifiers, key))
                {
                    UnregisterHotKey(IntPtr.Zero, id);
                    return true;
                }
            }
            return false;
        }

        private static bool TestDwm()
        {
            try
            {
                bool enabled;
                return DwmIsCompositionEnabled(out enabled) == 0;
            }
            catch (DllNotFoundException)
            {
                return false;
            }
            catch (EntryPointNotFoundException)
            {
                return false;
            }
        }

        private static void OnTick(object sender, EventArgs eventArgs)
        {
            result.MessageLoop = true;
            WriteHeartbeat(true);
            string stopPath = Path.Combine(stateRoot, "stop.request");
            if (File.Exists(stopPath))
            {
                Application.ExitThread();
                return;
            }
            if (autoStopMilliseconds > 0 && stopwatch.ElapsedMilliseconds >= autoStopMilliseconds)
            {
                Application.ExitThread();
            }
        }

        private static void WriteHeartbeat(bool ready)
        {
            try
            {
                string json = String.Format(
                    CultureInfo.InvariantCulture,
                    "{{\"Ready\":{0},\"ProcessId\":{1},\"HookInstalled\":{2},\"HotkeyAvailable\":{3},\"MessageLoop\":{4},\"UtcTicks\":{5}}}",
                    ready ? "true" : "false",
                    result.ProcessId,
                    result.HookInstalled ? "true" : "false",
                    result.HotkeyAvailable ? "true" : "false",
                    result.MessageLoop ? "true" : "false",
                    DateTime.UtcNow.Ticks);
                string target = Path.Combine(stateRoot, "heartbeat.json");
                string temporary = target + ".tmp";
                File.WriteAllText(temporary, json, new UTF8Encoding(false));
                if (File.Exists(target))
                {
                    File.Delete(target);
                }
                File.Move(temporary, target);
            }
            catch
            {
                // The probe must never interfere with input because status I/O failed.
            }
        }
    }
}
