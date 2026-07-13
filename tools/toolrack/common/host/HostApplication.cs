using System;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ToolRack
{
    internal sealed class HostLog
    {
        private readonly string path;
        private readonly object sync = new object();

        internal HostLog(string stateRoot)
        {
            DirectoryInfo state = new DirectoryInfo(stateRoot);
            string localRoot = state.FullName;
            if (state.Parent != null && String.Equals(state.Parent.Name, "state", StringComparison.OrdinalIgnoreCase) &&
                state.Parent.Parent != null)
            {
                localRoot = state.Parent.Parent.FullName;
            }
            string directory = Path.Combine(localRoot, "log");
            Directory.CreateDirectory(directory);
            path = Path.Combine(directory, "host.log");
        }

        internal void Write(string message)
        {
            lock (sync)
            {
                RotateIfNeeded();
                File.AppendAllText(path, DateTime.UtcNow.ToString("o") + " " + message + Environment.NewLine, new UTF8Encoding(false));
            }
        }

        private void RotateIfNeeded()
        {
            FileInfo info = new FileInfo(path);
            if (!info.Exists || info.Length < 1024 * 1024) { return; }
            string second = path + ".2";
            string first = path + ".1";
            if (File.Exists(second)) { File.Delete(second); }
            if (File.Exists(first)) { File.Move(first, second); }
            File.Move(path, first);
        }
    }

    internal sealed class ToolRackApplicationContext : ApplicationContext
    {
        private readonly object sync = new object();
        private readonly string pipeName;
        private readonly bool testMode;
        private readonly HostLog log;
        private readonly HostNotifier notifier;
        private readonly HostMessageWindow window;
        private readonly ActivationWorker worker;
        private readonly HostPipeServer pipe;
        private readonly HotkeyManager hotkeys;
        private readonly MouseGestureManager mouse;
        private readonly HostConfigReloader reloader;
        private HostStateBundle bundle;
        private HostStateBundle pendingBundle;

        internal ToolRackApplicationContext(HostStateBundle stateBundle, string name, HostLog hostLog, bool isTestMode)
        {
            bundle = stateBundle;
            pipeName = name;
            log = hostLog;
            testMode = isTestMode;
            notifier = new HostNotifier(testMode, log.Write);
            worker = new ActivationWorker(bundle.State.Root, log.Write, notifier.Notify, testMode ? 8 : 64);
            window = new HostMessageWindow(ExitThread, OnHotkey, OnMouseActivation, OnConfigReload);
            hotkeys = new HotkeyManager(new Win32HotkeyBackend(window.Handle), worker, log.Write);
            hotkeys.Reload(bundle.Config.Active);
            mouse = new MouseGestureManager(new Win32MouseHookFactory(), worker, log.Write, window.Handle);
            mouse.Reload(bundle.Config.Active);
            reloader = new HostConfigReloader(bundle, QueueConfigReload, log.Write, notifier.Notify);
            pipe = new HostPipeServer(pipeName, GetStatus, reloader.RequestReload, window.PostShutdown,
                testMode ? (Func<string, string>)HandleTestCommand : null, log.Write);
        }

        internal void Start()
        {
            reloader.Start();
            pipe.Start();
            log.Write("ready pid=" + Process.GetCurrentProcess().Id);
        }

        internal HostStatus GetStatus()
        {
            lock (sync)
            {
                HostStatus status = HostApplication.CreateStatus(bundle, pipeName, true, hotkeys.ActiveCount, hotkeys.InactiveCount,
                    mouse.HookActive, mouse.ActiveCount, mouse.InactiveCount);
                status.ActivationQueueDepth = worker.QueueDepth;
                status.ActivationQueued = worker.QueuedCount;
                status.ActivationCompleted = worker.CompletedCount;
                status.ActivationFailed = worker.FailedCount;
                status.ActivationRejected = worker.RejectedCount;
                status.WorkerGeneration = worker.Generation;
                status.ReloadSucceeded = reloader.Succeeded;
                status.ReloadFailed = reloader.Failed;
                status.NotificationCount = notifier.Count;
                status.ConfigSourceSha256 = bundle.Config.SourceConfigSha256;
                status.LastError = FirstNonEmpty(worker.LastError, reloader.LastError, notifier.LastMessage);
                return status;
            }
        }

        private static string FirstNonEmpty(params string[] values)
        {
            foreach (string value in values)
            {
                if (!String.IsNullOrWhiteSpace(value)) { return value; }
            }
            return String.Empty;
        }

        private void OnHotkey(int registrationId)
        {
            lock (sync) { hotkeys.HandleHotkey(registrationId); }
        }

        private void OnMouseActivation(int activationId)
        {
            lock (sync) { mouse.HandleActivation(activationId); }
        }

        private void QueueConfigReload(HostStateBundle next)
        {
            lock (sync) { pendingBundle = next; }
            window.PostConfigReload();
        }

        private void OnConfigReload()
        {
            lock (sync)
            {
                if (pendingBundle == null) { return; }
                HostStateBundle next = pendingBundle;
                pendingBundle = null;
                hotkeys.Reload(next.Config.Active);
                mouse.Reload(next.Config.Active);
                bundle = next;
            }
        }

        private string HandleTestCommand(string command)
        {
            if (!testMode) { return "{\"ok\":false,\"queued\":false}"; }
            if (String.Equals(command, "test-pause-worker", StringComparison.Ordinal))
            {
                worker.PauseForTest();
                return "{\"ok\":true}";
            }
            if (String.Equals(command, "test-resume-worker", StringComparison.Ordinal))
            {
                worker.ResumeForTest();
                return "{\"ok\":true}";
            }
            if (String.Equals(command, "test-throw-next", StringComparison.Ordinal))
            {
                worker.ThrowNextForTest();
                return "{\"ok\":true}";
            }
            if (String.Equals(command, "test-break-worker", StringComparison.Ordinal))
            {
                worker.BreakRunspaceForTest();
                return "{\"ok\":true}";
            }

            string[] parts = command.Split('|');
            ResolvedBindingData binding = null;
            lock (sync)
            {
                if (parts.Length == 3 && String.Equals(parts[0], "test-activate", StringComparison.Ordinal))
                {
                    foreach (ResolvedBindingData item in bundle.Config.Active)
                    {
                        if (item != null && item.Invoke != null &&
                            String.Equals(item.Invoke.Tool, parts[1], StringComparison.Ordinal) &&
                            String.Equals(item.Invoke.Action, parts[2], StringComparison.Ordinal))
                        {
                            binding = item;
                            break;
                        }
                    }
                }
                else if (parts.Length == 2 && String.Equals(parts[0], "test-activate-id", StringComparison.Ordinal))
                {
                    foreach (ResolvedBindingData item in bundle.Config.Active)
                    {
                        if (item != null && String.Equals(item.Id, parts[1], StringComparison.Ordinal))
                        {
                            binding = item;
                            break;
                        }
                    }
                }
            }
            bool queued = binding != null && worker.TryActivate(binding);
            return "{\"ok\":" + JsonBoolean(queued) + ",\"queued\":" + JsonBoolean(queued) + "}";
        }

        private static string JsonBoolean(bool value)
        {
            return value ? "true" : "false";
        }

        protected override void ExitThreadCore()
        {
            reloader.Dispose();
            pipe.Stop();
            hotkeys.Dispose();
            mouse.Dispose();
            worker.Dispose();
            window.Dispose();
            log.Write("stopped");
            base.ExitThreadCore();
        }
    }

    public static class HostApplication
    {
        public static string SelfTest(string stateRoot)
        {
            HostStateBundle bundle = HostStateBundle.Load(stateRoot);
            ActivationWorker.ValidateLaunchCore(bundle.State.Root);
            string pipeName = GetPipeName(bundle.State.Namespace);
            return CreateStatus(bundle, pipeName, false, 0, 0, false, 0, 0).ToJson();
        }

        public static int Run(string stateRoot, bool testMode)
        {
            HostStateBundle bundle = HostStateBundle.Load(stateRoot);
            string mutexName = GetMutexName(bundle.State.Namespace);
            string pipeName = GetPipeName(bundle.State.Namespace);
            using (Mutex mutex = new Mutex(false, mutexName))
            {
                bool owns = false;
                try
                {
                    try { owns = mutex.WaitOne(0, false); }
                    catch (AbandonedMutexException) { owns = true; }
                    if (!owns) { return 2; }

                    HostLog log = new HostLog(bundle.StateRoot);
                    using (ToolRackApplicationContext context = new ToolRackApplicationContext(bundle, pipeName, log, testMode))
                    {
                        context.Start();
                        Application.Run(context);
                    }
                    return 0;
                }
                finally
                {
                    if (owns) { mutex.ReleaseMutex(); }
                }
            }
        }

        internal static HostStatus CreateStatus(HostStateBundle bundle, string pipeName, bool ready, int activeHotkeys,
            int inactiveHotkeys, bool mouseHookActive, int activeMouseBindings, int inactiveMouseBindings)
        {
            return new HostStatus
            {
                Ok = true,
                Ready = ready,
                ProcessId = Process.GetCurrentProcess().Id,
                Root = bundle.State.Root,
                Version = bundle.State.Version,
                ActiveBindings = bundle.Config.Active.Count,
                RejectedBindings = bundle.Config.Rejected.Count,
                PipeName = pipeName,
                PipeCurrentUserOnly = HostPipeServer.CurrentUserAclIsExclusive(),
                ActiveHotkeys = activeHotkeys,
                InactiveHotkeys = inactiveHotkeys,
                MouseHookActive = mouseHookActive,
                ActiveMouseBindings = activeMouseBindings,
                InactiveMouseBindings = inactiveMouseBindings,
                LastError = String.Empty,
                ConfigSourceSha256 = bundle.Config.SourceConfigSha256
            };
        }

        public static string GetPipeName(string nameSpace)
        {
            return "ToolRackHost-" + GetCurrentSid() + "-" + nameSpace;
        }

        private static string GetMutexName(string nameSpace)
        {
            return "Local\\ToolRackHost-" + GetCurrentSid() + "-" + nameSpace;
        }

        private static string GetCurrentSid()
        {
            return WindowsIdentity.GetCurrent().User.Value;
        }
    }
}
