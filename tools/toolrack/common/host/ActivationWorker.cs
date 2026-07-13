using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace ToolRack
{
    internal sealed class HostNotifier
    {
        private readonly bool testMode;
        private readonly Action<string> log;
        private readonly object sync = new object();
        private readonly HashSet<string> shown = new HashSet<string>(StringComparer.Ordinal);
        private int count;
        private string lastMessage = String.Empty;

        internal HostNotifier(bool isTestMode, Action<string> logger)
        {
            testMode = isTestMode;
            log = logger;
        }

        internal int Count { get { lock (sync) { return count; } } }
        internal string LastMessage { get { lock (sync) { return lastMessage; } } }

        internal void Notify(string message)
        {
            if (String.IsNullOrWhiteSpace(message)) { return; }
            lock (sync)
            {
                lastMessage = message;
                if (!shown.Add(message)) { return; }
                count++;
            }
            try { log("notification: " + message); }
            catch { }
            if (testMode) { return; }
            ThreadPool.QueueUserWorkItem(delegate
            {
                try
                {
                    MessageBox.Show(message, "Tool Rack", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch { }
            });
        }
    }

    internal sealed class ActivationRequest
    {
        internal ResolvedBindingData Binding;
    }

    internal sealed class ActivationWorker : IActivationSink, IDisposable
    {
        private const string InvokeScript =
            "param($toolDir,$actionId) " +
            "$result = Invoke-HostAction -ToolDir $toolDir -ActionId $actionId; " +
            "if (-not $result.Ok) { throw (($result.Errors | ForEach-Object { [string]$_ }) -join [Environment]::NewLine) }; " +
            "[pscustomobject]@{ ok = $true; pid = [int]$result.ProcessId }";

        private readonly string launchCorePath;
        private readonly Action<string> log;
        private readonly Action<string> notify;
        private readonly BlockingCollection<ActivationRequest> queue;
        private readonly Thread thread;
        private readonly ManualResetEvent ready = new ManualResetEvent(false);
        private readonly ManualResetEvent testGate = new ManualResetEvent(true);
        private readonly object runspaceSync = new object();
        private Runspace runspace;
        private PowerShell currentPipeline;
        private Exception startupError;
        private volatile bool accepting = true;
        private int disposed;
        private int queuedCount;
        private int completedCount;
        private int failedCount;
        private int rejectedCount;
        private int generation;
        private int throwNext;
        private int breakNext;
        private string lastError = String.Empty;

        internal ActivationWorker(string repositoryRoot, Action<string> logger, Action<string> notification, int capacity)
        {
            if (capacity < 1) { throw new ArgumentOutOfRangeException("capacity"); }
            launchCorePath = Path.Combine(Path.GetFullPath(repositoryRoot), "common", "launch-core.ps1");
            if (!File.Exists(launchCorePath)) { throw new FileNotFoundException("launch-core.ps1 not found", launchCorePath); }
            log = logger;
            notify = notification;
            queue = new BlockingCollection<ActivationRequest>(capacity);
            thread = new Thread(WorkerMain);
            thread.IsBackground = true;
            thread.Name = "ToolRackActivationWorker";
            thread.SetApartmentState(ApartmentState.MTA);
            thread.Start();
            if (!ready.WaitOne(5000))
            {
                Dispose();
                throw new TimeoutException("activation worker readiness timed out");
            }
            if (startupError != null)
            {
                Dispose();
                throw new InvalidOperationException("activation worker failed: " + startupError.Message, startupError);
            }
        }

        internal int QueueDepth { get { return queue.Count; } }
        internal int QueuedCount { get { return Volatile.Read(ref queuedCount); } }
        internal int CompletedCount { get { return Volatile.Read(ref completedCount); } }
        internal int FailedCount { get { return Volatile.Read(ref failedCount); } }
        internal int RejectedCount { get { return Volatile.Read(ref rejectedCount); } }
        internal int Generation { get { return Volatile.Read(ref generation); } }
        internal string LastError { get { return Volatile.Read(ref lastError); } }

        internal static void ValidateLaunchCore(string repositoryRoot)
        {
            string path = Path.Combine(Path.GetFullPath(repositoryRoot), "common", "launch-core.ps1");
            if (!File.Exists(path)) { throw new FileNotFoundException("launch-core.ps1 not found", path); }
            using (Runspace probe = RunspaceFactory.CreateRunspace())
            {
                probe.Open();
                LoadLaunchCore(probe, path);
            }
        }

        public void Activate(ResolvedBindingData binding)
        {
            TryActivate(binding);
        }

        internal bool TryActivate(ResolvedBindingData binding)
        {
            if (!accepting || binding == null || binding.Invoke == null || String.IsNullOrWhiteSpace(binding.ToolDir))
            {
                Reject("activation rejected: invalid request");
                return false;
            }
            ActivationRequest request = new ActivationRequest { Binding = binding };
            if (!queue.TryAdd(request))
            {
                Reject("activation rejected: queue capacity exceeded");
                return false;
            }
            Interlocked.Increment(ref queuedCount);
            return true;
        }

        internal void PauseForTest()
        {
            testGate.Reset();
        }

        internal void ResumeForTest()
        {
            testGate.Set();
        }

        internal void ThrowNextForTest()
        {
            Interlocked.Exchange(ref throwNext, 1);
        }

        internal void BreakRunspaceForTest()
        {
            Interlocked.Exchange(ref breakNext, 1);
        }

        private void Reject(string message)
        {
            Interlocked.Increment(ref rejectedCount);
            SetLastError(message);
            try { log(message); }
            catch { }
            try { notify(message); }
            catch { }
        }

        private void WorkerMain()
        {
            try
            {
                EnsureRunspace();
            }
            catch (Exception exception)
            {
                startupError = exception;
                ready.Set();
                return;
            }
            ready.Set();
            try
            {
                foreach (ActivationRequest request in queue.GetConsumingEnumerable())
                {
                    testGate.WaitOne();
                    ProcessRequest(request);
                }
            }
            catch (Exception exception)
            {
                SetLastError("activation worker stopped: " + exception.Message);
                try { log(LastError); }
                catch { }
            }
            finally
            {
                DisposeRunspace();
            }
        }

        private void ProcessRequest(ActivationRequest request)
        {
            try
            {
                if (Interlocked.Exchange(ref throwNext, 0) != 0)
                {
                    throw new InvalidOperationException("test request failure");
                }
                if (Interlocked.Exchange(ref breakNext, 0) != 0)
                {
                    DisposeRunspace();
                }
                EnsureRunspace();
                Collection<PSObject> output;
                using (PowerShell pipeline = PowerShell.Create())
                {
                    pipeline.Runspace = runspace;
                    pipeline.AddScript(InvokeScript, false)
                        .AddArgument(request.Binding.ToolDir)
                        .AddArgument(request.Binding.Invoke.Action);
                    lock (runspaceSync) { currentPipeline = pipeline; }
                    try { output = pipeline.Invoke(); }
                    finally { lock (runspaceSync) { currentPipeline = null; } }
                    if (pipeline.HadErrors)
                    {
                        throw new InvalidOperationException(JoinErrors(pipeline.Streams));
                    }
                }
                if (output == null || output.Count != 1)
                {
                    throw new InvalidOperationException("launch core returned an invalid result");
                }
                Interlocked.Increment(ref completedCount);
                try { log("activation started " + request.Binding.Id); }
                catch { }
            }
            catch (Exception exception)
            {
                Interlocked.Increment(ref failedCount);
                string message = "activation failed " + SafeBindingId(request.Binding) + ": " + exception.Message;
                SetLastError(message);
                try { log(message); }
                catch { }
                try { notify(message); }
                catch { }
                if (runspace != null && runspace.RunspaceStateInfo.State != RunspaceState.Opened)
                {
                    DisposeRunspace();
                }
            }
        }

        private static string SafeBindingId(ResolvedBindingData binding)
        {
            if (binding == null || String.IsNullOrWhiteSpace(binding.Id)) { return "<unknown>"; }
            return binding.Id;
        }

        private static string JoinErrors(PSDataStreams streams)
        {
            StringBuilder builder = new StringBuilder();
            foreach (ErrorRecord item in streams.Error)
            {
                if (builder.Length > 0) { builder.AppendLine(); }
                builder.Append(item.ToString());
            }
            return builder.Length == 0 ? "PowerShell invocation failed" : builder.ToString();
        }

        private void EnsureRunspace()
        {
            if (runspace != null && runspace.RunspaceStateInfo.State == RunspaceState.Opened) { return; }
            DisposeRunspace();
            Runspace created = RunspaceFactory.CreateRunspace();
            try
            {
                created.Open();
                LoadLaunchCore(created, launchCorePath);
                runspace = created;
                Interlocked.Increment(ref generation);
            }
            catch
            {
                created.Dispose();
                throw;
            }
        }

        private static void LoadLaunchCore(Runspace target, string path)
        {
            using (PowerShell loader = PowerShell.Create())
            {
                loader.Runspace = target;
                loader.AddScript("param($path) . $path", false).AddArgument(path);
                loader.Invoke();
                if (loader.HadErrors)
                {
                    throw new InvalidOperationException(JoinErrors(loader.Streams));
                }
            }
        }

        private void DisposeRunspace()
        {
            Runspace old = runspace;
            runspace = null;
            if (old != null)
            {
                try { old.Close(); }
                catch { }
                try { old.Dispose(); }
                catch { }
            }
        }

        private void SetLastError(string message)
        {
            Volatile.Write(ref lastError, message ?? String.Empty);
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref disposed, 1) != 0) { return; }
            accepting = false;
            queue.CompleteAdding();
            testGate.Set();
            if (thread.IsAlive && !thread.Join(2000))
            {
                lock (runspaceSync)
                {
                    if (currentPipeline != null)
                    {
                        try { currentPipeline.Stop(); }
                        catch { }
                    }
                }
                thread.Join(500);
            }
            DisposeRunspace();
            queue.Dispose();
            ready.Dispose();
            testGate.Dispose();
        }
    }

    internal sealed class HostConfigReloader : IDisposable
    {
        private readonly string stateRoot;
        private readonly string repositoryRoot;
        private readonly string sourcePath;
        private readonly string resolverPath;
        private readonly Action<HostStateBundle> apply;
        private readonly Action<string> log;
        private readonly Action<string> notify;
        private readonly object sync = new object();
        private readonly System.Threading.Timer timer;
        private FileSystemWatcher watcher;
        private bool running;
        private bool pending;
        private bool disposed;
        private int succeeded;
        private int failed;
        private string lastError = String.Empty;

        internal HostConfigReloader(HostStateBundle initial, Action<HostStateBundle> applyBundle,
            Action<string> logger, Action<string> notification)
        {
            stateRoot = initial.StateRoot;
            repositoryRoot = initial.State.Root;
            sourcePath = Path.GetFullPath(initial.Config.SourceConfigPath);
            resolverPath = Path.Combine(repositoryRoot, "common", "resolve-host-config.ps1");
            apply = applyBundle;
            log = logger;
            notify = notification;
            timer = new System.Threading.Timer(ReloadTimer, null, Timeout.Infinite, Timeout.Infinite);
        }

        internal int Succeeded { get { return Volatile.Read(ref succeeded); } }
        internal int Failed { get { return Volatile.Read(ref failed); } }
        internal string LastError { get { return Volatile.Read(ref lastError); } }

        internal void Start()
        {
            string directory = Path.GetDirectoryName(sourcePath);
            if (String.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            {
                throw new DirectoryNotFoundException("bindings directory not found: " + directory);
            }
            if (!File.Exists(resolverPath))
            {
                throw new FileNotFoundException("resolve-host-config.ps1 not found", resolverPath);
            }
            watcher = new FileSystemWatcher(directory, Path.GetFileName(sourcePath));
            watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.CreationTime;
            watcher.Changed += OnChanged;
            watcher.Created += OnChanged;
            watcher.Deleted += OnChanged;
            watcher.Renamed += OnRenamed;
            watcher.EnableRaisingEvents = true;
        }

        internal bool RequestReload()
        {
            return Schedule(1);
        }

        private void OnChanged(object sender, FileSystemEventArgs args)
        {
            Schedule(500);
        }

        private void OnRenamed(object sender, RenamedEventArgs args)
        {
            Schedule(500);
        }

        private bool Schedule(int delay)
        {
            lock (sync)
            {
                if (disposed) { return false; }
                timer.Change(delay, Timeout.Infinite);
                return true;
            }
        }

        private void ReloadTimer(object ignored)
        {
            lock (sync)
            {
                if (disposed) { return; }
                if (running)
                {
                    pending = true;
                    return;
                }
                running = true;
            }
            try { PerformReload(); }
            finally
            {
                lock (sync)
                {
                    running = false;
                    if (pending && !disposed)
                    {
                        pending = false;
                        timer.Change(500, Timeout.Infinite);
                    }
                }
            }
        }

        private void PerformReload()
        {
            string temporary = Path.Combine(stateRoot, "bindings.reload." + Guid.NewGuid().ToString("N") + ".json");
            try
            {
                ProcessStartInfo start = new ProcessStartInfo();
                start.FileName = "powershell.exe";
                start.Arguments = JoinArguments(new[] {
                    "-NoProfile", "-ExecutionPolicy", "Bypass", "-WindowStyle", "Hidden", "-File", resolverPath,
                    "-Root", repositoryRoot, "-BindingsPath", sourcePath, "-OutputPath", temporary
                });
                start.UseShellExecute = false;
                start.CreateNoWindow = true;
                start.WindowStyle = ProcessWindowStyle.Hidden;
                start.RedirectStandardOutput = true;
                start.RedirectStandardError = true;
                string standardError;
                int exitCode;
                using (Process process = Process.Start(start))
                {
                    process.StandardOutput.ReadToEnd();
                    standardError = process.StandardError.ReadToEnd();
                    if (!process.WaitForExit(15000))
                    {
                        try { process.Kill(); }
                        catch { }
                        throw new TimeoutException("binding resolver timed out");
                    }
                    exitCode = process.ExitCode;
                }
                if (exitCode != 0)
                {
                    throw new InvalidOperationException("binding resolver failed: " + standardError.Trim());
                }
                ResolvedConfigData candidate = HostStateBundle.LoadResolvedConfig(temporary, repositoryRoot);
                string actualHash = ComputeSha256(sourcePath);
                if (!String.Equals(candidate.SourceConfigSha256, actualHash, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException("resolved source hash does not match bindings.json");
                }
                string destination = Path.Combine(stateRoot, "bindings.resolved.json");
                if (File.Exists(destination)) { File.Replace(temporary, destination, null); }
                else { File.Move(temporary, destination); }
                HostStateBundle next = HostStateBundle.Load(stateRoot);
                apply(next);
                Interlocked.Increment(ref succeeded);
                Volatile.Write(ref lastError, String.Empty);
                try { log("bindings reloaded " + actualHash); }
                catch { }
            }
            catch (Exception exception)
            {
                Interlocked.Increment(ref failed);
                string message = "bindings reload failed: " + exception.Message;
                Volatile.Write(ref lastError, message);
                try { log(message); }
                catch { }
                try { notify(message); }
                catch { }
            }
            finally
            {
                try { if (File.Exists(temporary)) { File.Delete(temporary); } }
                catch { }
            }
        }

        private static string ComputeSha256(string path)
        {
            using (SHA256 sha = SHA256.Create())
            using (FileStream stream = File.OpenRead(path))
            {
                return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", String.Empty).ToLowerInvariant();
            }
        }

        private static string JoinArguments(IEnumerable<string> items)
        {
            List<string> result = new List<string>();
            foreach (string item in items) { result.Add(QuoteArgument(item)); }
            return String.Join(" ", result.ToArray());
        }

        private static string QuoteArgument(string value)
        {
            if (String.IsNullOrEmpty(value)) { return "\"\""; }
            if (value.IndexOfAny(new[] { ' ', '\t', '\n', '\v', '"' }) < 0) { return value; }
            StringBuilder result = new StringBuilder();
            result.Append('"');
            int backslashes = 0;
            foreach (char character in value)
            {
                if (character == '\\')
                {
                    backslashes++;
                }
                else if (character == '"')
                {
                    result.Append('\\', (backslashes * 2) + 1);
                    result.Append('"');
                    backslashes = 0;
                }
                else
                {
                    if (backslashes > 0) { result.Append('\\', backslashes); }
                    backslashes = 0;
                    result.Append(character);
                }
            }
            if (backslashes > 0) { result.Append('\\', backslashes * 2); }
            result.Append('"');
            return result.ToString();
        }

        public void Dispose()
        {
            lock (sync)
            {
                if (disposed) { return; }
                disposed = true;
                timer.Change(Timeout.Infinite, Timeout.Infinite);
            }
            if (watcher != null)
            {
                watcher.EnableRaisingEvents = false;
                watcher.Dispose();
                watcher = null;
            }
            timer.Dispose();
        }
    }
}
