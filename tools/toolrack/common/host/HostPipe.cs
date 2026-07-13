using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Text;
using System.Threading;

namespace ToolRack
{
    internal sealed class HostPipeServer : IDisposable
    {
        private const int MaximumCommandBytes = 512;
        private const int CommandReadTimeoutMilliseconds = 2000;
        private readonly string pipeName;
        private readonly Func<HostStatus> getStatus;
        private readonly Func<bool> requestReload;
        private readonly Action requestShutdown;
        private readonly Func<string, string> handleExtendedCommand;
        private readonly Action<string> log;
        private Thread thread;
        private volatile bool stopping;

        internal HostPipeServer(string name, Func<HostStatus> status, Func<bool> reload, Action shutdown,
            Func<string, string> extendedCommand, Action<string> logger)
        {
            pipeName = name;
            getStatus = status;
            requestReload = reload;
            requestShutdown = shutdown;
            handleExtendedCommand = extendedCommand;
            log = logger;
        }

        internal static bool CurrentUserAclIsExclusive()
        {
            PipeSecurity security = CreateCurrentUserSecurity();
            AuthorizationRuleCollection rules = security.GetAccessRules(true, true, typeof(SecurityIdentifier));
            SecurityIdentifier current = WindowsIdentity.GetCurrent().User;
            int allows = 0;
            foreach (AuthorizationRule item in rules)
            {
                PipeAccessRule rule = item as PipeAccessRule;
                if (rule != null && rule.AccessControlType == AccessControlType.Allow)
                {
                    allows++;
                    if (!current.Equals(rule.IdentityReference)) { return false; }
                }
            }
            return security.AreAccessRulesProtected && allows == 1;
        }

        internal void Start()
        {
            thread = new Thread(Run);
            thread.IsBackground = true;
            thread.Name = "ToolRackHostPipe";
            thread.Start();
        }

        private void Run()
        {
            while (!stopping)
            {
                try
                {
                    using (NamedPipeServerStream server = CreateServer())
                    {
                        server.WaitForConnection();
                        if (stopping) { return; }
                        using (StreamWriter writer = new StreamWriter(server, new UTF8Encoding(false), 1024, true))
                        {
                            writer.AutoFlush = true;
                            string command = ReadCommand(server);
                            if (String.Equals(command, "status", StringComparison.Ordinal))
                            {
                                writer.WriteLine(getStatus().ToJson());
                            }
                            else if (String.Equals(command, "reload", StringComparison.Ordinal))
                            {
                                bool accepted = requestReload != null && requestReload();
                                writer.WriteLine("{\"ok\":" + JsonBoolean(accepted) + ",\"reloaded\":" + JsonBoolean(accepted) + "}");
                            }
                            else if (String.Equals(command, "shutdown", StringComparison.Ordinal))
                            {
                                writer.WriteLine("{\"ok\":true,\"shutdown\":true}");
                                requestShutdown();
                            }
                            else if (handleExtendedCommand != null && command != null && command.StartsWith("test-", StringComparison.Ordinal))
                            {
                                writer.WriteLine(handleExtendedCommand(command));
                            }
                            else
                            {
                                writer.WriteLine("{\"ok\":false,\"error\":\"unknown command\"}");
                            }
                        }
                    }
                }
                catch (Exception exception)
                {
                    if (!stopping)
                    {
                        log("pipe error: " + exception.Message);
                        Thread.Sleep(100);
                    }
                }
            }
        }

        private static string ReadCommand(NamedPipeServerStream server)
        {
            Stopwatch watch = Stopwatch.StartNew();
            using (MemoryStream bytes = new MemoryStream())
            {
                while (true)
                {
                    int remaining = CommandReadTimeoutMilliseconds - (int)watch.ElapsedMilliseconds;
                    if (remaining <= 0)
                    {
                        throw new TimeoutException("pipe command read timed out");
                    }

                    byte[] next = new byte[1];
                    IAsyncResult pending = server.BeginRead(next, 0, 1, null, null);
                    WaitHandle wait = pending.AsyncWaitHandle;
                    if (!wait.WaitOne(remaining))
                    {
                        try { server.Close(); }
                        catch { }
                        try { server.EndRead(pending); }
                        catch { }
                        wait.Close();
                        throw new TimeoutException("pipe command read timed out");
                    }

                    int read;
                    try { read = server.EndRead(pending); }
                    finally { wait.Close(); }
                    if (read == 0) { break; }
                    if (next[0] == 10) { break; }
                    if (bytes.Length >= MaximumCommandBytes)
                    {
                        throw new InvalidDataException("pipe command is too long");
                    }
                    bytes.WriteByte(next[0]);
                }

                string command = new UTF8Encoding(false, true).GetString(bytes.ToArray());
                if (command.EndsWith("\r", StringComparison.Ordinal))
                {
                    command = command.Substring(0, command.Length - 1);
                }
                if (command.IndexOf('\r') >= 0)
                {
                    throw new InvalidDataException("pipe command contains an invalid line break");
                }
                return command;
            }
        }

        private static string JsonBoolean(bool value)
        {
            return value ? "true" : "false";
        }

        private NamedPipeServerStream CreateServer()
        {
            return new NamedPipeServerStream(
                pipeName,
                PipeDirection.InOut,
                1,
                PipeTransmissionMode.Byte,
                PipeOptions.Asynchronous,
                4096,
                4096,
                CreateCurrentUserSecurity());
        }

        private static PipeSecurity CreateCurrentUserSecurity()
        {
            SecurityIdentifier sid = WindowsIdentity.GetCurrent().User;
            PipeSecurity security = new PipeSecurity();
            security.SetAccessRuleProtection(true, false);
            security.SetOwner(sid);
            security.AddAccessRule(new PipeAccessRule(sid, PipeAccessRights.FullControl, AccessControlType.Allow));
            return security;
        }

        internal void Stop()
        {
            stopping = true;
            try
            {
                using (NamedPipeClientStream client = new NamedPipeClientStream(".", pipeName, PipeDirection.Out))
                {
                    client.Connect(200);
                }
            }
            catch { }
            if (thread != null && thread.IsAlive)
            {
                thread.Join(2000);
            }
        }

        public void Dispose()
        {
            Stop();
        }
    }
}
