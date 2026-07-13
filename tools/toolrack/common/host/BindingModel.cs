using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Text.RegularExpressions;

namespace ToolRack
{
    [DataContract]
    public sealed class HostStateData
    {
        [DataMember(Name = "schema")]
        public int Schema { get; set; }

        [DataMember(Name = "version")]
        public string Version { get; set; }

        [DataMember(Name = "root")]
        public string Root { get; set; }

        [DataMember(Name = "namespace")]
        public string Namespace { get; set; }
    }

    [DataContract]
    public sealed class ResolvedConfigData
    {
        [DataMember(Name = "schema")]
        public int Schema { get; set; }

        [DataMember(Name = "root")]
        public string Root { get; set; }

        [DataMember(Name = "sourceConfigPath")]
        public string SourceConfigPath { get; set; }

        [DataMember(Name = "sourceConfigSha256")]
        public string SourceConfigSha256 { get; set; }

        [DataMember(Name = "active")]
        public List<ResolvedBindingData> Active { get; set; }

        [DataMember(Name = "rejected")]
        public List<RejectedBindingData> Rejected { get; set; }
    }

    [DataContract]
    public sealed class ResolvedBindingData
    {
        [DataMember(Name = "id")]
        public string Id { get; set; }

        [DataMember(Name = "trigger")]
        public TriggerData Trigger { get; set; }

        [DataMember(Name = "invoke")]
        public InvokeData Invoke { get; set; }

        [DataMember(Name = "toolDir")]
        public string ToolDir { get; set; }
    }

    [DataContract]
    public sealed class TriggerData
    {
        [DataMember(Name = "type")]
        public string Type { get; set; }

        [DataMember(Name = "key", EmitDefaultValue = false)]
        public string Key { get; set; }

        [DataMember(Name = "button", EmitDefaultValue = false)]
        public string Button { get; set; }

        [DataMember(Name = "modifiers")]
        public List<string> Modifiers { get; set; }
    }

    [DataContract]
    public sealed class InvokeData
    {
        [DataMember(Name = "tool")]
        public string Tool { get; set; }

        [DataMember(Name = "action")]
        public string Action { get; set; }
    }

    [DataContract]
    public sealed class RejectedBindingData
    {
        [DataMember(Name = "id")]
        public string Id { get; set; }

        [DataMember(Name = "reason")]
        public string Reason { get; set; }
    }

    public sealed class HostStateBundle
    {
        public string StateRoot { get; private set; }
        public HostStateData State { get; private set; }
        public ResolvedConfigData Config { get; private set; }

        public static HostStateBundle Load(string stateRoot)
        {
            if (String.IsNullOrWhiteSpace(stateRoot))
            {
                throw new InvalidDataException("state root is required");
            }
            string fullStateRoot = Path.GetFullPath(stateRoot);
            string hostPath = Path.Combine(fullStateRoot, "host.json");
            string configPath = Path.Combine(fullStateRoot, "bindings.resolved.json");
            HostStateData state = ReadJson<HostStateData>(hostPath, "host.json");

            if (state.Schema != 1)
            {
                throw new InvalidDataException("host.json schema must be 1");
            }
            if (String.IsNullOrWhiteSpace(state.Version))
            {
                throw new InvalidDataException("host.json version is required");
            }
            if (String.IsNullOrWhiteSpace(state.Root))
            {
                throw new InvalidDataException("host.json root is required");
            }
            if (String.IsNullOrWhiteSpace(state.Namespace) ||
                !Regex.IsMatch(state.Namespace, "^[A-Za-z0-9-]{1,80}$", RegexOptions.CultureInvariant))
            {
                throw new InvalidDataException("host.json namespace is invalid");
            }
            string stateRootPath = Path.GetFullPath(state.Root);
            ResolvedConfigData config = LoadResolvedConfig(configPath, stateRootPath);
            state.Root = stateRootPath;

            return new HostStateBundle
            {
                StateRoot = fullStateRoot,
                State = state,
                Config = config
            };
        }

        internal static ResolvedConfigData LoadResolvedConfig(string path, string expectedRoot)
        {
            ResolvedConfigData config = ReadJson<ResolvedConfigData>(path, "bindings.resolved.json");
            if (config.Schema != 1)
            {
                throw new InvalidDataException("bindings.resolved.json schema must be 1");
            }
            string expectedRootPath = Path.GetFullPath(expectedRoot);
            string configRootPath;
            try { configRootPath = Path.GetFullPath(config.Root ?? String.Empty); }
            catch (Exception exception)
            {
                throw new InvalidDataException("bindings.resolved.json root is invalid", exception);
            }
            if (!String.Equals(expectedRootPath.TrimEnd('\\'), configRootPath.TrimEnd('\\'), StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidDataException("bindings.resolved.json root does not match host.json root");
            }
            if (String.IsNullOrWhiteSpace(config.SourceConfigPath))
            {
                throw new InvalidDataException("bindings.resolved.json source path is required");
            }
            try { config.SourceConfigPath = Path.GetFullPath(config.SourceConfigPath); }
            catch (Exception exception)
            {
                throw new InvalidDataException("bindings.resolved.json source path is invalid", exception);
            }
            if (String.IsNullOrWhiteSpace(config.SourceConfigSha256) ||
                !Regex.IsMatch(config.SourceConfigSha256, "^[A-Fa-f0-9]{64}$", RegexOptions.CultureInvariant))
            {
                throw new InvalidDataException("bindings.resolved.json source hash is invalid");
            }
            config.Root = configRootPath;
            if (config.Active == null) { config.Active = new List<ResolvedBindingData>(); }
            if (config.Rejected == null) { config.Rejected = new List<RejectedBindingData>(); }
            return config;
        }

        private static T ReadJson<T>(string path, string label)
        {
            if (!File.Exists(path))
            {
                throw new FileNotFoundException(label + " not found", path);
            }
            string text;
            try
            {
                using (StreamReader reader = new StreamReader(path, new UTF8Encoding(false, true), true))
                {
                    text = reader.ReadToEnd();
                }
                byte[] bytes = new UTF8Encoding(false, true).GetBytes(text);
                using (MemoryStream stream = new MemoryStream(bytes))
                {
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
                    return (T)serializer.ReadObject(stream);
                }
            }
            catch (Exception exception)
            {
                throw new InvalidDataException(label + " is invalid: " + exception.Message, exception);
            }
        }
    }

    [DataContract]
    public sealed class HostStatus
    {
        [DataMember(Name = "ok")]
        public bool Ok { get; set; }

        [DataMember(Name = "ready")]
        public bool Ready { get; set; }

        [DataMember(Name = "pid")]
        public int ProcessId { get; set; }

        [DataMember(Name = "root")]
        public string Root { get; set; }

        [DataMember(Name = "version")]
        public string Version { get; set; }

        [DataMember(Name = "activeBindings")]
        public int ActiveBindings { get; set; }

        [DataMember(Name = "rejectedBindings")]
        public int RejectedBindings { get; set; }

        [DataMember(Name = "pipeName")]
        public string PipeName { get; set; }

        [DataMember(Name = "pipeCurrentUserOnly")]
        public bool PipeCurrentUserOnly { get; set; }

        [DataMember(Name = "activeHotkeys")]
        public int ActiveHotkeys { get; set; }

        [DataMember(Name = "inactiveHotkeys")]
        public int InactiveHotkeys { get; set; }

        [DataMember(Name = "mouseHookActive")]
        public bool MouseHookActive { get; set; }

        [DataMember(Name = "activeMouseBindings")]
        public int ActiveMouseBindings { get; set; }

        [DataMember(Name = "inactiveMouseBindings")]
        public int InactiveMouseBindings { get; set; }

        [DataMember(Name = "activationQueueDepth")]
        public int ActivationQueueDepth { get; set; }

        [DataMember(Name = "activationQueued")]
        public int ActivationQueued { get; set; }

        [DataMember(Name = "activationCompleted")]
        public int ActivationCompleted { get; set; }

        [DataMember(Name = "activationFailed")]
        public int ActivationFailed { get; set; }

        [DataMember(Name = "activationRejected")]
        public int ActivationRejected { get; set; }

        [DataMember(Name = "workerGeneration")]
        public int WorkerGeneration { get; set; }

        [DataMember(Name = "reloadSucceeded")]
        public int ReloadSucceeded { get; set; }

        [DataMember(Name = "reloadFailed")]
        public int ReloadFailed { get; set; }

        [DataMember(Name = "notificationCount")]
        public int NotificationCount { get; set; }

        [DataMember(Name = "lastError")]
        public string LastError { get; set; }

        [DataMember(Name = "configSourceSha256")]
        public string ConfigSourceSha256 { get; set; }

        public string ToJson()
        {
            return JsonUtility.Serialize(this);
        }
    }

    public static class JsonUtility
    {
        public static string Serialize<T>(T value)
        {
            DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(T));
            using (MemoryStream stream = new MemoryStream())
            {
                serializer.WriteObject(stream, value);
                return Encoding.UTF8.GetString(stream.ToArray());
            }
        }
    }
}
