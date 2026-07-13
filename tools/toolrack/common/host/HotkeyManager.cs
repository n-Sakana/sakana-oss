using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;

namespace ToolRack
{
    public interface IHotkeyBackend
    {
        bool Register(int id, uint modifiers, uint virtualKey);
        void Unregister(int id);
    }

    public interface IActivationSink
    {
        void Activate(ResolvedBindingData binding);
    }

    public sealed class HotkeyReloadResult
    {
        public int ActiveCount;
        public int InactiveCount;
    }

    public sealed class HotkeyManager : IDisposable
    {
        private const uint ModAlt = 0x0001;
        private const uint ModControl = 0x0002;
        private const uint ModShift = 0x0004;
        private const uint ModNoRepeat = 0x4000;
        private const int MaximumRegistrationId = 0xBFFF;

        private readonly IHotkeyBackend backend;
        private readonly IActivationSink sink;
        private readonly Action<string> log;
        private readonly Dictionary<int, ResolvedBindingData> active = new Dictionary<int, ResolvedBindingData>();
        private int inactiveCount;

        public HotkeyManager(IHotkeyBackend hotkeyBackend, IActivationSink activationSink, Action<string> logger)
        {
            backend = hotkeyBackend;
            sink = activationSink;
            log = logger;
        }

        public int ActiveCount { get { return active.Count; } }
        public int InactiveCount { get { return inactiveCount; } }

        public HotkeyReloadResult Reload(IEnumerable<ResolvedBindingData> bindings)
        {
            UnregisterAll();
            inactiveCount = 0;
            HashSet<string> triggers = new HashSet<string>(StringComparer.Ordinal);
            HashSet<int> registrationIds = new HashSet<int>();
            foreach (ResolvedBindingData binding in bindings)
            {
                if (binding == null || binding.Trigger == null ||
                    !String.Equals(binding.Trigger.Type, "hotkey", StringComparison.Ordinal))
                {
                    continue;
                }
                string trigger = FormatTrigger(binding.Trigger);
                if (!triggers.Add(trigger))
                {
                    inactiveCount++;
                    log("hotkey inactive " + binding.Id + ": duplicate trigger " + trigger);
                    continue;
                }
                int registrationId = GetStableRegistrationId(binding.Id);
                if (!registrationIds.Add(registrationId))
                {
                    inactiveCount++;
                    log("hotkey inactive " + binding.Id + ": registration ID collision");
                    continue;
                }
                uint modifiers;
                uint virtualKey;
                if (!TryConvert(binding.Trigger, out modifiers, out virtualKey))
                {
                    inactiveCount++;
                    log("hotkey inactive " + binding.Id + ": invalid normalized trigger");
                    continue;
                }
                modifiers |= ModNoRepeat;
                if (!backend.Register(registrationId, modifiers, virtualKey))
                {
                    inactiveCount++;
                    log("hotkey inactive " + binding.Id + ": RegisterHotKey failed " + Marshal.GetLastWin32Error());
                    continue;
                }
                active[registrationId] = binding;
            }
            return new HotkeyReloadResult { ActiveCount = active.Count, InactiveCount = inactiveCount };
        }

        public bool HandleHotkey(int registrationId)
        {
            ResolvedBindingData binding;
            if (!active.TryGetValue(registrationId, out binding)) { return false; }
            try
            {
                sink.Activate(binding);
                return true;
            }
            catch (Exception exception)
            {
                ThreadPool.QueueUserWorkItem(delegate { log("hotkey dispatch failed: " + exception.Message); });
                return false;
            }
        }

        private void UnregisterAll()
        {
            foreach (int id in new List<int>(active.Keys))
            {
                try { backend.Unregister(id); }
                catch (Exception exception) { log("UnregisterHotKey failed " + id + ": " + exception.Message); }
            }
            active.Clear();
        }

        private static int GetStableRegistrationId(string bindingId)
        {
            uint hash = 2166136261;
            foreach (char character in bindingId)
            {
                hash ^= character;
                hash *= 16777619;
            }
            return (int)(hash % MaximumRegistrationId) + 1;
        }

        private static string FormatTrigger(TriggerData trigger)
        {
            return String.Join("+", trigger.Modifiers.ToArray()) + "+" + trigger.Key;
        }

        private static bool TryConvert(TriggerData trigger, out uint modifiers, out uint virtualKey)
        {
            modifiers = 0;
            virtualKey = 0;
            if (trigger.Modifiers == null || trigger.Modifiers.Count == 0) { return false; }
            foreach (string modifier in trigger.Modifiers)
            {
                if (String.Equals(modifier, "ctrl", StringComparison.Ordinal)) { modifiers |= ModControl; }
                else if (String.Equals(modifier, "alt", StringComparison.Ordinal)) { modifiers |= ModAlt; }
                else if (String.Equals(modifier, "shift", StringComparison.Ordinal)) { modifiers |= ModShift; }
                else { return false; }
            }
            string key = trigger.Key;
            if (String.IsNullOrEmpty(key)) { return false; }
            if (key.Length == 1 && key[0] >= 'A' && key[0] <= 'Z')
            {
                virtualKey = key[0];
                return true;
            }
            if (key.Length == 1 && key[0] >= '0' && key[0] <= '9')
            {
                virtualKey = key[0];
                return true;
            }
            if (key[0] == 'F')
            {
                int number;
                if (Int32.TryParse(key.Substring(1), out number) &&
                    ((number >= 1 && number <= 11) || (number >= 13 && number <= 24)))
                {
                    virtualKey = (uint)(0x70 + number - 1);
                    return true;
                }
            }
            return false;
        }

        public void Dispose()
        {
            UnregisterAll();
        }
    }

    internal sealed class Win32HotkeyBackend : IHotkeyBackend
    {
        private readonly IntPtr window;

        internal Win32HotkeyBackend(IntPtr handle)
        {
            window = handle;
        }

        public bool Register(int id, uint modifiers, uint virtualKey)
        {
            return NativeMethods.RegisterHotKey(window, id, modifiers, virtualKey);
        }

        public void Unregister(int id)
        {
            NativeMethods.UnregisterHotKey(window, id);
        }
    }

    internal sealed class LoggingActivationSink : IActivationSink
    {
        private readonly Action<string> log;

        internal LoggingActivationSink(Action<string> logger)
        {
            log = logger;
        }

        public void Activate(ResolvedBindingData binding)
        {
            ThreadPool.QueueUserWorkItem(delegate
            {
                try { log("activation queued " + binding.Id); }
                catch { }
            });
        }
    }
}
