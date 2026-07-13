using System;
using System.Collections.Generic;
using ToolRack;

namespace ToolRackTests
{
    public sealed class HarnessResult
    {
        public bool InitialRegistered;
        public bool DuplicateRejected;
        public bool FailureIsolated;
        public bool ReloadUnregistered;
        public bool NoRepeat;
        public bool DispatchQueued;
        public bool DisposeUnregistered;
        public bool StableIds;
    }

    internal sealed class FakeBackend : IHotkeyBackend
    {
        internal readonly Dictionary<int, Registration> Registrations = new Dictionary<int, Registration>();
        internal readonly List<int> Unregistered = new List<int>();
        internal uint FailVirtualKey;

        public bool Register(int id, uint modifiers, uint virtualKey)
        {
            if (virtualKey == FailVirtualKey) { return false; }
            Registrations[id] = new Registration { Id = id, Modifiers = modifiers, VirtualKey = virtualKey };
            return true;
        }

        public void Unregister(int id)
        {
            Unregistered.Add(id);
            Registrations.Remove(id);
        }
    }

    internal sealed class Registration
    {
        internal int Id;
        internal uint Modifiers;
        internal uint VirtualKey;
    }

    internal sealed class FakeSink : IActivationSink
    {
        internal readonly List<string> Activated = new List<string>();

        public void Activate(ResolvedBindingData binding)
        {
            Activated.Add(binding.Id);
        }
    }

    public static class HotkeyHarness
    {
        private const uint ModNoRepeat = 0x4000;

        public static HarnessResult Run()
        {
            HarnessResult result = new HarnessResult();
            FakeBackend backend = new FakeBackend();
            FakeSink sink = new FakeSink();
            HotkeyManager manager = new HotkeyManager(backend, sink, delegate(string ignored) { });

            ResolvedBindingData capture = Binding("capture-hotkey", "C", "ctrl", "alt");
            ResolvedBindingData transcribe = Binding("transcribe-hotkey", "T", "ctrl", "alt");
            HotkeyReloadResult initial = manager.Reload(new[] { capture, transcribe });
            result.InitialRegistered = initial.ActiveCount == 2 && initial.InactiveCount == 0 && backend.Registrations.Count == 2;
            result.NoRepeat = true;
            foreach (Registration registration in backend.Registrations.Values)
            {
                if ((registration.Modifiers & ModNoRepeat) == 0) { result.NoRepeat = false; }
            }
            Dictionary<uint, int> firstIds = IdsByKey(backend.Registrations);

            HotkeyReloadResult reversed = manager.Reload(new[] { transcribe, capture });
            Dictionary<uint, int> secondIds = IdsByKey(backend.Registrations);
            result.StableIds = reversed.ActiveCount == 2 && firstIds[0x43] == secondIds[0x43] && firstIds[0x54] == secondIds[0x54];
            result.ReloadUnregistered = backend.Unregistered.Count >= 2;

            ResolvedBindingData duplicate = Binding("duplicate", "C", "ctrl", "alt");
            HotkeyReloadResult duplicateResult = manager.Reload(new[] { capture, duplicate });
            result.DuplicateRejected = duplicateResult.ActiveCount == 1 && duplicateResult.InactiveCount == 1;

            backend.FailVirtualKey = 0x54;
            HotkeyReloadResult failed = manager.Reload(new[] { capture, transcribe });
            result.FailureIsolated = failed.ActiveCount == 1 && failed.InactiveCount == 1 && backend.Registrations.Count == 1;
            backend.FailVirtualKey = 0;

            manager.Reload(new[] { capture });
            int captureId = 0;
            foreach (Registration registration in backend.Registrations.Values)
            {
                if (registration.VirtualKey == 0x43) { captureId = registration.Id; }
            }
            result.DispatchQueued = captureId != 0 && manager.HandleHotkey(captureId) && sink.Activated.Count == 1 && sink.Activated[0] == "capture-hotkey";

            int unregisterBeforeDispose = backend.Unregistered.Count;
            manager.Dispose();
            result.DisposeUnregistered = backend.Unregistered.Count > unregisterBeforeDispose && backend.Registrations.Count == 0;
            return result;
        }

        private static Dictionary<uint, int> IdsByKey(Dictionary<int, Registration> registrations)
        {
            Dictionary<uint, int> result = new Dictionary<uint, int>();
            foreach (Registration registration in registrations.Values)
            {
                result[registration.VirtualKey] = registration.Id;
            }
            return result;
        }

        private static ResolvedBindingData Binding(string id, string key, params string[] modifiers)
        {
            return new ResolvedBindingData
            {
                Id = id,
                Trigger = new TriggerData
                {
                    Type = "hotkey",
                    Key = key,
                    Modifiers = new List<string>(modifiers)
                },
                Invoke = new InvokeData { Tool = "capture", Action = "default" },
                ToolDir = "C:\\capture"
            };
        }
    }
}
