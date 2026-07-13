using System;
using System.Collections.Generic;
using System.Diagnostics;
using ToolRack;

namespace ToolRackTests
{
    public sealed class MouseHarnessResult
    {
        public bool PlainPass;
        public bool ExactActivatesOnce;
        public bool ReleasedModifierStillSwallows;
        public bool ExtraModifierPasses;
        public bool UpOnlyPasses;
        public bool DoubleClickNoSecondActivation;
        public bool InjectedPasses;
        public bool NegativeCodePasses;
        public bool DisabledPasses;
        public bool ReloadDisposesOldHook;
        public bool ShutdownDisposesHook;
        public double P99Milliseconds;
    }

    internal sealed class FakeMouseHook : IMouseHookBackend
    {
        internal bool Disposed;
        public bool Active { get { return true; } }
        public string Error { get { return String.Empty; } }
        public void Dispose() { Disposed = true; }
    }

    internal sealed class FakeMouseFactory : IMouseHookFactory
    {
        internal readonly List<FakeMouseHook> Hooks = new List<FakeMouseHook>();
        public IMouseHookBackend Create(Dictionary<int, int> chords, IntPtr notifyWindow)
        {
            FakeMouseHook hook = new FakeMouseHook();
            Hooks.Add(hook);
            return hook;
        }
    }

    public static class MouseHarness
    {
        public static MouseHarnessResult Run()
        {
            MouseHarnessResult result = new MouseHarnessResult();
            Dictionary<int, int> chords = new Dictionary<int, int>();
            chords[MouseModifiers.Control] = 101;
            MouseGestureStateMachine machine = new MouseGestureStateMachine(chords);

            MouseDecision plainDown = machine.Process(0, MouseMessages.RightDown, 0, false, true);
            MouseDecision plainUp = machine.Process(0, MouseMessages.RightUp, 0, false, true);
            result.PlainPass = plainDown.Disposition == MouseDisposition.Pass && plainUp.Disposition == MouseDisposition.Pass;

            MouseDecision exactDown = machine.Process(0, MouseMessages.RightDown, MouseModifiers.Control, false, true);
            MouseDecision exactUp = machine.Process(0, MouseMessages.RightUp, MouseModifiers.Control, false, true);
            MouseDecision extraUp = machine.Process(0, MouseMessages.RightUp, MouseModifiers.Control, false, true);
            result.ExactActivatesOnce = exactDown.Disposition == MouseDisposition.Swallow &&
                exactUp.Disposition == MouseDisposition.SwallowActivate && exactUp.ActivationId == 101 &&
                extraUp.Disposition == MouseDisposition.Pass;

            machine.Process(0, MouseMessages.RightDown, MouseModifiers.Control, false, true);
            MouseDecision released = machine.Process(0, MouseMessages.RightUp, 0, false, true);
            result.ReleasedModifierStillSwallows = released.Disposition == MouseDisposition.SwallowActivate;

            MouseDecision extra = machine.Process(0, MouseMessages.RightDown, MouseModifiers.Control | MouseModifiers.Shift, false, true);
            result.ExtraModifierPasses = extra.Disposition == MouseDisposition.Pass;
            result.UpOnlyPasses = machine.Process(0, MouseMessages.RightUp, MouseModifiers.Control, false, true).Disposition == MouseDisposition.Pass;

            machine.Process(0, MouseMessages.RightDown, MouseModifiers.Control, false, true);
            MouseDecision firstUp = machine.Process(0, MouseMessages.RightUp, MouseModifiers.Control, false, true);
            MouseDecision doubleDown = machine.Process(0, MouseMessages.RightDouble, MouseModifiers.Control, false, true);
            MouseDecision doubleUp = machine.Process(0, MouseMessages.RightUp, MouseModifiers.Control, false, true);
            result.DoubleClickNoSecondActivation = firstUp.Disposition == MouseDisposition.SwallowActivate &&
                doubleDown.Disposition == MouseDisposition.Swallow && doubleUp.Disposition == MouseDisposition.Swallow;

            result.InjectedPasses = machine.Process(0, MouseMessages.RightDown, MouseModifiers.Control, true, true).Disposition == MouseDisposition.Pass;
            result.NegativeCodePasses = machine.Process(-1, MouseMessages.RightDown, MouseModifiers.Control, false, true).Disposition == MouseDisposition.Pass;
            result.DisabledPasses = machine.Process(0, MouseMessages.RightDown, MouseModifiers.Control, false, false).Disposition == MouseDisposition.Pass;

            FakeMouseFactory factory = new FakeMouseFactory();
            FakeSink sink = new FakeSink();
            MouseGestureManager manager = new MouseGestureManager(factory, sink, delegate(string ignored) { }, IntPtr.Zero);
            ResolvedBindingData binding = MouseBinding("capture-mouse", "ctrl");
            manager.Reload(new[] { binding });
            FakeMouseHook first = factory.Hooks[0];
            manager.Reload(new[] { binding });
            result.ReloadDisposesOldHook = first.Disposed && factory.Hooks.Count == 2;
            FakeMouseHook second = factory.Hooks[1];
            manager.Dispose();
            result.ShutdownDisposesHook = second.Disposed;

            long[] elapsed = new long[100000];
            Stopwatch stopwatch = Stopwatch.StartNew();
            for (int index = 0; index < elapsed.Length; index++)
            {
                long before = Stopwatch.GetTimestamp();
                machine.Process(0, MouseMessages.LeftDown, 0, false, true);
                elapsed[index] = Stopwatch.GetTimestamp() - before;
            }
            stopwatch.Stop();
            Array.Sort(elapsed);
            long p99 = elapsed[(int)Math.Ceiling(elapsed.Length * 0.99) - 1];
            result.P99Milliseconds = p99 * 1000.0 / Stopwatch.Frequency;
            return result;
        }

        private static ResolvedBindingData MouseBinding(string id, params string[] modifiers)
        {
            return new ResolvedBindingData
            {
                Id = id,
                Trigger = new TriggerData { Type = "mouse", Button = "right", Modifiers = new List<string>(modifiers) },
                Invoke = new InvokeData { Tool = "capture", Action = "default" },
                ToolDir = "C:\\capture"
            };
        }
    }
}
