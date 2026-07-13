using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace ToolRack
{
    public static class MouseModifiers
    {
        public const int Control = 1;
        public const int Alt = 2;
        public const int Shift = 4;
    }

    public static class MouseMessages
    {
        public const int LeftDown = 0x0201;
        public const int RightDown = 0x0204;
        public const int RightUp = 0x0205;
        public const int RightDouble = 0x0206;
        public const int MiddleDown = 0x0207;
        public const int XDown = 0x020B;
    }

    public enum MouseDisposition
    {
        Pass,
        Swallow,
        SwallowActivate
    }

    public struct MouseDecision
    {
        public MouseDisposition Disposition;
        public int ActivationId;
    }

    public sealed class MouseGestureStateMachine
    {
        private readonly Dictionary<int, int> chords;
        private int armedActivationId;
        private bool activateOnUp;

        public MouseGestureStateMachine(Dictionary<int, int> modifierChords)
        {
            chords = new Dictionary<int, int>(modifierChords);
        }

        public MouseDecision Process(int hookCode, int message, int modifiers, bool injected, bool enabled)
        {
            if (hookCode < 0 || injected) { return Pass(); }
            if (!enabled)
            {
                armedActivationId = 0;
                activateOnUp = false;
                return Pass();
            }
            if (message == MouseMessages.RightDown)
            {
                int activationId;
                if (chords.TryGetValue(modifiers, out activationId))
                {
                    armedActivationId = activationId;
                    activateOnUp = true;
                    return Swallow();
                }
                armedActivationId = 0;
                activateOnUp = false;
                return Pass();
            }
            if (message == MouseMessages.RightDouble)
            {
                int activationId;
                if (chords.TryGetValue(modifiers, out activationId))
                {
                    armedActivationId = activationId;
                    activateOnUp = false;
                    return Swallow();
                }
                armedActivationId = 0;
                activateOnUp = false;
                return Pass();
            }
            if (message == MouseMessages.RightUp)
            {
                if (armedActivationId == 0) { return Pass(); }
                int activationId = armedActivationId;
                bool activate = activateOnUp;
                armedActivationId = 0;
                activateOnUp = false;
                return new MouseDecision
                {
                    Disposition = activate ? MouseDisposition.SwallowActivate : MouseDisposition.Swallow,
                    ActivationId = activate ? activationId : 0
                };
            }
            if (message == MouseMessages.LeftDown || message == MouseMessages.MiddleDown || message == MouseMessages.XDown)
            {
                armedActivationId = 0;
                activateOnUp = false;
            }
            return Pass();
        }

        private static MouseDecision Pass()
        {
            return new MouseDecision { Disposition = MouseDisposition.Pass, ActivationId = 0 };
        }

        private static MouseDecision Swallow()
        {
            return new MouseDecision { Disposition = MouseDisposition.Swallow, ActivationId = 0 };
        }
    }

    public interface IMouseHookBackend : IDisposable
    {
        bool Active { get; }
        string Error { get; }
    }

    public interface IMouseHookFactory
    {
        IMouseHookBackend Create(Dictionary<int, int> chords, IntPtr notifyWindow);
    }

    public sealed class MouseGestureManager : IDisposable
    {
        private const int MaximumActivationId = 0x7FFF;
        private readonly IMouseHookFactory factory;
        private readonly IActivationSink sink;
        private readonly Action<string> log;
        private readonly IntPtr notifyWindow;
        private readonly Dictionary<int, ResolvedBindingData> bindings = new Dictionary<int, ResolvedBindingData>();
        private IMouseHookBackend hook;
        private int inactiveCount;

        public MouseGestureManager(IMouseHookFactory hookFactory, IActivationSink activationSink, Action<string> logger, IntPtr window)
        {
            factory = hookFactory;
            sink = activationSink;
            log = logger;
            notifyWindow = window;
        }

        public int ActiveCount { get { return hook != null && hook.Active ? bindings.Count : 0; } }
        public int InactiveCount { get { return inactiveCount; } }
        public bool HookActive { get { return hook != null && hook.Active; } }

        public void Reload(IEnumerable<ResolvedBindingData> sourceBindings)
        {
            DisposeHook();
            bindings.Clear();
            inactiveCount = 0;
            Dictionary<int, int> chords = new Dictionary<int, int>();
            HashSet<int> activationIds = new HashSet<int>();
            foreach (ResolvedBindingData binding in sourceBindings)
            {
                if (binding == null || binding.Trigger == null ||
                    !String.Equals(binding.Trigger.Type, "mouse", StringComparison.Ordinal))
                {
                    continue;
                }
                int modifiers;
                if (!TryGetModifiers(binding.Trigger.Modifiers, out modifiers) || chords.ContainsKey(modifiers))
                {
                    inactiveCount++;
                    log("mouse binding inactive " + binding.Id + ": invalid or duplicate chord");
                    continue;
                }
                int activationId = GetStableActivationId(binding.Id);
                if (!activationIds.Add(activationId))
                {
                    inactiveCount++;
                    log("mouse binding inactive " + binding.Id + ": activation ID collision");
                    continue;
                }
                chords[modifiers] = activationId;
                bindings[activationId] = binding;
            }
            if (chords.Count == 0) { return; }
            hook = factory.Create(chords, notifyWindow);
            if (!hook.Active)
            {
                inactiveCount += bindings.Count;
                log("mouse hook inactive: " + hook.Error);
                bindings.Clear();
                DisposeHook();
            }
        }

        public bool HandleActivation(int activationId)
        {
            ResolvedBindingData binding;
            if (!bindings.TryGetValue(activationId, out binding)) { return false; }
            try
            {
                sink.Activate(binding);
                return true;
            }
            catch (Exception exception)
            {
                ThreadPool.QueueUserWorkItem(delegate { log("mouse dispatch failed: " + exception.Message); });
                return false;
            }
        }

        private static bool TryGetModifiers(List<string> source, out int modifiers)
        {
            modifiers = 0;
            if (source == null || source.Count == 0) { return false; }
            foreach (string item in source)
            {
                if (String.Equals(item, "ctrl", StringComparison.Ordinal)) { modifiers |= MouseModifiers.Control; }
                else if (String.Equals(item, "alt", StringComparison.Ordinal)) { modifiers |= MouseModifiers.Alt; }
                else if (String.Equals(item, "shift", StringComparison.Ordinal)) { modifiers |= MouseModifiers.Shift; }
                else { return false; }
            }
            return true;
        }

        private static int GetStableActivationId(string bindingId)
        {
            uint hash = 2166136261;
            foreach (char character in bindingId)
            {
                hash ^= character;
                hash *= 16777619;
            }
            return (int)(hash % MaximumActivationId) + 1;
        }

        private void DisposeHook()
        {
            if (hook != null)
            {
                hook.Dispose();
                hook = null;
            }
        }

        public void Dispose()
        {
            DisposeHook();
            bindings.Clear();
        }
    }

    internal sealed class Win32MouseHookFactory : IMouseHookFactory
    {
        public IMouseHookBackend Create(Dictionary<int, int> chords, IntPtr notifyWindow)
        {
            return new LowLevelMouseHook(chords, notifyWindow);
        }
    }

    internal sealed class LowLevelMouseHook : IMouseHookBackend
    {
        private const int WhMouseLl = 14;
        private const int WmQuit = 0x0012;
        private const int VkControl = 0x11;
        private const int VkMenu = 0x12;
        private const int VkShift = 0x10;
        private const uint LlInjected = 0x00000001;
        private const uint LlLowerIntegrityInjected = 0x00000002;

        private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct Point
        {
            internal int X;
            internal int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MouseHookData
        {
            internal Point Position;
            internal uint MouseData;
            internal uint Flags;
            internal uint Time;
            internal UIntPtr ExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int hookId, HookProc callback, IntPtr module, uint threadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hook);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int virtualKey);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool PostThreadMessage(uint threadId, int message, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string moduleName);

        private readonly MouseGestureStateMachine stateMachine;
        private readonly IntPtr notifyWindow;
        private readonly ManualResetEvent ready = new ManualResetEvent(false);
        private readonly Thread thread;
        private HookProc callback;
        private IntPtr hookHandle;
        private uint threadId;
        private volatile bool stopping;
        private volatile bool active;
        private string error = String.Empty;

        internal LowLevelMouseHook(Dictionary<int, int> chords, IntPtr window)
        {
            stateMachine = new MouseGestureStateMachine(chords);
            notifyWindow = window;
            thread = new Thread(ThreadMain);
            thread.IsBackground = true;
            thread.Name = "ToolRackMouseHook";
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            if (!ready.WaitOne(2000)) { error = "hook thread readiness timed out"; }
        }

        public bool Active { get { return active; } }
        public string Error { get { return error; } }

        private void ThreadMain()
        {
            threadId = GetCurrentThreadId();
            callback = HookCallback;
            try
            {
                using (Process process = Process.GetCurrentProcess())
                using (ProcessModule module = process.MainModule)
                {
                    hookHandle = SetWindowsHookEx(WhMouseLl, callback, GetModuleHandle(module.ModuleName), 0);
                }
                if (hookHandle == IntPtr.Zero)
                {
                    error = "SetWindowsHookEx failed " + Marshal.GetLastWin32Error();
                    ready.Set();
                    return;
                }
                active = true;
                ready.Set();
                Application.Run(new ApplicationContext());
            }
            catch (Exception exception)
            {
                error = exception.Message;
                ready.Set();
            }
            finally
            {
                active = false;
                if (hookHandle != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(hookHandle);
                    hookHandle = IntPtr.Zero;
                }
                callback = null;
                ready.Set();
            }
        }

        private IntPtr HookCallback(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code < 0) { return CallNextHookEx(hookHandle, code, wParam, lParam); }
            try
            {
                MouseHookData data = (MouseHookData)Marshal.PtrToStructure(lParam, typeof(MouseHookData));
                bool injected = (data.Flags & (LlInjected | LlLowerIntegrityInjected)) != 0;
                MouseDecision decision = stateMachine.Process(code, wParam.ToInt32(), GetModifiers(), injected, !stopping);
                if (decision.Disposition == MouseDisposition.Pass)
                {
                    return CallNextHookEx(hookHandle, code, wParam, lParam);
                }
                if (decision.Disposition == MouseDisposition.SwallowActivate)
                {
                    NativeMethods.PostMessage(notifyWindow, NativeMethods.WmAppMouseActivate, new IntPtr(decision.ActivationId), IntPtr.Zero);
                }
                return new IntPtr(1);
            }
            catch
            {
                return CallNextHookEx(hookHandle, code, wParam, lParam);
            }
        }

        private static int GetModifiers()
        {
            int modifiers = 0;
            if ((GetAsyncKeyState(VkControl) & 0x8000) != 0) { modifiers |= MouseModifiers.Control; }
            if ((GetAsyncKeyState(VkMenu) & 0x8000) != 0) { modifiers |= MouseModifiers.Alt; }
            if ((GetAsyncKeyState(VkShift) & 0x8000) != 0) { modifiers |= MouseModifiers.Shift; }
            return modifiers;
        }

        public void Dispose()
        {
            stopping = true;
            if (threadId != 0) { PostThreadMessage(threadId, WmQuit, IntPtr.Zero, IntPtr.Zero); }
            if (thread.IsAlive) { thread.Join(2000); }
            ready.Dispose();
        }
    }
}
