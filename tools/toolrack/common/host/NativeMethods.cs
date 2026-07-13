using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ToolRack
{
    internal static class NativeMethods
    {
        internal const int WmAppShutdown = 0x8001;
        internal const int WmHotkey = 0x0312;
        internal const int WmAppMouseActivate = 0x8002;
        internal const int WmAppConfigReload = 0x8003;
        internal static readonly IntPtr HwndMessage = new IntPtr(-3);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool PostMessage(IntPtr window, int message, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool RegisterHotKey(IntPtr window, int id, uint modifiers, uint virtualKey);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool UnregisterHotKey(IntPtr window, int id);
    }

    internal sealed class HostMessageWindow : NativeWindow, IDisposable
    {
        private readonly Action shutdown;
        private readonly Action<int> hotkey;
        private readonly Action<int> mouseActivation;
        private readonly Action configReload;

        internal HostMessageWindow(Action shutdownAction, Action<int> hotkeyAction, Action<int> mouseActivationAction,
            Action configReloadAction)
        {
            shutdown = shutdownAction;
            hotkey = hotkeyAction;
            mouseActivation = mouseActivationAction;
            configReload = configReloadAction;
            CreateParams parameters = new CreateParams();
            parameters.Caption = "ToolRackHostMessageWindow";
            parameters.Parent = NativeMethods.HwndMessage;
            CreateHandle(parameters);
        }

        internal void PostConfigReload()
        {
            NativeMethods.PostMessage(Handle, NativeMethods.WmAppConfigReload, IntPtr.Zero, IntPtr.Zero);
        }

        internal void PostShutdown()
        {
            NativeMethods.PostMessage(Handle, NativeMethods.WmAppShutdown, IntPtr.Zero, IntPtr.Zero);
        }

        protected override void WndProc(ref Message message)
        {
            if (message.Msg == NativeMethods.WmAppShutdown)
            {
                shutdown();
                return;
            }
            if (message.Msg == NativeMethods.WmHotkey)
            {
                hotkey(message.WParam.ToInt32());
                return;
            }
            if (message.Msg == NativeMethods.WmAppMouseActivate)
            {
                mouseActivation(message.WParam.ToInt32());
                return;
            }
            if (message.Msg == NativeMethods.WmAppConfigReload)
            {
                configReload();
                return;
            }
            base.WndProc(ref message);
        }

        public void Dispose()
        {
            if (Handle != IntPtr.Zero)
            {
                DestroyHandle();
            }
        }
    }
}
