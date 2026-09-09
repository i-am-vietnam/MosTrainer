using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace MosTrainer.Excel
{
    internal static class WinApiProcessHelper
    {
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        public static int GetProcessIdFromHwnd(IntPtr hwnd)
        {
            uint pid;
            GetWindowThreadProcessId(hwnd, out pid);
            return (int)pid;
        }

        public static bool IsProcessAlive(int pid)
        {
            if (pid <= 0) return false;
            try
            {
                var p = Process.GetProcessById(pid);
                return !p.HasExited;
            }
            catch
            {
                return false;
            }
        }

        public static bool WaitForExit(int pid, int timeoutMs)
        {
            if (pid <= 0) return true;
            try
            {
                var p = Process.GetProcessById(pid);
                return p.WaitForExit(timeoutMs);
            }
            catch
            {
                return true;
            }
        }

        public static void KillByPid(int pid)
        {
            if (pid <= 0) return;
            try
            {
                var p = Process.GetProcessById(pid);
                if (!p.HasExited)
                    p.Kill();
            }
            catch
            {
                // ignore
            }
        }
    }
}
