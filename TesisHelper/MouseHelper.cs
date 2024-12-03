using System.Runtime.InteropServices;

namespace TesisHelper
{
    internal static class MouseHelper
    {
        [DllImport("user32.dll")]
        public static extern bool SetCursorPos(int X, int Y);
        [DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
        [DllImport("user32.dll")]
        public static extern bool GetCursorPos(out CursorPos lpPoint);
        public struct CursorPos { public int X; public int Y; }

        public static void CLick()
        {
            const int LMBDown = 0x02;
            const int LMBUp = 0x04;

            if (GetCursorPos(out CursorPos position))
            {
                SetCursorPos(position.X - 1, position.Y);
                Thread.Sleep(200);
                SetCursorPos(position.X + 1, position.Y);
                //SetCursorPos(position.X, position.Y);
                Thread.Sleep(1000);
                mouse_event(LMBDown, position.X, position.Y, 0, 0);
                mouse_event(LMBUp, position.X, position.Y, 0, 0);
                Thread.Sleep(1000);
            }
        }
    }
}
