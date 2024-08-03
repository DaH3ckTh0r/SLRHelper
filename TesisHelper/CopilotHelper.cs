using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace TesisHelper
{
    internal static class CopilotHelper
    {
        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);

        public static Process LoadBrowser(string? uri = null, int waitingTime = 15000)
        {
            Console.WriteLine("Starting Microsoft Edge and enabling Copilot...");
            //Process.Start(new ProcessStartInfo("cmd", $"/c start microsoft-edge://?ux=copilot&tcp=1&source=taskbar") { CreateNoWindow = true });
            Process p = uri == null ? Process.Start(Settings.Constants.EDGE_DIRECTORY) : Process.Start(Settings.Constants.EDGE_DIRECTORY, $"--new-window {uri}");
            Thread.Sleep(3000);

            if (p != null)
            {
                IntPtr h = p.MainWindowHandle;
                SetForegroundWindow(h);
                SendKeys.SendWait("^+.");
            }

            Thread.Sleep(waitingTime);
            return p;
        }

        public static string? EvaluateQuestion(Process p, string question, bool usePdf = false,
            int waitingTime = 60000)
        {
            List<string> KeysForAbstract = new List<string> { Settings.Keys.SHIFT_TAB, Settings.Keys.SHIFT_TAB, Settings.Keys.SHIFT_TAB, Settings.Keys.SHIFT_TAB, Settings.Keys.SHIFT_TAB, Settings.Keys.SHIFT_TAB, Settings.Keys.SHIFT_TAB, Settings.Keys.ENTER };
            List<string> KeysForPdf = new List<string> { Settings.Keys.SHIFT_TAB, Settings.Keys.SHIFT_TAB, Settings.Keys.SHIFT_TAB, Settings.Keys.SHIFT_TAB, Settings.Keys.SHIFT_TAB, Settings.Keys.SHIFT_TAB, Settings.Keys.SHIFT_TAB, Settings.Keys.SHIFT_TAB, Settings.Keys.ENTER };

            string questionForCopilot = $"{(usePdf ? "@This page: " : "")}{question}{Environment.NewLine}";
            Console.WriteLine($"Sending question to Copilot: {questionForCopilot}");
            ClipboardHelper.SetText(questionForCopilot);
            SendKeys.SendWait("^V");
            SendKeys.SendWait("{ENTER}");
            Thread.Sleep(waitingTime);
            List<string> keys = !usePdf ? KeysForAbstract : KeysForPdf;
            int times = 3;
            do
            {
                foreach (var key in keys)
                {
                    SendKeys.SendWait(key);
                    Thread.Sleep(1000);
                }
                var copilotResponse = ClipboardHelper.GetText() ?? questionForCopilot;
                if (copilotResponse.Equals(questionForCopilot) || string.IsNullOrEmpty(copilotResponse))
                {
                    for (var i = keys.Count - 1; i >= 0; i--)
                    {
                        string key = GetOppositeKey(keys[i]);
                        SendKeys.SendWait(key);
                        Thread.Sleep(1000);
                    }
                    keys.Insert(0, Settings.Keys.SHIFT_TAB);
                }
                else
                {
                    return CleanUpResponse(copilotResponse);
                }
                times--;
            } while (times > 0);
            return string.Empty;
        }

        private static string GetOppositeKey(string key)
        {
            switch (key)
            {
                case Settings.Keys.ENTER: return Settings.Keys.ESC;
                case Settings.Keys.ESC: return Settings.Keys.ENTER;
                case Settings.Keys.SHIFT_TAB: return Settings.Keys.TAB;
                case Settings.Keys.TAB: return Settings.Keys.SHIFT_TAB;
            }
            return string.Empty;
        }

        private static string CleanUpResponse(string response)
        {
            string[] textoPorRemover = ["Si estás", "Si tienes", "Si deseas", "Si necesitas", "En resumen", "Source", "¿"];
            response = response.Replace("**", "").Replace("¹", "").Replace("²", "").Replace("³", "").Replace("⁴", "").Replace("😊", "");
            string regex = @"(\[.*?\])";
            response = Regex.Replace(response, regex, "");
            foreach (var texto in textoPorRemover)
            {
                if (response.Contains(texto, StringComparison.OrdinalIgnoreCase))
                {
                    response = response.Substring(0, response.IndexOf(texto, StringComparison.OrdinalIgnoreCase));
                }
            }
            return response.Trim();
        }
    }
}
