namespace TesisHelper
{
    internal static class ClipboardHelper
    {
        public static string? GetText()
        {
            string? response = null;
            Exception? threadEx = null;
            Thread staThread = new Thread(
                delegate ()
                {
                    try
                    {
                        response = Clipboard.GetText();
                    }

                    catch (Exception ex)
                    {
                        threadEx = ex;
                    }
                });
            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join();
            return response;
        }

        public static void SetText(string text)
        {
            Exception? threadEx = null;
            Thread staThread = new Thread(
                delegate ()
                {
                    try
                    {
                        Clipboard.SetText(text);
                    }

                    catch (Exception ex)
                    {
                        threadEx = ex;
                    }
                });
            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join();
        }
    }
}
