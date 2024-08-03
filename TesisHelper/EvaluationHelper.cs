using ClosedXML.Excel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace TesisHelper
{
    internal static class EvaluationHelper
    {
        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);

        public static void EvaluarSiArchivoEstaDisponible(IXLWorksheet worksheet, int numeroDelaColumnaDeLaPreguntaEvaluada,
            Func<int, int, bool>? condicionPorEvaluarParaEstaPregunta = null, int[]? idsDeLasPreguntasPorEvaluar = null)
        {
            if (worksheet == null) return;

            int numeroDeColumnaConIdRequerido = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + Settings.Columnas.Informacion.NumeroColumna(Settings.Constants.COLUMNA_ID) - 1;
            int numeroDeColumnaConNombreDelArchivoRequerido = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + Settings.Columnas.Informacion.NumeroColumna(Settings.Constants.COLUMNA_ARCHIVO) - 1;

            var filasPorEvaluar = worksheet.RangeUsed().RowsUsed().Skip(Settings.Constants.FILAS_SIN_USAR);
            foreach (var fila in filasPorEvaluar)
            {
                var numeroDeFilaActual = fila.RowNumber();
                if (SoloSeDebenEvaluarIdsRequeridos(worksheet, idsDeLasPreguntasPorEvaluar, numeroDeColumnaConIdRequerido, numeroDeFilaActual))
                    continue;

                if (condicionPorEvaluarParaEstaPregunta?.Invoke(numeroDeFilaActual, numeroDelaColumnaDeLaPreguntaEvaluada) ?? false)
                {
                    bool archivoExiste = !worksheet.Cell(numeroDeFilaActual, numeroDeColumnaConNombreDelArchivoRequerido).Value.ToString().Equals(Settings.Constants.ARCHIVO_NO_ENCONTRADO);
                    worksheet.Cell(numeroDeFilaActual, numeroDelaColumnaDeLaPreguntaEvaluada).Value = !archivoExiste ? Settings.Constants.SI : Settings.Constants.NO;
                }
                else
                {
                    worksheet.Cell(numeroDeFilaActual, numeroDelaColumnaDeLaPreguntaEvaluada).Value = string.Empty;
                }
                ExcelForPapersEvaluation.Save();
            }
        }

        public static void EvaluarPreguntaSegunAbstract(IXLWorksheet worksheet, string questionToEvaluate, int numeroDelaColumnaDeLaPreguntaEvaluada,
            Func<int, int, bool>? conditionToEvaluate = null, int[]? idsDeLasPreguntasPorEvaluar = null)
        {
            if (worksheet == null) return;

            int numeroDeColumnaConIdRequerido = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + Settings.Columnas.Informacion.NumeroColumna(Settings.Constants.COLUMNA_ID) - 1;
            int numeroDeColumnaDeAbstract = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + Settings.Columnas.Informacion.NumeroColumna(Settings.Constants.COLUMNA_ABSTRACT) - 1;

            var filasPorEvaluar = worksheet.RangeUsed().RowsUsed().Skip(Settings.Constants.FILAS_SIN_USAR);
            foreach (var fila in filasPorEvaluar)
            {
                var numeroDeFilaActual = fila.RowNumber();
                if (SoloSeDebenEvaluarIdsRequeridos(worksheet, idsDeLasPreguntasPorEvaluar, numeroDeColumnaConIdRequerido, numeroDeFilaActual))
                    continue;

                if (conditionToEvaluate?.Invoke(numeroDeFilaActual, numeroDelaColumnaDeLaPreguntaEvaluada) ?? false) continue;
                var abstractText = Regex.Replace(worksheet.Cell(numeroDeFilaActual, numeroDeColumnaDeAbstract).Value.ToString(), "[+^%~()]", "{$0}");
                Process browser = CopilotHelper.LoadBrowser();
                var copilotResponse = CopilotHelper.EvaluateQuestion(browser, $"Segun el siguiente texto: \"{abstractText}\" {questionToEvaluate}. No incluyas definiciones, ni agregues comentarios adicionales ni referencias externas.", waitingTime: 25000);
                KillEdgeProcess(browser);
                if (copilotResponse?.Equals(questionToEvaluate) ?? true) continue;
                SetResponse(worksheet, numeroDelaColumnaDeLaPreguntaEvaluada, numeroDeFilaActual, copilotResponse);
                ExcelForPapersEvaluation.Save();
            }
        }

        private static void SetResponse(IXLWorksheet worksheet, int numeroDelaColumnaDeLaPreguntaEvaluada, int numeroDeFilaActual, string? copilotResponse)
        {
            worksheet.Cell(numeroDeFilaActual, numeroDelaColumnaDeLaPreguntaEvaluada).Value = copilotResponse;
            worksheet.Cell(numeroDeFilaActual, numeroDelaColumnaDeLaPreguntaEvaluada).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            worksheet.Cell(numeroDeFilaActual, numeroDelaColumnaDeLaPreguntaEvaluada).Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
            worksheet.Cell(numeroDeFilaActual, numeroDelaColumnaDeLaPreguntaEvaluada).Style.Alignment.WrapText = true;
        }

        public static void EvaluarPreguntaSegunDocumento(IXLWorksheet worksheet, string questionToEvaluate, int isAvailableColumnNumber,
            int numeroDelaColumnaDeLaPreguntaEvaluada, Func<int, int, bool>? conditionToEvaluate = null, bool soloReprocesaErrores = false, int waitingTime = 30000,
            int[]? idsDeLasPreguntasPorEvaluar = null)
        {
            if (worksheet == null) return;

            int numeroDeColumnaConRutaDelArchivoRequerido = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + Settings.Columnas.Informacion.NumeroColumna(Settings.Constants.COLUMNA_RUTA) - 1;
            int numeroDeColumnaConIdRequerido = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + Settings.Columnas.Informacion.NumeroColumna(Settings.Constants.COLUMNA_ID) - 1;

            var filasPorEvaluar = worksheet.RangeUsed().RowsUsed().Skip(Settings.Constants.FILAS_SIN_USAR);
            foreach (var fila in filasPorEvaluar)
            {
                var numeroDeFilaActual = fila.RowNumber();
                if (SoloSeDebenEvaluarIdsRequeridos(worksheet, idsDeLasPreguntasPorEvaluar, numeroDeColumnaConIdRequerido, numeroDeFilaActual))
                    continue;

                string fileName = worksheet.Cell(numeroDeFilaActual, numeroDeColumnaConRutaDelArchivoRequerido).Value.ToString();
                if (string.IsNullOrEmpty(fileName)) continue;

                string respuestaActual = worksheet.Cell(numeroDeFilaActual, numeroDelaColumnaDeLaPreguntaEvaluada).Value.ToString();
                if (conditionToEvaluate?.Invoke(numeroDeFilaActual, isAvailableColumnNumber) ?? false)
                {
                    if (soloReprocesaErrores)
                    {
                        string[] errores = ["Parece que el archivo que intentas abrir no se encuentra disponible o ha sido movido.", "Parece que el archivo que intentas abrir no se encuentra disponible o ha sido movido, editado o eliminado.", "@This page:"];
                        if (!errores.Any(respuestaActual.StartsWith))
                        {
                            continue;
                        }
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(respuestaActual))
                        {
                            continue;
                        }
                    }

                    UriBuilder builder = new UriBuilder(fileName);
                    Uri uri = builder.Uri;

                    Process browser = CopilotHelper.LoadBrowser(uri.AbsoluteUri);
                    var copilotResponse = CopilotHelper.EvaluateQuestion(browser, questionToEvaluate, usePdf: true, waitingTime: waitingTime);
                    KillEdgeProcess(browser);
                    if (copilotResponse?.Equals(questionToEvaluate) ?? true) continue;
                    SetResponse(worksheet, numeroDelaColumnaDeLaPreguntaEvaluada, numeroDeFilaActual, copilotResponse);
                }
                else
                {
                    if (soloReprocesaErrores)
                    {
                        continue;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(respuestaActual))
                        {
                            continue;
                        }
                    }
                    SetResponse(worksheet, numeroDelaColumnaDeLaPreguntaEvaluada, numeroDeFilaActual, Settings.Constants.DOCUMENTO_NO_ENCONTRADO);
                }
                ExcelForPapersEvaluation.Save();
            }
        }

        private static bool SoloSeDebenEvaluarIdsRequeridos(IXLWorksheet worksheet, int[]? idsDeLasPreguntasPorEvaluar, int numeroDeColumnaConIdRequerido, int numeroDeFilaActual)
        {
            if (!(idsDeLasPreguntasPorEvaluar?.Any() ?? false)) return false;

            if (!int.TryParse(worksheet.Cell(numeroDeFilaActual, numeroDeColumnaConIdRequerido).Value.ToString(), out int itemNumber))
                return true;
            if (!idsDeLasPreguntasPorEvaluar.Contains(itemNumber)) return true;

            return false;
        }

        private static void KillEdgeProcess(Process browser)
        {
            Process[] Edge = Process.GetProcessesByName("msedge");
            foreach (Process Item in Edge)
            {
                try
                {
                    Item.Kill();
                    Item.WaitForExit(30000);
                }
                catch (Exception)
                {

                }
            }
            browser.Dispose();
        }
    }
}
