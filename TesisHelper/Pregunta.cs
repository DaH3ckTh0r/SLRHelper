using ClosedXML.Excel;
using TesisHelper;

internal abstract class Pregunta
{
    public Pregunta(int indice, bool evaluar, TipoPregunta tipoPregunta)
    {
        Indice = indice;
        Evaluar = evaluar;
        NumeroFilaEnTabla = Settings.Constants.NUMERO_FILA_PREGUNTAS;
        TipoPregunta = tipoPregunta;
    }

    public int Indice { get; private set; }
    public bool Evaluar { get; private set; }
    public int NumeroFilaEnTabla { get; private set; }
    public int NumeroColumnaEnTabla { get; protected set; }
    public TipoPregunta TipoPregunta { get; private set; }
};

internal sealed class PreguntaInvestigacion : Pregunta
{
    public PreguntaInvestigacion(int indice, TipoPregunta tipoPregunta, bool evaluar = true)
        : base(indice, evaluar, tipoPregunta)
    {
        NumeroColumnaEnTabla = Settings.Columnas.NumeroUltimaColumnaAntesDeLasPreguntas() + indice;
    }
}

internal sealed class PreguntaInclusion : Pregunta
{
    public PreguntaInclusion(int indice, TipoPregunta tipoPregunta, bool evaluar = true)
        : base(indice, evaluar, tipoPregunta)
    {
        NumeroColumnaEnTabla = Settings.Columnas.NumeroUltimaColumnaAntesDeLasPreguntas() +
            (Settings.PreguntasDeInvestigacion?.Length ?? 0) + indice;
    }
}

internal sealed class PreguntaExclusion : Pregunta
{
    public PreguntaExclusion(int indice, TipoPregunta tipoPregunta, bool evaluar = true)
        : base(indice, evaluar, tipoPregunta)
    {
        NumeroColumnaEnTabla = Settings.Columnas.NumeroUltimaColumnaAntesDeLasPreguntas() +
            (Settings.PreguntasDeInvestigacion?.Length ?? 0) + (Settings.PreguntasDeInclusion?.Length ?? 0) + indice;
    }
}

internal enum TipoPregunta
{
    EvaluaSegunAbstract,
    EvaluaSegunPdf,
    EvaluaSiArchivoEstaDisponible
}

internal static class Extensions
{
    private static Pregunta? ObtenerPregunta(this Pregunta[] preguntas, int indice)
    {
        return preguntas.FirstOrDefault(x => x.Indice == indice);
    }

    public static void Evaluar(this Pregunta[] preguntas, IXLWorksheet worksheet, bool soloReprocesaErrores = false, int[]? idsDeLasPreguntasPorEvaluar = null)
    {
        const int COLUMNA_CE5 = 0;
        const int COLUMNA_CE1 = 4;

        for (var i = 0; i < preguntas.Length; i++)
        {
            Pregunta? pregunta = preguntas.ObtenerPregunta(i + 1);
            if (!(pregunta?.Evaluar ?? false)) continue;

            var numeroColumnaDeLaCeldaEvaluada = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string textoPorEvaluar = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaDeLaCeldaEvaluada).Value.ToString();
            switch (pregunta.TipoPregunta)
            {
                case TipoPregunta.EvaluaSiArchivoEstaDisponible:
                    EvaluationHelper.EvaluarSiArchivoEstaDisponible(worksheet, numeroColumnaDeLaCeldaEvaluada,
                        (rowNumber, columnNumber) =>
                        {
                            string criterioDeExclusionPrevioRequerido = Settings.Columnas.ResultadosCriterioExclusion[COLUMNA_CE5];
                            var numeroDeLaColumnaDeLaCondicionPorEvaluar = GetNumeroDeLaColumnaDeLaExclusionionPorEvaluar(criterioDeExclusionPrevioRequerido);
                            string valorActual = worksheet.Cell(rowNumber, numeroDeLaColumnaDeLaCondicionPorEvaluar).Value.ToString();
                            return valorActual.Equals(Settings.Constants.SI) || valorActual.Equals(Settings.Constants.NO_DEFINIDO);
                        }, idsDeLasPreguntasPorEvaluar: idsDeLasPreguntasPorEvaluar);
                    break;
                case TipoPregunta.EvaluaSegunPdf:
                    var numeroDeLaColumnaDeLaCondicionPorEvaluar = GetNumeroDeLaColumnaDeLaCondicionionPorEvaluar(pregunta, numeroColumnaDeLaCeldaEvaluada);
                    EvaluationHelper.EvaluarPreguntaSegunDocumento(worksheet, textoPorEvaluar, numeroDeLaColumnaDeLaCondicionPorEvaluar,
                        numeroColumnaDeLaCeldaEvaluada, (rowNumber, columnNumber) =>
                        {
                            return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals(Settings.Constants.NO);
                        }, soloReprocesaErrores, idsDeLasPreguntasPorEvaluar: idsDeLasPreguntasPorEvaluar);
                    break;
                case TipoPregunta.EvaluaSegunAbstract:
                    EvaluationHelper.EvaluarPreguntaSegunAbstract(worksheet, textoPorEvaluar, numeroColumnaDeLaCeldaEvaluada,
                        (rowNumber, columnNumber) =>
                        {
                            string criterioDeExclusionPrevioRequerido = Settings.Columnas.ResultadosCriterioExclusion[COLUMNA_CE1];
                            var numeroDeLaColumnaDeLaCondicionPorEvaluar = GetNumeroDeLaColumnaDeLaExclusionionPorEvaluar(criterioDeExclusionPrevioRequerido);
                            string value = worksheet.Cell(rowNumber, columnNumber).Value.ToString();
                            return (!string.IsNullOrWhiteSpace(value) && !value.Equals(Settings.Constants.NO_DEFINIDO));
                        }, idsDeLasPreguntasPorEvaluar: idsDeLasPreguntasPorEvaluar);
                    break;
                default: break;
            }
        }
    }

    private static int GetNumeroDeLaColumnaDeLaCondicionionPorEvaluar(Pregunta pregunta, int numeroColumnaDeLaCeldaEvaluada)
    {
        return pregunta.GetType().Name switch
        {
            nameof(PreguntaExclusion) => numeroColumnaDeLaCeldaEvaluada - pregunta.Indice + 1,
            nameof(PreguntaInclusion) => numeroColumnaDeLaCeldaEvaluada + (Settings.PreguntasDeExclusion?.Length ?? 0) - pregunta.Indice + 1,
            nameof(PreguntaInvestigacion) => numeroColumnaDeLaCeldaEvaluada + (Settings.PreguntasDeInclusion?.Length ?? 0) + (Settings.PreguntasDeExclusion?.Length ?? 0) - pregunta.Indice + 2,
            _ => 0
        };
    }

    private static int GetNumeroDeLaColumnaDeLaExclusionionPorEvaluar(string tituloDeLaColumnaConElCriterioDeExclusionRequerido)
    {
        return Settings.Columnas.NumeroUltimaColumnaAntesDeLasPreguntas() +
                                        (Settings.PreguntasDeInvestigacion?.Length ?? 0) + (Settings.PreguntasDeInclusion?.Length ?? 0) +
                                        (Settings.PreguntasDeExclusion?.Length ?? 0) + Settings.Columnas.ResultadosCriterioExclusion.NumeroColumna(tituloDeLaColumnaConElCriterioDeExclusionRequerido);
    }
    /*
    #region Preguntas de Exclusion
    public static void EvaluarPreguntaExclusion1(this Pregunta[] preguntas, IXLWorksheet worksheet, int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(1);
        if (pregunta?.Evaluar ?? false)
        {
            var numeroColumnaEnTabla = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarSiArchivoEstaDisponible(worksheet, numeroColumnaEnTabla,
                (rowNumber, columnNumber) =>
                {
                    return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals("SI");
                }, idsDeLasPreguntasPorEvaluar: idsToReview);
        }
    }
    public static void EvaluarPreguntaExclusion2(this Pregunta[] preguntas, IXLWorksheet worksheet, bool soloReprocesaErrores = false,
        int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(2);
        if (pregunta?.Evaluar ?? false)
        {
            var numeroColumnaEnTabla = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarPreguntaSegunDocumento(worksheet, texto, numeroColumnaEnTabla - 1,
                numeroColumnaEnTabla, (rowNumber, columnNumber) =>
                {
                    return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals("NO");
                }, soloReprocesaErrores, idsDeLasPreguntasPorEvaluar: idsToReview);
        }
    }
    public static void EvaluarPreguntaExclusion3(this Pregunta[] preguntas, IXLWorksheet worksheet, bool soloReprocesaErrores = false,
        int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(3);
        if (pregunta?.Evaluar ?? false)
        {
            var numeroColumnaEnTabla = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarPreguntaSegunDocumento(worksheet, texto, numeroColumnaEnTabla - 2,
                numeroColumnaEnTabla, (rowNumber, columnNumber) =>
                {
                    return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals("NO");
                }, soloReprocesaErrores, idsDeLasPreguntasPorEvaluar: idsToReview);
        }
    }
    public static void EvaluarPreguntaExclusion4(this Pregunta[] preguntas, IXLWorksheet worksheet, bool soloReprocesaErrores = false,
        int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(4);
        if (pregunta?.Evaluar ?? false)
        {
            var numeroColumnaEnTabla = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarPreguntaSegunDocumento(worksheet, texto, numeroColumnaEnTabla - 3,
                numeroColumnaEnTabla, (rowNumber, columnNumber) =>
                {
                    return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals("NO");
                }, soloReprocesaErrores, idsDeLasPreguntasPorEvaluar: idsToReview);
        }
    }
    public static void EvaluarPreguntaExclusion5(this Pregunta[] preguntas, IXLWorksheet worksheet,
        int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(5);
        if (pregunta?.Evaluar ?? false)
        {
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, pregunta.NumeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarPreguntaSegunAbstract(worksheet, texto, pregunta.NumeroColumnaEnTabla,
                (rowNumber, columnNumber) =>
                {
                    string value = worksheet.Cell(rowNumber, columnNumber).Value.ToString();
                    return (!string.IsNullOrWhiteSpace(value) && !value.Equals("¿?"));
                }, idsDeLasPreguntasPorEvaluar: idsToReview);
        }
    }
    #endregion

    #region Preguntas de Inclusion
    public static void EvaluarPreguntaInclusion1(this Pregunta[] preguntas, IXLWorksheet worksheet, bool soloReprocesaErrores = false, int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(1);
        if (pregunta?.Evaluar ?? false)
        {
            var numeroColumnaEnTabla = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarPreguntaSegunDocumento(worksheet, texto, numeroColumnaEnTabla + 5,
                numeroColumnaEnTabla, (rowNumber, columnNumber) =>
                {
                    return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals("NO");
                }, soloReprocesaErrores, waitingTime: 40000, idsDeLasPreguntasPorEvaluar: idsToReview);
        }
    }
    public static void EvaluarPreguntaInclusion2(this Pregunta[] preguntas, IXLWorksheet worksheet, bool soloReprocesaErrores = false, int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(2);
        if (pregunta?.Evaluar ?? false)
        {
            var numeroColumnaEnTabla = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarPreguntaSegunDocumento(worksheet, texto, numeroColumnaEnTabla + 4,
                numeroColumnaEnTabla, (rowNumber, columnNumber) =>
                {
                    return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals("NO");
                }, soloReprocesaErrores, idsDeLasPreguntasPorEvaluar: idsToReview);
        }
    }
    public static void EvaluarPreguntaInclusion3(this Pregunta[] preguntas, IXLWorksheet worksheet, bool soloReprocesaErrores = false, int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(3);
        if (pregunta?.Evaluar ?? false)
        {
            var numeroColumnaEnTabla = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarPreguntaSegunDocumento(worksheet, texto, numeroColumnaEnTabla + 3,
                numeroColumnaEnTabla, (rowNumber, columnNumber) =>
                {
                    return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals("NO");
                }, soloReprocesaErrores, idsDeLasPreguntasPorEvaluar: idsToReview);
        }
    }
    public static void EvaluarPreguntaInclusion4(this Pregunta[] preguntas, IXLWorksheet worksheet, bool soloReprocesaErrores = false, int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(4);
        if (pregunta?.Evaluar ?? false)
        {
            var numeroColumnaEnTabla = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarPreguntaSegunDocumento(worksheet, texto, numeroColumnaEnTabla + 2,
                numeroColumnaEnTabla, (rowNumber, columnNumber) =>
                {
                    return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals("NO");
                }, soloReprocesaErrores, idsDeLasPreguntasPorEvaluar: idsToReview);
        }
    }
    public static void EvaluarPreguntaInclusion5(this Pregunta[] preguntas, IXLWorksheet worksheet, bool soloReprocesaErrores = false, int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(5);
        if (pregunta?.Evaluar ?? false)
        {
            var numeroColumnaEnTabla = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarPreguntaSegunDocumento(worksheet, texto, numeroColumnaEnTabla + 1,
                numeroColumnaEnTabla, (rowNumber, columnNumber) =>
                {
                    return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals("NO");
                }, soloReprocesaErrores, idsDeLasPreguntasPorEvaluar: idsToReview);
        }
    }
    #endregion

    #region Preguntas de Investigacion
    public static void EvaluarPreguntaInvestigacion1(this Pregunta[] preguntas, IXLWorksheet worksheet, bool soloReprocesaErrores = false, int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(1);
        if (pregunta?.Evaluar ?? false)
        {
            var numeroColumnaEnTabla = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarPreguntaSegunDocumento(worksheet, texto, numeroColumnaEnTabla + 28,
                numeroColumnaEnTabla, (rowNumber, columnNumber) =>
                {
                    return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals("Aprobado");
                }, soloReprocesaErrores, waitingTime: 40000, idsToReview);
        }
    }
    public static void EvaluarPreguntaInvestigacion2(this Pregunta[] preguntas, IXLWorksheet worksheet, bool soloReprocesaErrores = false, int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(2);
        if (pregunta?.Evaluar ?? false)
        {
            var numeroColumnaEnTabla = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarPreguntaSegunDocumento(worksheet, texto, numeroColumnaEnTabla + 27,
                numeroColumnaEnTabla, (rowNumber, columnNumber) =>
                {
                    return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals("Aprobado");
                }, soloReprocesaErrores, idsDeLasPreguntasPorEvaluar: idsToReview);
        }
    }
    public static void EvaluarPreguntaInvestigacion3(this Pregunta[] preguntas, IXLWorksheet worksheet, bool soloReprocesaErrores = false, int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(3);
        if (pregunta?.Evaluar ?? false)
        {
            var numeroColumnaEnTabla = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarPreguntaSegunDocumento(worksheet, texto, numeroColumnaEnTabla + 26,
                numeroColumnaEnTabla, (rowNumber, columnNumber) =>
                {
                    return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals("Aprobado");
                }, soloReprocesaErrores, idsDeLasPreguntasPorEvaluar: idsToReview);
        }
    }
    public static void EvaluarPreguntaInvestigacion4(this Pregunta[] preguntas, IXLWorksheet worksheet, bool soloReprocesaErrores = false,
        int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(4);
        if (pregunta?.Evaluar ?? false)
        {
            var numeroColumnaEnTabla = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarPreguntaSegunDocumento(worksheet, texto, numeroColumnaEnTabla + 25,
                numeroColumnaEnTabla, (rowNumber, columnNumber) =>
                {
                    return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals("Aprobado");
                }, soloReprocesaErrores, idsDeLasPreguntasPorEvaluar: idsToReview);
        }
    }
    public static void EvaluarPreguntaInvestigacion5(this Pregunta[] preguntas, IXLWorksheet worksheet, bool soloReprocesaErrores = false, int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(5);
        if (pregunta?.Evaluar ?? false)
        {
            var numeroColumnaEnTabla = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarPreguntaSegunDocumento(worksheet, texto, numeroColumnaEnTabla + 24,
                numeroColumnaEnTabla, (rowNumber, columnNumber) =>
                {
                    return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals("Aprobado");
                }, soloReprocesaErrores, idsDeLasPreguntasPorEvaluar: idsToReview);
        }
    }
    public static void EvaluarPreguntaInvestigacion6(this Pregunta[] preguntas, IXLWorksheet worksheet, bool soloReprocesaErrores = false,
        int[]? idsToReview = null)
    {
        Pregunta? pregunta = preguntas.ObtenerPregunta(6);
        if (pregunta?.Evaluar ?? false)
        {
            var numeroColumnaEnTabla = worksheet.RangeUsed().FirstColumnUsed().ColumnNumber() + pregunta.NumeroColumnaEnTabla - 2;
            string texto = worksheet.Cell(pregunta.NumeroFilaEnTabla, numeroColumnaEnTabla).Value.ToString();
            EvaluationHelper.EvaluarPreguntaSegunDocumento(worksheet, texto, numeroColumnaEnTabla + 23,
                numeroColumnaEnTabla, (rowNumber, columnNumber) =>
                {
                    return worksheet.Cell(rowNumber, columnNumber).Value.ToString().Equals("Aprobado");
                }, soloReprocesaErrores, idsDeLasPreguntasPorEvaluar: idsToReview);
        }
    }
    #endregion
    */
}