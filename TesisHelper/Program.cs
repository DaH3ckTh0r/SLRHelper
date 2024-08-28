using ClosedXML.Excel;
using TesisHelper;

ExcelSettings configuracionLibroExcel = ExcelSettingsHelper.CargarConfiguraciones();
IXLWorkbook? libroExcel = ExcelTools.CargarArchivoDeTrabajo(configuracionLibroExcel);
libroExcel?.CrearCabeceraEnTablaSiHojasEstanVacias(configuracionLibroExcel);

FileInfo finfo = new FileInfo(@$"{AppSettings.Constants.MAIN_PATH}\{AppSettings.Constants.EXCEL_FILE_NAME}");
if (!(finfo.Extension == ".xls" || finfo.Extension == ".xlsx" || finfo.Extension == ".xlt" || finfo.Extension == ".xlsm" || finfo.Extension == ".csv"))
    throw new Exception("Archivo no valido");

IXLWorksheet? worksheet = ExcelTools.CargarArchivoDeTrabajo(finfo.FullName);

if (worksheet != null)
{
    AppSettings.PreguntasDeInvestigacion =
        Enumerable.Range(1, 6).Select(i =>
        {
            TipoPregunta tipoPregunta = TipoPregunta.EvaluaSegunPdf;
            return new PreguntaInvestigacion(i, tipoPregunta);
        }).ToArray();

    AppSettings.PreguntasDeInclusion =
        Enumerable.Range(1, 5).Select(i =>
        {
            TipoPregunta tipoPregunta = TipoPregunta.EvaluaSegunPdf;
            return new PreguntaInclusion(i, tipoPregunta);
        }).ToArray();

    AppSettings.PreguntasDeExclusion =
        Enumerable.Range(1, 5).Select(i =>
        {
            TipoPregunta tipoPregunta = i switch
            {
                1 => TipoPregunta.EvaluaSiArchivoEstaDisponible,
                5 => TipoPregunta.EvaluaSegunAbstract,
                _ => TipoPregunta.EvaluaSegunPdf
            };
            return new PreguntaExclusion(i, tipoPregunta);
        }).ToArray();

    AppSettings.IdsAProcesar = [469, 472, 485];

    AppSettings.PreguntasDeExclusion.Evaluar(worksheet, idsDeLasPreguntasPorEvaluar: AppSettings.IdsAProcesar);
    AppSettings.PreguntasDeInclusion.Evaluar(worksheet, idsDeLasPreguntasPorEvaluar: AppSettings.IdsAProcesar);
    AppSettings.PreguntasDeInvestigacion.Evaluar(worksheet, idsDeLasPreguntasPorEvaluar: AppSettings.IdsAProcesar);
    MessageBox.Show("¡Terminé!", "Ya puedes revisar los archivos generados y el Excel", MessageBoxButtons.OK, MessageBoxIcon.Information);
}


/*
preguntasExclusion.EvaluarPreguntaExclusion5(worksheet, idsToReview: idsToProcess);
preguntasExclusion.EvaluarPreguntaExclusion1(worksheet, idsToReview: idsToProcess);
preguntasExclusion.EvaluarPreguntaExclusion2(worksheet, idsToReview: idsToProcess);
preguntasExclusion.EvaluarPreguntaExclusion3(worksheet, idsToReview: idsToProcess);
preguntasExclusion.EvaluarPreguntaExclusion4(worksheet, idsToReview: idsToProcess);


preguntasInclusion.EvaluarPreguntaInclusion1(worksheet, idsToReview: idsToProcess);
preguntasInclusion.EvaluarPreguntaInclusion2(worksheet, idsToReview: idsToProcess);
preguntasInclusion.EvaluarPreguntaInclusion3(worksheet, idsToReview: idsToProcess);
preguntasInclusion.EvaluarPreguntaInclusion4(worksheet, idsToReview: idsToProcess);
preguntasInclusion.EvaluarPreguntaInclusion5(worksheet, idsToReview: idsToProcess);


preguntasInvestigacion.EvaluarPreguntaInvestigacion1(worksheet, idsToReview: idsToProcess);
preguntasInvestigacion.EvaluarPreguntaInvestigacion2(worksheet, idsToReview: idsToProcess);
preguntasInvestigacion.EvaluarPreguntaInvestigacion3(worksheet, idsToReview: idsToProcess);
preguntasInvestigacion.EvaluarPreguntaInvestigacion4(worksheet, idsToReview: idsToProcess);
preguntasInvestigacion.EvaluarPreguntaInvestigacion5(worksheet, idsToReview: idsToProcess);
preguntasInvestigacion.EvaluarPreguntaInvestigacion6(worksheet, idsToReview: idsToProcess);
*/