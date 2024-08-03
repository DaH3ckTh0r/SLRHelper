using ClosedXML.Excel;

namespace TesisHelper
{
    internal static class ExcelForPapersEvaluation
    {
        static XLWorkbook? _workbook;

        public static IXLWorksheet? LoadEvaluationTable(string fileName)
        {
            try
            {
                _workbook = _workbook ?? new XLWorkbook(fileName);
                return _workbook.Worksheet(1);
            }
            catch
            {
                MessageBox.Show("¡Cierra el archivo!", "Archivo de Excel está abierto", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;
            }
        }

        public static void Save()
        {
            try
            {
                _workbook?.Save();
            }
            catch
            {
                MessageBox.Show("¡Cierra el archivo!", "Archivo de Excel está abierto", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}

