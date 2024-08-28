using ClosedXML.Excel;

namespace TesisHelper
{
    internal static partial class ExcelTools
    {
        private static XLWorkbook? _workbook;

        public static XLWorkbook? CargarArchivoDeTrabajo(ExcelSettings settings)
        {
            string nombreDelArchivo = ObtenerNombreDelArchivo(settings, out bool existe);

            try
            {
                XLWorkbook libroExcel;
                if (existe)
                {
                    libroExcel = new XLWorkbook(nombreDelArchivo, new LoadOptions { RecalculateAllFormulas = true });
                    CrearHojasDeTrabajoSiNoExisten(libroExcel, settings);
                    libroExcel.Grabar();
                }
                else
                {
                    libroExcel = new XLWorkbook();
                    CrearHojasDeTrabajoSiNoExisten(libroExcel, settings);
                    libroExcel.Grabar(nombreDelArchivo);
                }

                return libroExcel;
            }
            catch
            {
                MessageBox.Show("¡Cierra el archivo!", "Archivo de Excel está abierto", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;
            }
        }

        private static string ObtenerNombreDelArchivo(ExcelSettings settings, out bool fileExists)
        {
            if (string.IsNullOrWhiteSpace(settings.RutaDelArchivo) || string.IsNullOrWhiteSpace(settings.NombreDelArchivo))
                throw new Exception("El nombre o ruta del archivo no ha sido configurada correctamente.");
            string nombreDelArchivo = Path.Combine(settings.RutaDelArchivo, settings.NombreDelArchivo);
            FileInfo finfo = new FileInfo(nombreDelArchivo);
            if (!(finfo.Extension == ".xls" || finfo.Extension == ".xlsx" || finfo.Extension == ".xlt" || finfo.Extension == ".xlsm" || finfo.Extension == ".csv"))
                throw new Exception("El nombre del archivo no es válido. Tiene que ser un libro de Excel.");
            fileExists = finfo.Exists;
            return nombreDelArchivo;
        }

        private static void CrearHojasDeTrabajoSiNoExisten(XLWorkbook libroExcel, ExcelSettings settings)
        {
            foreach (var tabla in settings.Tablas.OrderBy(t => t.Value.Posicion))
            {
                if (!libroExcel.Worksheets.Any(x => x.Name.Equals(tabla.Key)))
                {
                    libroExcel.Worksheets.Add(tabla.Key, tabla.Value.Posicion.Value);
                }
            }
        }

        public static void Grabar(this IXLWorkbook libroExcel, string? nombreDelArchivo = null)
        {
            try
            {
                if (nombreDelArchivo != null)
                    libroExcel.SaveAs(nombreDelArchivo);
                else
                    libroExcel.Save();
            }
            catch
            {
                MessageBox.Show("¡Cierra el archivo!", "Archivo de Excel está abierto", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        public static void CrearCabeceraEnTablaSiHojasEstanVacias(this IXLWorkbook libroExcel, ExcelSettings configuracionLibroExcel)
        {
            if (libroExcel == null) return;

            foreach (var tabla in configuracionLibroExcel.Tablas.OrderBy(t => t.Value.Posicion))
                libroExcel?.Worksheet(tabla.Key).CrearCabeceraEnTablaSiHojaEstaVacia(tabla.Value, configuracionLibroExcel);
        }

        public static IXLWorksheet? CargarArchivoDeTrabajo(string fileName)
        {
            try
            {
                _workbook = _workbook ?? new XLWorkbook(fileName, new LoadOptions { RecalculateAllFormulas = true });
                return _workbook.Worksheet(1);
            }
            catch
            {
                MessageBox.Show("¡Cierra el archivo!", "Archivo de Excel está abierto", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return null;
            }
        }

        public static void Grabar()
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

