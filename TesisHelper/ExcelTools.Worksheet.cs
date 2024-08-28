using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TesisHelper
{
    internal static partial class ExcelTools
    {
        public static void CrearCabeceraEnTablaSiHojaEstaVacia(this IXLWorksheet hojaExcel, TablaExcel tablaExcel, ExcelSettings configuracionLibroExcel)
        {
            if (hojaExcel.CellsUsed().Any()) return;

            int columnaInicioCabecera = 64 + tablaExcel.ColumnasNoIncluidas;
            int filaCabecera = 1 + tablaExcel.FilasNoIncluidas;
            int columnaActual = columnaInicioCabecera;
            foreach (var cabecera in tablaExcel.Columnas.OrderBy(t => t.Value.Posicion))
            {
                if (!cabecera.Value.Columnas?.Any() ?? true)
                {
                    PoblarCabeceraSinColumnas(hojaExcel, tablaExcel, configuracionLibroExcel, columnaInicioCabecera, filaCabecera, columnaActual, cabecera);
                    columnaActual++;
                }
                else
                {
                    PoblarCabeceraSinColumnas(hojaExcel, tablaExcel, configuracionLibroExcel, columnaInicioCabecera, filaCabecera, columnaActual, cabecera);
                    foreach (var columna in cabecera.Value.Columnas.OrderBy(t => t.Value.Posicion))
                    {
                        PoblarCabeceraConColumnas(hojaExcel, tablaExcel, configuracionLibroExcel, columnaInicioCabecera, filaCabecera + 1, ref columnaActual, cabecera, columna);
                    }
                }
            }
            hojaExcel.Workbook.Grabar();
        }

        private static void PoblarCabeceraConColumnas(IXLWorksheet hojaExcel, TablaExcel tablaExcel, ExcelSettings configuracionLibroExcel, int columnaInicioCabecera, int filaCabecera, ref int columnaActual, KeyValuePair<string, CabeceraTabla> cabecera, KeyValuePair<string, CabeceraTabla> columna)
        {
            if (!columna.Value.Columnas?.Any() ?? true)
            {
                PoblarCabeceraSinColumnas(hojaExcel, tablaExcel, configuracionLibroExcel, columnaInicioCabecera, filaCabecera, columnaActual, columna);
                columnaActual++;
            }
            else
            {
                PoblarCabeceraSinColumnas(hojaExcel, tablaExcel, configuracionLibroExcel, columnaInicioCabecera, filaCabecera, columnaActual, columna);
                foreach (var subcolumna in columna.Value.Columnas.OrderBy(t => t.Value.Posicion))
                {
                    PoblarCabeceraConColumnas(hojaExcel, tablaExcel, configuracionLibroExcel, columnaInicioCabecera, filaCabecera + 1, ref columnaActual, columna, subcolumna);
                }
            }
            //if (!configuracionLibroExcel.Estilos.TryGetValue(columna.Value.Estilo ?? "", out Estilo? estilo))
            //{
            //    if (!configuracionLibroExcel.Estilos.TryGetValue(cabecera.Value.Estilo ?? "", out estilo))
            //    {
            //        if (!configuracionLibroExcel.Estilos.TryGetValue(tablaExcel.Estilo ?? "", out estilo))
            //            estilo = configuracionLibroExcel.Estilos.First().Value;
            //    }
            //}
            //PoblarCabecera(hojaExcel, tablaExcel, estilo, columnaActual, columnaInicioCabecera, filaCabecera, columna.Value);
        }

        private static void PoblarCabeceraSinColumnas(IXLWorksheet hojaExcel, TablaExcel tablaExcel, ExcelSettings configuracionLibroExcel, int columnaInicioCabecera, int filaCabecera, int columnaActual, KeyValuePair<string, CabeceraTabla> cabecera)
        {
            if (!configuracionLibroExcel.Estilos.TryGetValue(cabecera.Value.Estilo ?? "", out Estilo? estilo))
            {
                if (!configuracionLibroExcel.Estilos.TryGetValue(tablaExcel.Estilo ?? "", out estilo))
                    estilo = configuracionLibroExcel.Estilos.First().Value;
            }
            PoblarCabecera(hojaExcel, tablaExcel, estilo, columnaActual, columnaInicioCabecera, filaCabecera, cabecera.Value);
        }

        private static void PoblarCabecera(IXLWorksheet hojaExcel, TablaExcel tablaExcel, Estilo estilo, int columnaActual, int columnaInicio, int filaCabecera, CabeceraTabla cabecera)
        {
            int cociente = (columnaActual - columnaInicio + 1) / 26;
            int residuo = (columnaActual - columnaInicio + 1) % 26;
            string columna = cociente > 0 ? $"{(char)(columnaInicio + cociente - 1)}{(char)(columnaInicio + residuo)}" : $"{(char)(columnaInicio + residuo)}";
            string celda = $"{columna}{filaCabecera}";
            hojaExcel.Cell(celda).Value = cabecera.Titulo;
            AjustarTamanoCelda(hojaExcel, cabecera, celda);
            IXLRange celdasUsadas = ObtenerCeldasUsadas(hojaExcel, tablaExcel, filaCabecera, cabecera, columna, celda);
            celdasUsadas.AplicaEstilo(estilo);
        }

        private static IXLRange ObtenerCeldasUsadas(IXLWorksheet hojaExcel, TablaExcel tablaExcel, int filaCabecera, CabeceraTabla cabecera, string columna, string celda)
        {
            if (cabecera.NumeroColumnasOcupadas > 1)
            {
                IXLCell celdaActual = hojaExcel.Cell(celda);
                int numeroColumnaFinal = celdaActual.WorksheetColumn().ColumnNumber() + cabecera.NumeroColumnasOcupadas - 1;
                int numeroFilaActual = celdaActual.WorksheetRow().RowNumber();
                IXLCell celdaFinal = hojaExcel.Cell(numeroFilaActual, numeroColumnaFinal);
                return hojaExcel.Range(celdaActual, celdaFinal).Merge();
            }
            if (tablaExcel.NumeroFilasOcupadas > 1)
            {
                int numeroFilasOcupadas = tablaExcel.NumeroFilasOcupadas - cabecera.NumeroFila;
                string celdaFinal = $"{columna}{filaCabecera + numeroFilasOcupadas}";
                return hojaExcel.Range(hojaExcel.Cell(celda), hojaExcel.Cell(celdaFinal)).Merge();
            }
            return hojaExcel.Range(celda);
        }

        private static void AjustarTamanoCelda(IXLWorksheet hojaExcel, CabeceraTabla cabecera, string celda)
        {
            if (cabecera.Ancho.HasValue)
                hojaExcel.Cell($"{celda}").WorksheetColumn().Width = cabecera.Ancho.Value;
            if (cabecera.Altura.HasValue)
                hojaExcel.Cell($"{celda}").WorksheetRow().Height = cabecera.Altura.Value;
        }

        private static void AplicaEstilo(this IXLRange celdas, Estilo estilo)
        {
            if (Enum.TryParse<AppSettings.AlineamientoHorizontal>(estilo.AlineadoHorizontal, out AppSettings.AlineamientoHorizontal alineadoHorizontalSeleccionado))
            {
                if (Enum.TryParse<XLAlignmentHorizontalValues>(alineadoHorizontalSeleccionado.GetStringValue(), out XLAlignmentHorizontalValues alineadoHorizontal))
                    celdas.Style.Alignment.Horizontal = alineadoHorizontal;
            }
            if (Enum.TryParse<AppSettings.AlineamientoVertical>(estilo.AlineadoVertical, out AppSettings.AlineamientoVertical alineadoVerticalSeleccionado))
            {
                if (Enum.TryParse<XLAlignmentVerticalValues>(alineadoVerticalSeleccionado.GetStringValue(), out XLAlignmentVerticalValues alineadoVertical))
                    celdas.Style.Alignment.Vertical = alineadoVertical;
            }
            celdas.Style.Alignment.WrapText = estilo.AjustarTexto;
            celdas.Style.Font.Bold = estilo.Fuente.Negrita;
            celdas.Style.Font.Italic = estilo.Fuente.Cursiva;
            celdas.Style.Font.Underline = estilo.Fuente.Subrayado ? XLFontUnderlineValues.Single : XLFontUnderlineValues.None;
            if (Enum.TryParse<AppSettings.Color>(estilo.Fuente.Color, out AppSettings.Color colorSeleccionado))
                celdas.Style.Font.FontColor = XLColor.FromName(colorSeleccionado.GetStringValue());
            celdas.Style.Font.FontName = estilo.Fuente.Nombre;
            celdas.Style.Font.FontSize = estilo.Fuente.Tamano;
            if (celdas.FirstCellUsed() != null)
            {
                if (estilo.Mayusculas && !celdas.FirstCellUsed().Value.IsBlank && celdas.FirstCellUsed().Value.IsText)
                    celdas.Value = celdas.FirstCellUsed().Value.ToString().ToUpper();
                if (estilo.Minusculas && !celdas.FirstCellUsed().Value.IsBlank && celdas.FirstCellUsed().Value.IsText)
                    celdas.Value = celdas.FirstCellUsed().Value.ToString().ToLower();
            }
            if (Enum.TryParse<AppSettings.TipoBorde>(estilo.Borde, out AppSettings.TipoBorde bordeSeleccionado))
            {
                if (Enum.TryParse<XLBorderStyleValues>(bordeSeleccionado.GetStringValue(), out XLBorderStyleValues estiloBorde))
                {
                    celdas.Style.Border.TopBorder = estiloBorde;
                    celdas.Style.Border.BottomBorder = estiloBorde;
                    celdas.Style.Border.LeftBorder = estiloBorde;
                    celdas.Style.Border.RightBorder = estiloBorde;
                }
            }
            if (Enum.TryParse<AppSettings.Color>(estilo.ColorFondo, out colorSeleccionado))
                celdas.Style.Fill.BackgroundColor = XLColor.FromName(colorSeleccionado.GetStringValue());
        }
    }
}
