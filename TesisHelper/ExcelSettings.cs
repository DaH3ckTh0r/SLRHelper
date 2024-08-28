using System.ComponentModel.DataAnnotations;

namespace TesisHelper
{
    internal class ExcelSettings
    {
        [Required]
        public string? NombreDelArchivo { get; set; }
        [Required]
        public string? RutaDelArchivo { get; set; }
        public Dictionary<string, TablaExcel> Tablas { get; set; } = new Dictionary<string, TablaExcel>();
        public Dictionary<string, Estilo> Estilos { get; set; } = new Dictionary<string, Estilo> { { AppSettings.Constants.ESTILO_POR_DEFECTO, new Estilo() } };
    }

    internal class TablaExcel
    {
        [Required]
        public int? Posicion { get; set; }
        public int FilasNoIncluidas { get; set; }
        public int ColumnasNoIncluidas { get; set; }
        public string? Estilo { get; set; }
        public int NumeroFilasOcupadas { get; set; }
        public int NumeroColumnasOcupadas { get; set; }
        public Dictionary<string, CabeceraTabla> Columnas { get; set; } = new Dictionary<string, CabeceraTabla>();
    }

    internal class CabeceraTabla
    {
        [Required]
        public string? Titulo { get; set; }
        [Required]
        public int? Posicion { get; set; }
        public double? Ancho { get; set; }
        public double? Altura { get; set; }
        public int NumeroFila { get; set; }
        public int NumeroColumna { get; set; }
        public int NumeroFilasOcupadas { get; set; }
        public int NumeroColumnasOcupadas { get; set; }
        public Dictionary<string, CabeceraTabla>? Columnas { get; set; }
        public string? Estilo { get; set; }
    }

    internal class Estilo
    {
        public bool Mayusculas { get; set; } = false;
        public bool Minusculas { get; set; } = false;
        public string AlineadoHorizontal { get; set; } = AppSettings.AlineamientoHorizontal.Izquierda.ToString();
        public string AlineadoVertical { get; set; } = AppSettings.AlineamientoVertical.Arriba.ToString();
        public bool AjustarTexto { get; set; } = false;
        public string ColorFondo { get; set; } = AppSettings.Color.Ninguno.ToString();
        public string Borde { get; set; } = AppSettings.TipoBorde.Ninguno.ToString();
        public Fuente Fuente { get; set; } = new Fuente();
    }

    internal class Fuente
    {
        public string Nombre { get; set; } = AppSettings.TipoFuente.TimesNewRoman.ToString();
        public double Tamano { get; set; } = 10;
        public string Color { get; set; } = AppSettings.Color.Negro.ToString();
        public bool Cursiva { get; set; } = false;
        public bool Subrayado { get; set; } = false;
        public bool Negrita { get; set; } = false;
    }
}
