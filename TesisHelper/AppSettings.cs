namespace TesisHelper
{
    internal static class AppSettings
    {
        public static class Constants
        {
            public const string EDGE_DIRECTORY = @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe";
            public const string MAIN_PATH = @"C:\Users\hecto\OneDrive\Documents\Maestria\Direccion de Tesis I";
            public const string PAPERS_PATH = @$"{MAIN_PATH}\Papers";
            public const string REVIEWS_PATH = @$"{MAIN_PATH}\Reviews";
            public const string EXCEL_FILE_NAME = "Evaluacion de Papers.xlsx";
            public const int NUMERO_FILA_PREGUNTAS = 4;
            public const int FILAS_SIN_USAR = 3;

            public const string COLUMNA_ABSTRACT = "ABSTRACT";
            public const string COLUMNA_ARCHIVO = "ARCHIVO";
            public const string COLUMNA_RUTA = "RUTA";
            public const string COLUMNA_ID = "#";

            public const string SI = "SI";
            public const string NO = "NO";
            public const string NO_DEFINIDO = "¿?";
            public const string ARCHIVO_NO_ENCONTRADO = "Not Found";
            public const string DOCUMENTO_NO_ENCONTRADO = "Document not available";

            public const string ESTILO_POR_DEFECTO = "Normal";
        }

        public enum TipoFuente
        {
            [StringValue("Times New Roman")]
            TimesNewRoman,
            [StringValue("Arial")]
            Arial,
            [StringValue("Tahoma")]
            Tahoma
        }

        public enum Color
        {
            [StringValue("NoColor")]
            Ninguno,
            [StringValue("Black")]
            Negro,
            [StringValue("White")]
            Blanco,
            [StringValue("Blue")]
            Azul,
            [StringValue("Red")]
            Rojo,
            [StringValue("Green")]
            Verde,
            [StringValue("DarkGreen")]
            VerdeOscuro,
            [StringValue("DarkGreen1")]
            VerdeOscuro1,
            [StringValue("TealGreen")]
            VerdeTurquesa,
            [StringValue("TealBlue")]
            AzulTurquesa,
            [StringValue("Teal")]
            Turquesa
        }

        public enum TipoBorde
        {
            [StringValue("None")]
            Ninguno,
            [StringValue("Thin")]
            Fino,
            [StringValue("Double")]
            Doble,
            [StringValue("Medium")]
            Medio,
            [StringValue("Thick")]
            Grueso
        }

        public enum AlineamientoHorizontal
        {
            [StringValue("Left")]
            Izquierda,
            [StringValue("Center")]
            Centro,
            [StringValue("Right")]
            Derecha,
            [StringValue("Justified")]
            Justificado
        }

        public enum AlineamientoVertical
        {
            [StringValue("Top")]
            Arriba,
            [StringValue("Bottom")]
            Abajo,
            [StringValue("Center")]
            Centro,
            [StringValue("Distributed")]
            Distribuido
        }

        public static class Keys
        {
            public const string SHIFT_TAB = "+{TAB}";
            public const string TAB = "{TAB}";
            public const string ESC = "{ESC}";
            public const string ENTER = "{ENTER}";
        }

        public static class Columnas
        {
            public static string[] Informacion = ["#", "BASE DE DATOS", "TIPO", "EntryKey", "TÍTULO", "ABSTRACT", "AUTORES", "PUBLICACIÓN", "VOLUMEN", "NÚMERO", "PÁGINAS", "AÑO", "DOI", "URL", "ARCHIVO", "RUTA"];
            public static string[] ResultadosCriterioExclusion = ["CE5", "CE4", "CE3", "CE2", "CE1"];
            public static string[] ResultadosCriterioInclusion = ["CI5", "CI4", "CI3", "CI2", "CI1"];

            public static int NumeroUltimaColumnaAntesDeLasPreguntas()
            {
                return Informacion.Length + 1;
            }
        }

        public static PreguntaInvestigacion[]? PreguntasDeInvestigacion;
        public static PreguntaInclusion[]? PreguntasDeInclusion;
        public static PreguntaExclusion[]? PreguntasDeExclusion;
        public static int[]? IdsAProcesar;

        public static int NumeroColumna(this string[] array, string titulo)
        {
            for (var i = 0; i < array.Length; i++)
            {
                if (array[i] == titulo) return i + 1;
            }
            return -1;
        }
    }
}
