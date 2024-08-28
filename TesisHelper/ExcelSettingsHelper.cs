using Microsoft.Extensions.Configuration;

namespace TesisHelper
{
    internal static class ExcelSettingsHelper
    {
        public static ExcelSettings CargarConfiguraciones(IConfigurationRoot? configuration = null)
        {
            configuration = configuration ?? LeerConfiguraciones();
            const string SectionName = "Excel";
            IConfigurationSection section = configuration.GetSection(SectionName);
            return section == null ? new ExcelSettings() : CargarConfiguracionesExcel(section);
        }

        private static ExcelSettings CargarConfiguracionesExcel(IConfigurationSection section)
        {
            ExcelSettings settings = new ExcelSettings();
            IEnumerable<IConfigurationSection> children = section.GetChildren();
            settings.NombreDelArchivo = children?.FirstOrDefault(x => x.Key.Equals(nameof(settings.NombreDelArchivo)))?.Value;
            settings.RutaDelArchivo = children?.FirstOrDefault(x => x.Key.Equals(nameof(settings.RutaDelArchivo)))?.Value;
            CargarEstilos(settings, children);
            CargarTablas(settings, children);
            CalcularNivelesDeProfundidadYAmplitud(settings.Tablas);
            return settings;
        }

        private static void CalcularNivelesDeProfundidadYAmplitud(Dictionary<string, TablaExcel> tablas)
        {
            foreach (var tabla in tablas)
            {
                tabla.Value.NumeroFilasOcupadas = CalcularNivelesDeProfundidad(tabla.Value.Columnas);
                tabla.Value.NumeroColumnasOcupadas = CalcularNivelesDeAmplitud(tabla.Value.Columnas);
            }
        }

        private static int CalcularNivelesDeProfundidad(Dictionary<string, CabeceraTabla> columnas, int nivelInicial = 0)
        {
            int nivelActual = nivelInicial + 1;
            int nivelProfundidad = 0;

            foreach (var columna in columnas.OrderBy(p => p.Value.Posicion))
            {
                columna.Value.NumeroFilasOcupadas = 1;
                if (columna.Value.Columnas?.Any() ?? false)
                {
                    columna.Value.NumeroFilasOcupadas = CalcularNivelesDeProfundidad(columna.Value.Columnas, nivelActual);
                    nivelProfundidad = Math.Max(columna.Value.NumeroFilasOcupadas, nivelProfundidad);
                }
                columna.Value.NumeroFila = nivelActual;
            }
            return Math.Max(nivelProfundidad, nivelActual);
        }

        private static int CalcularNivelesDeAmplitud(Dictionary<string, CabeceraTabla> columnas, int nivelAntecesor = 0)
        {
            int nivelActual = 0;

            foreach (var columna in columnas.OrderBy(p => p.Value.Posicion))
            {
                if (columna.Value.Columnas?.Any() ?? false)
                {
                    columna.Value.NumeroColumna = nivelAntecesor + nivelActual + 1;
                    columna.Value.NumeroColumnasOcupadas = CalcularNivelesDeAmplitud(columna.Value.Columnas, columna.Value.NumeroColumna - 1);
                    nivelActual += columna.Value.NumeroColumnasOcupadas;
                    continue;
                }
                columna.Value.NumeroColumnasOcupadas = 1;
                columna.Value.NumeroColumna = nivelAntecesor + ++nivelActual;
            }
            return nivelActual;
        }

        private static void CargarTablas(ExcelSettings settings, IEnumerable<IConfigurationSection>? children)
        {
            if (!children?.Any(x => x.Key.Equals(nameof(settings.Tablas))) ?? false) return;

            foreach (IConfigurationSection child in children?.FirstOrDefault(x => x.Key.Equals(nameof(settings.Tablas)))?.GetChildren())
            {
                TablaExcel tablaExcel = new TablaExcel();
                child.Bind(tablaExcel);
                settings.Tablas.Add(child.Key, tablaExcel);
            }
        }

        private static void CargarEstilos(ExcelSettings settings, IEnumerable<IConfigurationSection>? children)
        {
            if (!children?.Any(x => x.Key.Equals(nameof(settings.Estilos))) ?? false) return;

            foreach (IConfigurationSection child in (children?.FirstOrDefault(x => x.Key.Equals(nameof(settings.Estilos)))?.GetChildren()))
            {
                Estilo estilo = new Estilo();
                child.Bind(estilo);
                settings.Estilos.Add(child.Key, estilo);
            }
        }

        public static IConfigurationRoot LeerConfiguraciones(string nombreArchivo = "settings.json")
        {
            var builder = new ConfigurationBuilder();
            builder.SetBasePath(Directory.GetCurrentDirectory())
                   .AddJsonFile(nombreArchivo, optional: false, reloadOnChange: true);
            return builder.Build();
        }
    }
}
