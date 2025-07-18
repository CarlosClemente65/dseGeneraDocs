using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace dseGeneraDocs
{
    public class Opciones
    {
        public Parametros Parametros { get; set; } = new Parametros();
        public Dictionary<string, string> Marcadores { get; set; } = new Dictionary<string, string>();
        public List<TablaDatos> Tablas { get; set; } = new List<TablaDatos>();

    }

    public class Parametros
    {
        public string Plantilla { get; set; }
        public string Salida { get; set; }
    }

    public class TablaDatos
    {
        public string Nombre { get; set; }                  // Ej: "#tabla1#"
        public string EncabezadoClave { get; set; }         // Ej: "Entidad concedente"
        public List<List<string>> Filas { get; set; } = new List<List<string>>(); // Filas a insertar en la tabla
    }

}
