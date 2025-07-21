using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace dseGeneraDocs
{
    public class DatosGuion
    {
        public Parametros Parametros { get; set; } = new Parametros();
        public Dictionary<string, string> Marcadores { get; set; } = new Dictionary<string, string>(); // Marcadores que deben localizarse en la plantilla para sustituir por su valor. Formato #marcador#=valor a reemplazar
        public List<TablaDatos> Tablas { get; set; } = new List<TablaDatos>();

    }

    public class Parametros
    {
        // Clase que recoje los parametros que se puede pasar desde el guion
        public string Plantilla { get; set; }
        public string Salida { get; set; }

        public bool PDF { get; set; } = false;
    }

    public class TablaDatos
    {
        // Clase para recoger los datos que luego deben incluirse en las tablas del documento
        public string Nombre { get; set; } // Nombre que se asigna a la tabla de forma temporal
        public string EncabezadoClave { get; set; } // Texto que debe aparecer en la celda A1 de la tabla para localizarla y poder insertar las filas
        public List<List<string>> Filas { get; set; } = new List<List<string>>(); // Filas a insertar en la tabla
    }

}
