using System.Collections.Generic;
using System.IO;
using System.Linq;
using UtilidadesDiagram;

namespace dseGeneraDocs
{
    internal class Program
    {
        Dictionary<string, List<List<string>>> DatosTablas;

        static void Main(string[] args)
        {
            string guion = args[0];
            if(File.Exists(guion))
            {
                Opciones opciones = new Opciones();
                LeerGuion(guion, opciones);

                ProcesarPlantilla procesarPlantilla = new ProcesarPlantilla(opciones);
                procesarPlantilla.AbrirPlantilla();
                procesarPlantilla.ReemplazarMarcadores();
                procesarPlantilla.InsertarDatosTablas();
                procesarPlantilla.GuardarDocumento();
            }
            else
            {
                File.WriteAllText("salida.err", "Fichero guion no encontrado");
            }

        }

        public static void LeerGuion(string guion, Opciones opciones)
        {
            //var docEntrada = new Opciones();

            string[] lineas = File.ReadAllLines(guion);

            string seccionActual = "";
            TablaDatos tablaActual = null;

            foreach(var lineaRaw in lineas)
            {
                string linea = lineaRaw.Trim();

                if(string.IsNullOrWhiteSpace(linea))
                {
                    continue; // Evita lineas vacias
                }

                // Cambiar de sección
                if(linea.StartsWith("[") && linea.EndsWith("]"))
                {
                    seccionActual = linea.ToLower();
                    continue;
                }

                switch(seccionActual)
                {
                    // Asignacion de parametros (se pueden añadir mas al switch)
                    case "[parametros]":
                        if(linea.Contains("="))
                        {
                            // Líneas con parametro=valor                           
                            (string clave, string valor) = Utilidades.DivideCadena(linea, '='); // Divide la linea en dos

                            switch(clave)
                            {
                                case "plantilla":
                                    opciones.Parametros.Plantilla = valor;
                                    break;

                                case "salida":
                                    opciones.Parametros.Salida = valor;
                                    break;
                            }

                        }
                        break;

                    case "[marcadores]":
                        if(linea.StartsWith("#") && linea.Contains("="))
                        {
                            (string clave, string valor) = Utilidades.DivideCadena(linea, '='); // Divide la linea en dos
                            opciones.Marcadores[clave] = valor;
                        }
                        break;

                    case "[tablas]":
                        if(linea.StartsWith("#") && linea.Contains("="))
                        {
                            (string nombreTabla, string encabezado) = Utilidades.DivideCadena(linea, '='); // Obtiene el texto a buscar en el encabezado de la tabla. El nombre de la tabla va entre '#'

                            tablaActual = new TablaDatos // Crea una instancia para una nueva table
                            {
                                Nombre = nombreTabla,
                                EncabezadoClave = encabezado
                            };

                            opciones.Tablas.Add(tablaActual);
                        }
                        else if(tablaActual != null)
                        {
                            var fila = linea.Split('|').Select(c => c.Trim()).ToList();
                            tablaActual.Filas.Add(fila);
                        }
                        break;
                }
            }
        }
    }
}
