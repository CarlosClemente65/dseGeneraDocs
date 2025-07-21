using System;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using UtilidadesDiagram;
using Word = Microsoft.Office.Interop.Word;

namespace dseGeneraDocs
{
    internal class Program
    {
        static Word.Application instanciaWord = null; // Definicion de la instancia de Word para procesar la plantilla 
        static Word.Document documento = null; // Definicion del documento que contendra la plantilla a procesar

        static void Main(string[] args)
        {
            string pruebas = "Word"; // Sirve para ejecutar con las librerias propias de Word o con las de Open XML (lo dejo por si fuera necesario en un futuro modificar algo).

            string guion = args[0];

            if(File.Exists(guion)) // Si el fichero del guion se encuentra se inicia el proceso.
            {
                DatosGuion DatosGuion = new DatosGuion(); // Instancia de la clase DatosGuion para poder usar sus propiedades.

                LeerGuion(guion, DatosGuion); // Carga los datos del guion en la clase Opciones.

                // Se genera un fichero de resultado por defecto poniendolo en la misma ruta que el fichero de salida
                string ficheroResultado = Path.Combine(Path.GetDirectoryName(DatosGuion.Parametros.Salida), "salida.txt");

                try
                {
                    // Antes de continuar se chequea si la plantilla ya esta abierta y no seguir con el resto
                    if(ArchivoBloqueado(DatosGuion.Parametros.Plantilla))
                    {
                        // Se muestra un mensaje por pantalla
                        MessageBox.Show("El fichero con la plantilla esta en uso. Cierrela para continuar.", "Error al abrir la plantilla", MessageBoxButton.OK, MessageBoxImage.Exclamation);

                        // En todo caso se genera una excepcion para que se grabe en el ficheroResultado el mensaje.
                        throw new IOException($"El fichero con la plantilla {DatosGuion.Parametros.Plantilla} esta en uso.\nCierrela para continuar.");
                    }

                    switch(pruebas) // Dos tipos de ejecucion (con Interop.Word y con OpenXML)
                    {
                        case "Word":
                            // Crea una nueva instancia para procesar los datos.
                            ProcesarWord procesarWord = new ProcesarWord(DatosGuion);

                            // Se abre el documento por referencia, pasando la instancia del Word creada mas arriba y el documento.
                            procesarWord.AbrirDocumento(ref instanciaWord, ref documento);

                            // Se hace el procesado de los marcadores, pasando la instancia del Word creada mas arriba y el documento.
                            procesarWord.ProcesarMarcadoresTexto(instanciaWord, documento);

                            // Se hace el procesado de las tablas que se hayan incluido en el guion, pasando la instancia del Word creada mas arriba y el documento.
                            procesarWord.InsertarFilasTablas(instanciaWord, documento);

                            // Se guarda el documento, pasando la instancia del Word creada mas arriba y el documento.
                            procesarWord.GuardarDocumento(ref instanciaWord, ref documento);
                            break;

                        case "OpenXml":
                            // Sin uso actualmente, ya que da problemas con las tablas y algunos marcadores
                            ProcesarPlantilla procesarPlantilla = new ProcesarPlantilla(DatosGuion);
                            procesarPlantilla.AbrirPlantilla();
                            procesarPlantilla.ReemplazarMarcadores();
                            procesarPlantilla.InsertarDatosTablas();
                            procesarPlantilla.GuardarDocumento();
                            break;
                    }

                    // Cuando acaba sin errores se graba el ficheroResultado con un OK
                    GrabarSalida(ficheroResultado, "OK");
                }

                // Captura las posibles excepciones y graba un fichero con el mensaje que corresponda.
                catch(Exception ex)
                {
                    GrabarSalida(ficheroResultado, $"Errores en el proceso:\n{ex.Message}");
                }

                // En cualquier caso, siempre librea recursos del Word
                finally
                {
                    LiberarWord(ref instanciaWord, ref documento);
                }
            }
            else
            {
                // Solo en el caso de que no se haya pasado el fichero del guion o no se localice.
                GrabarSalida("salida.txt", "Fichero guion no encontrado");
            }
        }

        public static void LeerGuion(string guion, DatosGuion opciones)
        {
            // Metodo para cargar en la clase DatosGuion todos los parametros del guion

            // Carga todas las lineas
            string[] lineas = File.ReadAllLines(guion);

            // Permite controlar que seccion del guion se esta procesando
            string seccionActual = "";

            TablaDatos tablaActual = null; // Define una variable para localizar la tabla en la que se esta buscando el texto para posteriormente añadir las filas.

            // Procesado de todas las lineas del fichero
            foreach(var lineaRaw in lineas)
            {
                string linea = lineaRaw.Trim();

                if(string.IsNullOrWhiteSpace(linea))
                {
                    continue; // Evita lineas vacias
                }

                // Las secciones siempre se identifican con el formato [seccion]
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

                            switch(clave.ToUpper())
                            {
                                case "PLANTILLA":
                                    opciones.Parametros.Plantilla = valor; // Asigna el nombre de la plantilla

                                    string ruta = (Path.GetDirectoryName(opciones.Parametros.Plantilla));
                                    string fichero = Path.GetFileNameWithoutExtension(opciones.Parametros.Plantilla)+"_salida.docx";
                                    opciones.Parametros.Salida = Path.Combine(ruta, fichero); // Se asigna un nombre de salida por defecto en caso de que no se pase en el guion.
                                    break;

                                case "SALIDA":
                                    opciones.Parametros.Salida = valor; // Asigna el nombre del fichero de salida
                                    break;

                                case "PDF":
                                    if (valor.ToUpper() == "SI")
                                    {
                                        opciones.Parametros.PDF = true;
                                    }
                                    break;
                            }
                        }
                        break;

                    case "[marcadores]":
                        // Procesado de la seccion de marcadores. Todos los marcadores tienen el formato #marcador#=valor a reemplazar
                        if(linea.StartsWith("#") && linea.Contains("="))
                        {
                            (string clave, string valor) = Utilidades.DivideCadena(linea, '='); // Divide la linea en dos
                            opciones.Marcadores[clave] = valor; // Graba en el diccionario de marcadores, la clave y valor que luego se utilizaran para hacer los reemplazos
                        }
                        break;

                    case "[tablas]":
                        // Procesado de la seccion de tablas.
                        if(linea.StartsWith("#") && linea.Contains("=")) // La primera linea de la tabla identifica con un nombre y el texto a localizar en la celda A1 con el formato #nombreTabla#=Texto a localizar
                        {
                            (string nombreTabla, string encabezado) = Utilidades.DivideCadena(linea, '='); 

                            tablaActual = new TablaDatos // Crea una instancia para una nueva tabla
                            {
                                Nombre = nombreTabla,
                                EncabezadoClave = encabezado
                            };

                            // Añade una nueva tabla
                            opciones.Tablas.Add(tablaActual);
                        }

                        // Si se ha asignado una tabla, procesa las siguientes lineas que son las que finalmente se insertaran en la tabla
                        else if(tablaActual != null)
                        {
                            // Crea una lista con los valores de la linea que estan separados por el caracter '|'
                            var fila = linea.Split('|').Select(c => c.Trim()).ToList();
                            tablaActual.Filas.Add(fila); // Se añade la fila con los valores a la tabla.
                        }
                        break;
                }
            }
        }

        public static void GrabarSalida(string rutaFichero, string texto)
        {
            // Metodo para grabar el fichero con el resultado
            File.WriteAllText(rutaFichero, texto);
        }

        public static void LiberarWord(ref Word.Application instanciaWord, ref Word.Document documento)
        {
            // Liberar objetos COM para evitar procesos WINWORD colgados
            if(documento != null) // Solo cuando haya algun documento cargado lo libera
            {
                documento.Close(); // Cierra el documento
                Marshal.ReleaseComObject(documento);
                documento = null;
            }

            if(instanciaWord != null) // Solo cuando haya una instancia de Word cargada la libera
            {
                instanciaWord.Quit(); // Cierra Word
                Marshal.ReleaseComObject(instanciaWord);
                instanciaWord = null;
            }

            // Libera el resto de posibles instancias que puedan quedar sin liberar
            GC.Collect(); 
            GC.WaitForPendingFinalizers();
        }

        public static bool ArchivoBloqueado(string rutaArchivo)
        {
            // Control para si el fichero con la plantilla esta en uso y no continuar con el proceso.
            try
            {
                using(FileStream stream = File.Open(rutaArchivo, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    // Si puede abrirse con FileShare.None, no está en uso
                    return false;
                }
            }
            catch(IOException)
            {
                // El archivo está bloqueado por otro proceso
                return true;
            }
        }
    }
}
