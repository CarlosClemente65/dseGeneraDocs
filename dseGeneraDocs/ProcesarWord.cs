using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;


namespace dseGeneraDocs
{
    public class ProcesarWord
    {
        private Parametros parametros;
        private DatosGuion datosGuion;

        public ProcesarWord(DatosGuion opciones)
        {
            datosGuion = opciones;
            parametros = opciones.Parametros;
        }

        // Abrir documento y preparar Word. Se reciben los parametros por referencia ya que se han definido fuera de la clase
        public void AbrirDocumento(ref Word.Application instanciaWord, ref Word.Document documento)
        {
            try
            {
                instanciaWord = new Word.Application(); // Se crea la instancia para una nueva aplicacion de Word
                instanciaWord.Visible = false; // Se pone como no visible

                documento = instanciaWord.Documents.Open(parametros.Plantilla); // Se crea el documento cargando la plantilla
            }
            catch(Exception ex)
            {
                // Control de la excepcion cuando se genera desde la aplicacion COM de Word
                if(ex is COMException && ex.Message.Contains("encontrar el archivo"))
                {
                    throw new FileNotFoundException($"No existe el fichero con la plantilla ({parametros.Plantilla}).");
                }

                // En otro caso, se devuelve el mensaje generico que ha provocado la excepcion
                throw new Exception($"Error al abrir el documento word: {ex.Message}");
            }
        }

        public void ProcesarMarcadoresTexto(Word.Application instanciaWord, Word.Document documento)
        {
            // Metodo para procesar los marcadores que se han incluido en el guion
            try
            {
                if(documento == null) // Evita un error si no esta abierto el documento
                {
                    throw new InvalidOperationException("Documento no abierto. Llama primero a AbrirDocumento().");
                }


                // Recorre cada seccion del documento para procesar los marcadores de encabezados, pies de pagina y cuerpo del documento
                foreach(Word.Section seccion in documento.Sections)
                {
                    // Procesa los encabezados
                    foreach(Word.HeaderFooter encabezado in seccion.Headers)
                    {
                        if(encabezado.Exists)
                        {
                            // Procesa los marcadores de cada encabezado
                            ReemplazarMarcadores(encabezado.Range, datosGuion.Marcadores);
                        }
                    }

                    // Procesa los pie de pagina
                    foreach(Word.HeaderFooter pie in seccion.Footers)
                    {
                        if(pie.Exists)
                        {
                            // Procesa los marcadores de cada pie de pagina
                            ReemplazarMarcadores(pie.Range, datosGuion.Marcadores);
                        }
                    }

                    // Procesa el cuerpo del documento
                    ReemplazarMarcadores(documento.Content, datosGuion.Marcadores);
                }
            }
            catch(Exception ex)
            {
                // Si hay un error se lanza una excepcion para generar un fichero de salida
                throw new Exception("Error al procesar marcadores: " + ex.Message);
            }
        }

        private void ReemplazarMarcadores(Word.Range rango, Dictionary<string, string> marcadores)
        {
            // Procesado de todos los marcadores del guion
            foreach(var marcador in datosGuion.Marcadores)
            {
                Word.Find findObject = rango.Find; // Crea el proceso para hacer la busqueda
                findObject.ClearFormatting(); // Limpia el formato de busqueda para evitar que pueda haber formatos de negrita que impidan encontrar los textos.
                findObject.Text = marcador.Key; // Texto a buscar en el documento
                findObject.Replacement.ClearFormatting(); // Limpia los valores de formato en el proceso de reemplazo para que se sustituya con el formato que tenga en el documento
                findObject.Replacement.Text = marcador.Value; // Texto por el que sera reemplazado el marcador

                findObject.Execute(Replace: Word.WdReplace.wdReplaceAll); // Ejecuta el proceso de reemplado en todo el documento.
            }
        }

        public void InsertarFilasTablas(Word.Application instanciaWord, Word.Document documento)
        {
            // Metodo para hacer la insercion de filas en las tablas que vengan indicadas en el guion

            // Chequea que haya datos en el guion antes de continuar
            if(datosGuion.Tablas == null || datosGuion.Tablas.Count == 0)
            {
                return;
            }

            // Chequea que se haya creado la instancia del documento Word con la plantilla
            if(documento == null)
            {
                throw new InvalidOperationException("Documento no abierto. Llama primero a AbrirDocumento().");
            }

            // Procesado de los datos de las tablas que se han pasado en el guion
            foreach(var tabla in datosGuion.Tablas)
            {
                // Texto que debe aparecer en la celda A1 de la tabla para localizarla
                string encabezadoClave = tabla.EncabezadoClave;
                var datos = tabla.Filas; // Carga todas las filas de la tabla del guion que se este procesando

                // Procesa todas las tablas del documento para localizar el texto del encabezado
                foreach(Word.Table tablaWord in documento.Tables)
                {
                    // Almacenamos el valor de la celda A1 (fila 1, columna 1)
                    string textoCelda = tablaWord.Cell(1, 1).Range.Text.TrimEnd('\r', '\a');

                    // Si se localiza el texto que estamos buscando en la tabla que se esta procesando, se insertan los datos
                    if(textoCelda.Equals(encabezadoClave, StringComparison.OrdinalIgnoreCase))
                    {
                        int numColumnas = tablaWord.Columns.Count;

                        // Evitar errores si la tabla no tiene una segunda fila para insertar datos.
                        if(tablaWord.Rows.Count < 2)
                        {
                            break;
                        }

                        // Copiamos el formato de la segunda fila
                        Word.Row filaPlantilla = tablaWord.Rows.Count >= 2 ? tablaWord.Rows[2] : null;

                        // Insertar filas a partir de los datos
                        int ultimaFila = 1; // Contador de filas que tendra la tabla sin contar la cabecera
                        foreach(var filaDatos in datos)
                        {
                            // Insertamos una nueva fila a continuacion del encabezado
                            Word.Row nuevaFila = tablaWord.Rows.Add(filaPlantilla);
                            ultimaFila++; // Se incrementa el numero de filas

                            // Se insertan todos los datos en las celdas de la fila actual
                            for(int i = 0; i < filaDatos.Count && i < numColumnas; i++)
                            {
                                nuevaFila.Cells[i + 1].Range.Text = filaDatos[i];
                            }
                        }

                        // Proceso para eliminar las filas que tuviera la tabla antes de insertar los nuevos datos (se insertan siempre entre la cabecera y la primera fila).
                        while(tablaWord.Rows.Count > ultimaFila) // Se procesa mientras haya mas filas que la ultima que se inserto
                        {
                            tablaWord.Rows[ultimaFila + 1].Delete(); // Se borra la fila siguiente que no es valida
                        }

                        break; // Ya insertamos en la tabla correspondiente                    
                    }
                }
            }
        }

        public void GuardarDocumento(ref Word.Application instanciaWord, ref Word.Document documento)
        {
            // Metodo para guardar el documento con los datos ya reemplazados
            try
            {
                if(documento != null)
                {
                    // Guardar documento en la ruta indicada
                    documento.SaveAs2(datosGuion.Parametros.Salida);

                    if(parametros.PDF)
                    {
                        // Guarda una copia del documento en formato PDF (por si fuera necesario)
                        string ficheroPdf = Path.Combine(Path.GetDirectoryName(datosGuion.Parametros.Salida), Path.GetFileNameWithoutExtension(datosGuion.Parametros.Salida) + ".pdf");

                        documento.SaveAs2(ficheroPdf, Word.WdSaveFormat.wdFormatPDF);
                    }
                }
            }
            catch(Exception ex)
            {
                // Si se generar un error se lanza una excepcion para grabarla en el fichero de resultado
                throw new Exception("Error al guardar el documento: " + ex.Message);
            }
            finally
            {
                // Siempre se liberan recursos de la instancia de Word y del documento
                Program.LiberarWord(ref instanciaWord, ref documento);
            }
        }
    }
}
