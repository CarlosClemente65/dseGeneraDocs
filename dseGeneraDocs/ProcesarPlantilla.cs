using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;


namespace dseGeneraDocs
{
    public class ProcesarPlantilla
    {
        private WordprocessingDocument documento;
        private string ficheroPlantilla;
        private string ficheroSalida;

        private Parametros parametros;
        private Opciones datosGuion;

        public ProcesarPlantilla(Opciones opciones)
        {
            datosGuion = opciones;
            parametros = opciones.Parametros;
        }

        public void AbrirPlantilla()
        {
            ficheroPlantilla = parametros.Plantilla;
            ficheroSalida = parametros.Salida;

            File.Copy(ficheroPlantilla, ficheroSalida, overwrite: true); // Se copia la plantilla con el nombre de salida para hacer la modificacion directamente en la copia.
            documento = WordprocessingDocument.Open(ficheroSalida, true);
        }

        public void ReemplazarMarcadores()
        {
            var marcadores = datosGuion.Marcadores;
            var body = documento.MainDocumentPart.Document.Body; // Contiene todo el contenido visible del documento

            // Recorremos todos los textos del documento
            foreach(var texto in body.Descendants<Text>()) // Recorremos todos los elementos de texto (textos, tablas, etc)
            {
                foreach(var kvp in marcadores) // Recorremos todos los marcadores pasados para hacer la sustitucion
                {
                    string marcador = kvp.Key; // Ejemplo: "#empresa#"
                    string valor = kvp.Value; // Ejemplo: DIAGRAM SOFTWARE EUROPA S.L.

                    if(texto.Text.Contains(marcador)) // Si lo encuenta, lo reemplaza
                    {
                        texto.Text = texto.Text.Replace(marcador, valor);
                    }
                }
            }

            // Guarda los cambios realizados en el documento de la memoria
            documento.MainDocumentPart.Document.Save();
        }

        public void InsertarDatosTablas()
        {
            List<TablaDatos> tablasDatos = datosGuion.Tablas;
            var body = documento.MainDocumentPart.Document.Body;

            foreach(var tablaDatos in tablasDatos)
            {
                // Buscar tabla que tenga en su fila de encabezado el texto indicado
                var tabla = body.Elements<Table>()
                    .FirstOrDefault(t =>
                        t.Elements<TableRow>().Any(r =>
                            r == t.Elements<TableRow>().First() && // Solo primera fila (encabezado)
                            r.Elements<TableCell>().Any(c =>
                                c.InnerText.Contains(tablaDatos.EncabezadoClave))));

                if(tabla == null) continue;

                var filas = tabla.Elements<TableRow>().ToList();

                //if(filas.Count < 2)
                //    throw new Exception("La tabla debe tener al menos una fila de encabezado y una fila plantilla.");

                // Copiamos la segunda fila para usarla como plantilla (índice 1)
                var filaPlantilla = filas[1];

                // Limpiar todas las filas excepto la primera (encabezado)
                for(int i = filas.Count - 1; i >= 1; i--)
                {
                    filas[i].Remove();
                }

                foreach(var filaDatos in tablaDatos.Filas)
                {
                    // Creamos una nueva fila con el formato de la fila que usamos como plantilla
                    var nuevaFila = (TableRow)filaPlantilla.CloneNode(true);

                    var celdas = nuevaFila.Elements<TableCell>().ToList();

                    // Insertamos en cada celda los datos de la fila que estamos procesando
                    for(int i = 0; i < filaDatos.Count && i < celdas.Count; i++)
                    {
                        var parrafo = celdas[i].GetFirstChild<Paragraph>();
                        var run = parrafo?.GetFirstChild<Run>();
                        var text = run?.GetFirstChild<Text>();

                        if(text != null)
                        {
                            text.Text = filaDatos[i];
                        }

                        
                        //var celda = celdas[i];

                        //// Limpiar contenido previo
                        //celda.RemoveAllChildren<Paragraph>();

                        //// Crear nuevo párrafo con el texto, pero usando el estilo del párrafo de plantilla
                        //var parrafoPlantilla = filaPlantilla.Elements<TableCell>().ElementAt(i).Elements<Paragraph>().FirstOrDefault();

                        //Paragraph nuevoParrafo;

                        //if(parrafoPlantilla != null)
                        //{
                        //    // Clonamos el párrafo de plantilla para mantener estilos
                        //    nuevoParrafo = (Paragraph)parrafoPlantilla.CloneNode(true);

                        //    // Reemplazamos el texto del primer Run/Text en el párrafo
                        //    var run = nuevoParrafo.GetFirstChild<Run>();
                        //    var text = run?.GetFirstChild<Text>();
                        //    if(text != null)
                        //    {
                        //        text.Text = filaDatos[i];
                        //    }
                        //    else
                        //    {
                        //        // Si no hay Text, limpiar párrafo y agregar nuevo Run con texto
                        //        // Eliminar todo contenido del párrafo clonado
                        //        nuevoParrafo.RemoveAllChildren();

                        //        // Crear nuevo Run
                        //        run = new Run();
                        //        text = new Text(filaDatos[i]);

                        //        run.Append(text);
                        //        nuevoParrafo.Append(run);
                        //    }
                        //}
                        //else
                        //{
                        //    // Si no hay párrafo plantilla, creamos uno básico
                        //    nuevoParrafo = new Paragraph(new Run(new Text(filaDatos[i])));
                        //}

                        // Añadimos el nuevo párrafo a la celda
                        //celda.AppendChild(nuevoParrafo);
                    }

                    // Añadimos la nueva fila a la tabla
                    tabla.AppendChild(nuevaFila);
                }
            }

            // Guardamos los cambios realizados en el documento de la memoria
            documento.MainDocumentPart.Document.Save();
        }

        public void GuardarDocumento()
        {
            // Se guarda el documento
            documento.MainDocumentPart.Document.Save();

            // Se liberan recursos
            documento.Dispose();

        }

    }
}
