using SpreadsheetLight;

namespace ExportarDocumentos.Models
{
    /// <summary>
    /// Clase para importar datos.
    /// </summary>
    public class ImportarDatos
    {
        /// <summary>
        /// documenton 
        /// </summary>
        public IFormFile doc { get; set; }

        /// <summary>
        /// para retornar el bytes[]
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public async Task<byte[]> ConvertToArrayBytes()
        {
            byte[] array = new byte[0];

            try
            {
                using (var ms = new MemoryStream())
                {
                    await this.doc.CopyToAsync(ms);

                    array = ms.ToArray();
                }
            }
            catch (Exception ee)
            {
                Console.WriteLine(ee.Message);

            }
            return array;
        }


        #region get list from excel

        /// <summary>
        /// Para obtener la lista de elementos en base a un excel
        /// </summary>
        /// <returns></returns>
        public  List<Datos> GetListFromExcel()
        {
            List<Datos> list = new();
            try
            {
                // ==> Primero creamos un lector de documentos.. a este le pasamos nuestro IFormFile. sin embargo sirve con indicandole una ruta.
                SLDocument doc = new(this.doc.OpenReadStream());
                /// es es desde que columna vamos a empezar a tomar los datos.
                /// utiliza el 4 pues segun mi formato desde alli vienen los datos
                int i = 4;
                // le digo que recorra el excel mientras no aya nulos o vacios
                while (!string.IsNullOrEmpty(doc.GetCellValueAsString(i, 1)))
                {
                    try
                    {
                        // creo un objeto y obtengo los elementos segun su celda.
                        // en este ejemple leo vario tipos de datos
                        Datos flag = new()
                        {
                            id = doc.GetCellValueAsInt32(i, 1),
                            userId = doc.GetCellValueAsInt32(i, 2),
                            title = doc.GetCellValueAsString(i, 3),
                            complete = false,
                        };
                        flag.setComplete(doc.GetCellValueAsString(i, 4));
                        // genero mi lista de lectura
                        list.Add(flag);
                    }
                    catch (Exception ee)
                    {
                        Console.WriteLine(ee.Message);
                    }

                    i++;
                }
            }
            catch (Exception ee)
            {
                Console.WriteLine(ee.Message);
            }
            // al final retorno la lista de objetos con los datos obtenidos desde el documento
            return list;
        }
    }
    #endregion
} 
}
