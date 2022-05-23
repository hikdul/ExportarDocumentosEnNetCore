using Newtonsoft.Json;
using OfficeOpenXml;
using SpreadsheetLight;

namespace ExportarDocumentos.Models
{
    /// <summary>
    /// data class
    /// </summary>
    public class Datos
    {

        #region props
        /// <summary>
        /// user id
        /// </summary>
        public int userId { get; set; }
        /// <summary>
        /// id
        /// </summary>
        public int id { get; set; }
        /// <summary>
        ///  title
        /// </summary>
        public string title { get; set; }
        /// <summary>
        /// is completed?
        /// </summary>
        public bool complete { get; set; }

        #endregion

        public string getComplete()
        {
            return this.complete ? "SI" : "NO";
        }

        public void setComplete(string complete)
        {
            this.complete = complete.ToUpper().Contains("SI");
        }
        #region get list
        /// <summary>
        /// Para obtener una lista de datos falsos
        /// </summary>
        /// <returns></returns>
        public static async Task<List<Datos>> GetList()
        {
            string url = "https://jsonplaceholder.typicode.com/todos/";

            HttpClient client = new();
            try
            {
                HttpResponseMessage response = await client.GetAsync(url);
                response.EnsureSuccessStatusCode();
                string responseBody = await response.Content.ReadAsStringAsync();
                Console.WriteLine(responseBody);
                var resp = JsonConvert.DeserializeObject<List<Datos>>(responseBody);
                return resp;
            }
            catch (Exception ee)
            {
                Console.WriteLine(ee.Message);
            }
            return new();
        }




        #endregion
    }

}




