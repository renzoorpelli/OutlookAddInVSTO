using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace Entidades2.Services
{
    public class ProyectoService
    {
        public static List<Proyectos> ObtenerProyectos()
        {
            try
            {
                using (HttpClient client = new HttpClient())
                {
                    string URL = "https://localhost:7274/proyectos";

                    var response = client.GetStringAsync(URL).Result; //nos devuelve el json en un string


                    var listaProyectos = JsonSerializer.Deserialize<List<Proyectos>>(response); //deserializamos en una lista de objetos

                    if (!(response is null))
                    {
                        //retorno la lista
                        return listaProyectos;
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            
            return null;
        }
    }
}
