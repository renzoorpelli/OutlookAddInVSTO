using Entidades2.Services;
using Newtonsoft.Json;
using System.Net.Http.Json;
using System.Text.Json.Serialization;

namespace ProbarEndPoint
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            Console.WriteLine("Hello, World!");

            ProyectoService.ObtenerProyectos();
        }
        




    }
}